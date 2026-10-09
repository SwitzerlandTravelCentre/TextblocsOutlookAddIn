import {
  AlertCircle,
  Check,
  Clipboard,
  CornerDownLeft,
  Download,
  Loader2,
  Search,
  Star,
  Upload
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import { TextBlock } from "../models/textBlock";
import { copyTextToClipboard, insertTextIntoEmail } from "../services/officeService";
import { TextBlockDataService } from "../services/textBlockDataService";
import {
  allFilter,
  departmentLabel,
  fallbackDepartment,
  LanguageFilter,
  normalizeCategoryKey,
  normalizeDepartmentKey,
  searchTextBlocks
} from "../utils/searchTextBlocks";

const appVersion = "1.1.0";
const favoriteStorageKey = "textblocsfit:favorites:v1";
const departmentStorageKey = "textblocsfit:lastDepartment:v1";

const languages: Array<{ code: LanguageFilter; label: string }> = [
  { code: allFilter, label: "All" },
  { code: "DE", label: "DE" },
  { code: "FR", label: "FR" },
  { code: "EN", label: "EN" },
  { code: "IT", label: "IT" }
];

type NoticeKind = "success" | "error";

interface Notice {
  kind: NoticeKind;
  message: string;
}

type FavoriteMode = "ALL" | "FAVORITES";

interface FilterOption {
  key: string;
  label: string;
}

interface FavoritesBackup {
  schemaVersion: number;
  favorites?: string[];
  defaultDepartment?: string;
}

const getFriendlyInsertError = (error: unknown): string => {
  console.error("Could not insert text into Outlook.", error);
  return "Could not insert text into Outlook. Make sure a message compose window is active.";
};

const textBlockService = new TextBlockDataService();

const readStorage = (key: string): string | null =>
  typeof window === "undefined" ? null : window.localStorage.getItem(key);

const writeStorage = (key: string, value: string): void => {
  if (typeof window !== "undefined") {
    window.localStorage.setItem(key, value);
  }
};

const readFavorites = (): Set<string> => {
  const stored = readStorage(favoriteStorageKey);

  if (!stored) {
    return new Set();
  }

  try {
    const parsed = JSON.parse(stored) as unknown;
    return new Set(Array.isArray(parsed) ? parsed.filter((value) => typeof value === "string") : []);
  } catch {
    return new Set();
  }
};

const buildDepartmentOptions = (blocks: TextBlock[]): FilterOption[] => {
  const options = new Map<string, string>();

  blocks.forEach((block) => {
    const label = departmentLabel(block);
    const key = normalizeDepartmentKey(label);

    if (key && !options.has(key)) {
      options.set(key, label);
    }
  });

  return [...options.entries()]
    .map(([key, label]) => ({ key, label }))
    .sort((a, b) => a.label.localeCompare(b.label));
};

const buildCategoryOptions = (blocks: TextBlock[]): FilterOption[] => {
  const options = new Map<string, string>();

  blocks.forEach((block) => {
    const label = block.category.trim();
    const key = normalizeCategoryKey(label);

    if (key && !options.has(key)) {
      options.set(key, label);
    }
  });

  return [...options.entries()]
    .map(([key, label]) => ({ key, label }))
    .sort((a, b) => a.label.localeCompare(b.label));
};

const blockMatchesDepartment = (block: TextBlock, department: string): boolean =>
  department === allFilter || normalizeDepartmentKey(block.department) === department;

const blockMatchesLanguage = (block: TextBlock, language: LanguageFilter): boolean =>
  language === allFilter || block.language === language;

const downloadJson = (fileName: string, payload: unknown): void => {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

function App() {
  const [blocks, setBlocks] = useState<TextBlock[]>([]);
  const [query, setQuery] = useState("");
  const [language, setLanguage] = useState<LanguageFilter>(allFilter);
  const [department, setDepartment] = useState(() => {
    const storedDepartment = readStorage(departmentStorageKey);
    return storedDepartment && storedDepartment !== allFilter
      ? normalizeDepartmentKey(storedDepartment)
      : allFilter;
  });
  const [category, setCategory] = useState<string | "ALL">(allFilter);
  const [favoriteMode, setFavoriteMode] = useState<FavoriteMode>("ALL");
  const [favorites, setFavorites] = useState(readFavorites);
  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [notice, setNotice] = useState<Notice | null>(null);
  const [busyBlockId, setBusyBlockId] = useState<string | null>(null);
  const importInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    let isMounted = true;

    async function loadBlocks() {
      try {
        const loadedBlocks = await textBlockService.getTextBlocks();

        if (isMounted) {
          setBlocks(loadedBlocks);
          setLoadError("");
        }
      } catch (error) {
        console.error("Could not load text block data.", error);

        if (isMounted) {
          setLoadError("Could not load text block data.");
        }
      } finally {
        if (isMounted) {
          setIsLoading(false);
        }
      }
    }

    loadBlocks();

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    writeStorage(favoriteStorageKey, JSON.stringify([...favorites].sort()));
  }, [favorites]);

  useEffect(() => {
    writeStorage(departmentStorageKey, department);
  }, [department]);

  const departments = useMemo(() => buildDepartmentOptions(blocks), [blocks]);
  const departmentKeys = useMemo(() => new Set(departments.map((item) => item.key)), [departments]);

  useEffect(() => {
    if (department !== allFilter && departmentKeys.size > 0 && !departmentKeys.has(department)) {
      setDepartment(allFilter);
    }
  }, [department, departmentKeys]);

  const departmentScopedBlocks = useMemo(
    () => blocks.filter((block) => blockMatchesDepartment(block, department)),
    [blocks, department]
  );

  const categories = useMemo(
    () => buildCategoryOptions(departmentScopedBlocks.filter((block) => blockMatchesLanguage(block, language))),
    [departmentScopedBlocks, language]
  );
  const categoryKeys = useMemo(() => new Set(categories.map((item) => item.key)), [categories]);

  useEffect(() => {
    if (category !== allFilter && !categoryKeys.has(category)) {
      setCategory(allFilter);
    }
  }, [category, categoryKeys]);

  const results = useMemo(() => {
    const searchedBlocks = searchTextBlocks(blocks, query, language, category, department);

    if (favoriteMode === "FAVORITES") {
      return searchedBlocks.filter((block) => favorites.has(block.id));
    }

    return searchedBlocks;
  }, [blocks, category, department, favoriteMode, favorites, language, query]);

  const handleInsert = async (block: TextBlock) => {
    setBusyBlockId(block.id);
    setNotice(null);

    try {
      await insertTextIntoEmail(block.text, block.formattedText);
      setNotice({ kind: "success", message: "Text inserted." });
    } catch (error) {
      setNotice({ kind: "error", message: getFriendlyInsertError(error) });
    } finally {
      setBusyBlockId(null);
    }
  };

  const handleCopy = async (block: TextBlock) => {
    setBusyBlockId(block.id);
    setNotice(null);

    try {
      await copyTextToClipboard(block.text);
      setNotice({ kind: "success", message: "Text copied." });
    } catch (error) {
      console.error("Clipboard copy failed.", error);
      setNotice({ kind: "error", message: "Clipboard copy failed." });
    } finally {
      setBusyBlockId(null);
    }
  };

  const toggleFavorite = (blockId: string) => {
    setFavorites((current) => {
      const next = new Set(current);

      if (next.has(blockId)) {
        next.delete(blockId);
      } else {
        next.add(blockId);
      }

      return next;
    });
  };

  const handleExportFavorites = () => {
    downloadJson(`textblocsfit-favorites-${new Date().toISOString().slice(0, 10)}.json`, {
      schemaVersion: 1,
      exportedAt: new Date().toISOString(),
      favorites: [...favorites].sort(),
      defaultDepartment:
        department === allFilter
          ? allFilter
          : departments.find((item) => item.key === department)?.label || fallbackDepartment
    });
  };

  const handleImportFavorites = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) {
      return;
    }

    try {
      const backup = JSON.parse(await file.text()) as FavoritesBackup;

      if (backup.schemaVersion !== 1 || !Array.isArray(backup.favorites)) {
        throw new Error("Unsupported favorites backup.");
      }

      setFavorites(new Set(backup.favorites.filter((item) => typeof item === "string")));

      if (backup.defaultDepartment) {
        const key =
          backup.defaultDepartment === allFilter
            ? allFilter
            : normalizeDepartmentKey(backup.defaultDepartment);

        if (key === allFilter || departmentKeys.has(key)) {
          setDepartment(key);
        }
      }

      setNotice({ kind: "success", message: "Favorites imported." });
    } catch (error) {
      console.error("Could not import favorites.", error);
      setNotice({ kind: "error", message: "Could not import favorites backup." });
    }
  };

  return (
    <main className="app-shell">
      <header className="pane-header">
        <div>
          <p className="eyebrow">TextBlocsFIT</p>
          <h1>STC Textblocs v{appVersion}</h1>
        </div>
      </header>

      <section className="controls" aria-label="Text block filters">
        <label className="search-field">
          <Search aria-hidden="true" size={18} />
          <input
            type="search"
            value={query}
            placeholder="Search text blocks..."
            onChange={(event) => setQuery(event.target.value)}
          />
        </label>

        <label className="select-field">
          <span className="filter-label">Department</span>
          <select value={department} onChange={(event) => setDepartment(event.target.value)}>
            <option value={allFilter}>All departments</option>
            {departments.map((item) => (
              <option key={item.key} value={item.key}>
                {item.label}
              </option>
            ))}
          </select>
        </label>

        <div className="filter-group" aria-label="Language">
          <span className="filter-label">Language</span>
          <div className="segmented-control">
            {languages.map((item) => (
              <button
                key={item.code}
                type="button"
                className={item.code === language ? "is-active" : ""}
                aria-pressed={item.code === language}
                onClick={() => setLanguage(item.code)}
              >
                {item.label}
              </button>
            ))}
          </div>
        </div>

        <label className="select-field">
          <span className="filter-label">Category</span>
          <select value={category} onChange={(event) => setCategory(event.target.value)}>
            <option value={allFilter}>All categories</option>
            {categories.map((item) => (
              <option key={item.key} value={item.key}>
                {item.label}
              </option>
            ))}
          </select>
        </label>

        <div className="filter-group" aria-label="View">
          <span className="filter-label">View</span>
          <div className="segmented-control view-control">
            <button
              type="button"
              className={favoriteMode === "ALL" ? "is-active" : ""}
              aria-pressed={favoriteMode === "ALL"}
              onClick={() => setFavoriteMode("ALL")}
            >
              All
            </button>
            <button
              type="button"
              className={favoriteMode === "FAVORITES" ? "is-active" : ""}
              aria-pressed={favoriteMode === "FAVORITES"}
              onClick={() => setFavoriteMode("FAVORITES")}
            >
              Favorites
            </button>
          </div>
        </div>

        <details className="backup-panel">
          <summary>Backup favorites</summary>
          <div className="backup-actions">
            <button type="button" className="tertiary-button" onClick={handleExportFavorites}>
              <Download aria-hidden="true" size={16} />
              <span>Export</span>
            </button>
            <button type="button" className="tertiary-button" onClick={() => importInputRef.current?.click()}>
              <Upload aria-hidden="true" size={16} />
              <span>Import</span>
            </button>
            <input
              ref={importInputRef}
              className="visually-hidden"
              type="file"
              accept="application/json,.json"
              onChange={handleImportFavorites}
            />
          </div>
        </details>
      </section>

      {notice ? (
        <div className={`notice notice-${notice.kind}`} role="status">
          {notice.kind === "success" ? <Check aria-hidden="true" size={16} /> : <AlertCircle aria-hidden="true" size={16} />}
          <span>{notice.message}</span>
        </div>
      ) : null}

      <section className="results-section" aria-live="polite">
        <div className="results-heading">
          <h2>Results</h2>
          <span>{isLoading ? "Loading" : `${results.length} found`}</span>
        </div>

        {isLoading ? (
          <div className="empty-state">
            <Loader2 aria-hidden="true" className="spin" size={22} />
            <span>Loading text blocks...</span>
          </div>
        ) : null}

        {!isLoading && loadError ? (
          <div className="empty-state is-error">
            <AlertCircle aria-hidden="true" size={22} />
            <span>{loadError}</span>
          </div>
        ) : null}

        {!isLoading && !loadError && results.length === 0 ? (
          <div className="empty-state">
            <Search aria-hidden="true" size={22} />
            <span>No matching text blocks found.</span>
          </div>
        ) : null}

        {!isLoading && !loadError && results.length > 0 ? (
          <div className="result-list">
            {results.map((block) => (
              <article className="result-card" key={block.id}>
                <div className="result-meta">
                  <span>{departmentLabel(block)} · {block.category || "Uncategorized"}</span>
                  <strong>{block.topic || "Untitled"}</strong>
                  <small>Usage: {block.usage || "General"}</small>
                </div>

                <div className="card-badges">
                  <button
                    type="button"
                    className={`favorite-button ${favorites.has(block.id) ? "is-active" : ""}`}
                    aria-label={favorites.has(block.id) ? "Remove favorite" : "Add favorite"}
                    aria-pressed={favorites.has(block.id)}
                    onClick={() => toggleFavorite(block.id)}
                  >
                    <Star aria-hidden="true" size={16} fill={favorites.has(block.id) ? "currentColor" : "none"} />
                  </button>
                  <div className="language-pill">{block.language}</div>
                </div>

                <p className="preview-text">{block.text}</p>

                <div className="card-actions">
                  <button
                    type="button"
                    className="primary-button"
                    disabled={busyBlockId === block.id}
                    onClick={() => handleInsert(block)}
                  >
                    <CornerDownLeft aria-hidden="true" size={16} />
                    <span>Insert</span>
                  </button>
                  <button
                    type="button"
                    className="secondary-button"
                    disabled={busyBlockId === block.id}
                    onClick={() => handleCopy(block)}
                  >
                    <Clipboard aria-hidden="true" size={16} />
                    <span>Copy</span>
                  </button>
                </div>
              </article>
            ))}
          </div>
        ) : null}
      </section>
    </main>
  );
}

export default App;
