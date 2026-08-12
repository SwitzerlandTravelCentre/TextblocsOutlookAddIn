import { AlertCircle, Check, Clipboard, CornerDownLeft, Loader2, Search } from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import { TextBlock, TextBlockLanguage } from "../models/textBlock";
import { copyTextToClipboard, insertTextIntoEmail } from "../services/officeService";
import { TextBlockDataService } from "../services/textBlockDataService";
import { searchTextBlocks } from "../utils/searchTextBlocks";

const languages: Array<{ code: TextBlockLanguage | "ALL"; label: string }> = [
  { code: "ALL", label: "All" },
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

const getFriendlyInsertError = (error: unknown): string => {
  console.error("Could not insert text into Outlook.", error);
  return "Could not insert text into Outlook. Make sure a message compose window is active.";
};

const textBlockService = new TextBlockDataService();

function App() {
  const [blocks, setBlocks] = useState<TextBlock[]>([]);
  const [query, setQuery] = useState("");
  const [language, setLanguage] = useState<TextBlockLanguage | "ALL">("ALL");
  const [category, setCategory] = useState<string | "ALL">("ALL");
  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [notice, setNotice] = useState<Notice | null>(null);
  const [busyBlockId, setBusyBlockId] = useState<string | null>(null);

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

  const categories = useMemo(
    () =>
      Array.from(new Set(blocks.map((block) => block.category).filter(Boolean))).sort((a, b) =>
        a.localeCompare(b)
      ),
    [blocks]
  );

  const results = useMemo(
    () => searchTextBlocks(blocks, query, language, category),
    [blocks, category, language, query]
  );

  const handleInsert = async (block: TextBlock) => {
    setBusyBlockId(block.id);
    setNotice(null);

    try {
      await insertTextIntoEmail(block.text);
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

  return (
    <main className="app-shell">
      <header className="pane-header">
        <div>
          <p className="eyebrow">TextBlocsFIT</p>
          <h1>STC Textblocs v1.0.1</h1>
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
            <option value="ALL">All categories</option>
            {categories.map((item) => (
              <option key={item} value={item}>
                {item}
              </option>
            ))}
          </select>
        </label>
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
                  <span>{block.category || "Uncategorized"}</span>
                  <strong>{block.topic || "Untitled"}</strong>
                  <small>Usage: {block.usage || "General"}</small>
                </div>

                <div className="language-pill">{block.language}</div>

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
