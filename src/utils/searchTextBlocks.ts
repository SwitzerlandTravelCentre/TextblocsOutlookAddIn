import { TextBlock, TextBlockLanguage } from "../models/textBlock";

export type LanguageFilter = TextBlockLanguage | "ALL";
export const allFilter = "ALL";
export const fallbackDepartment = "General";

type CategoryFilter = string | "ALL";
type DepartmentFilter = string | "ALL";

const normalize = (value: string): string =>
  value
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

export const departmentLabel = (block: Pick<TextBlock, "department">): string =>
  block.department?.trim() || fallbackDepartment;

export const normalizeDepartmentKey = (value?: string): string =>
  normalize(value || fallbackDepartment);

export const normalizeCategoryKey = (value?: string): string => normalize(value || "");

const searchableText = (block: TextBlock): string =>
  [block.department, block.category, block.topic, block.usage, block.languageLabel, block.text]
    .join(" ")
    .toLowerCase();

const relevanceScore = (block: TextBlock, query: string): number => {
  if (!query) {
    return 0;
  }

  const checks = [block.topic, block.usage, block.category, block.text, block.languageLabel];
  const index = checks.findIndex((value) => value.toLowerCase().includes(query));

  return index === -1 ? checks.length : index;
};

const stableLabel = (block: TextBlock): string =>
  [departmentLabel(block), block.category, block.topic, block.usage, block.language].join("|");

export function searchTextBlocks(
  blocks: TextBlock[],
  query: string,
  language: LanguageFilter = allFilter,
  category: CategoryFilter = allFilter,
  department: DepartmentFilter = allFilter
): TextBlock[] {
  const normalizedQuery = normalize(query);
  const normalizedCategory = category === allFilter ? allFilter : normalizeCategoryKey(category);
  const normalizedDepartment =
    department === allFilter ? allFilter : normalizeDepartmentKey(department);
  const hasQuery = normalizedQuery.length > 0;

  return blocks
    .filter((block) => language === allFilter || block.language === language)
    .filter(
      (block) =>
        normalizedDepartment === allFilter ||
        normalizeDepartmentKey(block.department) === normalizedDepartment
    )
    .filter(
      (block) =>
        normalizedCategory === allFilter ||
        normalizeCategoryKey(block.category) === normalizedCategory
    )
    .filter((block) => !hasQuery || searchableText(block).includes(normalizedQuery))
    .sort((a, b) => {
      if (!hasQuery) {
        const weight = b.weight - a.weight;

        if (weight !== 0) {
          return weight;
        }
      }

      const relevance = relevanceScore(a, normalizedQuery) - relevanceScore(b, normalizedQuery);

      if (relevance !== 0) {
        return relevance;
      }

      return stableLabel(a).localeCompare(stableLabel(b));
    });
}
