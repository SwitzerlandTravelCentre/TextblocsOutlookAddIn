import { TextBlock, TextBlockLanguage } from "../models/textBlock";

type LanguageFilter = TextBlockLanguage | "ALL";
type CategoryFilter = string | "ALL";

const normalize = (value: string): string => value.trim().toLowerCase();

const searchableText = (block: TextBlock): string =>
  [block.category, block.topic, block.usage, block.languageLabel, block.text]
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
  [block.category, block.topic, block.usage, block.language].join("|");

export function searchTextBlocks(
  blocks: TextBlock[],
  query: string,
  language: LanguageFilter = "ALL",
  category: CategoryFilter = "ALL"
): TextBlock[] {
  const normalizedQuery = normalize(query);
  const hasQuery = normalizedQuery.length > 0;

  return blocks
    .filter((block) => language === "ALL" || block.language === language)
    .filter((block) => category === "ALL" || block.category === category)
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
