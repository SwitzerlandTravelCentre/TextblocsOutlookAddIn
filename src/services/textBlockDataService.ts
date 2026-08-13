import {
  TextBlock,
  TextBlockDataFile,
  TextBlockFormattedRun,
  TextBlockLanguage,
  TextBlockProvider
} from "../models/textBlock";

const languages: TextBlockLanguage[] = ["DE", "FR", "EN", "IT"];
const textBlockDataRelativePath = "../../data/textblocks.json";

function getTextBlockDataUrl(): string {
  if (typeof window === "undefined") {
    return "/data/textblocks.json";
  }

  // SharePoint may host the add-in in a nested folder. Resolve the JSON file
  // relative to taskpane.html so the same build works on localhost and SharePoint.
  // Change this path only if textblocks.json is moved outside the built dist/data folder.
  return new URL(textBlockDataRelativePath, window.location.href).toString();
}

function isFormattedTextRun(value: unknown): value is TextBlockFormattedRun {
  if (!value || typeof value !== "object") {
    return false;
  }

  const run = value as Partial<TextBlockFormattedRun>;
  const hasValidBold = run.bold === undefined || typeof run.bold === "boolean";
  const hasValidHighlight = run.highlight === undefined || run.highlight === "required";
  const hasValidHref = run.href === undefined || typeof run.href === "string";

  return typeof run.text === "string" && hasValidBold && hasValidHighlight && hasValidHref;
}

function isTextBlock(value: unknown): value is TextBlock {
  if (!value || typeof value !== "object") {
    return false;
  }

  const block = value as Partial<TextBlock>;
  const hasValidWeight = block.weight === undefined || typeof block.weight === "number";
  const hasValidFormattedText =
    block.formattedText === undefined ||
    (Array.isArray(block.formattedText) && block.formattedText.every(isFormattedTextRun));

  return (
    typeof block.id === "string" &&
    typeof block.sourceItemId === "string" &&
    typeof block.category === "string" &&
    typeof block.topic === "string" &&
    typeof block.usage === "string" &&
    typeof block.text === "string" &&
    typeof block.languageLabel === "string" &&
    hasValidWeight &&
    hasValidFormattedText &&
    !!block.language &&
    languages.includes(block.language)
  );
}

function withDefaultWeight(block: TextBlock): TextBlock {
  return {
    ...block,
    weight: block.weight ?? 0
  };
}

function parseTextBlockDataFile(value: unknown): TextBlock[] {
  if (!value || typeof value !== "object") {
    throw new Error("Text block data file is not an object.");
  }

  const data = value as Partial<TextBlockDataFile>;

  if (data.schemaVersion !== 1 || !Array.isArray(data.blocks)) {
    throw new Error("Text block data file has an unsupported schema.");
  }

  const blocks = data.blocks.filter(isTextBlock);

  if (blocks.length !== data.blocks.length) {
    throw new Error("Text block data file contains invalid records.");
  }

  return blocks.map(withDefaultWeight);
}

export class TextBlockDataService implements TextBlockProvider {
  async getTextBlocks(): Promise<TextBlock[]> {
    const response = await fetch(getTextBlockDataUrl(), {
      cache: "no-cache",
      headers: {
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      throw new Error(`Text block data request failed with status ${response.status}.`);
    }

    return parseTextBlockDataFile(await response.json());
  }
}
