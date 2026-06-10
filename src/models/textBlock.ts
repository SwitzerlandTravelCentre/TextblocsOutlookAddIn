export type TextBlockLanguage = "DE" | "FR" | "EN" | "IT";

export interface TextBlock {
  id: string;
  sourceItemId: string;
  sourceSheet?: string;
  sourceRow?: number;
  category: string;
  topic: string;
  usage: string;
  language: TextBlockLanguage;
  languageLabel: string;
  weight: number;
  text: string;
}

export interface TextBlockProvider {
  getTextBlocks(): Promise<TextBlock[]>;
}

export interface TextBlockDataFile {
  schemaVersion: number;
  blockCount: number;
  blocks: TextBlock[];
}
