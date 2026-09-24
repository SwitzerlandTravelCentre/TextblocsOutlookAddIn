import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { TextBlockDataFile } from "../models/textBlock";

describe("textblocks.json", () => {
  const data = JSON.parse(
    readFileSync(resolve("public/data/textblocks.json"), "utf-8")
  ) as TextBlockDataFile;

  it("contains the generated Excel data in the expected schema", () => {
    expect(data.schemaVersion).toBe(1);
    expect(data.blocks).toHaveLength(data.blockCount);
    expect(data.blockCount).toBeGreaterThan(0);
    expect(new Set(data.blocks.map((block) => block.id)).size).toBe(data.blockCount);
  });

  it("contains all supported languages", () => {
    expect(new Set(data.blocks.map((block) => block.language))).toEqual(
      new Set(["DE", "FR", "EN", "IT"])
    );
  });

  it("includes a numeric weight for every block", () => {
    expect(data.blocks.every((block) => typeof block.weight === "number")).toBe(true);
  });

  it("preserves UTF-8 text content", () => {
    expect(data.blocks.some((block) => block.text.includes("f\u00FCr"))).toBe(true);
    expect(data.blocks.some((block) => block.text.includes("r\u00E9servation"))).toBe(true);
  });
});
