import { afterEach, describe, expect, it, vi } from "vitest";
import { TextBlockDataService } from "./textBlockDataService";

describe("TextBlockDataService", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("loads valid blocks from the static JSON data file", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        schemaVersion: 1,
        blockCount: 1,
        blocks: [
          {
            id: "hotel-confirmation-2-de",
            sourceItemId: "Template Outlook!2",
            category: "Hotel",
            topic: "Confirmation",
            usage: "",
            language: "DE",
            languageLabel: "German",
            text: "Danke.",
            formattedText: [
              { text: "Danke", bold: true },
              { text: ".", highlight: "required", href: "https://example.com" }
            ]
          }
        ]
      })
    });
    vi.stubGlobal("fetch", fetchMock);

    const blocks = await new TextBlockDataService().getTextBlocks();

    expect(blocks).toHaveLength(1);
    expect(blocks[0].weight).toBe(0);
    expect(blocks[0].formattedText).toEqual([
      { text: "Danke", bold: true },
      { text: ".", highlight: "required", href: "https://example.com" }
    ]);
    expect(fetchMock).toHaveBeenCalledWith(
      "/data/textblocks.json",
      expect.objectContaining({ cache: "no-cache" })
    );
  });

  it("rejects unsupported data schemas", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ schemaVersion: 99, blocks: [] })
      })
    );

    await expect(new TextBlockDataService().getTextBlocks()).rejects.toThrow(
      "unsupported schema"
    );
  });
});
