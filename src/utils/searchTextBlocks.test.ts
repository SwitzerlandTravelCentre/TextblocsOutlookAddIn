import { describe, expect, it } from "vitest";
import { TextBlock } from "../models/textBlock";
import { searchTextBlocks } from "./searchTextBlocks";

const blocks: TextBlock[] = [
  {
    id: "1-de",
    sourceItemId: "1",
    category: "Bahn Tickets",
    topic: "Gueltigkeit Bahntickets",
    usage: "Information",
    language: "DE",
    languageLabel: "German",
    weight: 0,
    text: "Die Bahntickets sind gueltig."
  },
  {
    id: "2-en",
    sourceItemId: "2",
    category: "Payments",
    topic: "Payment link",
    usage: "Outstanding amount",
    language: "EN",
    languageLabel: "English",
    weight: 0,
    text: "Please use the payment link."
  },
  {
    id: "3-fr",
    sourceItemId: "3",
    category: "Payments",
    topic: "General information",
    usage: "Standard",
    language: "FR",
    languageLabel: "French",
    weight: 10,
    text: "Merci."
  }
];

describe("searchTextBlocks", () => {
  it("searches case-insensitively across metadata and text", () => {
    expect(searchTextBlocks(blocks, "bahntickets")).toHaveLength(1);
    expect(searchTextBlocks(blocks, "PAYMENT")).toHaveLength(2);
  });

  it("filters by language and category", () => {
    const result = searchTextBlocks(blocks, "", "EN", "Payments");

    expect(result).toHaveLength(1);
    expect(result[0].id).toBe("2-en");
  });

  it("sorts the default list by weight first", () => {
    const result = searchTextBlocks(blocks, "");

    expect(result[0].id).toBe("3-fr");
  });

  it("keeps search result relevance ahead of weight", () => {
    const result = searchTextBlocks(blocks, "payment");

    expect(result.map((block) => block.id)).toEqual(["2-en", "3-fr"]);
  });
});
