import { describe, expect, it } from "vitest";
import { textToOutlookHtml } from "./officeService";

describe("textToOutlookHtml", () => {
  it("preserves line breaks and escapes HTML-sensitive characters", () => {
    expect(textToOutlookHtml("Line 1\n\nLine <2> & \"3\"")).toBe(
      "Line 1<br><br>Line &lt;2&gt; &amp; &quot;3&quot;"
    );
  });
});
