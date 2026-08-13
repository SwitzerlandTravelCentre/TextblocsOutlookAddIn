import { describe, expect, it } from "vitest";
import { textToOutlookHtml } from "./officeService";

describe("textToOutlookHtml", () => {
  it("formats blank lines as paragraphs and inline line breaks as br tags", () => {
    expect(textToOutlookHtml("Line 1\nLine 2\n\nLine 3")).toBe(
      "<p style=\"margin:0 0 12px 0;\">Line 1<br>Line 2</p><p style=\"margin:0 0 12px 0;\">Line 3</p>"
    );
  });

  it("converts supported links to black anchors", () => {
    expect(textToOutlookHtml("See https://www.sbb.ch/de. Email mailto:test@example.com")).toBe(
      '<p style="margin:0 0 12px 0;">See <a href="https://www.sbb.ch/de" style="color:#000;text-decoration:underline;">https://www.sbb.ch/de</a>. Email <a href="mailto:test@example.com" style="color:#000;text-decoration:underline;">mailto:test@example.com</a></p>'
    );
  });

  it("keeps closing punctuation outside linked URLs", () => {
    expect(textToOutlookHtml("Terms (https://example.com/gtc).")).toBe(
      '<p style="margin:0 0 12px 0;">Terms (<a href="https://example.com/gtc" style="color:#000;text-decoration:underline;">https://example.com/gtc</a>).</p>'
    );
  });

  it("renders exported Excel rich text formatting", () => {
    expect(
      textToOutlookHtml("Bold Link Change\n\nNext", [
        { text: "Bold", bold: true },
        { text: " " },
        { text: "Link", href: "https://example.com" },
        { text: " " },
        { text: "Change", highlight: "required" },
        { text: "\n\nNext" }
      ])
    ).toBe(
      '<p style="margin:0 0 12px 0;"><strong>Bold</strong> <a href="https://example.com" style="color:#000;text-decoration:underline;">Link</a> <span style="background-color:#fff2cc;color:#000;">Change</span></p><p style="margin:0 0 12px 0;">Next</p>'
    );
  });

  it("escapes HTML-sensitive characters instead of trusting source HTML", () => {
    expect(textToOutlookHtml("Line <2> & \"3\"\n<script>alert('x')</script>")).toBe(
      "<p style=\"margin:0 0 12px 0;\">Line &lt;2&gt; &amp; &quot;3&quot;<br>&lt;script&gt;alert(&#39;x&#39;)&lt;/script&gt;</p>"
    );
  });
});
