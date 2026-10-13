import { TextBlockFormattedRun } from "../models/textBlock";

function normalizeLineBreaks(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

function toPlainTextForOutlook(text: string): string {
  return normalizeLineBreaks(text).replace(/\n/g, "\r\n");
}

function escapeHtmlText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;");
}

function trimUrlPunctuation(value: string): { url: string; suffix: string } {
  let url = value;
  let suffix = "";

  while (/[.,!?;:]$/.test(url)) {
    suffix = url.charAt(url.length - 1) + suffix;
    url = url.slice(0, -1);
  }

  while (url.endsWith(")") && !url.includes("(")) {
    suffix = ")" + suffix;
    url = url.slice(0, -1);
  }

  while (url.endsWith("]") && !url.includes("[")) {
    suffix = "]" + suffix;
    url = url.slice(0, -1);
  }

  return { url, suffix };
}

const outlookParagraphStyle = "margin:0 0 12px 0;";
const outlookListStyle = "margin:0 0 12px 22px;padding-left:18px;";
const outlookListItemStyle = "margin:0 0 4px 0;";
const outlookLinkStyle = "color:#000;text-decoration:underline;";
const outlookHighlightStyle = "background-color:#fff2cc;color:#000;";
const markdownBulletPrefixPattern = /^[ \t]*(?:\u2022[\t ]+|-[ \t]+)/;

function isAllowedHref(value: string): boolean {
  return /^(?:https?:\/\/|mailto:)/i.test(value);
}

function anchorHtml(href: string, contentHtml: string): string {
  const escapedHref = escapeHtmlText(href);
  return `<a href="${escapedHref}" style="${outlookLinkStyle}">${contentHtml}</a>`;
}

function textLineToHtml(text: string): string {
  const urlPattern = /\b(?:https?:\/\/|mailto:)[^\s<]+/gi;
  let html = "";
  let lastIndex = 0;

  for (const match of text.matchAll(urlPattern)) {
    const matchedUrl = match[0];
    const matchIndex = match.index ?? 0;
    const { url, suffix } = trimUrlPunctuation(matchedUrl);

    html += escapeHtmlText(text.slice(lastIndex, matchIndex));

    if (url) {
      const escapedUrl = escapeHtmlText(url);
      html += anchorHtml(url, escapedUrl);
    }

    html += escapeHtmlText(suffix);
    lastIndex = matchIndex + matchedUrl.length;
  }

  return html + escapeHtmlText(text.slice(lastIndex));
}

function inlinePlainTextToHtml(text: string): string {
  return text.split("\n").map(textLineToHtml).join("<br>");
}

function inlineDisplayTextToHtml(text: string): string {
  return text.split("\n").map(escapeHtmlText).join("<br>");
}

function formattedRunToHtml(run: TextBlockFormattedRun): string {
  const href = run.href && isAllowedHref(run.href) ? run.href : "";
  let html = href ? inlineDisplayTextToHtml(run.text) : inlinePlainTextToHtml(run.text);

  if (href) {
    html = anchorHtml(href, html);
  }

  if (run.highlight === "required") {
    html = `<span style="${outlookHighlightStyle}">${html}</span>`;
  }

  if (run.bold) {
    html = `<strong>${html}</strong>`;
  }

  return html;
}

function runsToPlainText(runs: TextBlockFormattedRun[]): string {
  return runs.map((run) => run.text).join("");
}

function splitRunsByLine(runs: TextBlockFormattedRun[]): TextBlockFormattedRun[][] {
  const lines: TextBlockFormattedRun[][] = [[]];

  runs.forEach((run) => {
    normalizeLineBreaks(run.text)
      .split("\n")
      .forEach((part, index) => {
        if (index > 0) {
          lines.push([]);
        }

        if (part) {
          lines[lines.length - 1].push({ ...run, text: part });
        }
      });
  });

  return lines;
}

function dropLeadingCharactersFromRuns(
  runs: TextBlockFormattedRun[],
  characterCount: number
): TextBlockFormattedRun[] {
  let remainingCharacters = characterCount;
  const output: TextBlockFormattedRun[] = [];

  runs.forEach((run) => {
    if (remainingCharacters >= run.text.length) {
      remainingCharacters -= run.text.length;
      return;
    }

    const text = remainingCharacters > 0 ? run.text.slice(remainingCharacters) : run.text;
    remainingCharacters = 0;

    if (text) {
      output.push({ ...run, text });
    }
  });

  return output;
}

function stripMarkdownBulletPrefixFromRuns(
  runs: TextBlockFormattedRun[]
): TextBlockFormattedRun[] | null {
  const match = markdownBulletPrefixPattern.exec(runsToPlainText(runs));

  if (!match) {
    return null;
  }

  return dropLeadingCharactersFromRuns(runs, match[0].length);
}

function splitFormattedParagraphs(runs: TextBlockFormattedRun[]): TextBlockFormattedRun[][] {
  const paragraphs: TextBlockFormattedRun[][] = [[]];

  runs.forEach((run) => {
    normalizeLineBreaks(run.text)
      .split(/(\n{2,})/)
      .forEach((part) => {
        if (!part) {
          return;
        }

        if (/^\n{2,}$/.test(part)) {
          if (paragraphs[paragraphs.length - 1].some((item) => item.text.trim().length > 0)) {
            paragraphs.push([]);
          }
          return;
        }

        paragraphs[paragraphs.length - 1].push({ ...run, text: part });
      });
  });

  return paragraphs.filter((paragraph) =>
    paragraph.some((run) => run.text.trim().length > 0)
  );
}

function formattedLinesToParagraphHtml(lines: TextBlockFormattedRun[][]): string {
  return `<p style="${outlookParagraphStyle}">${lines
    .map((line) => line.map(formattedRunToHtml).join(""))
    .join("<br>")}</p>`;
}

function formattedListToHtml(items: TextBlockFormattedRun[][]): string {
  return `<ul style="${outlookListStyle}">${items
    .map(
      (item) =>
        `<li style="${outlookListItemStyle}">${item.map(formattedRunToHtml).join("")}</li>`
    )
    .join("")}</ul>`;
}

function formattedParagraphToOutlookHtml(runs: TextBlockFormattedRun[]): string {
  const blocks: string[] = [];
  let paragraphLines: TextBlockFormattedRun[][] = [];
  let listItems: TextBlockFormattedRun[][] = [];

  const flushParagraph = () => {
    if (!paragraphLines.length) {
      return;
    }

    blocks.push(formattedLinesToParagraphHtml(paragraphLines));
    paragraphLines = [];
  };

  const flushList = () => {
    if (!listItems.length) {
      return;
    }

    blocks.push(formattedListToHtml(listItems));
    listItems = [];
  };

  splitRunsByLine(runs).forEach((line) => {
    if (!line.some((run) => run.text.trim().length > 0)) {
      return;
    }

    const listItem = stripMarkdownBulletPrefixFromRuns(line);

    if (listItem) {
      flushParagraph();
      listItems.push(listItem);
      return;
    }

    if (listItems.length) {
      listItems[listItems.length - 1].push({ text: "\n" }, ...line);
      return;
    }

    flushList();
    paragraphLines.push(line);
  });

  flushParagraph();
  flushList();

  return blocks.join("");
}

function formattedTextToOutlookHtml(runs: TextBlockFormattedRun[]): string {
  return splitFormattedParagraphs(runs)
    .map((paragraph) => formattedParagraphToOutlookHtml(paragraph))
    .join("");
}

function stripMarkdownBulletPrefixFromText(text: string): string | null {
  const match = markdownBulletPrefixPattern.exec(text);

  if (!match) {
    return null;
  }

  return text.slice(match[0].length);
}

function plainLinesToParagraphHtml(lines: string[]): string {
  return `<p style="${outlookParagraphStyle}">${lines.map(textLineToHtml).join("<br>")}</p>`;
}

function plainListToHtml(items: string[][]): string {
  return `<ul style="${outlookListStyle}">${items
    .map(
      (item) =>
        `<li style="${outlookListItemStyle}">${item.map(textLineToHtml).join("<br>")}</li>`
    )
    .join("")}</ul>`;
}

function plainParagraphToOutlookHtml(paragraph: string): string {
  const blocks: string[] = [];
  let paragraphLines: string[] = [];
  let listItems: string[][] = [];

  const flushParagraph = () => {
    if (!paragraphLines.length) {
      return;
    }

    blocks.push(plainLinesToParagraphHtml(paragraphLines));
    paragraphLines = [];
  };

  const flushList = () => {
    if (!listItems.length) {
      return;
    }

    blocks.push(plainListToHtml(listItems));
    listItems = [];
  };

  paragraph.split("\n").forEach((line) => {
    if (!line.trim()) {
      return;
    }

    const listItem = stripMarkdownBulletPrefixFromText(line);

    if (listItem !== null) {
      flushParagraph();
      listItems.push([listItem]);
      return;
    }

    if (listItems.length) {
      listItems[listItems.length - 1].push(line);
      return;
    }

    flushList();
    paragraphLines.push(line);
  });

  flushParagraph();
  flushList();

  return blocks.join("");
}

export function textToOutlookHtml(text: string, formattedText?: TextBlockFormattedRun[]): string {
  if (formattedText?.some((run) => run.text.trim().length > 0)) {
    return formattedTextToOutlookHtml(formattedText);
  }

  const paragraphs = normalizeLineBreaks(text)
    .split(/\n[ \t]*\n+/)
    .filter((paragraph) => paragraph.trim().length > 0);

  return paragraphs
    .map((paragraph) => plainParagraphToOutlookHtml(paragraph))
    .join("");
}

function setSelectedBodyData(
  body: Office.Body,
  data: string,
  coercionType: Office.CoercionType
): Promise<void> {
  return new Promise((resolve, reject) => {
    body.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(result.error ?? new Error("Could not insert text into Outlook."));
    });
  });
}

export function insertTextIntoEmail(
  text: string,
  formattedText?: TextBlockFormattedRun[]
): Promise<void> {
  const body = typeof Office === "undefined" ? undefined : Office.context?.mailbox?.item?.body;

  if (!body) {
    return Promise.reject(new Error("Outlook compose body is not available."));
  }

  return setSelectedBodyData(body, textToOutlookHtml(text, formattedText), Office.CoercionType.Html).catch(
    async (error) => {
      console.warn("HTML insertion failed. Falling back to plain text insertion.", error);
      await setSelectedBodyData(body, toPlainTextForOutlook(text), Office.CoercionType.Text);
    }
  );
}

export async function copyTextToClipboard(text: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.style.position = "fixed";
  textArea.style.left = "-9999px";
  textArea.style.top = "0";
  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();

  try {
    const copied = document.execCommand("copy");

    if (!copied) {
      throw new Error("Browser copy command failed.");
    }
  } finally {
    document.body.removeChild(textArea);
  }
}
