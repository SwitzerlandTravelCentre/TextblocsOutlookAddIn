# TextBlocsFIT Outlook Add-in

Internal Outlook task pane add-in for searching, previewing, copying, and inserting multilingual STC email text blocks.

## Current Architecture

- React + TypeScript + Vite Office.js task pane.
- Static JSON data file served with the add-in from `public/data/textblocks.json`.
- Excel workbook is the editable source of truth for content.
- One workbook row is converted into one selectable card per available language.
- Searches category, topic, usage, language label, and text.
- Filters by language and category.
- Shows the default list by descending `weight`; searched results still use search relevance first.
- Inserts selected text with `Office.context.mailbox.item.body.setSelectedDataAsync` as controlled HTML.
- Includes clipboard fallback.
- Supports Outlook task pane pinning in clients that support Mailbox requirement set 1.5 and `VersionOverridesV1_1`.

No external data source, Microsoft Graph, MSAL, Entra app registration, or runtime data permissions are required for the current static JSON version.

## Setup

1. Install dependencies:

   ```powershell
   npm install
   ```

2. Install the trusted local HTTPS certificate once:

   ```powershell
   npm run certs:install
   ```

3. Start the HTTPS dev server:

   ```powershell
   npm run dev
   ```

4. Sideload `manifest.xml` in Outlook with the task pane URL:

   ```text
   https://localhost:3000/src/taskpane/taskpane.html
   ```

## Updating Text Blocks

The source workbook lives in `source/STC_Textblocs_Source.xlsx`. The runtime data lives here:

```text
public/data/textblocks.json
```

To regenerate it from the Excel workbook:

```powershell
npm run convert:textblocks
```

This uses `tools/convert-xlsx-to-json-runner.txt`, which keeps UTF-8 formatting intact and avoids Windows or OneDrive cases where Python cannot read local `.py` files directly. The direct Python command is still available on machines where direct script execution works:

```powershell
python tools/convert-xlsx-to-json.py
```

The converter requires Python with `openpyxl` available. If `openpyxl` is missing, install it with:

```powershell
python -m pip install openpyxl
```

The converter expects these workbook columns:

```text
Category
Topic
German
French
English
Italian
```

`Usage` and `Weight` are optional. `Usage` defaults to an empty value and `Weight` defaults to `0`. Use higher weights, for example `9` or `10`, to place often-used text blocks higher in the default list. `Englisch` is also supported as an English column name for compatibility with earlier exports.

Line breaks are preserved when the JSON text contains newline characters. In Excel, add line breaks inside a cell with `Alt+Enter`; the converter writes them as `\n`, and the add-in formats them when inserting into Outlook. A blank line creates a new HTML paragraph with Outlook-friendly spacing, while a single line break inside a paragraph becomes an HTML line break. Keep the Excel source readable and do not type `<p>`, `<br>`, `<a>`, or other HTML tags in the workbook.

URLs in the Excel text can be maintained as plain `https://`, `http://`, or `mailto:` values. During Outlook insertion, the add-in turns them into clickable black links and keeps punctuation after the URL outside the link.

Markdown-style formatting in the Excel language cells is exported into the JSON for Outlook insertion:

- `**Text**` is inserted as bold text.
- `==Text==` is inserted with a yellow change-required highlight.
- `[Link text](https://...)` is inserted as a black underlined link and stores the URL as the link target.
- Lines starting with `- ` are inserted with a bullet marker.

Excel rich-text formatting remains supported as a fallback, but Markdown-style source text is preferred because it is easier to review and less fragile across Excel clients.

Recommended maintainer workflow:

1. Maintain text blocks in the Excel workbook as readable plain text.
2. Use blank lines for paragraph breaks and `Alt+Enter` for line breaks inside a cell.
3. Add links as `[visible label](https://...)`; add plain URLs only when the URL itself should be visible.
4. Generate `public/data/textblocks.json` from the workbook.
5. Upload the generated JSON to the hosted add-in location, for example `data/textblocks.json` in SharePoint. This can run weekly and may also be triggered manually.

## Deployment Notes

- Outlook permission is scoped to `ReadWriteItem` so the add-in can insert into the current compose body.
- The add-in does not send email, read mailbox contents, change recipients, modify subject, touch attachments, or alter signatures directly.
- For SharePoint hosting, upload the built `dist` files to the SharePoint folder and replace every `https://localhost:3000` URL in `manifest.xml` with that SharePoint folder URL.
- The production `AppDomain` in `manifest.xml` should be the SharePoint origin only, for example `https://stctravel.sharepoint.com`.
- The JSON data path is configured in `src/services/textBlockDataService.ts` as `textBlockDataRelativePath`; change it only if `textblocks.json` is moved outside the built `dist/data` folder.
- Office command icons are PNG files because Outlook manifest validation rejects SVG icon URLs.
- Text changes require regenerating `textblocks.json` and redeploying the hosted add-in files. End users do not update anything locally.

## Files

- `public/data/textblocks.json` contains the generated static text block data.
- `tools/convert-xlsx-to-json.py` converts the Excel source workbook to JSON.
- `tools/convert-xlsx-to-json-runner.txt` is the Windows/OneDrive-safe runner used by `npm run convert:textblocks`.
- `docs/xls-markdown-guide.md` explains how maintainers format Excel text blocks with Markdown.
- `src/models/textBlock.ts` contains runtime text block types.
- `src/services/textBlockDataService.ts` loads and validates the static JSON data.
- `src/utils/searchTextBlocks.ts` handles filtering and relevance sorting.
- `src/services/officeService.ts` wraps Office.js insertion and clipboard copy.
- `src/taskpane/App.tsx` contains the task pane UI.
