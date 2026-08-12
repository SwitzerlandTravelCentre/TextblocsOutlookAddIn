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

The runtime data lives here:

```text
public/data/textblocks.json
```

To regenerate it from the Excel workbook:

```powershell
python tools/convert-xlsx-to-json.py "C:\path\to\Bausteine E-Mails (1).xlsx" --output public\data\textblocks.json
```

The converter requires Python with `openpyxl` available.

The converter expects these workbook columns:

```text
Category
Topic
Usage
Weight
German
French
English
Italian
```

`Weight` is optional and defaults to `0`. Use higher values, for example `9` or `10`, to place often-used text blocks higher in the default list. `Englisch` is also supported as an English column name for compatibility with earlier exports.

Line breaks are preserved when the JSON text contains newline characters. In Excel, add line breaks inside a cell with `Alt+Enter`; the converter writes them as `\n`, and the add-in formats them when inserting into Outlook. A blank line creates a new HTML paragraph with Outlook-friendly spacing, while a single line break inside a paragraph becomes an HTML line break. Keep the Excel source as plain text: do not maintain `<p>`, `<br>`, `<a>`, or other HTML tags in the workbook.

URLs in the Excel text can be maintained as plain `https://`, `http://`, or `mailto:` values. During Outlook insertion, the add-in turns them into clickable links and keeps punctuation after the URL outside the link.

Recommended maintainer workflow:

1. Maintain text blocks in the Excel workbook as readable plain text.
2. Use blank lines for paragraph breaks and `Alt+Enter` for line breaks inside a cell.
3. Add links as normal URLs, not as HTML.
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
- `src/models/textBlock.ts` contains runtime text block types.
- `src/services/textBlockDataService.ts` loads and validates the static JSON data.
- `src/utils/searchTextBlocks.ts` handles filtering and relevance sorting.
- `src/services/officeService.ts` wraps Office.js insertion and clipboard copy.
- `src/taskpane/App.tsx` contains the task pane UI.
