# TextBlocs Add-in Instructions

This file intentionally uses placeholders so the deployment can be reused in another environment without changing the code.

## Placeholders

| Placeholder | Meaning |
| --- | --- |
| `<SOURCE_XLSX_PATH>` | Local or synced filesystem path to the source workbook. |
| `<SOURCE_XLSX_URL>` | Browser URL of the source workbook, for documentation only. |
| `<ADDIN_HOST_URL>` | HTTPS folder URL where the built add-in files are hosted. |
| `<SHAREPOINT_ORIGIN>` | SharePoint origin, for example `https://tenant.sharepoint.com`. |
| `<OUTPUT_JSON_PATH>` | Generated JSON output path. Default: `public\data\textblocks.json`. |

## Data Update

The Outlook add-in does not read the Excel workbook at runtime. It reads the generated JSON file from the hosted add-in folder:

```text
data/textblocks.json
```

Download or sync the workbook from `<SOURCE_XLSX_URL>` so it is available as `<SOURCE_XLSX_PATH>`, then generate the JSON:

```powershell
npm run convert:textblocks -- "<SOURCE_XLSX_PATH>" --output "<OUTPUT_JSON_PATH>"
```

Typical repository-local output:

```powershell
npm run convert:textblocks -- "<SOURCE_XLSX_PATH>" --output public\data\textblocks.json
```

After generation, rebuild and redeploy the hosted files:

```powershell
npm run build
```

Upload the contents of `dist` to `<ADDIN_HOST_URL>`.

## Manifest Update

Before central deployment, replace every local development URL in `manifest.xml`:

```text
https://localhost:3000
```

with the production add-in host URL:

```text
<ADDIN_HOST_URL>
```

Set every production `AppDomain` entry to the SharePoint origin only:

```text
<SHAREPOINT_ORIGIN>
```

The URL must point to the folder where these built paths are reachable:

```text
assets/
data/
src/taskpane/taskpane.html
```

## Content Rules

Maintain text blocks in the Excel language columns with Markdown:

- `**Text**` for bold text.
- `==Text==` for yellow "must edit" highlights.
- `**==Text==**` for bold highlighted placeholders.
- `[visible label](https://example.com)` for links.
- `- ` at the beginning of a line for bullet points.

Do not enter HTML tags in the workbook. Use `Alt+Enter` for line breaks inside a cell and blank lines for paragraph breaks.

## Validation

Before deployment, run:

```powershell
npm test
npm run build
```

If only the text block content changed, regenerate `textblocks.json`, rebuild, upload the new `dist` files, and keep the manifest unchanged unless `<ADDIN_HOST_URL>` or command metadata changed.
