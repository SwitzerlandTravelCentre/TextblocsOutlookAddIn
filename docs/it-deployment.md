# IT Deployment Notes

## Development Sideloading

1. Install dependencies with `npm install`.
2. Run `npm run certs:install` once so Outlook trusts the local HTTPS dev server.
3. Run `npm run dev`.
4. Sideload `manifest.xml` into a test Outlook account.
5. Open Outlook Windows Classic, create a new email, and open the `Text Blocks` command.

## Data Source

The add-in currently uses a static JSON file:

```text
public/data/textblocks.json
```

The JSON is generated from the Excel source workbook with:

```powershell
python tools/convert-xlsx-to-json.py "C:\path\to\Bausteine E-Mails (1).xlsx" --output public\data\textblocks.json
```

The converter requires Python with `openpyxl` available.

The optional `Weight` workbook column is written to JSON as `weight`. Missing values default to `0`; higher values move entries higher in the default, non-search list. Search results are still ordered by search relevance.

Line breaks inside workbook cells should be entered with `Alt+Enter`. They are preserved in JSON as newline characters and inserted into Outlook as HTML line breaks.

The add-in loads the file at runtime from:

```text
/data/textblocks.json
```

No external data source, Microsoft Graph, MSAL, Entra app registration, or runtime data permissions are required for this version.

## Production Deployment

1. Build the app with `npm run build`.
2. Upload the contents of `dist` to the configured SharePoint HTTPS folder.
3. Replace every `https://localhost:3000` URL in `manifest.xml` with the SharePoint folder URL where the built files are reachable:

   ```text
   https://stctravel.sharepoint.com/sites/STCGlobal/SiteAssets/TextblocsOutlook
   ```

   Use the folder URL where `assets`, `data`, and `src/taskpane/taskpane.html` are reachable. Do not add a trailing slash.

4. Set the production `AppDomain` in `manifest.xml` to the SharePoint origin only:

   ```text
   https://stctravel.sharepoint.com
   ```

5. Use the adjusted manifest for sideloading or centralized deployment.
6. Keep manifest command icons as PNG files; Outlook manifest validation rejects SVG icon URLs.
7. Validate sideloading with a 2-5 user pilot group.
8. Deploy centrally through Microsoft 365 Admin Center to the selected users or groups.

The JSON path is configured in `src/services/textBlockDataService.ts`:

```typescript
const textBlockDataRelativePath = "../../data/textblocks.json";
```

Change this only if `textblocks.json` is moved outside the built `dist/data` folder.

If only text block content changes, regenerate `textblocks.json`, rebuild/redeploy the hosted files, and keep the manifest unchanged unless URLs or command metadata changed.

## Scope Boundaries

The add-in uses `ReadWriteItem` to insert selected text into the active compose body. It does not send emails, scan mailboxes, alter recipients, change subjects, modify attachments, or rewrite signatures.

## Pinning

The manifest includes both `VersionOverridesV1_0` and `VersionOverridesV1_1`. The v1.1 override enables `<SupportsPinning>true</SupportsPinning>` for the compose task pane, so users can pin the pane and keep it open across new compose windows in supported Outlook clients.
