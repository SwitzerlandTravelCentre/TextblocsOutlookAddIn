# Wöchentliche Aktualisierung der Textbausteine

Stand 24.09.2026: Clemens Bartlome hat `STC_Textblocs_Source.xlsx` in SharePoint als finale,
verbindliche Quelle bestätigt. Wochenworkflow im Originalrepository eingerichtet und aktiv;
Details und Nachweise im Abschnitt "Inbetriebnahme 24.09.2026" am Ende.

## Belegter Ausgangspunkt

- Repository-Commit: `fed41d80c7287a5ec453f2a119c5ba135649c487` (27.08.2026).
- Clemens bestätigt erfolgreiche Tests; offen ist der zugesagte wöchentliche Betrieb.
- Live: `https://calm-rock-03c40f810.7.azurestaticapps.net`, JSON vom 27.05.2026, 74 Quellzeilen / 296 Bausteine.
- Im Repository liegt eine neuere Excel-Datei. Sie ist eine lokale Testquelle und ersetzt nicht die von Clemens genannte SharePoint-Datei.
- Die bisherige Testdatei erzwingt exakt 296 Bausteine und würde jede legitime Mengenänderung blockieren. Sie prüft jetzt Nicht-Leere, Konsistenz und eindeutige IDs.

## Ablauf

`refresh-textblocks.yml` läuft montags um 04:17 UTC und kann manuell gestartet werden.
Es lädt exakt das konfigurierte SharePoint-Element herunter, prüft eine unveränderte
ETag während des Downloads, konvertiert mit Clemens' Skript und validiert das Ergebnis.
Danach laufen Anwendungstests und Build. Erst dann wird die bestehende Static Web App
aktualisiert. Abschliessend wird die öffentliche JSON gegen den SHA256 des Kandidaten geprüft.

Die private Arbeitsmappe wird nicht in das Repository eingecheckt, als Actions-Artefakt
hochgeladen oder in Logs ausgegeben. Das gebaute Add-in enthält die für das bisherige
statische Betriebsmodell erforderlichen Textbausteine. Sie bleiben wie beim bestehenden
Hosting öffentlich abrufbar; dieses Verfahren fügt keine Benutzeranmeldung hinzu.

Bei Download-, Konvertierungs-, Test- oder Buildfehlern wird nicht veröffentlicht und
der bisherige Live-Stand bleibt erhalten. Nach einem Deployment-/Nachprüfungsfehler ist
der tatsächliche Live-Stand zu prüfen; ein Rücksprung ist dann ein gesonderter Schritt,
keine bereits implementierte automatische Garantie. GitHub Actions zeigt den fehlgeschlagenen
Lauf; die zuständigen Betreiber müssen die Workflow-Fehlerbenachrichtigungen abonnieren.

## Einmalige Einrichtung im STC-Tenant und Originalrepository

1. Die SharePoint-Quelle mit einem berechtigten STC-Konto öffnen und prüfen.
   Drive-ID und Item-ID des tatsächlichen `STC_Textblocs_Source.xlsx` ermitteln.
   Diese IDs bleiben bei Umbenennung innerhalb derselben Bibliothek stabil; Löschen und
   Neuerstellen der Datei erfordert eine neue Zuordnung.
2. Eigene Entra-App für den unbeaufsichtigten Quellzugriff einrichten. Bevorzugt
   `Files.SelectedOperations.Selected` mit Rolle `read` nur auf dieser Datei.
   Administratorkonsens allein gewährt noch keinen Dateizugriff; die explizite Zuweisung
   ist zusätzlich nötig. Eine Datei-Zuweisung unterbricht die Vererbungsstruktur auf
   diesem Element und muss bei der Rechtepflege berücksichtigt werden.
3. Föderierte Identität: Issuer `https://token.actions.githubusercontent.com`, Audience
   `api://AzureADTokenExchange`, Subject
   `repo:SwitzerlandTravelCentre/TextblocsOutlookAddIn:ref:refs/heads/main`.
   Damit erhält nur ein Lauf auf `main` ein Token (ein GitHub-Environment mit
   Branch-Schutz bräuchte Admin-Rechte auf dem Repository und wird nicht verwendet).
   Keine Benutzerkennwörter oder langfristigen Graph-Client-Secrets im Workflow.
4. Repository-Variablen (Settings → Secrets and variables → Actions → Variables):
   - `TEXTBLOCKS_TENANT_ID`: `19af225c-a760-483f-b9b6-f0771d9406da`
   - `TEXTBLOCKS_READER_CLIENT_ID`: ID der neuen Lese-App
   - `TEXTBLOCKS_DRIVE_ID`: tatsächliche Bibliotheks-ID
   - `TEXTBLOCKS_ITEM_ID`: tatsächliche Datei-ID
5. Das Deployment-Token der vorhandenen Static Web App als Repository-Secret
   `AZURE_STATIC_WEB_APPS_API_TOKEN` hinterlegen. Keine neue Website erzeugen.
   Wird das Token in Azure neu erzeugt, muss dieses Secret nachgeführt werden.
6. Die hier vorbereiteten Änderungen im Originalrepository übernehmen.
7. Manuellen Lauf mit der echten SharePoint-Quelle ausführen und Resultat verifizieren.
   Erst danach den wöchentlichen Betrieb als aktiv bestätigen.
8. Produktionsmanifest aus `dist/manifest.xml` prüfen. Repository-Manifest 1.0.4.0
   enthält noch localhost-URLs; der Build der Pipeline ersetzt diese nur im ausgelieferten
   Manifest durch die vorhandene Azure-Adresse. Die zentrale M365-Zuweisung wird dadurch
   nicht automatisch aktualisiert. Bestehende Manifest-ID und Testgruppe erhalten.

## Abnahme und Rückweg

- Vor dem ersten Deployment die letzte funktionierende vollständige Build-Ausgabe
  und das zentral verwendete Manifest sichern; eine JSON-Sicherung allein reicht bei
  gleichzeitiger Änderung des Anwendungscodes nicht.
- Eine vereinbarte Änderung in der SharePoint-Excel-Datei auslösen, dann deren
  Darstellung und Einfügen (Absätze, Fett, Links, Markierungen) im Outlook-Testbenutzer
  nachweisen. Der lokal vorhandene Quellstand darf nicht als SharePoint-Abnahme gelten.
- Zusätzlich einen planmässigen Lauf verifizieren und Fehlerbenachrichtigung prüfen.
- Bei einem Fehler nach Veröffentlichung vorherigen vollständigen Build mit demselben
  SWA-Deploymentweg erneut veröffentlichen und Live-Hash prüfen. Keine Quellarbeitsmappe
  überschreiben; für Quellenfehler SharePoint-Versionierung gezielt verwenden.
- Erst nach Ende-zu-Ende-Abnahme ist das Ticket abschliessbar.

## Lokaler Test

```powershell
python -m pip install -r tools/requirements-refresh.txt
python -m unittest discover -s tools -p 'test_refresh_*.py'
python tools/refresh_textblocks.py --source source/STC_Textblocs_Source.xlsx
npm ci --ignore-scripts --no-audit --no-fund
npm test
npm run build
```

Grundlagen:
- https://learn.microsoft.com/en-us/graph/permissions-selected-overview
- https://learn.microsoft.com/en-us/azure/static-web-apps/build-configuration

## Prüfung der Quelle (08.09.2026)

- Echte Quelle: SharePoint STC Global, `Document Library/Operations/B2C/Global/SourceData/STC_Textblocs_Source.xlsx`.
- Drive-ID: `b!wAdGkIkOiEy2G1cLVAJ6PcCXtLdApetCqdwIRgx7dvtMQNCWg_1KRoLE7jHSwmbl`.
- Item-ID: `01HM5O6M32R2AQW4PYHVDI4S2AKSLV4JKH`.
- Authentifizierter Download über Excel im Browser erfolgreich. Direkter Download mit dem vorhandenen delegierten Azure-CLI-Token lieferte HTTP 401; Metadatenzugriff allein ist kein Nachweis für automatisierten Dateidownload. Deshalb liest der Wochenlauf mit einer eigenen App-only-Identität (siehe Inbetriebnahme).
- Die echte SharePoint-Quelle ergibt 41 befüllte Zeilen, 161 Bausteine, davon 120 mit Formatierungen; alle vier Sprachen vorhanden. Die Repository-Kopie ergibt dagegen 76 Zeilen / 292 Bausteine. Die SharePoint-Datei ist seit 24.09.2026 die verbindliche Quelle; `source/STC_Textblocs_Source.xlsx` im Repository ist nur noch eine lokale Testkopie.
- 7 Fehlerfalltests und 15 Anwendungstests mit der echten SharePoint-Konvertierung erfolgreich, Build erfolgreich.
- Fehlende Root-Startdatei für SWA ergänzt: `public/index.html` entspricht der bisher produktiv ausgelieferten Infoseite (unverändertes Verhalten); Manifest-ID und Taskpane-Pfad bleiben erhalten.

## Veröffentlichte Vorschau

- URL: https://calm-rock-03c40f810-ticket122827.centralus.7.azurestaticapps.net/src/taskpane/taskpane.html
- Azure-Umgebung `ticket122827`, Deployment `19369dc5-a9b2-4f4a-9dff-b31b3e522876`, erfolgreich.
- Vorschau: 161 Bausteine, JSON-SHA256 `5bd6c8c67966f2f4d130897906742012837536bb53dd06a26a9c648e12d48499`; Kandidat und ausgelieferte Datei identisch.
- Produktions-JSON nach dem Probelauf unverändert: 296 Bausteine, SHA256 `31c30d8381f9eae9d7611b7a31dd3a5510b9a618bf81bc4e501596eb969a5451`.
- Oberfläche im Browser sichtbar. Einfügen im echten Outlook noch nicht geprüft.

## Inbetriebnahme 24.09.2026

- Quelle von Clemens Bartlome als final bestätigt (Datei wird nicht mehr verschoben).
- Entra-App `STC Textblocs Weekly Refresh (GitHub Actions)`, Client-ID
  `ef41b352-d798-43ff-8231-9d7918970866`: kein Secret, nur föderierte Identität für
  `ref:refs/heads/main`; Graph-Anwendungsrecht `Files.SelectedOperations.Selected`
  mit Administratorkonsens; Rolle `read` ausschliesslich auf `STC_Textblocs_Source.xlsx`.
- Die Datei-Zuweisung wurde über eine kurzlebige Hilfs-App (Sites.FullControl.All,
  1-h-Secret nur im Speicher) gesetzt; die Hilfs-App wurde direkt danach gelöscht.
  Entzug: Dateiberechtigung der App entfernen oder den Administratorkonsens widerrufen.
- Repository-Variablen und das Secret `AZURE_STATIC_WEB_APPS_API_TOKEN` gesetzt.
- Produktionsstand vor dem ersten Lauf vollständig gesichert (Taskpane, JS/CSS, Icons,
  Infoseite, JSON mit 296 Bausteinen, SHA256 `31c30d83…a5451`); der Rücksprungweg
  (erneutes Hochladen der Sicherung mit demselben SWA-Deploymentweg und Hashprüfung)
  wurde auf der Vorschau-Umgebung erfolgreich getestet. Die Sicherung liegt bei oneICT.
