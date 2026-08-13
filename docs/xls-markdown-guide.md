# Anleitung für Alessia: Textbausteine im Excel mit Markdown pflegen

Diese Anleitung beschreibt, wie die Texte im Excel so gepflegt werden, dass sie nach der JSON-Generierung korrekt in Outlook eingefügt werden.

Die Markdown-Zeichen bleiben im Excel sichtbar. Im eingefügten Outlook-Text werden sie entfernt und als Formatierung umgesetzt.

## Wo wird gearbeitet?

Die Textquelle ist:

```text
source/STC_Textblocs_Source.xlsx
```

Pro Zeile gibt es einen Textbaustein. Gepflegt werden vor allem diese Spalten:

```text
Category
Topic
German
French
English
Italian
```

Die Markdown-Formatierung gehört nur in die Sprachspalten, also `German`, `French`, `English` und `Italian`.

## Grundregeln

1. Kein HTML im Excel verwenden. Also keine Tags wie `<p>`, `<br>`, `<a>`, `<strong>` oder `<span>`.
2. Absatzwechsel mit einer Leerzeile machen.
3. Zeilenumbrüche innerhalb einer Excel-Zelle mit `Alt+Enter` erfassen.
4. Formatierungen als Markdown direkt in den Zelltext schreiben.
5. Nach dem Speichern wird aus dem Excel ein JSON generiert. Das JSON nicht manuell bearbeiten.

## Markdown-Zeichen

| Gewünschtes Resultat | Im Excel schreiben |
| --- | --- |
| Fett | `**Inkludierte Leistungen:**` |
| Gelbe Markierung für Text, der angepasst werden muss | `==Link==` oder `==xx.xx.xxxx==` |
| Fett und gelb markiert | `**==23. April 2026==**` |
| Link mit sichtbarem Text | `[allgemeinen Vertragsbedingungen](https://switzerlandtravelcentre.com/de-ch/agb)` |
| Bullet-Liste | Jede Zeile mit `- ` beginnen |

## Beispiele

### Titel fett markieren

Im Excel:

```text
**Warum bei uns buchen?**
```

Resultat in Outlook: Der Satz `Warum bei uns buchen?` wird fett eingefügt.

### Platzhalter gelb markieren

Im Excel:

```text
Über den folgenden Link können Sie Ihr persönliches Reiseprogramm bequem einsehen: ==Link==
```

Resultat in Outlook: `Link` wird gelb markiert. Das bedeutet: Diese Stelle muss vor dem Versand angepasst werden.

### Fett und Platzhalter kombiniert

Im Excel:

```text
**Sichern Sie sich Ihre Reise, indem Sie Ihre Offerte (==Link==) verbindlich bestätigen.**
```

Resultat in Outlook: Der ganze Satz ist fett. Das Wort `Link` ist zusätzlich gelb markiert.

### Link mit lesbarem Linktext

Im Excel:

```text
Hier finden Sie unsere [allgemeinen Vertragsbedingungen](https://switzerlandtravelcentre.com/de-ch/agb).
```

Resultat in Outlook: Sichtbar ist nur `allgemeinen Vertragsbedingungen`. Die URL ist als Link hinterlegt und wird schwarz dargestellt.

### Bullet-Liste

Im Excel:

```text
**Inkludierte Leistungen:**

- An- und Rückreise ab/bis zu Ihrem Schweizer Wohnort
- Bahnfahrten gemäss Offerte
- Sitzplatzreservation Panoramazüge
- Hotelübernachtungen inkl. Frühstück
```

Wichtig: Für Bullet-Listen immer `- ` am Zeilenanfang verwenden, nicht das Excel-Bullet-Symbol.

## Musterbaustein

```text
Grüezi

Vielen Dank für Ihre Anfrage und Ihr Interesse an unseren Angeboten. Es freut uns sehr, Sie bei der Planung Ihrer Ferien in der Schweiz unterstützen zu dürfen.

Über den folgenden Link können Sie Ihr persönliches Reiseprogramm bequem einsehen: ==Link==

**Inkludierte Leistungen:**

- An- und Rückreise ab/bis zu Ihrem Schweizer Wohnort
- Bahnfahrten gemäss Offerte
- Sitzplatzreservation Panoramazüge
- Hotelübernachtungen inkl. Frühstück

**Nicht inkludiert:** Kurtaxen

**Warum bei uns buchen?**

Als spezialisierter Anbieter für Schweiz-Reisen stellen wir für Sie sorgfältig abgestimmte Reiseprogramme zusammen.

**Sichern Sie sich Ihre Reise, indem Sie Ihre Offerte (==Link==) verbindlich bestätigen.**

Bitte teilen Sie uns bis spätestens **==23. April 2026==** mit, ob das Angebot Ihren Vorstellungen entspricht.

Hier finden Sie ausserdem unsere [allgemeinen Vertragsbedingungen](https://switzerlandtravelcentre.com/de-ch/agb) sowie Antworten auf [häufige Fragen](https://switzerlandtravelcentre.com/de-ch/reiseinfos/haeufige-fragen).

Freundliche Grüsse
```

## Checkliste vor dem Speichern

- Sind alle Überschriften, die fett sein sollen, mit `**...**` markiert?
- Sind alle Stellen, die der User später anpassen muss, mit `==...==` markiert?
- Sind Links als `[sichtbarer Text](https://...)` erfasst?
- Beginnen Bullet-Zeilen mit `- `?
- Gibt es keine HTML-Tags im Text?
- Wurde das Excel gespeichert?