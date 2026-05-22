# PureLoX Kapazitätsplaner – PowerPoint Generator

Ein Excel-basiertes Tool zur automatischen Erstellung von PowerPoint-Kapazitätsplänen. Du pflegst Deine Projektdaten direkt in Excel, klickst auf einen Button – und das Tool generiert daraus eine fertige Präsentation im PureLoX-Design.

---

## Inhaltsverzeichnis

1. [Wie das Tool funktioniert](#wie-das-tool-funktioniert)
2. [Voraussetzungen](#voraussetzungen)
3. [Installation](#installation)
4. [Excel-Datei einrichten](#excel-datei-einrichten)
5. [Konfigurationsblatt ausfüllen](#konfigurationsblatt-ausfüllen)
6. [Projektdaten eingeben](#projektdaten-eingeben)
7. [PowerPoint exportieren](#powerpoint-exportieren)
8. [Makros im Überblick](#makros-im-überblick)
9. [Wie die VBA-Seite funktioniert](#wie-die-vba-seite-funktioniert)
10. [Wie die JavaScript-Seite funktioniert](#wie-die-javascript-seite-funktioniert)
11. [Fehlerbehebung](#fehlerbehebung)

---

## Wie das Tool funktioniert

```
Excel (VBA)  →  JSON-Datei  →  Node.js (JavaScript)  →  PPTX-Datei
```

1. Du pflegst Deine Kapazitätsdaten in Excel-Projektblättern (Gantt-Phasen + FTE-Werte).
2. Ein VBA-Makro liest diese Daten aus und schreibt sie als JSON-Konfigurationsdatei auf Deine Festplatte.
3. Das Makro startet automatisch das Node.js-Skript `PLX_generate_pptx.js`.
4. Das Skript liest die JSON-Datei und erzeugt daraus eine fertige `.pptx`-Datei.
5. Excel öffnet die Präsentation automatisch.

---

## Voraussetzungen

| Software | Version | Download |
|---|---|---|
| Microsoft Excel | 2016 oder neuer (Windows) | – |
| Node.js | 18 LTS oder neuer | https://nodejs.org |
| npm | wird mit Node.js mitgeliefert | – |

> **Hinweis:** Das Tool läuft nur unter **Windows**, da VBA Shell-Befehle verwendet werden, die Node.js direkt aufrufen.

---

## Installation

### 1. Node.js installieren

1. Lade Node.js von https://nodejs.org herunter (LTS-Version empfohlen).
2. Führe den Installer aus und bestätige alle Standardoptionen (inkl. „Add to PATH").
3. Öffne die Eingabeaufforderung (`Win + R` → `cmd`) und prüfe die Installation:

```cmd
node --version
npm --version
```

Du solltest eine Versionsnummer sehen (z. B. `v22.0.0`).

### 2. Projektdateien einrichten

Lege einen Ordner an, z. B. `C:\PLX_Tools\`, und kopiere folgende Datei dort hinein:

```
C:\PLX_Tools\
└── PLX_generate_pptx.js
```

### 3. Abhängigkeiten installieren

Öffne die Eingabeaufforderung und wechsle in den Ordner:

```cmd
cd C:\PLX_Tools
npm install pptxgenjs
```

Danach sieht der Ordner so aus:

```
C:\PLX_Tools\
├── PLX_generate_pptx.js
├── package.json
└── node_modules\
    └── pptxgenjs\
        └── ...
```

### 4. VBA-Code in Excel einbinden

1. Erstelle eine neue Excel-Datei (`.xlsm` – mit Makros) oder öffne Deine bestehende.
2. Öffne den VBA-Editor: `Alt + F11`
3. Im Projekt-Explorer: Rechtsklick auf „VBAProject" → **Einfügen** → **Modul**
4. Öffne die Datei `VBA.txt` in einem Texteditor, kopiere den gesamten Inhalt und füge ihn in das neue Modul ein.
5. Speichere die Excel-Datei als `.xlsm`.

---

## Excel-Datei einrichten

Führe das Einrichtungsmakro **einmalig** aus, um alle Blätter und die Struktur zu erstellen:

1. Drücke `Alt + F8`
2. Wähle das Makro `Erstelle_Komplettes_Planungs_Tool`
3. Klicke auf **Ausführen**

Das Makro erstellt automatisch folgende Blätter:

| Blatt | Beschreibung |
|---|---|
| `Konfiguration` | Einstellungen, Pfade, Zeitraum |
| `Legende` | Erklärung der Phasencodes und FTE-Farben |
| `Gesamtuebersicht` | Aggregierte FTE-Übersicht aller Projekte |
| `Projekt 01` – `Projekt 12` | Deine 12 Projektblätter |

> Dieses Makro nur **einmal** ausführen. Bei einem erneuten Aufruf wird die bestehende Struktur nicht überschrieben.

---

## Konfigurationsblatt ausfüllen

Wechsle auf das Blatt **„Konfiguration"** und fülle die orangefarbenen Felder aus:

### Präsentationsinhalte (Zeilen 4–8)

| Feld | Zeile | Beispiel |
|---|---|---|
| Headline (Titel) | 4 | `Portfolio Kapazitätsplanung` |
| Subheadline | 5 | `Übersicht FTE & Projektphasen` |
| Präsentation für | 6 | `Musterkunde GmbH` |
| Datum | 7 | `Mai 2026` |
| Ersteller / Team | 8 | `BOOST – Customer Solutions Team` |

### Ausgabe-Einstellungen (Zeilen 11–14)

| Feld | Zeile | Beispiel |
|---|---|---|
| Ausgabepfad | 11 | `C:\Users\Dein Name\Documents\` |
| Dateiname | 12 | `PLX_Kapazitaetsplanung` (ohne `.pptx`) |
| Node.js Pfad | 13 | `node` (wenn Node.js im PATH) oder `C:\Program Files\nodejs\node.exe` |
| Pfad zu PLX_generate_pptx.js | 14 | `C:\PLX_Tools\PLX_generate_pptx.js` |

### Planungszeitraum (Zeilen 17–18)

| Feld | Zeile | Beispiel |
|---|---|---|
| Startmonat | 17 | `04.2026` (Format: `MM.YYYY`) |
| Anzahl Monate | 18 | `14` (maximal 24) |

Nach dem Speichern der Startmonat/Anzahl-Felder aktualisiert das Tool die Monatsspalten in allen Projektblättern automatisch.

---

## Projektdaten eingeben

Wechsle auf ein Projektblatt (z. B. **„Projekt 01"**).

### Struktur eines Projektblatts

```
Zeile 1:       Projektname
Zeile 2:       Monatsspalten (automatisch generiert)
Zeilen 3–21:   Gantt-Bereich (3 Teilprojekte × 6 Zeilen)
Zeile 23:      "KAPAZITÄTSPLANUNG (FTE)" + Monatsspalten
Zeilen 24–27:  FTE-Werte (3 Teilprojekte + Summe)
```

### Gantt-Phasencodes

Trage in die Gantt-Zellen einen der folgenden Buchstaben ein:

| Code | Phase | Farbe |
|---|---|---|
| `p` | Pflichtenheft | Dunkelblau |
| `k` | Konfiguration | Hellblau |
| `t` | Test | Lila |
| `g` | GoLive | Rot/Pink |
| `?` | Meilenstein | Navy |

**Mehrere Phasen in einem Monat** sind möglich – trage sie kommagetrennt ein, z. B. `k,t`.

Lasse die Zelle **leer**, wenn keine Phase aktiv ist.

### FTE-Werte

Trage in die FTE-Zeilen (Zeilen 24–26) numerische Werte ein:

- `0.5` = halbe Stelle
- `1` = eine Vollzeitstelle
- `1.5`, `2`, usw.

Die Summe in Zeile 27 wird automatisch berechnet.

### FTE-Farbskala (Heatmap)

| Wert | Farbe |
|---|---|
| 0 / leer | Grau |
| > 0 bis 1 | Hellblau |
| > 1 bis 1,5 | Blau |
| > 1,5 bis 2 | Lila |
| > 2 | Rot/Pink |

---

## PowerPoint exportieren

Wenn alle Projektdaten gepflegt sind:

1. Drücke `Alt + F8`
2. Wähle das Makro `Exportiere_Als_JSON_Und_PPTX`
3. Klicke auf **Ausführen**

Das Makro:
1. Liest alle Projektblätter aus
2. Schreibt eine JSON-Konfigurationsdatei (gleicher Ordner wie `.pptx`)
3. Ruft Node.js auf: `node PLX_generate_pptx.js <config.json> <output.pptx>`
4. Öffnet die fertige `.pptx`-Datei automatisch

### Tipp: Export-Button direkt in Excel

Um den Export per Klick auszulösen, kannst Du einen Button einfügen:

1. Reiter **Entwicklertools** → **Einfügen** → **Schaltfläche (Formularsteuerelement)**
2. Zeichne die Schaltfläche auf einem Blatt
3. Weise das Makro `Exportiere_Als_JSON_Und_PPTX` zu

---

## Makros im Überblick

| Makro | Wann aufrufen |
|---|---|
| `Erstelle_Komplettes_Planungs_Tool` | Einmalig zur Ersteinrichtung |
| `Exportiere_Als_JSON_Und_PPTX` | Jedes Mal, wenn eine neue PPTX erstellt werden soll |
| `Aktualisiere_Monatszeile` | Wenn der Startmonat oder die Anzahl der Monate geändert wurde |
| `Berechne_Monate` | Intern aufgerufen – muss nicht manuell ausgeführt werden |

---

## Wie die VBA-Seite funktioniert

Die VBA-Datei (`VBA.txt`) enthält das komplette Excel-Makro-Backend. Die wichtigsten Abläufe:

### Einrichtung (`Erstelle_Komplettes_Planungs_Tool`)

- Erstellt das Konfigurationsblatt mit formatierten Eingabefeldern
- Legt 12 Projektblätter an mit vordefinierten Zeilen für Gantt und FTE
- Wendet bedingte Formatierungen an (Phasencodes → Farben, FTE-Werte → Heatmap)

### Monatsberechnung (`Berechne_Monate` / `Aktualisiere_Monatszeile`)

- Liest Startmonat (`MM.YYYY`) und Anzahl aus dem Konfigurationsblatt
- Berechnet die Monatsnamen dynamisch (z. B. `Apr 26`, `Mai 26`, ...) und behandelt Jahreswechsel automatisch
- Schreibt die neuen Monatsnamen in alle Projektblätter, ohne bestehende Eingaben zu löschen

### JSON-Export (`Exportiere_Als_JSON_Und_PPTX`)

- Iteriert über alle Projektblätter (alles außer `Konfiguration`, `Legende`, `Gesamtuebersicht`)
- Liest pro Projektblatt:
  - Projektname (Zeile 1)
  - 3 Teilprojekte mit je 5 Phasenzeilen (Gantt-Codes)
  - FTE-Werte pro Teilprojekt und Monat
  - FTE-Summe pro Monat
- Baut eine JSON-Struktur auf und schreibt sie als Datei (`<Dateiname>_config.json`)
- Startet Node.js per `Shell`-Befehl und prüft den Exit-Code
- Bei Fehler wird ein Fehlerprotokoll geschrieben und eine Fehlermeldung angezeigt

---

## Wie die JavaScript-Seite funktioniert

Das Skript `PLX_generate_pptx.js` übernimmt die JSON-Datei und baut daraus die Präsentation mit der Bibliothek `pptxgenjs`.

### Aufruf

```cmd
node PLX_generate_pptx.js <pfad-zur-config.json> <pfad-zur-ausgabe.pptx>
```

### Folien-Aufbau

| Folie | Inhalt |
|---|---|
| 1 | Titelfolie (Hintergrundbild, Headline, Kundenname, Datum) |
| 2 | Legende (Phasencodes + FTE-Farbskala) |
| 3+ | Projektfolien (Gantt + FTE-Tabelle pro Projekt) |
| Letzte | Gesamtübersicht: FTE aller Projekte zusammengefasst |

### Projektfolien-Logik

Für jedes Projekt prüft das Skript, ob Gantt und FTE auf **eine Folie** passen:

- **Eine Folie:** Wenn 2 oder weniger Teilprojekte vorhanden sind und die Zeilenhöhe ausreicht
- **Zwei Folien:** Bei 3 Teilprojekten oder wenn die Mindesthöhe unterschritten würde – dann kommt Gantt auf Folie A und FTE auf Folie B

### Gantt-Rendering

- Zeichnet pro Teilprojekt farbige „Pills" (abgerundete Rechtecke) für jede aktive Phase
- Filtert leere Phasenzeilen automatisch heraus
- Unterstützt mehrere Phasen pro Monatszelle

### FTE-Tabelle

- Kopfzeile mit Monatsnamen (Navy-Hintergrund)
- Eine Zeile pro Teilprojekt + eine Summenzeile (dunkel, fett)
- Zellenfarben nach FTE-Heatmap-Schema
- Fette Schrift ab FTE ≥ 1,5

### Corporate Identity

Das Skript enthält alle Grafikassets (Logos, Hintergründe, Footer) als Base64-kodierte Strings, sodass **keine externen Bilddateien** benötigt werden.

---

## Fehlerbehebung

### „Node.js wurde nicht gefunden"

- Prüfe, ob Node.js korrekt installiert ist: `node --version` in der Eingabeaufforderung
- Trage den vollständigen Pfad in das Feld **Node.js Pfad** ein, z. B.:
  `C:\Program Files\nodejs\node.exe`

### „pptxgenjs nicht gefunden" / `Cannot find module 'pptxgenjs'`

- Führe `npm install pptxgenjs` im Ordner von `PLX_generate_pptx.js` aus
- Prüfe, ob der Ordner `node_modules\pptxgenjs` existiert

### „Datei konnte nicht gespeichert werden"

- Stelle sicher, dass der **Ausgabepfad** im Konfigurationsblatt existiert
- Prüfe, ob die PPTX-Datei bereits in PowerPoint geöffnet ist (dann ist sie gesperrt)

### Monatszeilen stimmen nicht

- Führe das Makro `Aktualisiere_Monatszeile` aus, nachdem Du Startmonat oder Anzahl geändert hast

### Makros sind deaktiviert

- Stelle sicher, dass die Excel-Datei als `.xlsm` gespeichert ist
- Gehe zu **Datei → Optionen → Trust Center → Trust Center-Einstellungen → Makroeinstellungen** und wähle „Alle Makros mit Benachrichtigung deaktivieren", dann beim Öffnen der Datei Makros aktivieren

### Fehlerdatei auswerten

Bei Fehlern schreibt das Makro eine Protokolldatei `<Dateiname>_error.log` in den Ausgabeordner – dort findest Du die genaue Fehlermeldung von Node.js.

---

## Projektstruktur

```
PowerPointGenerator/
├── VBA.txt                  # VBA-Makrocode (in Excel einzubinden)
├── PLX_generate_pptx.js     # Node.js-Skript zur PPTX-Generierung
└── README.md                # Diese Dokumentation
```
