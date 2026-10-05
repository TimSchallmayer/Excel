# PLVS ULTRA Graphs

PLVS ULTRA Graphs ist ein kleines Werkzeug zur Analyse physikalischer Messreihen in Excel. Es erkennt lokale Größen und Einheiten, prüft passende Zusammenhänge und erstellt ein natives Excel-XY-Diagramm mit optionalen Trendlinien. Die Ausgabe ist kein PNG, sondern eine bearbeitbare Excel-Arbeitsmappe.

Die Excel-Oberfläche ist das eigene Ribbon-Register **PLVS ULTRA Graphs**. Analyse und Datenquellenauswahl werden im Menüband angezeigt; für die normale Benutzung werden keine Arbeitsblätter angelegt.

## Funktionen

- echte Excel-Diagramme mit nativem Scatter-/XY-Chart
- lokale physikalische Erkennung von Größen und Einheiten
- Erkennung und Auswahl von Excel-Tabellen und zusammenhängenden Datenbereichen
- Auswahl eines Zellbereichs direkt in Excel als manueller Fallback
- optionale KI-Unterstützung für Mehrdeutigkeiten
- kompakte Analyseinformationen direkt im Ribbon

## Unterstützte Umgebung

- Windows 10/11
- Excel 2024 oder Microsoft 365 Desktop für Windows
- Python 3.11 oder neuer

## Installation

Im Projektordner:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Falls PowerShell die Ausführung blockiert:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

## KI-Konfiguration

```powershell
Copy-Item config.json.example config.json
```

Beispiel:

```json
{
  "endpoint": "https://your-openai-compatible-server/v1/chat/completions",
  "api_key": "",
  "model": "your-model-name"
}
```

`endpoint` ist die OpenAI-kompatible Chat-Completions-URL, `model` der benötigte Modellname und `api_key` optional. Die Datei `config.json` bleibt lokal und wird nicht in die Arbeitsmappe geschrieben.

## Aktueller CLI-Aufruf

```powershell
python main.py messwerte.xlsx
python main.py messwerte.xlsx --x "Zeit [s]" --y "Strecke [m]"
python main.py messwerte.xlsx --ai
python main.py --test-ai
```

Die CLI unterstützt `.xlsx` und `.xlsm`. Die Originalarbeitsmappe wird nicht überschrieben; als Ausgabe wird eine neue Datei im gleichen Ordner erzeugt, z. B. `messwerte_PLVS_ULTRA_Graphs.xlsx`.

## Beispiel für Tabellenüberschriften

```text
Zeit [s] | Strecke [m] | Strecke Fehler [m]
0        | 0.0         | 0.1
1        | 2.1         | 0.1
2        | 4.3         | 0.2
```

Die Analyse funktioniert am zuverlässigsten mit klaren Überschriften in der ersten benutzten Zeile.

## Excel-Ribbon

Das Add-in bietet einen eigenen Tab **PLVS ULTRA Graphs** mit den Gruppen **Diagramm**, **KI**, **Analyse** und **Sonstiges**. **Datenquelle auswählen** erkennt Excel-Tabellen sowie zusammenhängende Messwertbereiche in sichtbaren Blättern. Gibt es genau eine geeignete Quelle, wird sie automatisch gewählt; bei mehreren Quellen erscheint eine Auswahl. Über **Manuell auswählen** kann ein Zellbereich direkt in Excel markiert werden.

**Diagramm erstellen** analysiert die ausgewählte Quelle automatisch und erzeugt ein natives Diagramm direkt auf deren Arbeitsblatt. Ein erneuter Aufruf ersetzt das zuvor von PLVS erzeugte Diagramm. Datenzellen werden nicht verändert; Analyse, Einstellungen und Quellbereich werden als ausgeblendete Arbeitsmappennamen gespeichert. Die normale Benutzung legt keine Analyse-, Dashboard-, Daten- oder Hilfsblätter an.

Die Ribbon-Gruppe **Analyse** zeigt nach dem Erstellen des Diagramms kompakt x-/y-Größe samt Einheit, Diagrammtyp und Fit sowie Messpunktzahl und Mittelwert. Confidence wird angezeigt, wenn sie in der vorhandenen Analyse vorliegt. Die Anzeige wird nach Auswahl und Diagrammerstellung automatisch aktualisiert. **KI-Einstellungen** fragt Endpoint, Modell und optionale KI-Nutzung über Excel-Dialoge ab; Zugangsdaten werden ausschließlich lokal in `config.json` gespeichert.

### Installation

PowerShell im Projektordner:

```powershell
python -m pip install -r requirements.txt
xlwings addin install
.\.venv\Scripts\python.exe excel_bridge.py create-addin
```

Der Build erzeugt bzw. aktualisiert immer dieselbe Datei `dist\PLVS ULTRA Graphs.xlam`. Excel muss beim Neuerzeugen geschlossen sein.

In Excel unter **Datei > Optionen > Add-Ins > Verwalten: Excel-Add-Ins > Gehe zu** alte PLVS-Einträge deaktivieren und gegebenenfalls entfernen. Danach über **Durchsuchen** genau `dist\PLVS ULTRA Graphs.xlam` auswählen und aktivieren. Das xlwings-Add-in muss geladen bleiben, da dessen `RunPython`-Funktion weiterhin benötigt wird. Falls eine ältere Kopie unter `XLSTART` liegt und den Tab **xlwings** noch anzeigt, synchronisiere sie mit der aktuellen xlwings-Installation:

```powershell
xlwings addin install
```

Die aktuelle Add-in-Definition blendet nur den xlwings-Ribbon-Tab mit `visible="false"` aus; das Add-in und `RunPython` bleiben erhalten. Excel vollständig schließen und neu starten, damit die Ribbon-Definition neu geladen wird.

Zur Versionskontrolle muss der ausgewählte Eintrag auf die gerade erzeugte Datei `dist\PLVS ULTRA Graphs.xlam` zeigen. In der Registrierung dürfen keine alten PLVS-Add-in-Pfade unter `HKCU\Software\Microsoft\Office\16.0\Excel\Options` (`OPEN`, `OPEN1`, `OPEN2` usw.) mehr aktiviert sein.

Das Add-in arbeitet immer mit der aktuell aufrufenden Arbeitsmappe (`xw.Book.caller()`), nicht mit einer festen Datei oder einer zusätzlichen Arbeitsmappe. Es verändert die ausgewählten Messdaten nicht und erzeugt keine zusätzlichen Arbeitsblätter. Makros müssen aktiviert sein. Der einmalige VBA-Projektzugriff im Trust Center wird benötigt, wenn das Add-in neu erzeugt wird; für die normale Verwendung der fertigen `.xlam`-Datei nicht.

## Tests

```powershell
$env:PHYSIK_EXCEL_COM_TESTS='1'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Der COM-Test benötigt Windows mit installiertem Excel. Er lädt die Add-ins und prüft die aktive Arbeitsmappe sowie das native Diagramm.

Der Ribbon-Test ist zusätzlich direkt in Excel zu prüfen: Excel vollständig schließen, neu starten, eine beliebige `Test.xlsx` öffnen und kontrollieren, dass **PLVS ULTRA Graphs** mit **Diagramm erstellen**, **Datenquelle auswählen**, **KI-Einstellungen**, **KI testen**, **Analyse** und **Hilfe** erscheint. Der xlwings-Tab soll nicht sichtbar sein; das xlwings-Add-in muss trotzdem geladen bleiben. Dafür muss keine spezielle `.xlsm` geöffnet sein.

## Was das Projekt aktuell erzeugt

Es werden native Excel-Diagramme erstellt, die in Excel direkt bearbeitet werden können. Fehlerbalken und Achsen werden nur dann gesetzt, wenn sie aus den Messdaten stabil abgeleitet werden können. Falls die Auswahl nicht eindeutig ist, bleibt eine manuelle Auswahl möglich.
