# PLVS ULTRA Graphs

PLVS ULTRA Graphs ist ein kleines Werkzeug zur Analyse physikalischer Messreihen in Excel. Es erkennt lokale Größen und Einheiten, prüft passende Zusammenhänge und erstellt ein natives Excel-XY-Diagramm mit optionalen Trendlinien. Die Ausgabe ist kein PNG, sondern eine bearbeitbare Excel-Arbeitsmappe.

Die Excel-Oberfläche ist ein eigenes Ribbon-Register neben der vorhandenen xlwings-Registerkarte. Das Blatt `PLVS ULTRA Graphs` bleibt als Analyse- und Statusanzeige erhalten.

## Funktionen

- echte Excel-Diagramme mit nativem Scatter-/XY-Chart
- lokale physikalische Erkennung von Größen und Einheiten
- Auswahl von unabhängiger/abhängiger Variable mit manuellem Fallback
- optionale KI-Unterstützung für Mehrdeutigkeiten
- einfache Excel-Startseite mit Grundfunktionen und Hilfsblättern

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

Das Ribbon-Add-in hat die Gruppen **Diagramm**, **KI** und **Analyse** mit den sechs Aktionen Analysieren, Diagramm erstellen, KI testen, KI-Einstellungen, Analyse anzeigen und Hilfe. Die VBA-Callbacks rufen die vorhandenen Python-Funktionen über xlwings `RunPython` auf; analysiert wird die aktive Arbeitsmappe.

### Installation

PowerShell im Projektordner:

```powershell
python -m pip install -r requirements.txt
xlwings addin install
.\.venv\Scripts\python.exe excel_bridge.py create-addin
```

In Excel unter **Datei > Optionen > Add-Ins > Verwalten: Excel-Add-Ins > Gehe zu > Durchsuchen** die Datei `dist\PLVS ULTRA Graphs Ribbon.xlam` hinzufügen und aktivieren. Das xlwings-Add-in muss ebenfalls aktiviert bleiben. Excel anschließend neu starten. Der PLVS-Tab wird per `insertAfterQ` direkt nach dem xlwings-Tab angeordnet.

Das Add-in speichert den Interpreterpfad auf die Projekt-`.venv`. Makros müssen aktiviert sein. Der einmalige VBA-Projektzugriff im Trust Center wird benötigt, wenn das Add-in neu erzeugt wird; für die normale Verwendung der fertigen `.xlam`-Datei nicht.

Das Blatt `PLVS ULTRA Graphs` wird bei Analyse/Diagrammerstellung bei Bedarf als Ergebnis- und Statusanzeige angelegt. Die sechs Aktionen bleiben im Ribbon.

## Tests

```powershell
$env:PHYSIK_EXCEL_COM_TESTS='1'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Der COM-Test benötigt Windows mit installiertem Excel. Er lädt die Add-ins und prüft die aktive Arbeitsmappe sowie das native Diagramm.

## Was das Projekt aktuell erzeugt

Es werden native Excel-Diagramme erstellt, die in Excel direkt bearbeitet werden können. Fehlerbalken und Achsen werden nur dann gesetzt, wenn sie aus den Messdaten stabil abgeleitet werden können. Falls die Auswahl nicht eindeutig ist, bleibt eine manuelle Auswahl möglich.
