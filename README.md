# PLVS ULTRA Graphs

PLVS ULTRA Graphs ist ein kleines Werkzeug zur Analyse physikalischer Messreihen in Excel. Es erkennt lokale Größen und Einheiten, prüft passende Zusammenhänge und erstellt ein natives Excel-XY-Diagramm mit optionalen Trendlinien. Die Ausgabe ist kein PNG, sondern eine bearbeitbare Excel-Arbeitsmappe.

Es gibt kein eigenes Ribbon. Die Oberfläche liegt direkt im Excel-Arbeitsblatt `PLVS ULTRA Graphs`.

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

## Excel-Integration

Die Oberfläche ist direkt im Excel-Arbeitsblatt `PLVS ULTRA Graphs` eingebettet. Es gibt dort die einfachen Aktionen:

- ANALYSIEREN
- DIAGRAMM ERSTELLEN
- KI-EINSTELLUNGEN
- ANALYSE
- HILFE

Die lokale Verbindung zu Excel läuft über `xlwings`. Die verwendete Arbeitsmappe und das aktive Blatt werden über `xw.Book.caller()` bzw. das aktive Sheet des Aufrufers erkannt.

### Excel-Setup

1. Python-Abhängigkeiten installieren:

```powershell
python -m pip install -r requirements.txt
```

2. In Excel über `xlwings` eine makrofähige Arbeitsmappe öffnen oder über die Python-Bridge-Funktionen eine generierte `.xlsm`-Datei verwenden.

3. Das Dashboardblatt `PLVS ULTRA Graphs` wird automatisch ergänzt, falls es noch fehlt.

4. Auf die Buttons klicken, um Analyse, Diagrammerstellung und KI-Einstellungen direkt im geöffneten Workbook zu starten.

## Tests

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Der Fokus liegt auf der Diagrammproduktion, der Physik-/Einheiten-Erkennung und der validierten Chart-Spezifikation.

## Was das Projekt aktuell erzeugt

Es werden native Excel-Diagramme erstellt, die in Excel direkt bearbeitet werden können. Fehlerbalken und Achsen werden nur dann gesetzt, wenn sie aus den Messdaten stabil abgeleitet werden können. Falls die Auswahl nicht eindeutig ist, bleibt eine manuelle Auswahl möglich.
