# PLVS ULTRA Graphs

Excel-Add-in für KI-gestützte Datenanalyse und bearbeitbare Diagramme.

## Installation

Voraussetzung: Windows, Excel Desktop und Python 3.11+.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item config.json.example config.json
```

`config.json` mit `endpoint` und `model` konfigurieren. Add-in bei geschlossenem Excel erstellen:

```powershell
xlwings addin install
.\.venv\Scripts\python.exe excel_bridge.py create-addin
```

`dist\PLVS ULTRA Graphs.xlam` und das xlwings-Add-in in Excel aktivieren.

## Verwendung

Daten auswählen und im Ribbon **Diagramm erstellen** klicken. Alternativ:

```powershell
python main.py messwerte.xlsx
```
