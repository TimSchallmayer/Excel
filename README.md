# PLVS ULTRA Graphs

Excel-Add-in für KI-gestützte Datenanalyse und bearbeitbare Diagramme.

## Installation aus dem Installer

Lade `dist\PLVS-ULTRA-Graphs-Setup.exe` herunter und starte sie. Der Installer
benötigt Microsoft Excel Desktop und installiert das Add-in samt eigener
Python-Runtime im gewählten Ordner (standardmäßig `%LocalAppData%\PLVS ULTRA
Graphs`). Er registriert nur PLVS ULTRA Graphs in Excel. Schließe Excel vor
Installation, Update und Deinstallation. Das Add-in wird in Excels Add-in-
Manager eingetragen und über einen eigenen Excel-Start-Eintrag geladen;
vorhandene Start-Einträge anderer Add-ins bleiben unverändert. Die xlwings-
Konfiguration in der `.xlam` wird bei der Installation auf den gebündelten
Interpreter und den Installationsordner gesetzt.

Damit Excel die VBA-Bridge ohne zusätzliche manuelle Trust-Center-Schritte
laden kann, wird ausschließlich der gewählte PLVS-Installationsordner als
Excel-Trusted-Location für den aktuellen Benutzer eingetragen. Bei der
Deinstallation werden nur der PLVS-Add-in-Eintrag, sein Excel-Start-Eintrag
und dieser vom Installer angelegte Trusted-Location-Eintrag entfernt.

Die gebündelte Runtime enthält die festgelegten Pakete aus
`requirements-runtime.txt`. Es wird kein separat installiertes Python benötigt.
Für KI-Funktionen muss der Benutzer seine API-Konfiguration weiterhin selbst
einrichten; persönliche Konfigurationsdateien werden nicht in den Installer
aufgenommen oder bei Updates überschrieben.

## Entwicklungsinstallation

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

## Release-Installer bauen

Voraussetzung: Windows, Excel Desktop, eine eingerichtete Entwicklungsumgebung
mit xlwings sowie Inno Setup 6. Der reproduzierbare Build pinnt die
Python-3.12.10-Runtime und Runtime-Abhängigkeiten, erstellt ein frisches
Stage-Verzeichnis und prüft es vor dem Installer-Build:

```powershell
.\installer\build_installer.ps1
```

Der Build schreibt `dist\PLVS ULTRA Graphs.xlam` und
`dist\PLVS-ULTRA-Graphs-Setup.exe`. Die temporäre Runtime-Staging-Struktur wird
nach dem Build entfernt. Enthalten sind nur die `.xlam` mit eingebettetem VBA-
und Ribbon-Inhalt, die für das Add-in benötigten Python-Module, CPython samt
Tcl/Tk sowie die fest gepinnten Pakete aus `requirements-runtime.txt`.

## Verwendung

Daten auswählen und im Ribbon **Diagramm erstellen** klicken. Alternativ:

```powershell
python main.py messwerte.xlsx
```
