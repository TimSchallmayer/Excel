# PLVS ULTRA Graphs

PLVS ULTRA Graphs ist ein flexibles Excel-Tool zur KI-gestützten Analyse allgemeiner Tabellen und zur Erstellung nativer Diagramme. Die KI interpretiert Spalten und Werte, wählt Achsen und Diagrammart und erkennt bei Bedarf auch nichtlineare Verläufe wie exponentielle, logarithmische oder Potenz-Zusammenhänge. Die Ausgabe bleibt ein bearbeitbares Excel-Diagramm in einer normalen Arbeitsmappe – keine Bilddatei.

Die Excel-Oberfläche ist das eigene Ribbon-Register **PLVS ULTRA Graphs**. Für den normalen Einsatz werden keine Arbeitsblätter angelegt.

## Funktionen

- echte Excel-Diagramme mit nativem Scatter-/XY-Chart
- KI-gestützte Analyse allgemeiner tabellarischer Daten
- Erkennung und Auswahl von Excel-Tabellen und zusammenhängenden Datenbereichen
- Auswahl eines Zellbereichs direkt in Excel als manueller Fallback
- KI-Entscheidung über Achsen, Diagrammart und passende lineare oder nichtlineare Trendlinie
- ausführliche Analyseinformationen und KI-Begründung direkt im Ribbon
- klare Fehlermeldung, falls die KI fehlt, nicht erreichbar ist oder keine verlässliche Analyse liefern kann

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

## CLI-Aufruf

```powershell
python main.py messwerte.xlsx
python main.py messwerte.xlsx --sheet "Tabelle1"
python main.py messwerte.xlsx --output diagramm.xlsx
python main.py --test-ai
```

Die KI entscheidet bei jeder Diagrammerstellung über Spalten, Achsen und Fit; lokale Regeln oder manuelle Achsen-/Fit-Overrides werden nicht als Ersatz verwendet. Die CLI unterstützt `.xlsx` und `.xlsm`. Die Originalarbeitsmappe wird nicht überschrieben; als Ausgabe wird eine neue Datei im gleichen Ordner erzeugt, z. B. `messwerte_PLVS_ULTRA_Graphs.xlsx`.

## Beispiel für Tabellenüberschriften

```text
Zeit [s] | Strecke [m] | Strecke Fehler [m]
0        | 0.0         | 0.1
1        | 2.1         | 0.1
2        | 4.3         | 0.2
```

Die Analyse funktioniert am zuverlässigsten mit klaren Überschriften in der ersten benutzten Zeile und aussagekräftigen Daten.

## Excel-Ribbon

Das Add-in bietet einen eigenen Tab **PLVS ULTRA Graphs** mit den Gruppen **Diagramm**, **KI** und **Sonstiges**. Die Gruppe **Analyse** erscheint direkt im kontextabhängigen Tab **Diagrammentwurf** der Excel-Diagrammtools, sobald ein Diagramm ausgewählt ist. **Diagramm erstellen** verwendet eine aktive Excel-Tabelle oder den markierten zusammenhängenden Datenbereich. Nur wenn keine passende Auswahl aktiv ist, wird im Rahmen dieses Befehls eine Datenquelle abgefragt; einen separaten Button zur Quellenauswahl gibt es nicht.

**Diagramm erstellen** analysiert die Quelle mit der konfigurierten KI und erzeugt ein natives Diagramm direkt auf deren Arbeitsblatt. Mehrere erzeugte Diagramme bleiben erhalten; beim Auswählen eines Diagramms zeigt das Menüband dessen eigene Analyse und aktualisiert sie beim Abwählen oder Wechseln. Bei fehlender KI-Konfiguration, Netzwerkfehlern oder zu geringer KI-Konfidenz wird mit einer verständlichen Meldung abgebrochen. Datenzellen werden nicht verändert; Analysen, Einstellungen und Quellbereich werden als ausgeblendete Arbeitsmappennamen gespeichert. Die normale Benutzung legt keine Analyse-, Dashboard-, Daten- oder Hilfsblätter an.

Die Ribbon-Gruppe **Analyse** im Tab **Diagrammentwurf** zeigt die Analyse des ausgewählten Diagramms: unabhängige und abhängige Größe, den erkannten Zusammenhang, Diagrammart und Trendlinie sowie gültige Messpunktzahl, Mittelwert, Median, Wertebereich und Streuung. Beim Erstellen wird das neue Diagramm automatisch ausgewählt; beim Wechseln zwischen Diagrammen aktualisiert sich die Analyse passend. Bei erkennbaren Kurs-/Preisreihen wird zusätzlich die absolute und prozentuale Veränderung zwischen dem kleinsten und größten x-Wert ausgegeben. Die KI-Konfidenz samt Begründung erscheint ebenfalls. Die KI bestimmt die unabhängige und abhängige Größe anhand der Spaltenbedeutung, nicht anhand ihrer Reihenfolge, und die angezeigte Richtung muss mit den gewählten Diagrammachsen übereinstimmen. Zellbereiche werden in der Quellenauswahl lesbar als `Tabelle1: A1:B8` statt mit Ausrufezeichen oder absoluten `$`-Koordinaten angezeigt. **KI-Einstellungen** speichert Endpoint, Modell und optionalen API-Key in `config.json`; die KI-Nutzung kann nicht deaktiviert werden.

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

Der Ribbon-Test ist zusätzlich direkt in Excel zu prüfen: Excel vollständig schließen, neu starten und eine beliebige `Test.xlsx` öffnen. **PLVS ULTRA Graphs** soll **Diagramm erstellen**, **KI-Einstellungen**, **KI testen** und **Hilfe** zeigen. Nach dem Erstellen oder Auswählen eines Diagramms muss die Gruppe **Analyse** direkt im Tab **Diagrammentwurf** erscheinen. Der xlwings-Tab soll nicht sichtbar sein; das xlwings-Add-in muss trotzdem geladen bleiben. Dafür muss keine spezielle `.xlsm` geöffnet sein.

## Was das Projekt aktuell erzeugt

Es werden native Excel-Diagramme erstellt, die in Excel direkt bearbeitet werden können. Trendlinie, Diagrammart, Achsenbeschriftungen und Achsengrenzen folgen der KI-Analyse und werden technisch auf vorhandene Spalten sowie gültige Werte geprüft.
