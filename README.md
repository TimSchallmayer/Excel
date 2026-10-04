# Physik-Diagramm-Assistent

Ein Windows-/Excel-Assistent für physikalische Messreihen. Er erkennt Spalten lokal, lässt bei Unsicherheit auswählen oder optional eine KI konsultieren und erstellt ein bearbeitbares, natives Excel-XY-Diagramm. Es wird kein PNG erzeugt.

## Unterstützte Umgebung

- Windows 10/11
- Excel 2024 oder Microsoft 365 Desktop für Windows
- Python 3.11 oder neuer (xlwings 0.37 benötigt mindestens Python 3.11)

## Installation

Python für Windows gibt es unter [python.org](https://www.python.org/downloads/). Beim Installieren **Add python.exe to PATH** aktivieren. Im PowerShell-Terminal im Projektordner:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Falls PowerShell die Aktivierung blockiert, kann sie nur für das aktuelle Terminal erlaubt werden:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

## Excel-Ribbon installieren

Das Add-in wird mit dem offiziellen xlwings-Custom-Add-in-Generator vorbereitet und als `dist\PhysikAssistent.xlam` gebaut. Excel muss geschlossen sein. Excel benötigt einmalig Zugriff auf VBA-Projekte, damit der Build die sechs Ribbon-Callbacks in das xlwings-Add-in einfügen darf:

1. Öffne Excel und wähle **Datei > Optionen > Trust Center > Einstellungen für das Trust Center > Makroeinstellungen**.
2. Aktiviere **Zugriff auf das VBA-Projektobjektmodell vertrauen**. Das ist eine sicherheitsrelevante Entwickleroption; aktiviere sie nur, wenn du diesem lokalen Projekt vertraust.
3. Schließe Excel vollständig.
4. Im aktivierten Projekt-Environment ausführen:

```powershell
python ribbon\build_addin.py --install
```

Das Script verwendet `xlwings quickstart --addin --ribbon`, bindet den Python-Interpreter aus `.venv` ein und installiert das fertige Add-in über `xlwings addin install --file`. Danach Excel neu starten. Die Registerkarte **Physik-Assistent** wird hinter **Ansicht** eingefügt.

Der Add-in-Build konnte in der aktuellen Entwicklungsumgebung noch nicht abgeschlossen werden: Excel blockiert dort den VBA-Projektzugriff aktuell mit „Der programmatische Zugriff auf das Visual Basic-Projekt ist nicht sicher“. Der Build stoppt deshalb absichtlich und lässt kein unvollständiges `.xlam` zurück. Nach der oben beschriebenen Trust-Center-Freigabe kann derselbe Build-Befehl erneut ausgeführt werden. Diese Einstellung ändert das Build-Script nicht automatisch.

## Excel verwenden

Im Ribbon stehen diese funktionsgebundenen Aktionen bereit:

- **Diagramm erstellen** analysiert das Datenblatt und erzeugt ein natives Excel-XY-Diagramm.
- **Tabelle analysieren** bestimmt Größen und Beziehungen und zeigt die Auswertung im Blatt `Physik-Analyse`.
- **KI testen** sendet eine minimale Chat-Completions-Anfrage an den konfigurierten Endpoint.
- **KI-Einstellungen** öffnet die Excel-Seite für Endpoint und Modell. Der API-Key bleibt in `config.json`.
- **Analyse anzeigen** zeigt die aktuelle Spezifikation und lokale Kandidaten.
- **Einstellungen** öffnet das Blatt `Physik-Assistent` mit Auswahlfeldern für Blatt, x/y, Diagramm, Trendlinie, Ursprung, KI und Ausgabe.

Der Standard-Ausgabemodus ist **Kopie (Standard)**: Excel speichert unter einem freien Namen wie `messwerte_physik.xlsx` und lässt die Quelldatei auf dem Datenträger unangetastet. **Aktuelle Datei** muss bewusst im Einstellungsblatt gewählt werden.

Für die verlässlichste Analyse sollten Überschriften in der ersten benutzten Tabellenzeile stehen, zum Beispiel:

| Zeit [s] | Strecke [m] | Strecke Fehler [m] |
|---:|---:|---:|
| 0 | 0.0 | 0.1 |
| 1 | 2.1 | 0.1 |
| 2 | 4.3 | 0.2 |
| 3 | 6.2 | 0.2 |

Auch der Kommandozeilenweg bleibt verfügbar:

```powershell
python main.py messwerte.xlsx
python main.py messwerte.xlsx --x "Zeit [s]" --y "Strecke [m]"
python main.py messwerte.xlsx --ai
```

Die CLI unterstützt `.xlsx` und `.xlsm`; die Originaldatei wird nicht überschrieben. Für `.xlsm` verwendet openpyxl `keep_vba=True`; VBA-Signaturen können nicht erhalten werden.

## Physik und Diagramme

Die lokale Datenbank deckt unter anderem Mechanik, Elektrizität, Magnetismus, Wärmelehre, Schwingungen, Wellen, Optik und Radioaktivität ab. Beziehungshinweise enthalten beispielsweise `s(t)`, `v(t)`, `a(t)`, `F(s)`, `U(I) / I(U)`, `p(V)`, `V(T)`, `λ(f)` und `E(f)`. Bei gleich plausiblen Richtungen, insbesondere U/I, wird keine Richtung blind gewählt: die Excel-Seite fragt nach x und y.

Excel erhält ein natives Diagrammobjekt auf dem Datenblatt. Es kann in Excel angeklickt und bearbeitet werden. Messpunkte sind Standard; lineare, quadratische und kubische native Excel-Trendlinien werden angeboten. Der `.xlsx`-CLI-Pfad kann benutzerdefinierte y-Fehlerbalken mit openpyxl schreiben. Excel 2024 weist den getesteten `Series.ErrorBar`-COM-Aufruf im Ribbon-Pfad auch für Standardwerte zurück; dieser Ablauf zeigt deshalb eine Warnung und erstellt das Diagramm ohne Fehlerbalken, statt Fehler in die Arbeitsmappe zu schreiben. Bereinigte Chartwerte werden auf einem ausgeblendeten Hilfsblatt gespeichert.

## KI-Konfiguration

Erstelle die lokale Datei:

```powershell
Copy-Item config.json.example config.json
```

Beispiel mit Platzhaltern:

```json
{
  "endpoint": "https://your-openai-compatible-server/v1/chat/completions",
  "api_key": "",
  "model": "your-model-name"
}
```

`endpoint` ist die vollständige OpenAI-kompatible Chat-Completions-URL, `model` der vom Server verlangte Modellname und `api_key` optional, sofern der Server keinen Schlüssel verlangt. OpenRouter, Ollama und andere kompatible Server können verwendet werden. Für Ollama ist je nach Installation üblicherweise ein lokaler `/v1/chat/completions`-Endpoint einzutragen.

`config.json` ist in `.gitignore` eingetragen. Echte Schlüssel gehören weder in Quellcode noch in Beispieldateien oder Git. Das Ribbon speichert den Schlüssel nicht im Workbook; es liest ihn aus der lokalen Datei. Bei versehentlicher Veröffentlichung den Schlüssel beim Anbieter widerrufen. KI-Daten werden auf höchstens 40 Spalten und 30 verteilte Zeilen begrenzt. KI-Ausgaben werden als JSON validiert; kein zurückgegebener Code wird ausgeführt.

Ohne `--ai` beziehungsweise ohne die Excel-Einstellung **KI bei Mehrdeutigkeit = Ja** arbeitet die Analyse lokal. KI-Aktionen übertragen Messwerte an den eingestellten Dienst; verwende sie nur, wenn das für die Daten zulässig ist.

## Pakete

- `pandas`: Tabellenanalyse
- `numpy`: numerische Daten und Fits
- `openpyxl`: Dateibasierte Excel-Ein-/Ausgabe und Diagrammtests
- `requests`: OpenAI-kompatible HTTP-Anfragen
- `xlwings`: Excel-COM, aktives Workbook und Custom-Ribbon-Add-in
- `pywin32`: Windows-COM-Abhängigkeit, wird von xlwings unter Windows installiert

`matplotlib` und `xlrd` werden nicht benötigt.

## Tests und Entwicklung

Die Offline-Tests verwenden temporäre Excel-Dateien und gemockte KI-Antworten:

```powershell
python -m unittest discover -s tests -v
```

Die Chart-Integrationstests öffnen gespeicherte Dateien erneut und prüfen native `ScatterChart`-Objekte. Die xlwings-COM-Schicht wurde zusätzlich lokal mit Excel 2024 gegen eine temporäre Arbeitsmappe geprüft. Für Python-Aufrufe aus dem Ribbon muss Excel installiert sein; für den einmaligen `.xlam`-Build muss der Trust-Center-Zugriff auf das VBA-Projekt erlaubt sein.

Der optionale End-to-End-Test mit einer installierten Windows-Excel-Version lässt sich separat aktivieren:

```powershell
$env:PHYSIK_EXCEL_COM_TESTS = "1"
python -m unittest tests.test_excel_addin_com -v
Remove-Item Env:PHYSIK_EXCEL_COM_TESTS
```

Relevante Dateien sind `excel_addin.py` (aktive Excel-Integration), `ribbon/customUI.xml` (Ribbon), `ribbon/RibbonPhysikAssistent.bas` (Excel-Callbacks) und `ribbon/build_addin.py` (Add-in-Build/Installation). Der CLI-Start bleibt `main.py`.

## Fehlerbehebung

- **Ribbon fehlt:** Excel komplett beenden und neu starten; unter **Datei > Optionen > Add-Ins** kontrollieren, dass `PhysikAssistent.xlam` aktiv ist.
- **Build meldet VBA-Projektzugriff:** Trust-Center-Freigabe wie oben setzen und Build erneut ausführen.
- **Python-Modul nicht gefunden:** Add-in mit `python ribbon\build_addin.py --install` aus der Projekt-`.venv` neu bauen; der Interpreterpfad wird dabei gesetzt.
- **Keine Messspalten:** Prüfen, dass die erste benutzte Tabellenzeile Überschriften hat und mindestens zwei Messspalten numerisch sind.
- **KI-Verbindung scheitert:** Endpoint, Modell und gegebenenfalls API-Key prüfen; `config.json` liegt im Projektordner.
- **Ausgabeort:** Standard ist ein neuer Name neben der Arbeitsmappe; ein bestehender Dateiname wird nicht überschrieben.
