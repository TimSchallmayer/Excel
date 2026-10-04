# Physik-Diagramm-Assistent

Ein Python-Prototyp, der Messwerte aus Excel einliest, lokale Physikregeln anwendet und eine Kopie der Arbeitsmappe mit einem **nativen Excel-XY-Streudiagramm** erstellt. Die Quelldatei bleibt unverändert. Optional kann bei mehrdeutiger Spaltenauswahl ein OpenAI-kompatibler KI-Endpunkt befragt werden.

## Voraussetzungen

- Windows 10 oder neuer
- Python 3.10 oder neuer

Python für Windows ist unter [python.org](https://www.python.org/downloads/) verfügbar. Beim Installer **Add python.exe to PATH** aktivieren und anschließend `python --version` prüfen.

## Einrichtung unter Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Falls PowerShell die Aktivierung blockiert, kann sie für das aktuelle Terminal freigegeben werden:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

In `cmd.exe` lautet der Aktivierungsbefehl `.venv\Scripts\activate.bat`.

## Excel-Diagramm erzeugen

```powershell
python main.py messwerte.xlsx
```

Das erzeugt standardmäßig `messwerte_physik.xlsx`. Die ursprüngliche Datei wird nicht verändert. Falls die Ausgabe schon existiert, wird ein freier Name wie `messwerte_physik_2.xlsx` verwendet.

Spalten können explizit gewählt werden:

```powershell
python main.py messwerte.xlsx --x "Zeit [s]" --y "Strecke [m]"
```

Weitere Optionen:

```powershell
python main.py messwerte.xlsx --sheet Messreihe --fit quadratic
python main.py messwerte.xlsx --output Ergebnisse.xlsx --fit linear
```

Unterstützt werden `.xlsx` und `.xlsm`. Bei `.xlsm` wird die Arbeitsmappe mit `keep_vba=True` gespeichert, damit VBA-Inhalte nach Möglichkeit erhalten bleiben. Openpyxl kann digitale VBA-Signaturen nicht erhalten. Alte `.xls`-Dateien werden nicht unterstützt.

Das Diagramm ist ein Excel-XY-Streudiagramm auf dem ursprünglichen Datenblatt, rechts neben den Daten. Ein verborgenes Hilfsblatt enthält bereinigte Diagrammwerte. Excel zeigt darin Messpunkte, beschriftete Achsen, sinnvolle Grenzen und bei passender Auswahl eine native lineare, quadratische oder kubische Trendlinie. Wenn Fehlerwerte erkannt werden, wird eine native benutzerdefinierte Fehlerbalkenreihe angelegt. Openpyxl bildet pro Datenreihe derzeit nur eine Fehlerbalkenrichtung ab; wenn x- und y-Unsicherheiten gleichzeitig vorhanden sind, wird die y-Richtung bevorzugt. Es findet keine PNG-Erzeugung statt.

## Lokale Analyse und KI

Lokale Physikregeln werden zuerst angewendet. Bei eindeutigen Fällen braucht `--ai` keine Anfrage; bei mehrdeutiger automatischer Auswahl kann die KI die Diagrammspezifikation liefern:

```powershell
python main.py messwerte.xlsx --ai
python main.py --test-ai
```

Kopiere `config.json.example` nach `config.json` und trage Endpoint und Modell des OpenAI-kompatiblen Dienstes ein:

```powershell
Copy-Item config.json.example config.json
```

```json
{
  "endpoint": "https://your-openai-compatible-server/v1/chat/completions",
  "api_key": "",
  "model": "your-model-name"
}
```

`api_key` ist optional, sofern der Dienst keinen Schlüssel verlangt. Ein gesetzter Schlüssel wird nur im Authorization-Header gesendet. `config.json` ist in `.gitignore` eingetragen. Echte Schlüssel gehören nicht in Python-Code, README, Beispieldateien oder Git. Bei versehentlicher Veröffentlichung den Schlüssel beim Anbieter widerrufen. KI-Antworten werden streng als JSON geprüft; Python- oder Excel-Code aus Antworten wird niemals ausgeführt.

Die Tabellenanfrage ist kompakt begrenzt und enthält höchstens 40 Spalten sowie 30 verteilte Messzeilen. Nutze `--ai` nur, wenn die Messdaten an den konfigurierten Anbieter übertragen werden dürfen.

## Pakete

- `pandas`: Excel-Messwerte für die Analyse einlesen
- `openpyxl`: Arbeitsmappen erhalten, native Excel-Diagramme erzeugen, `.xlsx` und `.xlsm` speichern
- `numpy`: numerische Spalten, endliche Werte und Fits
- `requests`: OpenAI-kompatible HTTP-Anfragen

Es wird kein Plotting-Paket benötigt.

## Testen

Die Tests verwenden temporäre Excel-Dateien und gemockte KI-Antworten; sie benötigen keine API-Zugangsdaten:

```powershell
python -m unittest discover -s tests -v
```

Die Integrationstests öffnen die gespeicherten Dateien mit openpyxl erneut und prüfen, dass im Arbeitsblatt ein echtes `ScatterChart` liegt.

## Beispiel für Messwerte

Eine Excel-Tabelle mit diesen Spalten kann direkt als `messwerte.xlsx` gespeichert werden:

| Zeit [s] | Strecke [m] | Strecke Fehler [m] |
|---:|---:|---:|
| 0 | 0.0 | 0.1 |
| 1 | 2.1 | 0.1 |
| 2 | 3.9 | 0.2 |
| 3 | 6.2 | 0.2 |

```powershell
python main.py messwerte.xlsx
```

Ergebnis: `messwerte_physik.xlsx` mit den ursprünglichen Messdaten und einem nativen Excel-Diagramm.
