"""PLVS ULTRA Graphs: KI-gesteuerte Diagrammerstellung für Excel-Tabellen."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from ai import AIServiceError, analyze_with_ai, test_ai_connection
from config import load_config
from excel import get_numeric_columns, load_workbook, save_workbook_with_chart


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description="PLVS ULTRA Graphs erstellt KI-gesteuerte Diagramme aus Excel-Tabellen.")
    parser.add_argument("datei", type=Path, nargs="?", help="Excel-Datei (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default="0", help="Tabellenblattname oder nullbasierter Index (Standard: 0)")
    parser.add_argument("--output", type=Path, help="Ausgabedatei; Standard: <Dateiname>_PLVS_ULTRA_Graphs.xlsx")
    parser.add_argument("--test-ai", action="store_true", help="Verbindung zur konfigurierten KI testen")
    args = parser.parse_args(argv)
    if args.test_ai:
        try:
            test_ai_connection(load_config())
            print("KI-Verbindung erfolgreich.")
            return 0
        except (AIServiceError, FileNotFoundError, ValueError, OSError) as exc:
            print(f"KI-Verbindung fehlgeschlagen: {exc}", file=sys.stderr)
            return 2
    if args.datei is None:
        parser.error("Eine Excel-Datei ist erforderlich, außer bei --test-ai.")
    if not args.datei.is_file():
        parser.error(f"Datei nicht gefunden: {args.datei}")
    if args.datei.suffix.lower() not in {".xlsx", ".xlsm"}:
        parser.error("Unterstützt werden .xlsx- und .xlsm-Dateien.")
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    try:
        frame = load_workbook(args.datei, sheet)
        print(f"Excel-Datei geladen: {args.datei}")
        if frame.empty:
            raise ValueError("Das Tabellenblatt enthält keine Daten.")
        columns = get_numeric_columns(frame)
        if len(columns) < 2:
            raise ValueError("Weniger als zwei numerische Messspalten gefunden.")

        spec = analyze_with_ai(frame, columns, load_config())
        output = args.output or args.datei.with_name(
            f"{args.datei.stem}_PLVS_ULTRA_Graphs{args.datei.suffix.lower()}"
        )
        print(f"KI-Diagnose (Konfidenz {spec.confidence:.0%}): {spec.reason}")
        print(f"x-Achse: {spec.x_label}")
        print(f"y-Achse: {spec.y_label}")
        print(f"Diagrammtyp: {spec.chart_type}")
        print(f"Trendlinie: {spec.trendline}")
        output = save_workbook_with_chart(args.datei, frame, sheet, spec, spec.trendline, output)
        print("Diagramm erstellt.")
        print(f"Ausgabedatei: {output.resolve()}")
        return 0
    except (AIServiceError, FileNotFoundError, ValueError, OSError, KeyError) as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
