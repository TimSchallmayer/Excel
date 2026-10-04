"""PLVS ULTRA Graphs: Standalone-Analyse für physikalische Excel-Messreihen."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import numpy as np
    import pandas as pd
except ImportError as exc:
    raise SystemExit(
        "Fehlendes Paket. Installiere alle Abhängigkeiten mit: "
        "python -m pip install -r requirements.txt"
    ) from exc

from ai import AIServiceError, analyze_with_ai, test_ai_connection
from config import load_config
from excel import get_numeric_columns, load_workbook, output_path_for, save_workbook_with_chart
from models import ChartSpecification
from physics import choose_columns, identify_column, linear_fit


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="backslashreplace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description="PLVS ULTRA Graphs erstellt native Excel-Physikdiagramme.")
    parser.add_argument("datei", type=Path, nargs="?", help="Excel-Datei (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default="0", help="Tabellenblattname oder nullbasierter Index (Standard: 0)")
    parser.add_argument("--x", help="Unabhängige Variable (Spaltenüberschrift)")
    parser.add_argument("--y", help="Abhängige Variable (Spaltenüberschrift)")
    parser.add_argument("--output", type=Path, help="Ausgabedatei; Standard: <Dateiname>_PLVS_ULTRA_Graphs.xlsx")
    parser.add_argument("--fit", choices=("auto", "none", "linear", "quadratic", "cubic"), default="auto")
    parser.add_argument("--zero", choices=("auto", "yes", "no"), default="auto", help="Nullpunkt der Achsen")
    parser.add_argument("--ai", action="store_true", help="Bei mehrdeutiger Spaltenauswahl KI-Diagnose verwenden")
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
        print("Erkannte Spalten:")
        for column in columns:
            quantity, unit, is_error = identify_column(column)
            description = quantity or "Größe nicht erkannt"
            if unit:
                description += f" [{unit}]"
            if is_error:
                description += " (mögliche Fehlerwerte)"
            print(f"  {column}: {description}")

        spec = None
        ai_used = False
        if args.ai and args.x is None and args.y is None:
            measurement_columns = [column for column in columns if not identify_column(column)[2]]
            if len(measurement_columns) == 2:
                time_columns = [
                    column for column in measurement_columns
                    if identify_column(column)[0] == "Zeit"
                ]
                if len(time_columns) == 1:
                    local_x = time_columns[0]
                    local_y = next(column for column in measurement_columns if column != local_x)
                    if identify_column(local_y)[0] is not None:
                        x_column, y_column = local_x, local_y
                    else:
                        x_column = y_column = ""
                else:
                    x_column = y_column = ""
            else:
                x_column = y_column = ""
            if not x_column or not y_column:
                config = load_config()
                spec = analyze_with_ai(frame, columns, config)
                ai_used = True
                print(f"KI-Diagnose (Konfidenz {spec.confidence:.0%}): {spec.reason}")
        if spec is None:
            x_column, y_column = choose_columns(frame, columns, args.x, args.y)
        else:
            x_column, y_column = spec.x_column, spec.y_column
        x_quantity, x_unit, _ = identify_column(x_column)
        y_quantity, y_unit, _ = identify_column(y_column)
        x_label = f"{x_quantity or x_column} [{x_unit}]" if x_unit else (x_quantity or x_column)
        y_label = f"{y_quantity or y_column} [{y_unit}]" if y_unit else (y_quantity or y_column)
        x_error = next((
            str(column) for column in frame.columns
            if identify_column(column)[2] and identify_column(column)[0] == x_quantity
        ), None)
        y_error = next((
            str(column) for column in frame.columns
            if identify_column(column)[2] and identify_column(column)[0] == y_quantity
        ), None)
        if spec is None:
            spec = ChartSpecification(
                x_column=x_column,
                y_column=y_column,
                x_label=x_label,
                y_label=y_label,
                x_unit=x_unit,
                y_unit=y_unit,
                origin=args.zero,
                x_error_column=x_error,
                y_error_column=y_error,
                x_quantity=x_quantity,
                y_quantity=y_quantity,
            )
        elif args.zero != "auto":
            spec.origin = args.zero

        fit = args.fit
        if ai_used and fit == "auto":
            fit = spec.trendline
        elif fit == "auto":
            x_values = pd.to_numeric(frame[x_column], errors="coerce").to_numpy(dtype=float)
            y_values = pd.to_numeric(frame[y_column], errors="coerce").to_numpy(dtype=float)
            valid = np.isfinite(x_values) & np.isfinite(y_values)
            fit = "none"
            if valid.sum() >= 3 and np.unique(x_values[valid]).size >= 2:
                if linear_fit(x_values[valid], y_values[valid])[2] >= 0.85:
                    fit = "linear"
        spec.trendline = fit
        output = args.output or args.datei.with_name(
            f"{args.datei.stem}_PLVS_ULTRA_Graphs{args.datei.suffix.lower()}"
        )
        print(f"x-Achse: {spec.x_label}")
        print(f"y-Achse: {spec.y_label}")
        print("Diagrammtyp: XY-Streudiagramm")
        print(f"Trendlinie: {fit}")
        output = save_workbook_with_chart(args.datei, frame, sheet, spec, fit, output)
        print("Diagramm erstellt.")
        print(f"Ausgabedatei: {output.resolve()}")
        return 0
    except (AIServiceError, FileNotFoundError, ValueError, OSError, KeyError, ImportError) as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
