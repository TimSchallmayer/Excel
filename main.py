"""Physik-Diagramm-Assistent für Excel-Messwerttabellen."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import matplotlib.pyplot as plt
    import numpy as np
    import pandas as pd
except ImportError as exc:
    raise SystemExit(
        "Fehlendes Paket. Installiere alle Abhängigkeiten mit: "
        "python -m pip install -r requirements.txt"
    ) from exc

from ai import AIServiceError, analyze_with_ai, test_ai_connection
from config import load_config
from excel import get_numeric_columns, load_workbook
from models import ChartSpecification
from physics import choose_columns, identify_column, linear_fit


def create_plot(frame: pd.DataFrame, spec: ChartSpecification, output: Path,
                fit: str, show: bool) -> None:
    data = pd.DataFrame({
        "x": pd.to_numeric(frame[spec.x_column], errors="coerce"),
        "y": pd.to_numeric(frame[spec.y_column], errors="coerce"),
    })
    if spec.x_error_column:
        data["xerr"] = pd.to_numeric(frame[spec.x_error_column], errors="coerce")
    if spec.y_error_column:
        data["yerr"] = pd.to_numeric(frame[spec.y_error_column], errors="coerce")
    data = data.replace([np.inf, -np.inf], np.nan).dropna(subset=["x", "y"]).sort_values("x")
    for error_name in ("xerr", "yerr"):
        if error_name in data:
            data[error_name] = data[error_name].where(
                np.isfinite(data[error_name]) & (data[error_name] >= 0)
            )
    if len(data) < 2:
        raise ValueError("Mindestens zwei vollständige Messwertpaare sind erforderlich.")
    x, y = data.x.to_numpy(dtype=float), data.y.to_numpy(dtype=float)
    if np.unique(x).size < 2:
        raise ValueError("Die x-Spalte muss mindestens zwei unterschiedliche Messwerte enthalten.")

    fig, ax = plt.subplots(figsize=(8, 5), constrained_layout=True)
    try:
        if spec.chart_type == "scatter":
            line_format = "o" if spec.show_points else "none"
        elif spec.chart_type == "line":
            line_format = "-o" if spec.show_points else "-"
        else:
            line_format = "-o" if spec.show_points else "-"
        ax.errorbar(
            x,
            y,
            xerr=data["xerr"] if spec.x_error_column and "xerr" in data else None,
            yerr=data["yerr"] if spec.y_error_column and "yerr" in data else None,
            fmt=line_format,
            capsize=3,
            color="#1769aa",
            ecolor="#777",
            label="Messwerte",
        )
        ax.set(
            xlabel=spec.x_label,
            ylabel=spec.y_label,
            title=f"{spec.y_label} in Abhängigkeit von {spec.x_label}",
        )
        ax.grid(True, linestyle=":", alpha=.65)
        if fit != "none":
            if fit == "linear":
                slope, intercept, r_squared = linear_fit(x, y)
                curve = lambda values: slope * values + intercept
                fit_label = f"Lineare Ausgleichsgerade (R² = {r_squared:.4f})"
                print(f"Ausgleich: y = {slope:.6g} x + {intercept:.6g}; R² = {r_squared:.6g}")
            else:
                degree = 2 if fit == "quadratic" else 3
                if len(x) <= degree or np.unique(x).size <= degree:
                    raise ValueError(
                        f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} "
                        "Messpunkte mit unterschiedlichen x-Werten."
                    )
                coefficients = np.polyfit(x, y, degree)
                curve = lambda values: np.polyval(coefficients, values)
                fit_label = f"Polynomfit (Grad {degree})"
                print("Polynomkoeffizienten:", coefficients)
            x_values = np.linspace(float(x.min()), float(x.max()), 200)
            ax.plot(x_values, curve(x_values), color="#d1495b", label=fit_label)
            ax.legend()

        use_zero = spec.origin is True or spec.origin == "yes" or (
            spec.origin == "auto"
            and spec.x_quantity != "Temperatur"
            and spec.y_quantity != "Temperatur"
            and x.min() >= 0
            and y.min() >= 0
        )
        if use_zero:
            if x.min() >= 0:
                ax.set_xlim(left=0)
            if y.min() >= 0:
                ax.set_ylim(bottom=0)
        if spec.x_min is not None or spec.x_max is not None:
            ax.set_xlim(left=spec.x_min, right=spec.x_max)
        if spec.y_min is not None or spec.y_max is not None:
            ax.set_ylim(bottom=spec.y_min, top=spec.y_max)

        output.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(output, dpi=180, bbox_inches="tight")
        print(f"Diagramm gespeichert: {output.resolve()}")
        if show:
            plt.show()
    finally:
        plt.close(fig)


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="backslashreplace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description="Erstellt ein Physikdiagramm aus Excel-Messwerten.")
    parser.add_argument("datei", type=Path, nargs="?", help="Excel-Datei (.xlsx/.xlsm/.xls)")
    parser.add_argument("--sheet", default="0", help="Tabellenblattname oder nullbasierter Index (Standard: 0)")
    parser.add_argument("--x", help="Unabhängige Variable (Spaltenüberschrift)")
    parser.add_argument("--y", help="Abhängige Variable (Spaltenüberschrift)")
    parser.add_argument("--output", type=Path, help="Ausgabedatei; Standard: <Dateiname>_diagramm.png")
    parser.add_argument("--fit", choices=("auto", "none", "linear", "quadratic", "cubic"), default="auto")
    parser.add_argument("--zero", choices=("auto", "yes", "no"), default="auto", help="Nullpunkt der Achsen")
    parser.add_argument("--show", action="store_true", help="Diagramm nach dem Speichern anzeigen")
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
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    try:
        frame = load_workbook(args.datei, sheet)
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
        output = args.output or args.datei.with_name(f"{args.datei.stem}_diagramm.png")
        create_plot(frame, spec, output, fit, args.show)
        return 0
    except (AIServiceError, FileNotFoundError, ValueError, OSError, KeyError, ImportError) as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
