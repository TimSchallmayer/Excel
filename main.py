"""Physik-Diagramm-Assistent für Excel-Messwerttabellen.

Installation: python -m pip install pandas openpyxl matplotlib
Aufruf: python main.py messwerte.xlsx [--sheet 0] [--x Zeit] [--y Spannung]

Die Erkennung ist heuristisch. Nicht erkannte Spalten lassen sich mit --x/--y
eindeutig festlegen. Die Excel-Datei wird nicht verändert.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from pathlib import Path

try:
	import matplotlib.pyplot as plt
	import numpy as np
	import pandas as pd
except ImportError as exc:
	raise SystemExit(
		"Fehlendes Paket. Installiere es mit: "
		"python -m pip install pandas openpyxl matplotlib"
	) from exc


# Aliasnamen decken gängige deutsche und englische Tabellenüberschriften ab.
QUANTITIES = {
	"zeit": ("Zeit", "s", ("zeit", "time", "dauer")),
	"strecke": ("Strecke", "m", ("strecke", "weg", "ort", "distance")),
	"geschwindigkeit": ("Geschwindigkeit", "m/s", ("geschwindigkeit", "velocity", "speed")),
	"beschleunigung": ("Beschleunigung", "m/s²", ("beschleunigung", "acceleration")),
	"kraft": ("Kraft", "N", ("kraft", "force")),
	"masse": ("Masse", "kg", ("masse", "mass")),
	"dehnung": ("Dehnung", "m", ("dehnung", "verlangerung", "auslenkung")),
	"spannung": ("Spannung", "V", ("spannung", "voltage")),
	"strom": ("Stromstärke", "A", ("strom", "stromstarke", "current")),
	"widerstand": ("Widerstand", "Ω", ("widerstand", "resistance", "ohm")),
	"leistung": ("Leistung", "W", ("leistung", "power")),
	"energie": ("Energie", "J", ("energie", "arbeit", "energy", "work")),
	"temperatur": ("Temperatur", "°C", ("temperatur", "temperature")),
	"druck": ("Druck", "Pa", ("druck", "pressure")),
	"volumen": ("Volumen", "m³", ("volumen", "volume")),
	"frequenz": ("Frequenz", "Hz", ("frequenz", "frequency")),
	"periode": ("Periodendauer", "s", ("periodendauer", "periode", "period")),
	"wellenlange": ("Wellenlänge", "m", ("wellenlange", "wavelength")),
	"winkel": ("Winkel", "°", ("winkel", "angle")),
	"ladung": ("Ladung", "C", ("ladung", "charge")),
}

UNIT_QUANTITY = {
	"s": ("Zeit", "s"), "ms": ("Zeit", "ms"), "min": ("Zeit", "min"),
	"h": ("Zeit", "h"), "m": ("Strecke", "m"), "cm": ("Strecke", "cm"),
	"mm": ("Strecke", "mm"), "km": ("Strecke", "km"), "m/s": ("Geschwindigkeit", "m/s"),
	"km/h": ("Geschwindigkeit", "km/h"), "m/s2": ("Beschleunigung", "m/s²"),
	"n": ("Kraft", "N"), "mn": ("Kraft", "mN"), "kg": ("Masse", "kg"),
	"g": ("Masse", "g"), "v": ("Spannung", "V"), "mv": ("Spannung", "mV"),
	"a": ("Stromstärke", "A"), "ma": ("Stromstärke", "mA"),
	"ω": ("Widerstand", "Ω"), "ohm": ("Widerstand", "Ω"),
	"w": ("Leistung", "W"), "mw": ("Leistung", "mW"),
	"j": ("Energie", "J"), "kj": ("Energie", "kJ"), "pa": ("Druck", "Pa"),
	"kpa": ("Druck", "kPa"), "m3": ("Volumen", "m³"), "ml": ("Volumen", "mL"),
	"hz": ("Frequenz", "Hz"), "khz": ("Frequenz", "kHz"),
	"°c": ("Temperatur", "°C"), "k": ("Temperatur", "K"),
	"°": ("Winkel", "°"), "c": ("Ladung", "C"),
}


def normalize(value: object) -> str:
	text = str(value).strip().lower().replace("ä", "a").replace("ö", "o").replace("ü", "u").replace("ß", "ss")
	return re.sub(r"\s+", " ", text)


def identify_column(column: object) -> tuple[str | None, str, bool]:
	"""Erkennt physikalische Größe, Einheit und mögliche Fehlerwert-Spalte."""
	raw = str(column).strip()
	name = normalize(raw)
	is_error = bool(re.search(r"fehler|unsicherheit|uncertainty|error|delta|±", name))
	match = re.search(r"[\[(]\s*([^\])]+)\s*[\])]", name)
	if not match:
		match = re.search(r"\s+([a-zµω°]+(?:/[a-z0-9²³]+)?)\s*$", name)
	unit = match.group(1).strip().replace("²", "2").replace("³", "3") if match else ""
	if unit.lower() in UNIT_QUANTITY:
		quantity, display_unit = UNIT_QUANTITY[unit.lower()]
		return quantity, display_unit, is_error
	base = normalize(re.sub(r"[\[(].*?[\])]|\s+[a-zµω°]+(?:/[a-z0-9²³]+)?$", "", raw, flags=re.I))
	for entry in QUANTITIES.values():
		if any(re.search(rf"\b{re.escape(alias)}\b", base) for alias in entry[2]):
			return entry[0], entry[1], is_error
	return None, "", is_error


def choose_columns(frame: pd.DataFrame, columns: list[str], x_arg: str | None, y_arg: str | None) -> tuple[str, str]:
	names = [str(c) for c in frame.columns]
	for requested in (x_arg, y_arg):
		if requested and requested not in names:
			raise ValueError(f"Spalte {requested!r} nicht gefunden. Verfügbar: {', '.join(names)}")
	x = x_arg or next((c for c in columns if identify_column(c)[0] == "Zeit"), columns[0])
	if y_arg:
		y = y_arg
	else:
		possible = [c for c in columns if c != x and not identify_column(c)[2]]
		recognized = [c for c in possible if identify_column(c)[0]]
		choices = recognized or possible
		if not choices:
			raise ValueError("Keine zweite numerische Messspalte gefunden.")
		y = choices[0]
		if len(choices) > 1:
			print("Mögliche y-Spalten: " + ", ".join(choices))
			answer = input(f"y-Spalte auswählen (Enter für {y}): ").strip()
			if answer:
				if answer not in choices:
					raise ValueError(f"Ungültige y-Spalte: {answer!r}")
				y = answer
	if x == y or x not in columns or y not in columns:
		raise ValueError("x- und y-Spalte müssen unterschiedliche numerische Spalten sein.")
	return x, y


def linear_fit(x: np.ndarray, y: np.ndarray) -> tuple[float, float, float]:
	slope, intercept = np.polyfit(x, y, 1)
	residual = float(np.sum((y - (slope * x + intercept)) ** 2))
	total = float(np.sum((y - np.mean(y)) ** 2))
	return float(slope), float(intercept), 1 - residual / total if total else float("nan")


def create_plot(frame: pd.DataFrame, x_col: str, y_col: str, output: Path,
				fit: str, zero: str, show: bool) -> None:
	x_name, x_unit, _ = identify_column(x_col)
	y_name, y_unit, _ = identify_column(y_col)
	x_label = f"{x_name or x_col} [{x_unit}]" if x_unit else (x_name or x_col)
	y_label = f"{y_name or y_col} [{y_unit}]" if y_unit else (y_name or y_col)
	data = pd.DataFrame({"x": pd.to_numeric(frame[x_col], errors="coerce"),
						 "y": pd.to_numeric(frame[y_col], errors="coerce")})
	x_error = next((str(c) for c in frame.columns if identify_column(c)[2] and identify_column(c)[0] == x_name), None)
	y_error = next((str(c) for c in frame.columns if identify_column(c)[2] and identify_column(c)[0] == y_name), None)
	if x_error:
		data["xerr"] = pd.to_numeric(frame[x_error], errors="coerce")
	if y_error:
		data["yerr"] = pd.to_numeric(frame[y_error], errors="coerce")
	data = data.replace([np.inf, -np.inf], np.nan).dropna(subset=["x", "y"]).sort_values("x")
	if len(data) < 2:
		raise ValueError("Mindestens zwei vollständige Messwertpaare sind erforderlich.")
	x, y = data.x.to_numpy(float), data.y.to_numpy(float)
	fig, ax = plt.subplots(figsize=(8, 5), constrained_layout=True)
	ax.errorbar(x, y, xerr=data["xerr"] if "xerr" in data else None,
				yerr=data["yerr"] if "yerr" in data else None,
				fmt="o", capsize=3, color="#1769aa", ecolor="#777", label="Messwerte")
	ax.set(xlabel=x_label, ylabel=y_label,
		   title=f"{y_name or y_col} in Abhängigkeit von {x_name or x_col}")
	ax.grid(True, linestyle=":", alpha=.65)
	if fit != "none":
		if fit == "linear":
			slope, intercept, r2 = linear_fit(x, y)
			curve = lambda values: slope * values + intercept
			fit_label = f"Lineare Ausgleichsgerade (R² = {r2:.4f})"
			print(f"Ausgleich: y = {slope:.6g} x + {intercept:.6g}; R² = {r2:.6g}")
		else:
			degree = 2 if fit == "quadratic" else 3
			if len(x) <= degree:
				raise ValueError(f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} Messpunkte.")
			coefficients = np.polyfit(x, y, degree)
			curve = lambda values: np.polyval(coefficients, values)
			fit_label = f"Polynomfit (Grad {degree})"
			print("Polynomkoeffizienten:", coefficients)
		xx = np.linspace(float(x.min()), float(x.max()), 200)
		ax.plot(xx, curve(xx), color="#d1495b", label=fit_label)
		ax.legend()
	if zero == "yes":
		use_zero = True
	elif zero == "no":
		use_zero = False
	else:
		use_zero = x_name != "Temperatur" and y_name != "Temperatur" and x.min() >= 0 and y.min() >= 0
	if use_zero:
		if x.min() >= 0:
			ax.set_xlim(left=0)
		if y.min() >= 0:
			ax.set_ylim(bottom=0)
	output.parent.mkdir(parents=True, exist_ok=True)
	fig.savefig(output, dpi=180, bbox_inches="tight")
	print(f"Diagramm gespeichert: {output.resolve()}")
	if show:
		plt.show()
	plt.close(fig)


def main(argv: list[str] | None = None) -> int:
	parser = argparse.ArgumentParser(description="Erstellt ein Physikdiagramm aus Excel-Messwerten.")
	parser.add_argument("datei", type=Path, help="Excel-Datei (.xlsx/.xls)")
	parser.add_argument("--sheet", default="0", help="Tabellenblattname oder nullbasierter Index (Standard: 0)")
	parser.add_argument("--x", help="Unabhängige Variable (Spaltenüberschrift)")
	parser.add_argument("--y", help="Abhängige Variable (Spaltenüberschrift)")
	parser.add_argument("--output", type=Path, help="Ausgabedatei; Standard: <Dateiname>_diagramm.png")
	parser.add_argument("--fit", choices=("auto", "none", "linear", "quadratic", "cubic"), default="auto")
	parser.add_argument("--zero", choices=("auto", "yes", "no"), default="auto", help="Nullpunkt der Achsen")
	parser.add_argument("--show", action="store_true", help="Diagramm nach dem Speichern anzeigen")
	args = parser.parse_args(argv)
	if not args.datei.is_file():
		parser.error(f"Datei nicht gefunden: {args.datei}")
	sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
	try:
		frame = pd.read_excel(args.datei, sheet_name=sheet)
		if frame.empty:
			raise ValueError("Das Tabellenblatt enthält keine Daten.")
		columns = []
		for column in frame.columns:
			values = pd.to_numeric(frame[column], errors="coerce")
			if values.notna().sum() >= 2 and values.nunique(dropna=True) > 1:
				columns.append(str(column))
		if len(columns) < 2:
			raise ValueError("Weniger als zwei numerische Messspalten gefunden.")
		print("Erkannte Spalten:")
		for column in columns:
			quantity, unit, error = identify_column(column)
			description = quantity or "Größe nicht erkannt"
			if unit:
				description += f" [{unit}]"
			if error:
				description += " (mögliche Fehlerwerte)"
			print(f"  {column}: {description}")
		x_col, y_col = choose_columns(frame, columns, args.x, args.y)
		fit = args.fit
		if fit == "auto":
			x = pd.to_numeric(frame[x_col], errors="coerce").to_numpy(float)
			y = pd.to_numeric(frame[y_col], errors="coerce").to_numpy(float)
			valid = np.isfinite(x) & np.isfinite(y)
			fit = "linear" if valid.sum() >= 3 and len(np.unique(x[valid])) >= 2 and linear_fit(x[valid], y[valid])[2] >= 0.85 else "none"
		output = args.output or args.datei.with_name(f"{args.datei.stem}_diagramm.png")
		create_plot(frame, x_col, y_col, output, fit, args.zero, args.show)
		return 0
	except (ValueError, OSError, KeyError) as exc:
		print(f"Fehler: {exc}", file=sys.stderr)
		return 2


if __name__ == "__main__":
	raise SystemExit(main())
