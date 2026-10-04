"""Lokale, nachvollziehbare Regeln zur Erkennung physikalischer Größen."""

from __future__ import annotations

import re

import numpy as np
import pandas as pd


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
	text = str(value).strip().lower()
	text = text.translate(str.maketrans({"ä": "a", "ö": "o", "ü": "u", "ß": "ss", "Ω": "ω"}))
	return re.sub(r"\s+", " ", text)


def identify_column(column: object) -> tuple[str | None, str, bool]:
	"""Gibt physikalische Größe, Einheit und Fehlerwert-Indikator zurück."""
	raw = str(column).strip()
	name = normalize(raw)
	is_error = bool(re.search(r"fehler|unsicherheit|uncertainty|error|delta|±|Δ", name, re.IGNORECASE))
	match = re.search(r"[\[(]\s*([^\])]+)\s*[\])]", raw)
	if not match:
		match = re.search(r"\s+([a-zµω°]+(?:/[a-z0-9²³]+)?)\s*$", name, re.IGNORECASE)
	unit = match.group(1).strip().replace("²", "2").replace("³", "3") if match else ""
	if unit.lower() in UNIT_QUANTITY:
		quantity, display_unit = UNIT_QUANTITY[unit.lower()]
		return quantity, display_unit, is_error
	base = normalize(re.sub(r"[\[(].*?[\])]", "", raw))
	base = re.sub(r"\s+[a-zµω°]+(?:/[a-z0-9²³]+)?$", "", base, flags=re.IGNORECASE)
	for entry in QUANTITIES.values():
		if any(re.search(rf"\b{re.escape(alias)}\b", base) for alias in entry[2]):
			return entry[0], entry[1], is_error
	return None, "", is_error


def choose_columns(
	frame: pd.DataFrame,
	columns: list[str],
	x_arg: str | None,
	y_arg: str | None,
) -> tuple[str, str]:
	names = [str(column) for column in frame.columns]
	for requested in (x_arg, y_arg):
		if requested and requested not in names:
			raise ValueError(f"Spalte {requested!r} nicht gefunden. Verfügbar: {', '.join(names)}")
	measurement_columns = [column for column in columns if not identify_column(column)[2]] or columns
	x_column = x_arg or next(
		(column for column in measurement_columns if identify_column(column)[0] == "Zeit"),
		measurement_columns[0],
	)
	if y_arg:
		y_column = y_arg
	else:
		possible = [column for column in columns if column != x_column and not identify_column(column)[2]]
		recognized = [column for column in possible if identify_column(column)[0]]
		choices = recognized or possible
		if not choices:
			raise ValueError("Keine zweite numerische Messspalte gefunden.")
		y_column = choices[0]
		if len(choices) > 1:
			print("Mögliche y-Spalten: " + ", ".join(choices))
			answer = input(f"y-Spalte auswählen (Enter für {y_column}): ").strip()
			if answer:
				if answer not in choices:
					raise ValueError(f"Ungültige y-Spalte: {answer!r}")
				y_column = answer
	if x_column == y_column or x_column not in columns or y_column not in columns:
		raise ValueError("x- und y-Spalte müssen unterschiedliche numerische Spalten sein.")
	return x_column, y_column


def linear_fit(x: np.ndarray, y: np.ndarray) -> tuple[float, float, float]:
	if len(x) < 2 or len(y) != len(x) or np.unique(x).size < 2:
		raise ValueError("Ein linearer Fit benötigt mindestens zwei unterschiedliche x-Werte.")
	slope, intercept = np.polyfit(x, y, 1)
	residual = float(np.sum((y - (slope * x + intercept)) ** 2))
	total = float(np.sum((y - np.mean(y)) ** 2))
	r_squared = 1 - residual / total if total else float("nan")
	return float(slope), float(intercept), r_squared