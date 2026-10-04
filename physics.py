"""Lokale, nachvollziehbare Regeln zur Erkennung physikalischer Größen."""

from __future__ import annotations

import re

import numpy as np
import pandas as pd

from models import ChartCandidate, ColumnInfo, RelationshipHint


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
	"impuls": ("Impuls", "kg·m/s", ("impuls", "momentum")),
	"drehmoment": ("Drehmoment", "N·m", ("drehmoment", "moment")),
	"winkelgeschwindigkeit": ("Winkelgeschwindigkeit", "rad/s", ("winkelgeschwindigkeit", "angular velocity")),
	"winkelbeschleunigung": ("Winkelbeschleunigung", "rad/s²", ("winkelbeschleunigung", "angular acceleration")),
	"dichte": ("Dichte", "kg/m³", ("dichte", "massendichte", "density")),
	"kapazitaet": ("Kapazität", "F", ("kapazitat", "kapazität", "capacitance")),
	"feldstaerke_elektrisch": ("Elektrische Feldstärke", "V/m", ("elektrische feldstarke", "elektrische feldstärke", "feldstarke elektrisch", "electric field")),
	"feldenergie": ("Feldenergie", "J", ("feldenergie", "field energy")),
	"flussdichte": ("Magnetische Flussdichte", "T", ("flussdichte", "flussdichte magnetisch", "magnetische flussdichte", "magnetic flux density")),
	"magnetischer_fluss": ("Magnetischer Fluss", "Wb", ("magnetischer fluss", "magnetfluss", "magnetic flux")),
	"magnetische_feldstaerke": ("Magnetische Feldstärke", "A/m", ("magnetische feldstarke", "magnetische feldstärke", "magnetic field strength")),
	"induktivitaet": ("Induktivität", "H", ("induktivitat", "induktivität", "inductance")),
	"waermemenge": ("Wärmemenge", "J", ("warmemenge", "wärmemenge", "warme", "wärme", "heat")),
	"waermekapazitaet": ("Spezifische Wärmekapazität", "J/(kg·K)", ("warmekapazitat", "wärmekapazität", "spezifische warme", "spezifische wärme", "specific heat capacity")),
	"waermeleitfaehigkeit": ("Wärmeleitfähigkeit", "W/(m·K)", ("warmeleitfahigkeit", "wärmeleitfähigkeit", "thermal conductivity")),
	"entropie": ("Entropie", "J/K", ("entropie", "entropy")),
	"wirkungsgrad": ("Wirkungsgrad", "", ("wirkungsgrad", "effizienz", "efficiency")),
	"amplitude": ("Amplitude", "m", ("amplitude",)),
	"phasendifferenz": ("Phasendifferenz", "rad", ("phasendifferenz", "phase difference", "phase")),
	"brennweite": ("Brennweite", "m", ("brennweite", "focal length")),
	"gegenstandsweite": ("Gegenstandsweite", "m", ("gegenstandsweite", "object distance")),
	"bildweite": ("Bildweite", "m", ("bildweite", "image distance")),
	"brechungsindex": ("Brechungsindex", "", ("brechungsindex", "refractive index")),
	"lichtgeschwindigkeit": ("Lichtgeschwindigkeit", "m/s", ("lichtgeschwindigkeit", "speed of light")),
	"intensitaet": ("Intensität", "W/m²", ("intensitat", "intensität", "lichtintensitat", "light intensity")),
	"photonenenergie": ("Photonenenergie", "J", ("photonenenergie", "photon energy", "energie photon")),
	"wellenzahl": ("Wellenzahl", "1/m", ("wellenzahl", "wave number", "wavevector")),
	"leitfaehigkeit": ("Leitfähigkeit", "S", ("leitfahigkeit", "leitfähigkeit", "conductivity", "electrical conductivity")),
	"halbwertszeit": ("Halbwertszeit", "s", ("halbwertszeit", "half-life")),
	"aktivitaet": ("Aktivität", "Bq", ("aktivitat", "aktivität", "radioaktivitat", "radioaktivität", "activity")),
	"dosis": ("Dosis", "Gy", ("dosis", "absorbierte dosis", "dose")),
	"auslenkung": ("Auslenkung", "m", ("auslenkung", "displacement")),
	"wellengeschwindigkeit": ("Wellengeschwindigkeit", "m/s", ("wellengeschwindigkeit", "wave speed")),
	"strahlungsfluss": ("Strahlungsfluss", "W", ("strahlungsfluss", "radiant flux")),
	"leistungsdichte": ("Leistungsdichte", "W/m²", ("leistungsdichte", "power density")),
	"magnetische_fluxdichte": ("Magnetische Flussdichte", "T", ("magnetische flussdichte", "flussdichte", "magnetic flux density")),
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
	"kg·m/s": ("Impuls", "kg·m/s"), "kg*m/s": ("Impuls", "kg·m/s"),
	"n·m": ("Drehmoment", "N·m"), "n*m": ("Drehmoment", "N·m"),
	"rad/s": ("Winkelgeschwindigkeit", "rad/s"), "rad/s2": ("Winkelbeschleunigung", "rad/s²"),
	"kg/m3": ("Dichte", "kg/m³"), "f": ("Kapazität", "F"),
	"v/m": ("Elektrische Feldstärke", "V/m"), "t": ("Magnetische Flussdichte", "T"),
	"wb": ("Magnetischer Fluss", "Wb"), "a/m": ("Magnetische Feldstärke", "A/m"),
	"j/k": ("Entropie", "J/K"),
	"j/(kg·k)": ("Spezifische Wärmekapazität", "J/(kg·K)"),
	"w/(m·k)": ("Wärmeleitfähigkeit", "W/(m·K)"),
	"rad": ("Winkel", "rad"), "bq": ("Aktivität", "Bq"),
	"gy": ("Dosis", "Gy"), "sv": ("Dosis", "Sv"),
	"w/m2": ("Intensität", "W/m²"), "m/s²": ("Beschleunigung", "m/s²"),
	"s": ("Leitfähigkeit", "S"), "ms": ("Leitfähigkeit", "mS"),
	"1/m": ("Wellenzahl", "1/m"),
}

RELATIONSHIP_RULES: tuple[RelationshipHint, ...] = (
	RelationshipHint("Zeit", "Strecke", "s(t)", "Bewegung: Strecke in Abhängigkeit von der Zeit", 0.92),
	RelationshipHint("Zeit", "Geschwindigkeit", "v(t)", "Bewegung: Geschwindigkeit in Abhängigkeit von der Zeit", 0.92),
	RelationshipHint("Zeit", "Beschleunigung", "a(t)", "Bewegung: Beschleunigung in Abhängigkeit von der Zeit", 0.92),
	RelationshipHint("Zeit", "Kraft", "F(t)", "Kraftmessung über die Zeit", 0.84),
	RelationshipHint("Strecke", "Kraft", "F(s)", "Kraft in Abhängigkeit von Strecke oder Verformung", 0.83),
	RelationshipHint("Dehnung", "Kraft", "F(s)", "Kraft in Abhängigkeit von Dehnung", 0.86),
	RelationshipHint("Spannung", "Stromstärke", "I(U) oder U(I)", "Elektrische Kennlinie; Richtung hängt vom Versuchsaufbau ab", 0.68),
	RelationshipHint("Stromstärke", "Widerstand", "R(I)", "Widerstand in Abhängigkeit von Strom", 0.68),
	RelationshipHint("Spannung", "Leistung", "P(U)", "Elektrische Leistung in Abhängigkeit von Spannung", 0.66),
	RelationshipHint("Stromstärke", "Leistung", "P(I)", "Elektrische Leistung in Abhängigkeit von Strom", 0.66),
	RelationshipHint("Zeit", "Temperatur", "T(t)", "Temperaturverlauf über die Zeit", 0.88),
	RelationshipHint("Volumen", "Druck", "p(V)", "Druck-Volumen-Beziehung; Richtung hängt vom Versuch ab", 0.68),
	RelationshipHint("Temperatur", "Volumen", "V(T)", "Volumen-Temperatur-Beziehung; Richtung hängt vom Versuch ab", 0.68),
	RelationshipHint("Frequenz", "Wellenlänge", "λ(f)", "Wellenbeziehung bei bekannter Wellengeschwindigkeit", 0.65),
	RelationshipHint("Frequenz", "Energie", "E(f)", "Photonenenergie ist proportional zur Frequenz", 0.72),
	RelationshipHint("Masse", "Energie", "E(m)", "Masse-Energie-Beziehung", 0.64),
	RelationshipHint("Gegenstandsweite", "Bildweite", "1/f = 1/g + 1/b", "Abbildung an einer dünnen Linse", 0.68),
	RelationshipHint("Zeit", "Aktivität", "A(t)", "Radioaktiver Zerfall über die Zeit", 0.68),
	RelationshipHint("Zeit", "Auslenkung", "x(t)", "Schwingung: Auslenkung in Abhängigkeit von der Zeit", 0.84),
)


def normalize(value: object) -> str:
	text = str(value).strip().lower()
	text = text.translate(str.maketrans({"ä": "a", "ö": "o", "ü": "u", "ß": "ss", "Ω": "ω"}))
	return re.sub(r"\s+", " ", text)


def identify_column(column: object) -> tuple[str | None, str, bool]:
	"""Gibt physikalische Größe, Einheit und Fehlerwert-Indikator zurück."""
	raw = str(column).strip()
	name = normalize(raw)
	is_error = bool(re.search(r"fehler|unsicherheit|uncertainty|error|delta|±|Δ", name, re.IGNORECASE))
	bracketed_unit = re.search(r"[\[(]\s*([^\])]+)\s*[\])]", raw)
	suffix_unit = None if bracketed_unit else re.search(
		r"\s+([a-zµω°]+(?:/[a-z0-9²³]+)?)\s*$", name, re.IGNORECASE
	)
	match = bracketed_unit or suffix_unit
	unit = match.group(1).strip().replace("²", "2").replace("³", "3") if match else ""
	base = normalize(re.sub(r"[\[(].*?[\])]", "", raw))
	if suffix_unit:
		base = re.sub(r"\s+[a-zµω°]+(?:/[a-z0-9²³]+)?$", "", base, flags=re.IGNORECASE)
	for entry in QUANTITIES.values():
		if any(re.search(rf"\b{re.escape(alias)}\b", base) for alias in entry[2]):
			unit_info = UNIT_QUANTITY.get(unit.lower())
			if unit_info and unit_info[0] in {entry[0], "Strecke" if entry[0] == "Dehnung" else entry[0]}:
				return entry[0], unit_info[1], is_error
			return entry[0], entry[1], is_error
	if unit.lower() in UNIT_QUANTITY:
		quantity, display_unit = UNIT_QUANTITY[unit.lower()]
		if bracketed_unit or len(unit) > 1 or unit in {"ω", "°"}:
			return quantity, display_unit, is_error
	return None, "", is_error


def inspect_columns(frame: pd.DataFrame) -> list[ColumnInfo]:
	"""Liefert erkannte Größen und robuste numerische Kennzahlen je Spalte."""
	columns: list[ColumnInfo] = []
	for raw_column in frame.columns:
		name = str(raw_column)
		quantity, unit, is_error = identify_column(name)
		values = pd.to_numeric(frame[raw_column], errors="coerce").to_numpy(dtype=float)
		finite_values = values[np.isfinite(values)]
		columns.append(ColumnInfo(
			name=name,
			quantity=quantity,
			unit=unit,
			numeric=finite_values.size >= 2 and np.unique(finite_values).size > 1,
			is_error=is_error,
			finite_count=int(finite_values.size),
			unique_count=int(np.unique(finite_values).size),
		))
	return columns


def relationship_hints(quantities: list[str | None]) -> list[RelationshipHint]:
	"""Gibt nur Regeln zurück, deren beide Größen im Versuch vorkommen."""
	present = set(quantities)
	return [
		rule for rule in RELATIONSHIP_RULES
		if rule.x_quantity in present and rule.y_quantity in present
	]


def infer_chart_candidates(frame: pd.DataFrame, columns: list[str]) -> list[ChartCandidate]:
	"""Ordnet mögliche x/y-Paare; nicht eindeutige Beziehungsrichtungen bleiben sichtbar."""
	infos = {info.name: info for info in inspect_columns(frame)}
	measurements = [
		column for column in columns
		if column in infos and infos[column].numeric and not infos[column].is_error
	]
	candidates: list[ChartCandidate] = []
	for left_index, left in enumerate(measurements):
		for right in measurements[left_index + 1:]:
			left_quantity = infos[left].quantity
			right_quantity = infos[right].quantity
			matching_rules = [
				rule for rule in RELATIONSHIP_RULES
				if {rule.x_quantity, rule.y_quantity} == {left_quantity, right_quantity}
			]
			if matching_rules:
				for rule in matching_rules:
					if rule.x_quantity == left_quantity and rule.y_quantity == right_quantity:
						x_column, y_column = left, right
					elif rule.x_quantity == right_quantity and rule.y_quantity == left_quantity:
						x_column, y_column = right, left
					else:
						continue
					confidence = rule.confidence
					candidates.append(ChartCandidate(
						x_column=x_column,
						y_column=y_column,
						relationship=rule.formula,
						reason=rule.context,
						confidence=confidence,
					))
				if any(rule.formula in {"I(U) oder U(I)"} for rule in matching_rules):
					for x_column, y_column in ((left, right), (right, left)):
						candidates.append(ChartCandidate(
							x_column=x_column,
							y_column=y_column,
							relationship="Kennlinie",
							reason="Elektrische Kennlinie: Versuchsanordnung bestimmt die unabhängige Größe.",
							confidence=0.62,
						))
				continue

			for x_column, y_column in ((left, right), (right, left)):
				x_info, y_info = infos[x_column], infos[y_column]
				confidence = 0.35
				reason = "Keine passende lokale Zusammenhangsregel; Versuchsaufbau oder Benutzerwahl erforderlich."
				if x_info.quantity == "Zeit":
					confidence = 0.76 if y_info.quantity else 0.58
					reason = "Zeit ist häufig die unabhängige Größe; Messkontext prüfen."
				elif x_info.quantity and not y_info.quantity:
					confidence = 0.45
				candidates.append(ChartCandidate(
					x_column=x_column,
					y_column=y_column,
					relationship=f"{y_info.quantity or y_column}({x_info.quantity or x_column})",
					reason=reason,
					confidence=confidence,
				))
	return sorted(candidates, key=lambda candidate: candidate.confidence, reverse=True)


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
	if x_arg and y_arg:
		x_column, y_column = x_arg, y_arg
	else:
		candidate_pairs: list[tuple[float, str, str]] = []
		for x_candidate in measurement_columns:
			for y_candidate in measurement_columns:
				if x_candidate == y_candidate:
					continue
				x_quantity, x_unit, _ = identify_column(x_candidate)
				y_quantity, y_unit, _ = identify_column(y_candidate)
				score = 0.0
				if x_quantity == "Zeit":
					score += 6.0
				if y_quantity == "Zeit":
					score -= 3.0
				if x_unit and y_unit and x_unit != y_unit:
					score += 1.5
				if x_quantity and y_quantity and {x_quantity, y_quantity} in ({"Zeit", "Strecke"}, {"Zeit", "Geschwindigkeit"}, {"Zeit", "Beschleunigung"}, {"Spannung", "Stromstärke"}, {"Frequenz", "Wellenlänge"}):
					score += 2.0
				if x_quantity and y_quantity:
					score += 0.5
				if any(token in normalize(x_candidate) for token in ("zeit", "time", "dauer")):
					score += 1.5
				if any(token in normalize(y_candidate) for token in ("fehler", "error", "delta", "unsicherheit")):
					score -= 10.0
				score += 0.25 * len([t for t in (x_candidate, y_candidate) if identify_column(t)[0]])
				candidate_pairs.append((score, x_candidate, y_candidate))
		if not candidate_pairs:
			raise ValueError("Keine zweite numerische Messspalte gefunden.")
		candidate_pairs.sort(key=lambda item: item[0], reverse=True)
		best_score, best_x, best_y = candidate_pairs[0]
		preferred = [(score, x, y) for score, x, y in candidate_pairs if score >= best_score - 0.1]
		if len(preferred) > 1:
			for _, x_candidate, y_candidate in preferred:
				if identify_column(x_candidate)[0] == "Zeit":
					best_x, best_y = x_candidate, y_candidate
					break
			else:
				best_x, best_y = preferred[0][1], preferred[0][2]
		else:
			best_x, best_y = best_x, best_y
		x_column, y_column = x_arg or best_x, y_arg or best_y
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