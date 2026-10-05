"""PLVS ULTRA Graphs Ribbon-Aktionen über xlwings und Excel COM."""

from __future__ import annotations

import math
from typing import Any

import numpy as np
import pandas as pd
import xlwings as xw

from ai import AIServiceError, analyze_with_ai, test_ai_connection
from config import load_config, save_config
from excel import get_numeric_columns
from models import ChartSpecification


PRODUCT_NAME = "PLVS ULTRA Graphs"
SOURCE_SHEET_NAME = "_PLVS_ULTRA_SourceSheet"
SOURCE_ADDRESS_NAME = "_PLVS_ULTRA_SourceAddress"
SOURCE_LABEL_NAME = "_PLVS_ULTRA_SourceLabel"
ANALYSIS_NAME = "_PLVS_ULTRA_Analysis"
CHART_NAME_PREFIX = "PLVS_ULTRA_Chart_"
CHART_SHEET_NAME = "_PLVS_ULTRA_ChartSheet"


def _active_book() -> xw.Book:
	try:
		return xw.Book.caller()
	except Exception as exc:
		raise RuntimeError(f"Keine aktive Arbeitsmappe gefunden. Starte die Aktion über das Blatt {PRODUCT_NAME}.") from exc


def _message(book: xw.Book, text: str, error: bool = False) -> None:
	try:
		if not book.app.visible:
			book.app.api.StatusBar = text
			return
		import ctypes
		flags = 0x10 if error else 0x40
		ctypes.windll.user32.MessageBoxW(int(book.app.api.Hwnd), text, PRODUCT_NAME, flags)
	except Exception:
		book.app.api.StatusBar = text


def _safe_action(action) -> None:
	book = None
	try:
		book = _active_book()
		action(book)
	except Exception as exc:
		message = str(exc) or "Unerwarteter Fehler in PLVS ULTRA Graphs."
		if book is None:
			raise
		_write_status(book, message)
		_message(book, message, error=True)


def _matrix(value: Any) -> list[list[Any]]:
	if value is None:
		return []
	if not isinstance(value, (list, tuple)):
		return [[value]]
	rows = [list(row) if isinstance(row, (list, tuple)) else [row] for row in value]
	while rows and not any(cell is not None for cell in rows[-1]):
		rows.pop()
	if rows:
		last_column = max((index for index in range(len(rows[0])) if any(row[index] is not None for row in rows if index < len(row))), default=-1)
		rows = [row[:last_column + 1] for row in rows]
	return rows


def _frame_from_values(value: Any) -> pd.DataFrame:
	rows = _matrix(value)
	if len(rows) < 2:
		raise ValueError("Der ausgewählte Bereich enthält keine Messwerttabelle mit mindestens einer Datenzeile.")
	width = max(len(row) for row in rows)
	rows = [row + [None] * (width - len(row)) for row in rows]
	headers: list[str] = []
	counts: dict[str, int] = {}
	for index, value in enumerate(rows[0], start=1):
		base = str(value).strip() if value is not None else f"Spalte {index}"
		if not base:
			base = f"Spalte {index}"
		counts[base] = counts.get(base, 0) + 1
		headers.append(base if counts[base] == 1 else f"{base} ({counts[base]})")
	frame = pd.DataFrame(rows[1:], columns=headers)
	frame = frame.dropna(how="all").reset_index(drop=True)
	if frame.empty:
		raise ValueError("Im ausgewählten Bereich wurden keine Messwerte gefunden.")
	return frame


def _frame_from_range(data_range: xw.Range) -> pd.DataFrame:
	return _frame_from_values(data_range.value)


def _range_key(sheet: xw.Sheet, data_range: xw.Range) -> str:
	return f"{sheet.name.casefold()}!{str(data_range.address).replace('$', '').casefold()}"


def _display_address(data_range: xw.Range) -> str:
	return str(data_range.address).replace("$", "")


def _source_label(sheet: xw.Sheet, data_range: xw.Range) -> str:
	return f"{sheet.name}: {_display_address(data_range)}"


def _name_item(book: xw.Book, name: str):
	names = book.api.Names
	for index in range(1, int(names.Count) + 1):
		item = names.Item(index)
		if str(item.Name).split("!")[-1].casefold() == name.casefold():
			return item
	return None


def _set_name(book: xw.Book, name: str, refers_to: str) -> None:
	item = _name_item(book, name)
	if item is None:
		book.api.Names.Add(Name=name, RefersTo=refers_to, Visible=False)
	else:
		item.RefersTo = refers_to
		item.Visible = False


def _set_text_name(book: xw.Book, name: str, value: str) -> None:
	chunks = [
		value[index:index + 120].replace('"', '""')
		for index in range(0, len(value), 120)
	] or [""]
	refers_to = "=" + "&".join(f'"{chunk}"' for chunk in chunks)
	_set_name(book, name, refers_to)


def _source_from_range(
	sheet: xw.Sheet,
	data_range: xw.Range,
	label: str,
) -> tuple[xw.Sheet, xw.Range, str] | None:
	try:
		frame = _frame_from_range(data_range)
	except ValueError:
		return None
	if len(get_numeric_columns(frame)) < 2:
		return None
	return sheet, data_range, label


def _overlaps(first: xw.Range, second: xw.Range) -> bool:
	first_top, first_left = int(first.api.Row), int(first.api.Column)
	second_top, second_left = int(second.api.Row), int(second.api.Column)
	first_bottom = first_top + int(first.api.Rows.Count) - 1
	second_bottom = second_top + int(second.api.Rows.Count) - 1
	first_right = first_left + int(first.api.Columns.Count) - 1
	second_right = second_left + int(second.api.Columns.Count) - 1
	return not (
		first_bottom < second_top
		or second_bottom < first_top
		or first_right < second_left
		or second_right < first_left
	)


def _connected_regions(value: Any) -> list[tuple[int, int, int, int]]:
	rows = _matrix(value)
	if not rows:
		return []
	width = max(map(len, rows))
	rows = [row + [None] * (width - len(row)) for row in rows]
	visited: set[tuple[int, int]] = set()
	regions: list[tuple[int, int, int, int]] = []
	for row_index, row in enumerate(rows):
		for column_index, cell in enumerate(row):
			if cell is None or cell == "" or (row_index, column_index) in visited:
				continue
			pending = [(row_index, column_index)]
			visited.add((row_index, column_index))
			region_rows: list[int] = []
			region_columns: list[int] = []
			while pending:
				current_row, current_column = pending.pop()
				region_rows.append(current_row)
				region_columns.append(current_column)
				for next_row, next_column in (
					(current_row - 1, current_column),
					(current_row + 1, current_column),
					(current_row, current_column - 1),
					(current_row, current_column + 1),
				):
					if (
						0 <= next_row < len(rows)
						and 0 <= next_column < width
						and rows[next_row][next_column] is not None
						and rows[next_row][next_column] != ""
						and (next_row, next_column) not in visited
					):
						visited.add((next_row, next_column))
						pending.append((next_row, next_column))
			regions.append((min(region_rows), min(region_columns), max(region_rows), max(region_columns)))
	return regions


def _discover_data_sources(book: xw.Book) -> list[tuple[xw.Sheet, xw.Range, str]]:
	sources: list[tuple[xw.Sheet, xw.Range, str]] = []
	seen: set[str] = set()
	for sheet in book.sheets:
		if int(sheet.api.Visible) != -1:
			continue
		ranges: list[xw.Range] = []
		tables = sheet.api.ListObjects
		for index in range(1, int(tables.Count) + 1):
			table = tables.Item(index)
			data_range = sheet.range(table.Range.Address)
			ranges.append(data_range)
			source = _source_from_range(
				sheet,
				data_range,
				_source_label(sheet, data_range),
			)
			if source is not None:
				key = _range_key(sheet, data_range)
				if key not in seen:
					sources.append(source)
					seen.add(key)
		used = sheet.used_range
		values = used.value
		for top, left, bottom, right in _connected_regions(values):
			data_range = sheet.range(
				(int(used.api.Row) + top, int(used.api.Column) + left),
				(int(used.api.Row) + bottom, int(used.api.Column) + right),
			)
			if any(_overlaps(data_range, table_range) for table_range in ranges):
				continue
			key = _range_key(sheet, data_range)
			if key in seen:
				continue
			label = _source_label(sheet, data_range)
			source = _source_from_range(sheet, data_range, label)
			if source is not None:
				sources.append(source)
				seen.add(key)

	current = _current_data_range(book)
	if current is not None:
		sheet, data_range = current
		key = _range_key(sheet, data_range)
		if key in seen:
			for index, (source_sheet, source_range, label) in enumerate(sources):
				if _range_key(source_sheet, source_range) == key:
					sources[index] = (source_sheet, source_range, label)
					break
		else:
			source = _source_from_range(
				sheet,
				data_range,
				_source_label(sheet, data_range),
			)
			if source is not None:
				sources.append(source)
	return sources


def _current_data_range(book: xw.Book) -> tuple[xw.Sheet, xw.Range] | None:
	sheet = book.sheets.active
	data_range = sheet.range(sheet.api.Application.ActiveCell.CurrentRegion.Address)
	if _source_from_range(sheet, data_range, "") is None:
		return None
	return sheet, data_range


def _store_source(book: xw.Book, sheet: xw.Sheet, data_range: xw.Range, label: str) -> None:
	_set_text_name(book, SOURCE_SHEET_NAME, sheet.name)
	_set_text_name(book, SOURCE_ADDRESS_NAME, data_range.address)
	_set_text_name(book, SOURCE_LABEL_NAME, label)
	_set_text_name(book, ANALYSIS_NAME, f"Datenbereich: {label} — Diagramm erstellen.")


def _manual_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range, str]:
	selected = book.app.api.InputBox(
		Prompt="Markiere den zusammenhängenden Zellbereich einschließlich Überschriften.",
		Title=f"{PRODUCT_NAME} – Bereich",
		Type=8,
	)
	if selected is False or selected is None:
		raise ValueError("Die Bereichsauswahl wurde abgebrochen.")
	if int(selected.Areas.Count) != 1:
		raise ValueError("Bitte einen zusammenhängenden Zellbereich auswählen.")
	if str(selected.Worksheet.Parent.FullName).casefold() != str(book.fullname).casefold():
		raise ValueError("Bitte einen Zellbereich aus der gerade verwendeten Arbeitsmappe auswählen.")
	sheet = book.sheets[str(selected.Worksheet.Name)]
	data_range = sheet.range(selected.Address)
	label = _source_label(sheet, data_range)
	source = _source_from_range(sheet, data_range, label)
	if source is None:
		raise ValueError("Der ausgewählte Bereich benötigt Überschriften und mindestens zwei numerische Messspalten.")
	return source


def _choose_data_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range]:
	sources = _discover_data_sources(book)
	if len(sources) == 1:
		sheet, data_range, label = sources[0]
		_store_source(book, sheet, data_range, label)
		return sheet, data_range
	options = [label for _, _, label in sources]
	options.append("Manuell auswählen …")
	if not sources:
		return _select_manual_source(book)
	prompt = (
		"Welche Datenquelle soll analysiert werden?\n\n"
		+ "\n".join(f"{index}. {label}" for index, label in enumerate(options, 1))
		+ "\n\nNummer eingeben:"
	)
	response = book.app.api.InputBox(
		Prompt=prompt,
		Title=f"{PRODUCT_NAME} – Datenquelle",
		Default="",
		Type=2,
	)
	if response is False or response is None:
		raise ValueError("Die Datenquellenauswahl wurde abgebrochen.")
	if not str(response).strip():
		raise ValueError("Die Datenquellenauswahl wurde abgebrochen.")
	try:
		selection = int(str(response).strip())
	except ValueError as exc:
		raise ValueError("Bitte die Nummer einer Datenquelle eingeben.") from exc
	if not 1 <= selection <= len(options):
		raise ValueError("Die ausgewählte Datenquellennummer ist ungültig.")
	if selection == len(options):
		return _select_manual_source(book)
	sheet, data_range, label = sources[selection - 1]
	_store_source(book, sheet, data_range, label)
	return sheet, data_range


def _select_manual_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range]:
	sheet, data_range, label = _manual_source(book)
	_store_source(book, sheet, data_range, label)
	return sheet, data_range


def _source_for_chart(book: xw.Book) -> tuple[xw.Sheet, xw.Range]:
	current = _current_data_range(book)
	if current is not None:
		sheet, data_range = current
		label = _source_label(sheet, data_range)
		_store_source(book, sheet, data_range, label)
		return sheet, data_range
	return _choose_data_source(book)


def _number(value: float) -> str:
	return f"{value:.4g}"


def _is_price_series(spec: ChartSpecification) -> bool:
	text = f"{spec.y_column} {spec.y_label} {spec.y_unit}".casefold()
	if any(term in text for term in ("rendite", "return", "performance", "prozent", "%")):
		return False
	return any(term in text for term in (
		"aktie", "stock", "share", "kurs", "preis", "price", "close", "schlusskurs",
		"eur", "usd", "gbp", "chf", "€", "$", "£", "¥",
	))


def _analysis_label(
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> str:
	x_label = spec.x_label.replace("|", "/")
	y_label = spec.y_label.replace("|", "/")
	y_unit = spec.y_unit.replace("|", "/").strip()
	unit_suffix = f" {y_unit}" if y_unit else ""
	independent_variable = (spec.independent_variable or x_label).replace("|", "/")
	dependent_variable = (spec.dependent_variable or y_label).replace("|", "/")
	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	x_data, y_data = x_values[valid], y_values[valid]
	fit_labels = {
		"none": "kein",
		"linear": "linear",
		"quadratic": "quadratisch",
		"cubic": "kubisch",
		"exponential": "exponentiell",
		"logarithmic": "logarithmisch",
		"power": "Potenz",
	}
	fit_label = fit_labels.get(spec.trendline, spec.trendline)
	chart_labels = {
		"scatter": "Punktdiagramm (XY)",
		"line": "Liniendiagramm",
		"line_scatter": "Punkt- und Liniendiagramm",
	}
	chart_label = chart_labels.get(spec.chart_type, spec.chart_type)
	mean_y = float(np.mean(y_data)) if y_data.size else None
	median_y = float(np.median(y_data)) if y_data.size else None
	min_y = float(np.min(y_data)) if y_data.size else None
	max_y = float(np.max(y_data)) if y_data.size else None
	std_y = float(np.std(y_data, ddof=1)) if y_data.size > 1 else None
	correlation = None
	if x_data.size > 1 and y_data.size > 1 and np.std(x_data) and np.std(y_data):
		try:
			correlation = float(np.corrcoef(x_data, y_data)[0, 1])
		except FloatingPointError:
			correlation = None
		if not math.isfinite(correlation):
			correlation = None
	parts = [
		f"Unabhängig: {x_label}; Abhängig: {y_label}",
		f"Zusammenhang: {independent_variable} → {dependent_variable}",
		f"Diagramm: {chart_label}; Trendlinie: {fit_label}",
		f"Messpunkte: {len(y_data)} von {len(frame)}; Mittelwert y: "
		f"{_number(mean_y) if mean_y is not None else 'n. v.'} {y_unit}".rstrip(),
	]
	statistics = []
	if median_y is not None:
		statistics.append(f"Median y: {_number(median_y)} {y_unit}".rstrip())
	if min_y is not None and max_y is not None:
		statistics.append(f"Min–Max y: {_number(min_y)}–{_number(max_y)} {y_unit}".rstrip())
	if std_y is not None:
		statistics.append(f"Stdabw. y: {_number(std_y)} {y_unit}".rstrip())
	if correlation is not None:
		statistics.append(f"Korrelation: {_number(correlation)}")
	if statistics:
		parts.append("Statistik: " + "; ".join(statistics))
	if _is_price_series(spec) and x_data.size:
		order = np.argsort(x_data, kind="stable")
		start_value = float(y_data[order[0]])
		end_value = float(y_data[order[-1]])
		change = end_value - start_value
		change_text = (
			f"Veränderung: Kurs {_number(start_value)} → {_number(end_value)}{unit_suffix}; "
			f"absolut {change:+.4g}{unit_suffix}"
		)
		if start_value != 0:
			percent_change = change / abs(start_value) * 100
			change_text += f"; Prozent: {percent_change:+.4g}%"
		else:
			change_text += "; Prozent: nicht definiert (Startwert 0)"
		parts.append(change_text)
	if spec.confidence is not None:
		reason = spec.reason.replace("|", "/").strip()
		if len(reason) > 140:
			reason = reason[:137].rstrip() + "..."
		parts.append(f"KI-Einschätzung: {spec.confidence:.0%} — {reason}")
	return " | ".join(parts)


def _write_analysis(
	book: xw.Book,
	frame: pd.DataFrame,
	spec: ChartSpecification,
	chart_name: str,
) -> None:
	analysis = _analysis_label(frame, spec)
	_set_text_name(book, ANALYSIS_NAME, analysis)
	_set_text_name(book, f"{ANALYSIS_NAME}_{chart_name}", analysis)


def show_help() -> None:
	def action(book: xw.Book) -> None:
		help_text = (
			f"{PRODUCT_NAME}\n\n"
			"- Analysiert allgemeine Tabellen und Messdaten mit der KI.\n"
			"- Die KI wählt Achsen, Diagrammart und einen passenden Kurvenfit.\n"
			"- Unterstützt werden unter anderem lineare, exponentielle, logarithmische und Potenz-Trends.\n"
			"- Eine aktive Tabelle oder ein markierter zusammenhängender Bereich wird direkt verwendet.\n"
			"- Die Quellenauswahl erscheint nur beim Erstellen eines Diagramms, wenn keine passende Auswahl aktiv ist.\n"
			"- Erstellt ein natives Excel-Diagramm und zeigt Analyse, Datenumfang und KI-Einschätzung im Menüband.\n"
			"- Bei fehlender oder unsicherer KI-Diagnose wird mit einer verständlichen Meldung abgebrochen.\n"
			"- Es werden keine Analyse- oder Hilfsblätter angelegt."
		)
		_message(book, help_text)
		_write_status(book, "Hilfe angezeigt.")
	_safe_action(action)


def _get_ai_config() -> dict[str, str]:
	try:
		config = load_config()
	except FileNotFoundError:
		config = {"endpoint": "", "api_key": "", "model": ""}
	return config


def _write_status(book: xw.Book, message: str) -> None:
	book.app.api.StatusBar = message


def open_ai_settings() -> None:
	def action(book: xw.Book) -> None:
		config = _get_ai_config()
		endpoint = book.app.api.InputBox(
			Prompt="OpenAI-kompatibler API-Endpunkt:",
			Title=f"{PRODUCT_NAME} – KI-Einstellungen",
			Default=config["endpoint"],
			Type=2,
		)
		if endpoint is False:
			return
		model = book.app.api.InputBox(
			Prompt="Modellname:",
			Title=f"{PRODUCT_NAME} – KI-Einstellungen",
			Default=config["model"],
			Type=2,
		)
		if model is False:
			return
		api_key = book.app.api.InputBox(
			Prompt="API-Key (optional; Eingabe wird lokal in config.json gespeichert):",
			Title=f"{PRODUCT_NAME} – KI-Einstellungen",
			Default="",
			Type=2,
		)
		if api_key is False:
			return
		config["endpoint"] = str(endpoint).strip()
		config["model"] = str(model).strip()
		if str(api_key).strip():
			config["api_key"] = str(api_key).strip()
		if not config["endpoint"] or not config["model"]:
			raise AIServiceError("Bitte Endpoint und Modell eintragen. Die KI ist für die Diagrammanalyse erforderlich.")
		path = save_config(config)
		_write_status(book, "KI-Einstellungen lokal gespeichert.")
		_message(book, f"KI-Einstellungen gespeichert. Zugangsdaten liegen ausschließlich lokal: {path}")
	_safe_action(action)


def _axis_bounds(values: np.ndarray, low: float | None, high: float | None, zero: bool) -> tuple[float, float]:
	data_low, data_high = float(values.min()), float(values.max())
	padding = (data_high - data_low) * 0.05 if data_high != data_low else max(abs(data_low) * 0.05, 1.0)
	minimum = float(low) if low is not None else data_low - padding
	maximum = float(high) if high is not None else data_high + padding
	if zero and data_low >= 0 and low is None:
		minimum = 0.0
	if not math.isfinite(minimum) or not math.isfinite(maximum) or minimum >= maximum:
		raise ValueError("Die Diagramm-Achsengrenzen sind ungültig.")
	return minimum, maximum


def _next_chart_name(book: xw.Book) -> str:
	last_number = 0
	for sheet in book.sheets:
		charts = sheet.api.ChartObjects()
		for index in range(1, int(charts.Count) + 1):
			name = str(charts.Item(index).Name)
			if name.startswith(CHART_NAME_PREFIX):
				suffix = name[len(CHART_NAME_PREFIX):]
				if suffix.isdigit():
					last_number = max(last_number, int(suffix))
	for index in range(1, int(book.api.Names.Count) + 1):
		name = str(book.api.Names.Item(index).Name).split("!")[-1]
		analysis_prefix = f"{ANALYSIS_NAME}_{CHART_NAME_PREFIX}"
		if name.startswith(analysis_prefix):
			suffix = name[len(analysis_prefix):]
			if suffix.isdigit():
				last_number = max(last_number, int(suffix))
	return f"{CHART_NAME_PREFIX}{last_number + 1}"


def _add_native_chart(
	book: xw.Book,
	sheet: xw.Sheet,
	data_range: xw.Range,
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> str:
	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	if int(valid.sum()) < 2 or np.unique(x_values[valid]).size < 2:
		raise ValueError("Mindestens zwei gültige Messwertpaare mit verschiedenen x-Werten sind erforderlich.")
	degree = {"linear": 1, "quadratic": 2, "cubic": 3}.get(spec.trendline)
	if spec.trendline != "none":
		if spec.trendline not in {"linear", "quadratic", "cubic", "exponential", "logarithmic", "power"}:
			raise ValueError(f"Unbekannte Trendlinie: {spec.trendline}")
		if degree is not None and np.unique(x_values[valid]).size <= degree:
			raise ValueError(f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} unterschiedliche x-Werte.")
		if spec.trendline == "exponential" and np.any(y_values[valid] <= 0):
			raise ValueError("Ein exponentieller Fit erfordert positive y-Werte.")
		if spec.trendline in {"logarithmic", "power"} and np.any(x_values[valid] <= 0):
			raise ValueError("Ein logarithmischer oder Potenz-Fit erfordert positive x-Werte.")
		if spec.trendline == "power" and np.any(y_values[valid] <= 0):
			raise ValueError("Ein Potenz-Fit erfordert positive y-Werte.")
	used = sheet.used_range.api
	left = float(used.Left) + float(used.Width) + 24
	top = float(used.Top)
	existing_charts = sheet.api.ChartObjects()
	for index in range(1, int(existing_charts.Count) + 1):
		existing_chart = existing_charts.Item(index)
		if str(existing_chart.Name).startswith("PLVS_ULTRA_Chart"):
			top = max(top, float(existing_chart.Top) + float(existing_chart.Height) + 24)
	chart_object = sheet.api.ChartObjects().Add(left, top, 500, 300)
	chart_object.Name = _next_chart_name(book)
	chart = chart_object.Chart
	chart.ChartType = 74 if spec.connect_points or spec.chart_type in {"line", "line_scatter"} else -4169
	chart.HasTitle = True
	chart.ChartTitle.Text = f"{spec.y_label} in Abhängigkeit von {spec.x_label}"
	chart.HasLegend = False
	series = chart.SeriesCollection().NewSeries()
	series.Name = "Messwerte"
	first_data_row = int(data_range.api.Row) + 1
	last_data_row = int(data_range.api.Row) + int(data_range.api.Rows.Count) - 1
	if last_data_row < first_data_row:
		raise ValueError("Der ausgewählte Bereich enthält keine Datenzeilen.")
	first_column = int(data_range.api.Column)
	x_column = first_column + int(frame.columns.get_loc(spec.x_column))
	y_column = first_column + int(frame.columns.get_loc(spec.y_column))
	series.XValues = sheet.range((first_data_row, x_column), (last_data_row, x_column)).api
	series.Values = sheet.range((first_data_row, y_column), (last_data_row, y_column)).api
	series.MarkerStyle = 8 if spec.show_points else -4142
	chart.DisplayBlanksAs = 1
	for axis_type, label, values, minimum, maximum in (
		(1, spec.x_label, x_values, spec.x_min, spec.x_max),
		(2, spec.y_label, y_values, spec.y_min, spec.y_max),
	):
		axis = chart.Axes(axis_type)
		axis.HasTitle = True
		axis.AxisTitle.Text = label
		valid_values = values[np.isfinite(values)]
		is_zero = spec.origin is True
		axis_minimum, axis_maximum = _axis_bounds(valid_values, minimum, maximum, is_zero)
		axis.MinimumScale = axis_minimum
		axis.MaximumScale = axis_maximum
	if spec.trendline != "none":
		trendlines = series.Trendlines()
		trendline_type = {
			"linear": -4132,
			"exponential": -4133,
			"logarithmic": -4134,
			"power": -4136,
		}
		if spec.trendline in trendline_type:
			trendlines.Add(Type=trendline_type[spec.trendline], Name=f"{spec.trendline.title()} Fit")
		else:
			trendlines.Add(Type=3, Order=degree, Name=f"Polynomfit Grad {degree}")
	_set_text_name(book, CHART_SHEET_NAME, sheet.name)
	return str(chart_object.Name)


def create_chart() -> None:
	def action(book: xw.Book) -> None:
		data_sheet, data_range = _source_for_chart(book)
		frame = _frame_from_range(data_range)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Der ausgewählte Bereich benötigt mindestens zwei numerische Messspalten.")
		spec = analyze_with_ai(frame, columns, _get_ai_config())
		chart_name = _add_native_chart(book, data_sheet, data_range, frame, spec)
		_write_analysis(book, frame, spec, chart_name)
		data_sheet.activate()
		data_sheet.api.ChartObjects(chart_name).Activate()
		message = f"Diagramm und Analyse erstellt: {spec.x_label} gegen {spec.y_label}."
		_write_status(book, message)
		_message(book, message)
	_safe_action(action)


def test_ai() -> None:
	def action(book: xw.Book) -> None:
		config = _get_ai_config()
		if not config["endpoint"] or not config["model"]:
			raise AIServiceError("KI-Einstellungen unvollständig. Endpoint und Modell über „KI-Einstellungen“ eintragen.")
		test_ai_connection(config)
		_write_status(book, "KI-Verbindung erfolgreich.")
		_message(book, "KI-Verbindung erfolgreich.")
	_safe_action(action)
