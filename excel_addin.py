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
from physics import identify_column, infer_chart_candidates, linear_fit


PRODUCT_NAME = "PLVS ULTRA Graphs"
SOURCE_SHEET_NAME = "_PLVS_ULTRA_SourceSheet"
SOURCE_ADDRESS_NAME = "_PLVS_ULTRA_SourceAddress"
SOURCE_LABEL_NAME = "_PLVS_ULTRA_SourceLabel"
ANALYSIS_NAME = "_PLVS_ULTRA_Analysis"
CHART_SHEET_NAME = "_PLVS_ULTRA_ChartSheet"
SETTING_PREFIX = "_PLVS_ULTRA_Setting_"


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
	escaped = value.replace('"', '""')
	_set_name(book, name, f'="{escaped}"')


def _get_text_name(book: xw.Book, name: str, default: str = "") -> str:
	item = _name_item(book, name)
	if item is None:
		return default
	value = book.app.api.Evaluate(item.RefersTo)
	return str(value) if value is not None else default


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
				f"Tabelle „{table.Name}“ — {sheet.name}!{data_range.address}",
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
			label = f"Datenbereich — {sheet.name}!{data_range.address}"
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
					sources[index] = (
						source_sheet,
						source_range,
						f"Aktueller Datenbereich — {label}",
					)
					break
		else:
			source = _source_from_range(
				sheet,
				data_range,
				f"Aktueller Datenbereich — {sheet.name}!{data_range.address}",
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
	_set_text_name(book, ANALYSIS_NAME, f"Datenquelle: {label} — Diagramm erstellen.")


def _selected_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range] | None:
	sheet_name = _get_text_name(book, SOURCE_SHEET_NAME)
	address = _get_text_name(book, SOURCE_ADDRESS_NAME)
	if not sheet_name and not address:
		return None
	if not sheet_name or not address:
		raise ValueError("Die gespeicherte Datenquelle ist unvollständig. Bitte erneut auswählen.")
	if sheet_name not in book.sheet_names:
		raise ValueError("Das Tabellenblatt der gespeicherten Datenquelle ist nicht mehr verfügbar. Bitte erneut auswählen.")
	sheet = book.sheets[sheet_name]
	return sheet, sheet.range(address)


def _manual_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range, str]:
	selected = book.app.api.InputBox(
		Prompt="Markiere den zusammenhängenden Zellbereich einschließlich Überschriften.",
		Title=f"{PRODUCT_NAME} – Datenquelle",
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
	label = f"Manuelle Auswahl — {sheet.name}!{data_range.address}"
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
	source = _selected_source(book)
	if source is not None:
		return source
	return _choose_data_source(book)


def select_data_source() -> None:
	def action(book: xw.Book) -> None:
		sheet, data_range = _choose_data_source(book)
		sheet.activate()
		data_range.select()
		label = _get_text_name(book, ANALYSIS_NAME)
		_write_status(book, label or f"Datenquelle ausgewählt: {sheet.name}!{data_range.address}")
	_safe_action(action)


def _prompt_choice(book: xw.Book, label: str, options: list[str], default: str | None = None) -> str:
	if not options:
		raise ValueError(f"Keine möglichen Werte für {label} gefunden.")
	default_text = f"\nVorschlag: {default}" if default else ""
	prompt = f"{label}:\n" + "\n".join(f"- {option}" for option in options) + default_text
	response = book.app.api.InputBox(prompt, PRODUCT_NAME, default or options[0], Type=2)
	if response is False or response is None:
		raise ValueError("Auswahl abgebrochen.")
	choice = str(response).strip()
	if choice not in options:
		raise ValueError(f"Ungültige Auswahl für {label}: {choice}")
	return choice


def _resolve_spec(book: xw.Book, frame: pd.DataFrame, columns: list[str], allow_ai: bool) -> tuple[ChartSpecification, list]:
	candidates = infer_chart_candidates(frame, columns)
	selected_x = _get_text_name(book, f"{SETTING_PREFIX}X", "automatisch").strip()
	selected_y = _get_text_name(book, f"{SETTING_PREFIX}Y", "automatisch").strip()
	best = candidates[0] if candidates else None
	ai_spec: ChartSpecification | None = None
	manual_confirmation_required = False
	if selected_x not in {"", "automatisch", "auto"} and selected_x not in columns:
		raise ValueError(f"Die eingestellte x-Spalte {selected_x!r} ist nicht numerisch oder fehlt.")
	if selected_y not in {"", "automatisch", "auto"} and selected_y not in columns:
		raise ValueError(f"Die eingestellte y-Spalte {selected_y!r} ist nicht numerisch oder fehlt.")

	if selected_x in columns and selected_y in columns and selected_x != selected_y:
		x_column, y_column = selected_x, selected_y
	elif selected_x in columns:
		x_column = selected_x
		possible_y = [column for column in columns if column != x_column and not identify_column(column)[2]]
		matching_y = [candidate for candidate in candidates if candidate.x_column == x_column and candidate.y_column in possible_y]
		if len(possible_y) == 1:
			y_column = possible_y[0]
		elif len(matching_y) == 1 and matching_y[0].confidence >= 0.84:
			y_column = matching_y[0].y_column
		else:
			y_column = _prompt_choice(book, "Abhängige Größe wählen", possible_y, matching_y[0].y_column if matching_y else None)
	elif selected_y in columns:
		y_column = selected_y
		possible_x = [column for column in columns if column != y_column and not identify_column(column)[2]]
		matching_x = [candidate for candidate in candidates if candidate.y_column == y_column and candidate.x_column in possible_x]
		if len(possible_x) == 1:
			x_column = possible_x[0]
		elif len(matching_x) == 1 and matching_x[0].confidence >= 0.84:
			x_column = matching_x[0].x_column
		else:
			x_column = _prompt_choice(book, "Unabhängige Größe wählen", possible_x, matching_x[0].x_column if matching_x else None)
	else:
		second_confidence = candidates[1].confidence if len(candidates) > 1 else 0.0
		is_clear = bool(best and best.confidence >= 0.84 and best.confidence - second_confidence >= 0.12)
		ambiguous = not is_clear
		use_ai = _get_text_name(book, f"{SETTING_PREFIX}UseAI", "Nein").strip().lower() in {"ja", "yes"}
		if ambiguous and allow_ai and use_ai:
			config = _get_ai_config()
			ai_spec = analyze_with_ai(frame, columns, config)
			if ai_spec.confidence is not None and ai_spec.confidence >= 0.9:
				ai_spec.reason = f"KI: {ai_spec.reason}"
				return ai_spec, candidates
			if ai_spec.confidence is not None and ai_spec.confidence >= 0.7:
				answer = book.app.api.MsgBox(
					f"Die KI schlägt {ai_spec.x_label} gegen {ai_spec.y_label} vor.\n\n"
					f"Confidence: {ai_spec.confidence:.0%}\n{ai_spec.reason}\n\n"
					"Entscheidung mit Vorbehalt übernehmen?",
					4 + 32,
					PRODUCT_NAME,
				)
				if answer == 6:
					ai_spec.reason = f"KI: {ai_spec.reason}"
					return ai_spec, candidates
				manual_confirmation_required = True
			else:
				book.app.api.MsgBox(
					"Die KI konnte keine ausreichend sichere Entscheidung treffen. Bitte x- und y-Größe auswählen.",
					48,
					PRODUCT_NAME,
				)
				manual_confirmation_required = True
		if is_clear and not manual_confirmation_required:
			x_column, y_column = best.x_column, best.y_column
		else:
			possible_x = list(dict.fromkeys(candidate.x_column for candidate in candidates)) or columns
			x_column = _prompt_choice(book, "Mehrere mögliche unabhängige Größen gefunden. X-Achse wählen", possible_x, best.x_column if best else None)
			possible_y = [column for column in columns if column != x_column and not identify_column(column)[2]]
			candidate_y = [candidate.y_column for candidate in candidates if candidate.x_column == x_column]
			default_y = candidate_y[0] if candidate_y else (possible_y[0] if possible_y else None)
			y_column = _prompt_choice(book, "Abhängige Größe wählen", possible_y, default_y)
		if x_column == y_column:
			raise ValueError("X- und Y-Achse müssen unterschiedliche Spalten verwenden.")

	x_quantity, x_unit, _ = identify_column(x_column)
	y_quantity, y_unit, _ = identify_column(y_column)
	x_label = f"{x_quantity} [{x_unit}]" if x_quantity and x_unit else x_quantity or x_column
	y_label = f"{y_quantity} [{y_unit}]" if y_quantity and y_unit else y_quantity or y_column
	fit = _get_text_name(book, f"{SETTING_PREFIX}Fit", "auto").strip().lower()
	if fit == "auto":
		x_values = pd.to_numeric(frame[x_column], errors="coerce").to_numpy(dtype=float)
		y_values = pd.to_numeric(frame[y_column], errors="coerce").to_numpy(dtype=float)
		valid = np.isfinite(x_values) & np.isfinite(y_values)
		fit = "none"
		if valid.sum() >= 3 and np.unique(x_values[valid]).size >= 2:
			if linear_fit(x_values[valid], y_values[valid])[2] >= 0.85:
				fit = "linear"
	error_for = {
		identify_column(str(column))[0]: str(column)
		for column in frame.columns
		if identify_column(str(column))[2]
	}
	if fit not in {"none", "linear", "quadratic", "cubic"}:
		raise ValueError(f"Unbekannter Trendlinientyp: {fit}")
	selected_candidate = next(
		(candidate for candidate in candidates if candidate.x_column == x_column and candidate.y_column == y_column),
		None,
	)
	spec = ChartSpecification(
		x_column=x_column,
		y_column=y_column,
		x_label=x_label,
		y_label=y_label,
		x_unit=x_unit,
		y_unit=y_unit,
		chart_type="line_scatter" if _get_text_name(book, f"{SETTING_PREFIX}ChartType", "XY-Streudiagramm").startswith("XY mit") else "scatter",
		origin={"ja": True, "nein": False, "auto": "auto"}.get(
			_get_text_name(book, f"{SETTING_PREFIX}Origin", "auto").strip().lower(),
			"auto",
		),
		trendline=fit,
		x_error_column=error_for.get(x_quantity),
		y_error_column=error_for.get(y_quantity),
		x_quantity=x_quantity,
		y_quantity=y_quantity,
		independent_variable=x_quantity,
		dependent_variable=y_quantity,
		confidence=selected_candidate.confidence if selected_candidate else None,
		reason=(f"KI mit Confidence {ai_spec.confidence:.0%} verworfen; Auswahl durch Benutzer." if ai_spec is not None else selected_candidate.reason if selected_candidate else "Benutzer hat x- und y-Spalte ausgewählt."),
	)
	return spec, candidates


def _number(value: float) -> str:
	return f"{value:.4g}"


def _analysis_label(
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> str:
	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	x_data, y_data = x_values[valid], y_values[valid]
	parts = [
		f"Unabhängig: {spec.x_label}; Abhängig: {spec.y_label}",
		(
			f"Diagramm: {'XY mit Linie' if spec.connect_points or spec.chart_type == 'line_scatter' else 'XY'}; "
			f"Fit: {'kein' if spec.trendline == 'none' else spec.trendline.capitalize()}"
		),
		(
			f"Messpunkte: {len(y_data)}; "
			f"Mittelwert: {_number(float(np.mean(y_data)))} {spec.y_unit}".rstrip()
		),
	]
	if spec.confidence is not None:
		parts[-1] += f"; Confidence: {spec.confidence:.2f}"
	return " | ".join(parts)


def _write_analysis(
	book: xw.Book,
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> None:
	_set_text_name(book, ANALYSIS_NAME, _analysis_label(frame, spec))


def show_help() -> None:
	def action(book: xw.Book) -> None:
		help_text = (
			f"{PRODUCT_NAME}\n\n"
			"- Analysiert die ausgewählte Excel-Tabelle oder den Zellbereich.\n"
			"- Erkennt physikalische Größen, Einheiten und potenzielle Beziehungen.\n"
			"- Erstellt ein natives Excel-Diagramm im geöffneten Workbook.\n"
			"- Zeigt die aktuelle Analyse direkt im Ribbon an.\n"
			"- Die KI wird nur bei mehrdeutigen Fällen verwendet.\n"
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
		use_ai_default = "Ja" if config["endpoint"] and config["model"] else "Nein"
		use_ai = book.app.api.InputBox(
			Prompt="KI bei mehrdeutigen Daten für die Analyse verwenden? Ja oder Nein:",
			Title=f"{PRODUCT_NAME} – KI-Einstellungen",
			Default=use_ai_default,
			Type=2,
		)
		if use_ai is False:
			return
		if str(use_ai).strip().lower() not in {"ja", "nein", "yes", "no"}:
			raise ValueError("Bitte für die KI-Nutzung „Ja“ oder „Nein“ eingeben.")
		if str(use_ai).strip().lower() in {"ja", "yes"} and not (config["endpoint"] and config["model"]):
			raise AIServiceError("Für die KI-Nutzung müssen Endpoint und Modell gesetzt sein.")
		path = save_config(config)
		_set_text_name(book, f"{SETTING_PREFIX}UseAI", "Ja" if str(use_ai).strip().lower() in {"ja", "yes"} else "Nein")
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


def _remove_generated_charts(book: xw.Book) -> None:
	for sheet in book.sheets:
		charts = sheet.api.ChartObjects()
		for index in range(int(charts.Count), 0, -1):
			chart_object = charts.Item(index)
			if str(chart_object.Name) == "PLVS_ULTRA_Chart":
				chart_object.Delete()


def _add_native_chart(
	book: xw.Book,
	sheet: xw.Sheet,
	data_range: xw.Range,
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> None:
	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	if int(valid.sum()) < 2 or np.unique(x_values[valid]).size < 2:
		raise ValueError("Mindestens zwei gültige Messwertpaare mit verschiedenen x-Werten sind erforderlich.")
	degree = {"linear": 1, "quadratic": 2, "cubic": 3}.get(spec.trendline)
	if spec.trendline != "none":
		if degree is None:
			raise ValueError(f"Unbekannte Trendlinie: {spec.trendline}")
		if np.unique(x_values[valid]).size <= degree:
			raise ValueError(f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} unterschiedliche x-Werte.")
	_remove_generated_charts(book)
	used = sheet.used_range.api
	left = float(used.Left) + float(used.Width) + 24
	top = float(used.Top)
	chart_object = sheet.api.ChartObjects().Add(left, top, 500, 300)
	chart_object.Name = "PLVS_ULTRA_Chart"
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
	for axis_type, label, values, minimum, maximum, quantity in (
		(1, spec.x_label, x_values, spec.x_min, spec.x_max, spec.x_quantity),
		(2, spec.y_label, y_values, spec.y_min, spec.y_max, spec.y_quantity),
	):
		axis = chart.Axes(axis_type)
		axis.HasTitle = True
		axis.AxisTitle.Text = label
		valid_values = values[np.isfinite(values)]
		is_zero = spec.origin is True or spec.origin == "ja" or (
			spec.origin == "auto" and quantity != "Temperatur" and valid_values.min() >= 0
		)
		axis_minimum, axis_maximum = _axis_bounds(valid_values, minimum, maximum, is_zero)
		axis.MinimumScale = axis_minimum
		axis.MaximumScale = axis_maximum
	if spec.trendline != "none":
		trendlines = series.Trendlines()
		if degree == 1:
			trendlines.Add(Type=-4132, Name="Lineare Ausgleichsgerade")
		else:
			trendlines.Add(Type=3, Order=degree, Name=f"Polynomfit Grad {degree}")
	_set_text_name(book, CHART_SHEET_NAME, sheet.name)


def create_chart() -> None:
	def action(book: xw.Book) -> None:
		data_sheet, data_range = _source_for_chart(book)
		frame = _frame_from_range(data_range)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Die ausgewählte Datenquelle benötigt mindestens zwei numerische Messspalten.")
		spec, _ = _resolve_spec(book, frame, columns, allow_ai=True)
		_add_native_chart(book, data_sheet, data_range, frame, spec)
		_write_analysis(book, frame, spec)
		data_sheet.activate()
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
