"""PLVS ULTRA Graphs Ribbon-Aktionen über xlwings und Excel COM."""

from __future__ import annotations

import math
import json
import os
from dataclasses import asdict
from typing import Any

import numpy as np
import pandas as pd
import xlwings as xw

from ai import (
    AIServiceError,
    analyze_with_ai,
    reanalyze_with_user_input,
    test_ai_connection,
)
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
CHART_STATE_PREFIX = "_PLVS_ULTRA_State_"
LAST_CHART_NAME = "_PLVS_ULTRA_LastChart"
PENDING_CHOICE_PREFIX = "_PLVS_ULTRA_PendingChoice"
PENDING_CHOICE_COUNT = f"{PENDING_CHOICE_PREFIX}_Count"
PENDING_CHOICE_LABEL = f"{PENDING_CHOICE_PREFIX}_Label"
PENDING_CHOICE_STATE = f"{PENDING_CHOICE_PREFIX}_State"
PENDING_CHOICE_ITEM_PREFIX = f"{PENDING_CHOICE_PREFIX}_Item_"
_TRENDLINE_TYPES = {
	"linear": -4132,
	"exponential": 5,
	"logarithmic": -4133,
	"power": 4,
	"quadratic": 3,
	"cubic": 3,
	"polynomial": 3,
	"moving_average": 6,
}
_TRENDLINE_CHART_TYPES = {-4169, 4, 51, 57, 65, 74, 75}
_MOVING_AVERAGE_PERIODS = (2, 3, 4, 5, 6, 7, 10, 12, 20)


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


def _read_text_name(book: xw.Book, name: str) -> str | None:
	item = _name_item(book, name)
	if item is None:
		return None
	value = book.app.api.Evaluate(str(item.RefersTo))
	return str(value)


def _clear_pending_choice(book: xw.Book) -> None:
	old_count = int(_read_text_name(book, PENDING_CHOICE_COUNT) or "0")
	for name in (
		PENDING_CHOICE_COUNT,
		PENDING_CHOICE_LABEL,
		PENDING_CHOICE_STATE,
		*(f"{PENDING_CHOICE_ITEM_PREFIX}{index}" for index in range(old_count)),
	):
		item = _name_item(book, name)
		if item is not None:
			item.Delete()


def _set_pending_choice(
	book: xw.Book,
	kind: str,
	label: str,
	options: list[str],
	state: dict[str, Any],
) -> None:
	if not options:
		raise ValueError("Es wurden keine anklickbaren Auswahlmöglichkeiten bereitgestellt.")
	old_count = int(_read_text_name(book, PENDING_CHOICE_COUNT) or "0")
	_set_text_name(book, PENDING_CHOICE_STATE, json.dumps(state, ensure_ascii=False, allow_nan=False))
	_set_text_name(book, PENDING_CHOICE_LABEL, f"{kind}|{label}")
	_set_text_name(book, PENDING_CHOICE_COUNT, str(len(options)))
	for index, option in enumerate(options):
		_set_text_name(book, f"{PENDING_CHOICE_ITEM_PREFIX}{index}", option)
	for index in range(len(options), old_count):
		item = _name_item(book, f"{PENDING_CHOICE_ITEM_PREFIX}{index}")
		if item is not None:
			item.Delete()
	_write_status(book, label)


def _source_from_range(
	sheet: xw.Sheet,
	data_range: xw.Range,
	label: str,
) -> tuple[xw.Sheet, xw.Range, str] | None:
	if _chartable_frame(data_range) is None:
		return None
	return sheet, data_range, label


def _chartable_frame(
	data_range: xw.Range,
) -> tuple[pd.DataFrame, list[str]] | None:
	try:
		frame = _frame_from_range(data_range)
	except ValueError:
		return None
	columns = get_numeric_columns(frame)
	if not columns:
		return None
	return frame, columns


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
		sheet, data_range, _, _ = current
		key = _range_key(sheet, data_range)
		if key in seen:
			for index, (source_sheet, source_range, label) in enumerate(sources):
				if _range_key(source_sheet, source_range) == key:
					sources[index] = (source_sheet, source_range, label)
					break
		else:
			sources.append((
				sheet,
				data_range,
				_source_label(sheet, data_range),
			))
	return sources


def _current_data_range(
	book: xw.Book,
) -> tuple[xw.Sheet, xw.Range, pd.DataFrame, list[str]] | None:
	selected = book.app.api.Selection
	try:
		if int(selected.Areas.Count) != 1 or int(selected.CountLarge) < 2:
			return None
		if str(selected.Worksheet.Parent.FullName).casefold() != str(book.fullname).casefold():
			return None
		sheet = book.sheets[str(selected.Worksheet.Name)]
		data_range = sheet.range(selected.Address)
	except (AttributeError, TypeError, ValueError):
		return None
	chartable = _chartable_frame(data_range)
	if chartable is None:
		return None
	frame, columns = chartable
	return sheet, data_range, frame, columns


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
		raise ValueError("Der ausgewählte Bereich benötigt Überschriften und mindestens eine numerische y-Messspalte.")
	return source


def _choose_data_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range] | None:
	sources = _discover_data_sources(book)
	if not sources:
		return _select_manual_source(book)
	selection = _show_data_source_dialog(book, [label for _, _, label in sources])
	if selection is None:
		return None
	if selection == len(sources):
		return _select_manual_source(book)
	sheet, data_range, label = sources[selection]
	_store_source(book, sheet, data_range, label)
	return sheet, data_range


def _show_data_source_dialog(book: xw.Book, labels: list[str]) -> int | None:
	import tkinter as tk

	root = tk.Tk()
	root.withdraw()
	dialog = tk.Toplevel(root)
	dialog.title(f"{PRODUCT_NAME} - Datenquelle auswählen")
	dialog.resizable(False, False)
	dialog.attributes("-topmost", True)
	dialog.protocol("WM_DELETE_WINDOW", lambda: finish(None))

	result: list[int | None] = [None]

	def finish(selection: int | None) -> None:
		result[0] = selection
		dialog.destroy()

	tk.Label(
		dialog,
		text="Datenquelle auswählen",
		font=("Segoe UI", 12, "bold"),
	).pack(padx=18, pady=(16, 10))
	for index, label in enumerate(labels):
		tk.Button(
			dialog,
			text=label,
			anchor="w",
			command=lambda selected=index: finish(selected),
		).pack(fill="x", padx=18, pady=3)
	tk.Button(
		dialog,
		text="Manuell Zellbereich auswählen …",
		anchor="w",
		command=lambda: finish(len(labels)),
	).pack(fill="x", padx=18, pady=(8, 3))
	tk.Button(
		dialog,
		text="Abbrechen",
		command=lambda: finish(None),
	).pack(anchor="e", padx=18, pady=(8, 14))
	dialog.update_idletasks()
	width, height = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
	try:
		left = int(book.app.api.Left)
		top = int(book.app.api.Top)
		excel_width = int(book.app.api.Width)
		excel_height = int(book.app.api.Height)
		x = left + max((excel_width - width) // 2, 0)
		y = top + max((excel_height - height) // 2, 0)
	except (AttributeError, TypeError, ValueError):
		x = (dialog.winfo_screenwidth() - width) // 2
		y = (dialog.winfo_screenheight() - height) // 2
	dialog.geometry(f"+{x}+{y}")
	dialog.grab_set()
	dialog.focus_force()
	root.wait_window(dialog)
	root.destroy()
	return result[0]


def _select_manual_source(book: xw.Book) -> tuple[xw.Sheet, xw.Range]:
	sheet, data_range, label = _manual_source(book)
	_store_source(book, sheet, data_range, label)
	return sheet, data_range


def _source_for_chart(
	book: xw.Book,
) -> tuple[xw.Sheet, xw.Range, pd.DataFrame, list[str]] | None:
	current = _current_data_range(book)
	if current is not None:
		sheet, data_range, frame, columns = current
		label = _source_label(sheet, data_range)
		_store_source(book, sheet, data_range, label)
		return sheet, data_range, frame, columns
	source = _choose_data_source(book)
	if source is None:
		return None
	sheet, data_range = source
	frame = _frame_from_range(data_range)
	return sheet, data_range, frame, get_numeric_columns(frame)


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
	x_is_numeric = np.any(np.isfinite(x_values))
	y_is_numeric = np.any(np.isfinite(y_values))
	if not y_is_numeric:
		valid = frame[spec.y_column].notna().to_numpy().copy()
		if x_is_numeric:
			valid &= np.isfinite(x_values)
			x_data = x_values[valid]
		else:
			valid &= frame[spec.x_column].notna().to_numpy()
			x_data = pd.factorize(frame.loc[valid, spec.x_column], sort=False)[0].astype(float) + 1
		y_data = np.array([], dtype=float)
	elif x_is_numeric:
		valid = np.isfinite(x_values) & np.isfinite(y_values)
		x_data = x_values[valid]
	else:
		valid = frame[spec.x_column].notna().to_numpy() & np.isfinite(y_values)
		x_data = pd.factorize(frame.loc[valid, spec.x_column], sort=False)[0].astype(float) + 1
	y_data = y_values[valid]
	fit_labels = {
		"none": "kein",
		"linear": "linear",
		"quadratic": "quadratisch",
		"cubic": "kubisch",
		"exponential": "exponentiell",
		"logarithmic": "logarithmisch",
		"power": "Potenz",
		"polynomial": f"polynomisch (Grad {spec.polynomial_degree})",
		"moving_average": f"gleitender Durchschnitt (Periode {spec.moving_average_period})",
	}
	fit_label = fit_labels.get(spec.trendline, spec.trendline)
	chart_labels = {
		"scatter": "Punktdiagramm (XY)",
		"scatter_lines": "Punktdiagramm (XY) mit Linien",
		"line": "Liniendiagramm",
		"line_scatter": "Punkt- und Liniendiagramm",
		"bar": "Balkendiagramm",
		"column": "Säulendiagramm",
	}
	chart_label = chart_labels.get(spec.chart_type, spec.chart_type)
	mean_y = float(np.mean(y_data)) if y_data.size else None
	median_y = float(np.median(y_data)) if y_data.size else None
	min_y = float(np.min(y_data)) if y_data.size else None
	max_y = float(np.max(y_data)) if y_data.size else None
	std_y = float(np.std(y_data, ddof=1)) if y_data.size > 1 else None
	correlation = None
	if (
		x_is_numeric
		and x_data.size > 1
		and y_data.size > 1
		and np.std(x_data)
		and np.std(y_data)
	):
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
		f"Messpunkte: {int(valid.sum())} von {len(frame)}; Mittelwert y: "
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


def _chart_state_name(chart_name: str) -> str:
	return f"{CHART_STATE_PREFIX}{chart_name}"


def _evaluated_text(value: Any) -> str:
	while isinstance(value, (tuple, list)) and len(value) == 1:
		value = value[0]
	if not isinstance(value, str):
		raise ValueError("Der gespeicherte Diagrammzustand ist kein Text.")
	return value


def _save_chart_state(
	book: xw.Book,
	chart_name: str,
	sheet_name: str,
	address: str,
	spec: ChartSpecification,
	settings: dict[str, Any] | None = None,
) -> None:
	spec_fields = (
		"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
		"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
		"trendline", "polynomial_degree", "moving_average_period",
		"x_error_column", "y_error_column",
		"independent_variable", "dependent_variable", "show_points",
		"connect_points", "confidence", "reason",
	)
	saved_spec = {field: getattr(spec, field) for field in spec_fields}
	if saved_spec["trendline"] in {"quadratic", "cubic"}:
		saved_spec["polynomial_degree"] = 2 if saved_spec["trendline"] == "quadratic" else 3
	state = {
		"sheet": sheet_name,
		"address": address,
		"spec": saved_spec,
		"settings": settings or {},
	}
	_set_text_name(
		book,
		_chart_state_name(chart_name),
		json.dumps(state, ensure_ascii=False, separators=(",", ":")),
	)
	_set_text_name(book, LAST_CHART_NAME, chart_name)


def _read_chart_state(book: xw.Book, chart_name: str) -> dict[str, Any]:
	item = _name_item(book, _chart_state_name(chart_name))
	if item is None:
		raise ValueError(
			"Für dieses Diagramm ist kein PLVS-Bearbeitungszustand gespeichert. "
			"Erstelle das Diagramm erneut mit PLVS ULTRA Graphs."
		)
	try:
		value = book.app.api.Evaluate(item.Name)
		state = json.loads(_evaluated_text(value))
	except (ValueError, TypeError) as exc:
		raise ValueError("Der gespeicherte Diagrammzustand ist ungültig.") from exc
	if not isinstance(state, dict) or not isinstance(state.get("spec"), dict):
		raise ValueError("Der gespeicherte Diagrammzustand ist unvollständig.")
	state["spec"].setdefault(
		"polynomial_degree",
		{"quadratic": 2, "cubic": 3}.get(state["spec"].get("trendline"), 2),
	)
	state["spec"].setdefault("moving_average_period", 3)
	return state


def _chart_by_name(
	book: xw.Book,
	chart_name: str,
	sheet_name: str | None = None,
) -> tuple[xw.Sheet, Any]:
	sheets = (book.sheets[sheet_name],) if sheet_name is not None else book.sheets
	for sheet in sheets:
		chart_objects = sheet.api.ChartObjects()
		for index in range(1, int(chart_objects.Count) + 1):
			chart_object = chart_objects.Item(index)
			if str(chart_object.Name) == chart_name:
				return sheet, chart_object.Chart
	raise ValueError("Das ausgewählte PLVS-Diagramm ist nicht mehr vorhanden.")


def _selected_chart(
	book: xw.Book,
	target_sheet: str | None = None,
	target_chart: str | None = None,
) -> tuple[str, xw.Sheet, Any, dict[str, Any]]:
	if target_sheet is not None or target_chart is not None:
		if not target_sheet or not target_chart:
			raise ValueError("Das Bearbeitungsziel des ausgewählten Diagramms ist unvollständig.")
		chart_name = target_chart
		sheet, chart = _chart_by_name(book, chart_name, target_sheet)
	else:
		chart_name = ""
		try:
			active_chart = book.app.api.ActiveChart
			if active_chart is not None:
				chart_name = str(active_chart.Parent.Name)
		except Exception:
			chart_name = ""
		if not chart_name:
			last_chart = _name_item(book, LAST_CHART_NAME)
			if last_chart is not None:
				value = book.app.api.Evaluate(last_chart.Name)
				chart_name = _evaluated_text(value)
		if not chart_name:
			raise ValueError("Wähle zuerst ein PLVS-Diagramm aus oder erstelle eines.")
		sheet, chart = _chart_by_name(book, chart_name)
	return chart_name, sheet, chart, _read_chart_state(book, chart_name)


def _chart_edit_book(book: xw.Book, target_workbook: str | None) -> xw.Book:
	if not target_workbook:
		return book
	target_path = os.path.normcase(os.path.abspath(target_workbook))
	for candidate in book.app.books:
		if os.path.normcase(os.path.abspath(candidate.fullname)) == target_path:
			return candidate
	raise ValueError("Die Arbeitsmappe des ausgewählten Diagramms ist nicht mehr geöffnet.")


def _frame_for_state(book: xw.Book, state: dict[str, Any]) -> tuple[xw.Sheet, xw.Range, pd.DataFrame]:
	sheet_name = state.get("sheet")
	address = state.get("address")
	if not isinstance(sheet_name, str) or not isinstance(address, str):
		raise ValueError("Im Diagrammzustand fehlt die Datenquelle.")
	if sheet_name not in book.sheet_names:
		raise ValueError(f"Das Quelldatenblatt {sheet_name!r} ist nicht mehr vorhanden.")
	sheet = book.sheets[sheet_name]
	data_range = sheet.range(address)
	return sheet, data_range, _frame_from_range(data_range)


def _remove_trendlines(series: Any) -> None:
	trendlines = series.Trendlines()
	for index in range(int(trendlines.Count), 0, -1):
		trendlines.Item(index).Delete()


def _trendline_data(
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> tuple[np.ndarray, np.ndarray]:
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	raw_x = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	x_is_numeric = bool(np.any(np.isfinite(raw_x)))
	if spec.chart_type in {"bar", "column", "line", "line_scatter"} or not x_is_numeric:
		x_values = np.arange(1, len(frame) + 1, dtype=float)
		valid_categories = frame[spec.x_column].notna().to_numpy()
	else:
		x_values = raw_x
		valid_categories = np.ones(len(frame), dtype=bool)
	valid = np.isfinite(x_values) & np.isfinite(y_values) & valid_categories
	return x_values[valid], y_values[valid]


def _available_trendlines(
	x_values: np.ndarray,
	y_values: np.ndarray,
	chart_type: int,
) -> tuple[tuple[str, ...], tuple[int, ...], tuple[int, ...]]:
	if chart_type not in _TRENDLINE_CHART_TYPES:
		return (), (), ()
	x = np.asarray(x_values, dtype=float)
	y = np.asarray(y_values, dtype=float)
	valid = np.isfinite(x) & np.isfinite(y)
	x = x[valid]
	y = y[valid]
	distinct_x = int(np.unique(x).size)
	count = int(x.size)
	options = ["none"]
	if distinct_x >= 2:
		options.append("linear")
	if count >= 3 and distinct_x >= 2 and np.all(y > 0):
		options.append("exponential")
	if distinct_x >= 2 and np.all(x > 0):
		options.append("logarithmic")
	if distinct_x >= 2 and np.all(x > 0) and np.all(y > 0):
		options.append("power")
	degrees = tuple(range(2, min(6, distinct_x - 1, count - 1) + 1))
	if degrees:
		options.append("polynomial")
	periods = tuple(period for period in _MOVING_AVERAGE_PERIODS if period < count)
	if periods:
		options.append("moving_average")
	return tuple(options), degrees, periods


def _actual_trendline_state(series: Any) -> tuple[str, int | None, int | None]:
	trendlines = series.Trendlines()
	count = int(trendlines.Count)
	if count == 0:
		return "none", None, None
	trendline = trendlines.Item(count)
	trendline_type = int(trendline.Type)
	if trendline_type == 3:
		return "polynomial", int(trendline.Order), None
	if trendline_type == 6:
		return "moving_average", None, int(trendline.Period)
	return {
		-4132: "linear",
		5: "exponential",
		-4133: "logarithmic",
		4: "power",
	}.get(trendline_type, f"unsupported:{trendline_type}"), None, None


def _set_trendline(
	series: Any,
	trendline: str,
	polynomial_degree: int = 2,
	moving_average_period: int = 3,
	x_values: np.ndarray | None = None,
	y_values: np.ndarray | None = None,
	chart_type: int | None = None,
) -> None:
	if trendline not in {"none", *_TRENDLINE_TYPES}:
		raise ValueError(f"Unbekannter Trendlinientyp: {trendline}")
	if trendline == "none":
		_remove_trendlines(series)
		if int(series.Trendlines().Count) != 0:
			raise RuntimeError("Excel konnte die vorhandene Trendlinie nicht entfernen.")
		return
	if chart_type is None:
		raise ValueError("Der Diagrammtyp muss vor dem Anwenden der Trendlinie geprüft werden.")
	if x_values is None or y_values is None:
		raise ValueError("Die Trendlinie benötigt die numerischen Daten des Diagramms.")
	options, degrees, periods = _available_trendlines(x_values, y_values, chart_type)
	aliases = {"quadratic": "polynomial", "cubic": "polynomial"}
	option = aliases.get(trendline, trendline)
	if option not in options:
		raise ValueError(
			f"Die Trendlinie {trendline!r} ist für Diagrammtyp und Daten nicht anwendbar."
		)
	if option == "polynomial":
		if trendline == "quadratic":
			polynomial_degree = 2
		elif trendline == "cubic":
			polynomial_degree = 3
		if polynomial_degree not in degrees:
			raise ValueError(
				f"Polynomgrad {polynomial_degree} ist mit den vorhandenen Messpunkten nicht möglich."
			)
	if option == "moving_average" and moving_average_period not in periods:
		raise ValueError(
			f"Periode {moving_average_period} ist mit den vorhandenen Messpunkten nicht möglich."
		)
	old_count = int(series.Trendlines().Count)
	kwargs: dict[str, Any] = {
		"Type": _TRENDLINE_TYPES[trendline],
		"Name": (
			f"Polynomfit Grad {polynomial_degree}"
			if option == "polynomial"
			else f"Gleitender Durchschnitt ({moving_average_period})"
			if option == "moving_average"
			else f"{trendline.title()} Fit"
		),
	}
	if option == "polynomial":
		kwargs["Order"] = polynomial_degree
	elif option == "moving_average":
		kwargs["Period"] = moving_average_period
	added = None
	try:
		added = series.Trendlines().Add(**kwargs)
		applied_type, applied_degree, applied_period = _actual_trendline_state(series)
		expected_type = option
		if (
			applied_type != expected_type
			or (option == "polynomial" and applied_degree != polynomial_degree)
			or (option == "moving_average" and applied_period != moving_average_period)
		):
			raise RuntimeError(
				"Excel hat den gewünschten Trendlinientyp, Grad oder Zeitraum nicht übernommen."
			)
	except Exception as exc:
		if added is not None:
			try:
				added.Delete()
			except Exception as rollback_error:
				raise RuntimeError(
					f"Excel konnte die Trendlinie nicht anwenden oder zurücksetzen: {rollback_error}"
				) from exc
		raise RuntimeError(f"Excel konnte die Trendlinie nicht anwenden: {exc}") from exc
	trendlines = series.Trendlines()
	for index in range(old_count, 0, -1):
		trendlines.Item(index).Delete()
	applied_type, applied_degree, applied_period = _actual_trendline_state(series)
	if (
		applied_type != option
		or (option == "polynomial" and applied_degree != polynomial_degree)
		or (option == "moving_average" and applied_period != moving_average_period)
	):
		raise RuntimeError("Excel hat den endgültigen Trendlinienzustand nicht beibehalten.")


def _chart_type_code(
	chart_type: str,
	show_points: bool,
	connect_points: bool,
	categorical: bool = False,
) -> int:
	if chart_type == "scatter":
		return -4169
	if chart_type == "scatter_lines":
		return 74 if show_points else 75
	if chart_type == "line":
		return 4
	if chart_type == "line_scatter":
		return 65
	if chart_type == "bar":
		return 57
	if chart_type == "column":
		return 51
	raise ValueError(f"Unbekannte Diagrammart: {chart_type}")


def _gridline_axis_types(chart_type: int, x_is_numeric: bool) -> tuple[int, ...]:
	if chart_type in {-4169, 65, 74, 75}:
		return (1, 2)
	if chart_type == 4:
		return (1, 2) if x_is_numeric else (2,)
	if chart_type in {51, 57}:
		# The category-axis gridlines rotate with the bar/column chart orientation.
		return (1,)
	raise ValueError(f"Gitternetzlinien für den Excel-Diagrammtyp {chart_type} werden nicht unterstützt.")


def _set_chart_gridlines(chart: Any, enabled: bool, x_is_numeric: bool) -> None:
	chart_type = int(chart.ChartType)
	axes = {axis_type: chart.Axes(axis_type) for axis_type in (1, 2)}
	for axis in axes.values():
		axis.HasMajorGridlines = False
		axis.HasMinorGridlines = False
	if enabled:
		for axis_type in _gridline_axis_types(chart_type, x_is_numeric):
			axes[axis_type].HasMajorGridlines = True

	expected_major_axes = set(_gridline_axis_types(chart_type, x_is_numeric)) if enabled else set()
	actual_major_axes = {
		axis_type
		for axis_type, axis in axes.items()
		if bool(axis.HasMajorGridlines)
	}
	if actual_major_axes != expected_major_axes or any(
		bool(axis.HasMinorGridlines) for axis in axes.values()
	):
		raise RuntimeError("Excel hat den gewünschten Gitternetzlinienzustand nicht übernommen.")


def _set_creation_gridlines(chart: Any, x_is_numeric: bool) -> None:
	_set_chart_gridlines(chart, True, x_is_numeric)


def _set_axis_title(
	chart: Any,
	spec: ChartSpecification,
	frame: pd.DataFrame,
	logical_axis: str,
) -> None:
	column = spec.x_column if logical_axis == "x" else spec.y_column
	label = spec.x_label if logical_axis == "x" else spec.y_label
	values = pd.to_numeric(frame[column], errors="coerce").to_numpy(dtype=float)
	is_categorical = not np.any(np.isfinite(values))
	if spec.chart_type in {"line", "line_scatter"}:
		axis_type = 1 if logical_axis == "x" else 2
	elif spec.chart_type in {"bar", "column"}:
		axis_type = 1 if is_categorical else 2
	else:
		axis_type = 1 if logical_axis == "x" else 2
	axis = chart.Axes(axis_type)
	axis.HasTitle = True
	axis.AxisTitle.Text = label


def _stored_color(value: Any) -> int:
	if isinstance(value, str) and value.startswith("#") and len(value) == 7:
		red, green, blue = (int(value[index:index + 2], 16) for index in (1, 3, 5))
		return red | (green << 8) | (blue << 16)
	if not isinstance(value, int) or isinstance(value, bool) or not 0 <= value <= 0xFFFFFF:
		raise ValueError("Der gespeicherte Diagrammfarbwert ist ungültig.")
	return value


def _set_series_line_color(series: Any, color: int, chart_type: int) -> None:
	if chart_type == 4:
		series.Border.Color = color
	else:
		series.Format.Line.ForeColor.RGB = color


def _restore_saved_series_format(series: Any, spec: ChartSpecification, settings: dict[str, Any]) -> None:
	marker_chart = spec.chart_type in {"scatter", "scatter_lines", "line", "line_scatter"}
	filled_chart = spec.chart_type in {"bar", "column"}
	if "series_color" in settings and filled_chart:
		rgb = _stored_color(settings["series_color"])
		series.Format.Fill.ForeColor.RGB = rgb
	elif "series_color" in settings and marker_chart:
		rgb = _stored_color(settings["series_color"])
		_set_series_line_color(series, rgb, 4 if spec.chart_type == "line" else 0)
		if spec.chart_type != "line":
			series.MarkerForegroundColor = rgb
			series.MarkerBackgroundColor = rgb
	if "line_color" in settings and marker_chart:
		_set_series_line_color(
			series,
			_stored_color(settings["line_color"]),
			4 if spec.chart_type == "line" else 0,
		)
	if "line_width" in settings:
		line_width = float(settings["line_width"])
		if math.isfinite(line_width) and 0.75 <= line_width <= 6:
			series.Format.Line.Weight = line_width
	if "point_size" in settings and marker_chart and spec.chart_type != "line":
		series.MarkerSize = float(settings["point_size"])


def _capture_series_format(series: Any, chart_type: int) -> dict[str, Any]:
	state: dict[str, Any] = {}
	if bool(series.Format.Line.Visible):
		line_color = int(series.Format.Line.ForeColor.RGB)
		if 0 <= line_color <= 0xFFFFFF:
			state["line_color"] = line_color
	line_width = float(series.Format.Line.Weight)
	if math.isfinite(line_width) and 0.75 <= line_width <= 6:
		state["line_width"] = line_width
	if chart_type in {51, 57}:
		fill_color = int(series.Format.Fill.ForeColor.RGB)
		if 0 <= fill_color <= 0xFFFFFF:
			state["fill_color"] = fill_color
	if chart_type in {-4169, 4, 65, 74, 75}:
		marker_foreground = int(series.MarkerForegroundColor)
		marker_background = int(series.MarkerBackgroundColor)
		if 0 <= marker_foreground <= 0xFFFFFF:
			state["marker_foreground"] = marker_foreground
		if 0 <= marker_background <= 0xFFFFFF:
			state["marker_background"] = marker_background
		if "line_color" not in state:
			marker_color = next(
				(
					color for color in (marker_foreground, marker_background)
					if 0 <= color <= 0xFFFFFF
				),
				None,
			)
			if marker_color is not None:
				state["line_color"] = marker_color
		marker_size = int(series.MarkerSize)
		if 2 <= marker_size <= 72:
			state["marker_size"] = marker_size
		state["marker_style"] = int(series.MarkerStyle)
	return state


def _apply_series_colors(series: Any, old_format: dict[str, Any], chart_type: int) -> None:
	if chart_type in {51, 57}:
		fill_color = old_format.get("fill_color", old_format.get("line_color"))
		if fill_color is not None:
			series.Format.Fill.ForeColor.RGB = fill_color
		if "line_color" in old_format:
			series.Format.Line.ForeColor.RGB = old_format["line_color"]
	elif chart_type in {-4169, 4, 65, 74, 75}:
		if "line_color" in old_format:
			_set_series_line_color(series, old_format["line_color"], chart_type)
		if chart_type != 4:
			if "marker_foreground" in old_format:
				series.MarkerForegroundColor = old_format["marker_foreground"]
			if "marker_background" in old_format:
				series.MarkerBackgroundColor = old_format["marker_background"]


def _apply_series_format(series: Any, old_format: dict[str, Any], chart_type: int) -> None:
	_apply_series_colors(series, old_format, chart_type)
	line_width = old_format.get("line_width")
	if chart_type in {51, 57}:
		if line_width is not None:
			series.Format.Line.Weight = line_width
	elif chart_type in {-4169, 4, 65, 74, 75}:
		if line_width is not None:
			series.Format.Line.Weight = line_width
		if chart_type != 4:
			if "marker_size" in old_format:
				series.MarkerSize = old_format["marker_size"]
			if "marker_style" in old_format:
				series.MarkerStyle = old_format["marker_style"]


def _sync_chart_state_from_excel(chart: Any, spec: ChartSpecification, settings: dict[str, Any]) -> None:
	chart_type = int(chart.ChartType)
	series = chart.SeriesCollection(1)
	if chart_type == -4169:
		spec.chart_type = "scatter"
	elif chart_type in {74, 75}:
		spec.chart_type = "scatter_lines"
	elif chart_type == 4:
		spec.chart_type = "line"
	elif chart_type == 65:
		spec.chart_type = "line_scatter"
	elif chart_type == 51:
		spec.chart_type = "column"
	elif chart_type == 57:
		spec.chart_type = "bar"
	else:
		raise ValueError(f"Der aktuelle Excel-Diagrammtyp {chart_type} wird nicht unterstützt.")
	if chart_type in {-4169, 4, 65, 74, 75}:
		spec.show_points = int(series.MarkerStyle) != -4142
		spec.connect_points = bool(series.Format.Line.Visible)
	settings["show_points"] = spec.show_points
	settings["connect_points"] = spec.connect_points
	settings["legend"] = bool(chart.HasLegend)
	if chart_type in {51, 57}:
		settings["series_color"] = int(series.Format.Fill.ForeColor.RGB)
		settings.pop("point_color", None)
	else:
		line_color = int(series.Format.Line.ForeColor.RGB)
		if bool(series.Format.Line.Visible) and 0 <= line_color <= 0xFFFFFF:
			settings["line_color"] = line_color
		elif "line_color" not in settings:
			settings.pop("line_color", None)
		settings.pop("point_color", None)
	raw_line_width = series.Format.Line.Weight
	try:
		line_width = float(raw_line_width)
	except (TypeError, ValueError, OverflowError):
		settings.pop("line_width", None)
	else:
		if math.isfinite(line_width) and 0.75 <= line_width <= 6:
			settings["line_width"] = line_width
		else:
			settings.pop("line_width", None)
	if chart_type in {-4169, 4, 65, 74, 75}:
		settings["point_size"] = int(series.MarkerSize)
		trendline, polynomial_degree, moving_average_period = _actual_trendline_state(series)
		spec.trendline = trendline
		if polynomial_degree is not None:
			spec.polynomial_degree = polynomial_degree
		if moving_average_period is not None:
			spec.moving_average_period = moving_average_period


def _selection(index: int | None, options: tuple[Any, ...], setting: str) -> Any:
	if not isinstance(index, int) or isinstance(index, bool) or not 0 <= index < len(options):
		raise ValueError(f"Bitte einen gültigen Eintrag für {setting} im Ribbon auswählen.")
	return options[index]


def _selected_text_element(chart: Any, spec: ChartSpecification, frame: pd.DataFrame, settings: dict[str, Any]) -> Any:
	target = settings.get("text_target", "title")
	if target == "title":
		if not bool(chart.HasTitle):
			raise ValueError("Der Diagrammtitel ist ausgeblendet. Blende ihn zuerst ein.")
		return chart.ChartTitle
	if target == "x_axis":
		column, label = spec.x_column, spec.x_label
	elif target == "y_axis":
		column, label = spec.y_column, spec.y_label
	elif target == "legend":
		if not bool(chart.HasLegend):
			raise ValueError("Die Legende ist ausgeblendet. Blende sie zuerst ein.")
		return chart.Legend
	else:
		raise ValueError("Das ausgewählte Textelement ist ungültig.")
	values = pd.to_numeric(frame[column], errors="coerce").to_numpy(dtype=float)
	is_categorical = not np.any(np.isfinite(values))
	if spec.chart_type in {"bar", "column"}:
		axis_type = 1 if is_categorical else 2
	else:
		axis_type = 1 if target == "x_axis" else 2
	axis = chart.Axes(axis_type)
	if not bool(axis.HasTitle):
		axis.HasTitle = True
		axis.AxisTitle.Text = label
	return axis.AxisTitle


def _selected_color(value: int | None) -> int:
	if not isinstance(value, int) or isinstance(value, bool) or not 0 <= value <= 0xFFFFFF:
		raise ValueError("Excel hat keinen gültigen Farbwert aus der Farbpalette übergeben.")
	return value


def _ask_text(book: xw.Book, prompt: str, title: str, default: str = "") -> str | None:
	value = book.app.api.InputBox(
		Prompt=prompt,
		Title=f"{PRODUCT_NAME} – {title}",
		Default=default,
		Type=2,
	)
	if value is False or value is None:
		return None
	return str(value).strip()


def _source_reference(sheet: xw.Sheet, data_range: xw.Range, column_name: str) -> Any:
	frame = _frame_from_range(data_range)
	if column_name not in frame.columns:
		raise ValueError(f"Die Diagrammspalte {column_name!r} fehlt im gespeicherten Datenbereich.")
	first_row = int(data_range.api.Row) + 1
	last_row = int(data_range.api.Row) + int(data_range.api.Rows.Count) - 1
	column = int(data_range.api.Column) + int(frame.columns.get_loc(column_name))
	return sheet.range((first_row, column), (last_row, column)).api


def _clear_error_bars(series: Any) -> None:
	if not bool(series.HasErrorBars):
		return
	series.ErrorBars.Delete()


def _apply_error_bars(
	series: Any,
	sheet: xw.Sheet,
	data_range: xw.Range,
	spec: ChartSpecification,
) -> None:
	for column_name, direction in (
		(spec.x_error_column, -4168),
		(spec.y_error_column, 1),
	):
		if column_name is None:
			continue
		reference = _source_reference(sheet, data_range, column_name)
		series.ErrorBar(
			Direction=direction,
			Include=1,
			Type=-4114,
			Amount=reference,
			MinusValues=reference,
		)


def _swap_chart_axes(
	book: xw.Book,
	chart_name: str,
	chart: Any,
	state: dict[str, Any],
) -> ChartSpecification:
	spec = ChartSpecification(**state["spec"])
	sheet, data_range, frame = _frame_for_state(book, state)
	x_numeric = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	categorical = spec.chart_type in {"bar", "column"} or not np.any(np.isfinite(x_numeric))
	series = chart.SeriesCollection(1)
	series_format = {
		"fill_color": series.Format.Fill.ForeColor.RGB,
		"line_color": series.Format.Line.ForeColor.RGB,
		"line_width": series.Format.Line.Weight,
	}
	if spec.chart_type in {"scatter", "scatter_lines", "line", "line_scatter"}:
		series_format.update({
			"marker_foreground": series.MarkerForegroundColor,
			"marker_background": series.MarkerBackgroundColor,
			"marker_size": series.MarkerSize,
			"marker_style": series.MarkerStyle,
		})
	if categorical:
		if spec.chart_type == "bar":
			spec.chart_type = "column"
			chart.ChartType = 51
		else:
			spec.chart_type = "bar"
			chart.ChartType = 57
	else:
		series.XValues = _source_reference(sheet, data_range, spec.y_column)
		series.Values = _source_reference(sheet, data_range, spec.x_column)
		spec.chart_type = "scatter_lines" if spec.connect_points else "scatter"
		chart.ChartType = _chart_type_code(
			spec.chart_type,
			spec.show_points,
			spec.connect_points,
		)
	spec.x_column, spec.y_column = spec.y_column, spec.x_column
	spec.x_label, spec.y_label = spec.y_label, spec.x_label
	spec.x_unit, spec.y_unit = spec.y_unit, spec.x_unit
	spec.x_min, spec.y_min = spec.y_min, spec.x_min
	spec.x_max, spec.y_max = spec.y_max, spec.x_max
	spec.x_error_column, spec.y_error_column = spec.y_error_column, spec.x_error_column
	spec.independent_variable, spec.dependent_variable = (
		spec.dependent_variable,
		spec.independent_variable,
	)
	series = chart.SeriesCollection(1)
	if not categorical:
		for axis_type, minimum, maximum in (
			(1, spec.x_min, spec.x_max),
			(2, spec.y_min, spec.y_max),
		):
			axis = chart.Axes(axis_type)
			if minimum is None:
				axis.MinimumScaleIsAuto = True
			else:
				axis.MinimumScale = minimum
			if maximum is None:
				axis.MaximumScaleIsAuto = True
			else:
				axis.MaximumScale = maximum
	for key, color_format in (
		("fill_color", series.Format.Fill.ForeColor),
		("line_color", series.Format.Line.ForeColor),
	):
		color = series_format[key]
		if isinstance(color, int) and 0 <= color <= 0xFFFFFF:
			color_format.RGB = color
	line_width = series_format["line_width"]
	if isinstance(line_width, (int, float)) and 0 < line_width <= 10:
		series.Format.Line.Weight = line_width
	if not categorical and "marker_foreground" in series_format and spec.chart_type in {
		"scatter", "scatter_lines", "line", "line_scatter"
	}:
		series.MarkerForegroundColor = series_format["marker_foreground"]
		series.MarkerBackgroundColor = series_format["marker_background"]
		series.MarkerSize = series_format["marker_size"]
		series.MarkerStyle = series_format["marker_style"]
	_clear_error_bars(series)
	_apply_error_bars(series, sheet, data_range, spec)
	if not categorical:
		chart.ChartType = _chart_type_code(
			spec.chart_type,
			spec.show_points,
			spec.connect_points,
		)
		series = chart.SeriesCollection(1)
		for key, color_format in (
			("fill_color", series.Format.Fill.ForeColor),
			("line_color", series.Format.Line.ForeColor),
		):
			color = series_format[key]
			if isinstance(color, int) and 0 <= color <= 0xFFFFFF:
				color_format.RGB = color
		if isinstance(line_width, (int, float)) and 0 < line_width <= 10:
			series.Format.Line.Weight = line_width
		if "marker_foreground" in series_format:
			series.MarkerForegroundColor = series_format["marker_foreground"]
			series.MarkerBackgroundColor = series_format["marker_background"]
			series.MarkerSize = series_format["marker_size"]
		series.MarkerStyle = 8 if spec.show_points else -4142
		_restore_saved_series_format(series, spec, dict(state.get("settings", {})))
		if not spec.connect_points:
			series.Format.Line.Visible = False
			if spec.show_points and "marker_foreground" in series_format:
				series.MarkerForegroundColor = series_format["marker_foreground"]
				series.MarkerBackgroundColor = series_format["marker_background"]
	state_settings = dict(state.get("settings", {}))
	_restore_saved_series_format(chart.SeriesCollection(1), spec, state_settings)
	chart.HasTitle = True
	chart.ChartTitle.Text = f"{spec.y_label} in Abhängigkeit von {spec.x_label}"
	state_settings["title"] = str(chart.ChartTitle.Text)
	_set_axis_title(chart, spec, frame, "x")
	_set_axis_title(chart, spec, frame, "y")
	if isinstance(state_settings.get("gridlines_enabled"), bool):
		x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
		_set_chart_gridlines(
			chart,
			state_settings["gridlines_enabled"],
			bool(np.any(np.isfinite(x_values))),
		)
	_write_analysis(book, frame, spec, chart_name)
	_save_chart_state(
		book,
		chart_name,
		str(state["sheet"]),
		str(state["address"]),
		spec,
		state_settings,
	)
	_write_status(book, f"Achsen getauscht: X = {spec.x_label}; Y = {spec.y_label}.")
	return spec


def _edit_chart_settings(
	book: xw.Book,
	action: str,
	value: int | str | None = None,
	target_workbook: str | None = None,
	target_sheet: str | None = None,
	target_chart: str | None = None,
) -> None:
	book = _chart_edit_book(book, target_workbook)
	chart_name, _, chart, state = _selected_chart(book, target_sheet, target_chart)
	spec = ChartSpecification(**state["spec"])
	sheet, data_range, frame = _frame_for_state(book, state)
	series = chart.SeriesCollection(1)
	settings = dict(state.get("settings", {}))
	if action == "chart_type":
		spec.chart_type = _selection(
			value if isinstance(value, int) else None,
			("scatter", "scatter_lines", "line", "line_scatter", "bar", "column"),
			"Diagrammtyp",
		)
		x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
		if spec.chart_type in {"scatter", "scatter_lines"} and not np.any(np.isfinite(x_values)):
			raise ValueError("Scatter benötigt eine numerische X-Spalte; wähle für Kategorien Line, Bar oder Column.")
		if spec.chart_type == "scatter":
			spec.show_points, spec.connect_points = True, False
		elif spec.chart_type == "scatter_lines":
			spec.show_points, spec.connect_points = True, True
		elif spec.chart_type == "line":
			spec.show_points, spec.connect_points = False, True
		elif spec.chart_type == "line_scatter":
			spec.show_points, spec.connect_points = True, True
		else:
			spec.show_points, spec.connect_points = False, False
		target_chart_type = _chart_type_code(
			spec.chart_type,
			spec.show_points,
			spec.connect_points,
			categorical=not np.any(np.isfinite(x_values)),
		)
		trendline_x_values, trendline_y_values = _trendline_data(frame, spec)
		if spec.trendline != "none":
			options, degrees, periods = _available_trendlines(
				trendline_x_values,
				trendline_y_values,
				target_chart_type,
			)
			option = {"quadratic": "polynomial", "cubic": "polynomial"}.get(
				spec.trendline,
				spec.trendline,
			)
			if (
				option not in options
				or (option == "polynomial" and spec.polynomial_degree not in degrees)
				or (option == "moving_average" and spec.moving_average_period not in periods)
			):
				raise ValueError(
					"Der gewählte Diagrammtyp unterstützt die aktuelle Trendlinie mit diesen Daten nicht. "
					"Entferne zuerst die Trendlinie oder wähle eine andere."
				)
		old_chart_type = int(chart.ChartType)
		old_format = _capture_series_format(series, old_chart_type)
		if "line_color" in settings:
			old_format["line_color"] = _stored_color(settings["line_color"])
		elif "series_color" in settings and old_chart_type in {-4169, 4, 65, 74, 75}:
			old_format["line_color"] = _stored_color(settings["series_color"])
		old_legend = bool(chart.HasLegend)
		chart.ChartType = target_chart_type
		if int(chart.ChartType) != target_chart_type:
			chart.ChartType = old_chart_type
			raise RuntimeError(
				f"Excel hat den gewünschten Diagrammtyp {target_chart_type} nicht übernommen."
			)
		series = chart.SeriesCollection(1)
		_apply_series_format(series, old_format, target_chart_type)
		if target_chart_type in {-4169, 4, 65, 74, 75}:
			series.Format.Line.Visible = bool(spec.connect_points)
			if "line_color" in old_format:
				series.Format.Line.ForeColor.RGB = old_format["line_color"]
			if "line_width" in old_format:
				series.Format.Line.Weight = old_format["line_width"]
			if "point_size" in settings and spec.chart_type != "line":
				series.MarkerSize = float(settings["point_size"])
			series.MarkerStyle = 8 if spec.show_points else -4142
			if int(series.MarkerStyle) != (8 if spec.show_points else -4142):
				raise RuntimeError("Excel hat den Markerzustand beim Diagrammtypwechsel nicht beibehalten.")
			if bool(series.Format.Line.Visible) != spec.connect_points:
				raise RuntimeError("Excel hat den Verbindungszustand beim Diagrammtypwechsel nicht beibehalten.")
		chart.HasLegend = old_legend
		if bool(chart.HasLegend) != old_legend:
			raise RuntimeError("Excel hat den Legendenzustand beim Diagrammtypwechsel nicht beibehalten.")
		if int(chart.ChartType) != target_chart_type:
			raise RuntimeError("Excel hat den Diagrammtyp nach der Formatübernahme zurückgesetzt.")
		series = chart.SeriesCollection(1)
		_apply_series_colors(series, old_format, target_chart_type)
		_restore_saved_series_format(
			series,
			spec,
			{
				key: settings[key]
				for key in ("series_color", "line_color")
				if key in settings
			},
		)
		if isinstance(settings.get("gridlines_enabled"), bool):
			_set_chart_gridlines(chart, settings["gridlines_enabled"], bool(np.any(np.isfinite(x_values))))
	elif action in {"series_color", "line_color"}:
		rgb = _selected_color(value if isinstance(value, int) else None)
		native_type = int(chart.ChartType)
		marker_chart = native_type in {-4169, 4, 65, 74, 75}
		filled_chart = native_type in {51, 57}
		if action == "series_color":
			if filled_chart:
				series.Format.Fill.ForeColor.RGB = rgb
			elif marker_chart:
				_set_series_line_color(series, rgb, native_type)
				if native_type != 4:
					series.MarkerForegroundColor = rgb
					series.MarkerBackgroundColor = rgb
			else:
				raise ValueError("Die Datenreihenfarbe ist für diesen Diagrammtyp nicht verfügbar.")
			settings["series_color"] = rgb
		elif action == "line_color":
			if not marker_chart:
				raise ValueError("Eine Linienfarbe ist für diesen Diagrammtyp nicht verfügbar.")
			_set_series_line_color(series, rgb, native_type)
			settings["line_color"] = rgb
	elif action in {"line_width", "point_size"}:
		if action == "line_width":
			number = _selection(value if isinstance(value, int) else None, (0.75, 1, 1.5, 2, 3), "Linienbreite")
			original_chart_type = int(chart.ChartType)
			original_connect_points = bool(series.Format.Line.Visible)
			original_marker_style = int(series.MarkerStyle)
			series.Format.Line.Weight = number
			if int(chart.ChartType) != original_chart_type:
				chart.ChartType = original_chart_type
				series = chart.SeriesCollection(1)
				series.Format.Line.Weight = number
			series.Format.Line.Visible = original_connect_points
			series.MarkerStyle = original_marker_style
			if original_chart_type in {-4169, 4, 65, 74, 75}:
				line_color = settings.get("line_color", settings.get("series_color"))
				if line_color is not None:
					series.Format.Line.ForeColor.RGB = _stored_color(line_color)
				if original_chart_type != 4 and "series_color" in settings:
					rgb = _stored_color(settings["series_color"])
					series.MarkerForegroundColor = rgb
					series.MarkerBackgroundColor = rgb
				if "point_size" in settings:
					series.MarkerSize = float(settings["point_size"])
			settings["line_width"] = number
		else:
			number = _selection(value if isinstance(value, int) else None, (3, 5, 7, 9, 11), "Punktgröße")
			series.MarkerSize = number
			settings["point_size"] = number
	elif action in {"points", "connections", "connections_state"}:
		if spec.chart_type not in {"scatter", "scatter_lines", "line", "line_scatter"}:
			raise ValueError("Punkte und Verbindungslinien sind für Balken- und Säulendiagramme nicht verfügbar.")
		if action == "points":
			spec.show_points = not spec.show_points
		elif action == "connections_state":
			if not isinstance(value, bool):
				raise ValueError("Der Verbindungszustand muss ein Boolean sein.")
			spec.connect_points = value
		else:
			spec.connect_points = not spec.connect_points
		x_is_numeric = bool(np.any(np.isfinite(
			pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
		)))
		if spec.chart_type in {"scatter", "scatter_lines"}:
			spec.chart_type = "scatter_lines" if spec.connect_points else "scatter"
		old_format = _capture_series_format(series, int(chart.ChartType))
		if "line_color" in settings:
			old_format["line_color"] = _stored_color(settings["line_color"])
		elif "series_color" in settings:
			old_format["line_color"] = _stored_color(settings["series_color"])
		target_chart_type = _chart_type_code(
			spec.chart_type, spec.show_points, spec.connect_points, not x_is_numeric
		)
		chart.ChartType = target_chart_type
		series = chart.SeriesCollection(1)
		_apply_series_format(series, old_format, target_chart_type)
		series.Format.Line.Visible = bool(spec.connect_points)
		if "line_color" in old_format:
			series.Format.Line.ForeColor.RGB = old_format["line_color"]
		if "line_width" in old_format:
			series.Format.Line.Weight = old_format["line_width"]
		series.MarkerStyle = 8 if spec.show_points else -4142
		if int(series.MarkerStyle) != (8 if spec.show_points else -4142):
			raise RuntimeError("Excel hat den Markerzustand bei der Verbindungsänderung nicht beibehalten.")
		if bool(series.Format.Line.Visible) != spec.connect_points:
			raise RuntimeError("Excel hat den Verbindungszustand nicht übernommen.")
		settings["show_points"] = spec.show_points
		settings["connect_points"] = spec.connect_points
	elif action == "trendline":
		x_values, y_values = _trendline_data(frame, spec)
		options, _, _ = _available_trendlines(x_values, y_values, int(chart.ChartType))
		spec.trendline = _selection(
			value if isinstance(value, int) else None,
			options,
			"Trendlinie",
		)
		if spec.trendline in {"quadratic", "cubic"}:
			spec.polynomial_degree = 2 if spec.trendline == "quadratic" else 3
			spec.trendline = "polynomial"
		_set_trendline(
			series,
			spec.trendline,
			spec.polynomial_degree,
			spec.moving_average_period,
			x_values,
			y_values,
			int(chart.ChartType),
		)
	elif action == "polynomial_degree":
		if spec.trendline in {"quadratic", "cubic"}:
			spec.polynomial_degree = 2 if spec.trendline == "quadratic" else 3
			spec.trendline = "polynomial"
		if spec.trendline != "polynomial":
			raise ValueError("Der Polynomgrad kann nur für eine aktive polynomische Trendlinie geändert werden.")
		x_values, y_values = _trendline_data(frame, spec)
		_, degrees, _ = _available_trendlines(x_values, y_values, int(chart.ChartType))
		spec.polynomial_degree = _selection(
			value if isinstance(value, int) else None,
			degrees,
			"Polynomgrad",
		)
		_set_trendline(
			series,
			spec.trendline,
			spec.polynomial_degree,
			spec.moving_average_period,
			x_values,
			y_values,
			int(chart.ChartType),
		)
	elif action == "moving_average_period":
		if spec.trendline != "moving_average":
			raise ValueError("Die Periode kann nur für einen aktiven gleitenden Durchschnitt geändert werden.")
		x_values, y_values = _trendline_data(frame, spec)
		_, _, periods = _available_trendlines(x_values, y_values, int(chart.ChartType))
		spec.moving_average_period = _selection(
			value if isinstance(value, int) else None,
			periods,
			"Gleitender-Durchschnitt-Periode",
		)
		_set_trendline(
			series,
			spec.trendline,
			spec.polynomial_degree,
			spec.moving_average_period,
			x_values,
			y_values,
			int(chart.ChartType),
		)
	elif action == "title_toggle":
		chart.HasTitle = not bool(chart.HasTitle)
		if chart.HasTitle and not str(chart.ChartTitle.Text).strip():
			chart.ChartTitle.Text = f"{spec.y_label} in Abhängigkeit von {spec.x_label}"
		settings["title_visible"] = bool(chart.HasTitle)
	elif action in {"title", "x_label", "y_label"}:
		default = str(chart.ChartTitle.Text) if action == "title" else (
			spec.x_label if action == "x_label" else spec.y_label
		)
		label = _ask_text(book, "Neuen Text eingeben:", "Diagrammtext", default)
		if label is None:
			return
		if action == "title":
			chart.HasTitle = True
			chart.ChartTitle.Text = label
			settings["title"] = label
		elif action == "x_label":
			spec.x_label = label
			_set_axis_title(chart, spec, frame, "x")
		else:
			spec.y_label = label
			_set_axis_title(chart, spec, frame, "y")
	elif action in {"legend", "legend_state"}:
		if action == "legend_state":
			if not isinstance(value, bool):
				raise ValueError("Der Legendenzustand muss ein Boolean sein.")
			expected_legend = value
		else:
			expected_legend = not bool(chart.HasLegend)
		chart.HasLegend = expected_legend
		if bool(chart.HasLegend) != expected_legend:
			raise RuntimeError("Excel hat die Legendenumschaltung nicht übernommen.")
		settings["legend"] = expected_legend
	elif action == "gridlines":
		gridlines_visible = any(
			bool(getattr(chart.Axes(axis_type), property_name))
			for axis_type in (1, 2)
			for property_name in ("HasMajorGridlines", "HasMinorGridlines")
		)
		gridlines_enabled = not gridlines_visible
		x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
		_set_chart_gridlines(chart, gridlines_enabled, bool(np.any(np.isfinite(x_values))))
		settings.pop("gridlines", None)
		settings["gridlines_enabled"] = gridlines_enabled
	elif action == "text_target":
		settings["text_target"] = _selection(
			value if isinstance(value, int) else None,
			("title", "x_axis", "y_axis", "legend"),
			"Textelement",
		)
	elif action in {"font_name", "font_size", "font_color", "font_bold", "font_italic"}:
		element = _selected_text_element(chart, spec, frame, settings)
		font = element.Font
		target = str(settings.get("text_target", "title"))
		text_formats = dict(settings.get("text_formats", {}))
		text_format = dict(text_formats.get(target, {}))
		if action == "font_name":
			value = _selection(
				value if isinstance(value, int) else None,
				("Aptos", "Arial", "Calibri", "Times New Roman", "Cambria", "Verdana", "Tahoma", "Georgia", "Trebuchet MS", "Courier New", "Consolas", "Segoe UI"),
				"Schriftart",
			)
			font.Name = value
			text_format["name"] = value
		elif action == "font_size":
			value = _selection(
				value if isinstance(value, int) else None,
				(8, 9, 10, 11, 12, 14, 16, 18, 20, 24),
				"Schriftgröße",
			)
			font.Size = value
			text_format["size"] = value
		elif action == "font_color":
			value = _selected_color(value if isinstance(value, int) else None)
			font.Color = value
			text_format["color"] = value
		elif action == "font_bold":
			font.Bold = not bool(font.Bold)
			text_format["bold"] = bool(font.Bold)
		else:
			font.Italic = not bool(font.Italic)
			text_format["italic"] = bool(font.Italic)
		text_formats[target] = text_format
		settings["text_formats"] = text_formats
	else:
		raise ValueError(f"Unbekannte Diagrammbearbeitung: {action}")
	_sync_chart_state_from_excel(chart, spec, settings)
	chart.Parent.Activate()
	_write_analysis(book, frame, spec, chart_name)
	_save_chart_state(
		book,
		chart_name,
		str(state["sheet"]),
		str(state["address"]),
		spec,
		settings,
	)
	_write_status(book, f"Diagramm aktualisiert: {action}.")


def _run_chart_edit(
	action: str,
	value: int | str | None = None,
	target_workbook: str | None = None,
	target_sheet: str | None = None,
	target_chart: str | None = None,
) -> None:
	def edit(book: xw.Book) -> None:
		book = _chart_edit_book(book, target_workbook)
		if action == "swap_axes":
			chart_name, _, chart, state = _selected_chart(book, target_sheet, target_chart)
			_swap_chart_axes(book, chart_name, chart, state)
		else:
			_edit_chart_settings(
				book,
				action,
				value,
				target_workbook,
				target_sheet,
				target_chart,
			)
	_safe_action(edit)


def swap_axes(
	target_workbook: str | None = None,
	target_sheet: str | None = None,
	target_chart: str | None = None,
) -> None:
	_run_chart_edit("swap_axes", None, target_workbook, target_sheet, target_chart)


def change_chart_type(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("chart_type", selection_index, target_workbook, target_sheet, target_chart)


def set_series_color(color: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("series_color", color, target_workbook, target_sheet, target_chart)


def set_line_color(color: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("line_color", color, target_workbook, target_sheet, target_chart)


def set_line_width(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("line_width", selection_index, target_workbook, target_sheet, target_chart)


def set_point_size(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("point_size", selection_index, target_workbook, target_sheet, target_chart)


def toggle_points(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("points", None, target_workbook, target_sheet, target_chart)


def toggle_connections(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("connections", None, target_workbook, target_sheet, target_chart)


def set_connections(connected: bool, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("connections_state", connected, target_workbook, target_sheet, target_chart)


def change_trendline(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("trendline", selection_index, target_workbook, target_sheet, target_chart)


def change_polynomial_degree(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("polynomial_degree", selection_index, target_workbook, target_sheet, target_chart)


def change_moving_average_period(selection_index: int | None = None, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("moving_average_period", selection_index, target_workbook, target_sheet, target_chart)


def toggle_title(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("title_toggle", None, target_workbook, target_sheet, target_chart)


def set_text_target(selection_index: int, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("text_target", selection_index, target_workbook, target_sheet, target_chart)


def set_text_font_name(selection_index: int, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("font_name", selection_index, target_workbook, target_sheet, target_chart)


def set_text_font_size(selection_index: int, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("font_size", selection_index, target_workbook, target_sheet, target_chart)


def set_text_color(color: int, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("font_color", color, target_workbook, target_sheet, target_chart)


def toggle_text_bold(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("font_bold", None, target_workbook, target_sheet, target_chart)


def toggle_text_italic(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("font_italic", None, target_workbook, target_sheet, target_chart)


def edit_chart_title(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("title", None, target_workbook, target_sheet, target_chart)


def edit_x_axis_title(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("x_label", None, target_workbook, target_sheet, target_chart)


def edit_y_axis_title(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("y_label", None, target_workbook, target_sheet, target_chart)


def toggle_legend(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("legend", None, target_workbook, target_sheet, target_chart)


def set_legend(visible: bool, target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("legend_state", visible, target_workbook, target_sheet, target_chart)


def toggle_gridlines(target_workbook: str | None = None, target_sheet: str | None = None, target_chart: str | None = None) -> None:
	_run_chart_edit("gridlines", None, target_workbook, target_sheet, target_chart)


def show_help() -> None:
	def action(book: xw.Book) -> None:
		help_text = (
			f"{PRODUCT_NAME}\n\n"
			"- Analysiert allgemeine Tabellen und Messdaten mit der KI.\n"
			"- Die KI wählt Achsen, Diagrammart und einen passenden Kurvenfit.\n"
			"- Im Bereich „Diagramm bearbeiten“ lassen sich ausgewählte PLVS-Diagramme nachträglich anpassen.\n"
			"- Achsentausch, Diagrammtyp, Darstellung, Farben, Trendlinie und Beschriftungen bearbeiten nur das Diagramm, nicht die Quelldaten.\n"
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
	x_is_numeric = np.any(np.isfinite(x_values))
	if not x_is_numeric and spec.chart_type in {"scatter", "scatter_lines"}:
		raise ValueError("Ein Scatter-/XY-Diagramm benötigt eine numerische x-Spalte.")
	if spec.chart_type == "scatter":
		spec.show_points, spec.connect_points = True, False
	elif spec.chart_type == "scatter_lines":
		spec.show_points, spec.connect_points = True, True
	elif spec.chart_type == "line":
		spec.show_points, spec.connect_points = False, True
	elif spec.chart_type == "line_scatter":
		spec.show_points, spec.connect_points = True, True
	else:
		spec.show_points, spec.connect_points = False, False
	if x_is_numeric:
		valid = np.isfinite(x_values) & np.isfinite(y_values)
	else:
		valid = frame[spec.x_column].notna().to_numpy() & np.isfinite(y_values)
		x_values = pd.factorize(frame[spec.x_column], sort=False)[0].astype(float) + 1
	if int(valid.sum()) < 2 or (
		x_is_numeric
		and spec.chart_type not in {"bar", "column"}
		and np.unique(x_values[valid]).size < 2
	):
		raise ValueError("Mindestens zwei gültige Messwertpaare mit verschiedenen x-Werten sind erforderlich.")
	native_chart_type = _chart_type_code(
		spec.chart_type,
		spec.show_points,
		spec.connect_points,
		categorical=not x_is_numeric,
	)
	trendline_x_values, trendline_y_values = _trendline_data(frame, spec)
	if spec.trendline != "none":
		options, degrees, periods = _available_trendlines(
			trendline_x_values,
			trendline_y_values,
			native_chart_type,
		)
		option = {"quadratic": "polynomial", "cubic": "polynomial"}.get(
			spec.trendline,
			spec.trendline,
		)
		if option not in options:
			raise ValueError(
				f"Die Trendlinie {spec.trendline!r} ist für Diagrammtyp und Daten nicht anwendbar."
			)
		degree = (
			2 if spec.trendline == "quadratic"
			else 3 if spec.trendline == "cubic"
			else spec.polynomial_degree
		)
		if option == "polynomial" and degree not in degrees:
			raise ValueError(
				f"Polynomgrad {degree} benötigt ausreichend verschiedene x-Werte."
			)
		if option == "moving_average" and spec.moving_average_period not in periods:
			raise ValueError(
				f"Periode {spec.moving_average_period} ist für die vorhandenen Messpunkte nicht möglich."
			)
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
	chart.ChartType = native_chart_type
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
	for error_column, direction in (
		(spec.x_error_column, -4168),
		(spec.y_error_column, 1),
	):
		if error_column is None:
			continue
		error_column_index = first_column + int(frame.columns.get_loc(error_column))
		error_values = sheet.range(
			(first_data_row, error_column_index),
			(last_data_row, error_column_index),
		).api
		series.ErrorBar(
			Direction=direction,
			Include=1,
			Type=-4114,
			Amount=error_values,
			MinusValues=error_values,
		)
	series.MarkerStyle = 8 if spec.show_points else -4142
	chart.DisplayBlanksAs = 1
	axis_values = [(2, spec.y_label, y_values, spec.y_min, spec.y_max)]
	if x_is_numeric and spec.chart_type not in {"bar", "column", "line", "line_scatter"}:
		axis_values.insert(0, (1, spec.x_label, x_values, spec.x_min, spec.x_max))
	for axis_type, label, values, minimum, maximum in axis_values:
		axis = chart.Axes(axis_type)
		axis.HasTitle = True
		axis.AxisTitle.Text = label
		valid_values = values[np.isfinite(values)]
		is_zero = spec.origin is True
		axis_minimum, axis_maximum = _axis_bounds(valid_values, minimum, maximum, is_zero)
		axis.MinimumScale = axis_minimum
		axis.MaximumScale = axis_maximum
	_set_creation_gridlines(chart, x_is_numeric)
	if spec.trendline != "none":
		try:
			_set_trendline(
				series,
				spec.trendline,
				spec.polynomial_degree,
				spec.moving_average_period,
				trendline_x_values,
				trendline_y_values,
				native_chart_type,
			)
		except Exception:
			chart_object.Delete()
			raise
	_set_text_name(book, CHART_SHEET_NAME, sheet.name)
	return str(chart_object.Name)

def _queue_ai_question(
	book: xw.Book,
	spec: ChartSpecification,
	sheet: xw.Sheet,
	data_range: xw.Range,
	decisions: dict[str, str] | None = None,
) -> None:
	options = spec.user_input_options
	if not options:
		raise AIServiceError("Die KI hat eine Rückfrage ohne anklickbare Optionen geliefert.")
	spec_state = asdict(spec)
	spec_state.pop("ai_response", None)
	state = {
		"kind": "ai",
		"sheet": sheet.name,
		"address": str(data_range.address),
		"spec": spec_state,
		"decisions": decisions or {},
		"question_index": 0,
	}
	question = options[0]
	choices = question.get("options")
	label = str(question.get("question", "")).strip()
	if not label or not isinstance(choices, list) or not choices:
		raise AIServiceError("Die KI hat eine ungültige Rückfrage geliefert.")
	_set_pending_choice(book, "ai", label, [str(choice) for choice in choices], state)


def _create_chart_from_spec(
	book: xw.Book,
	data_sheet: xw.Sheet,
	data_range: xw.Range,
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> None:
	chart_name = _add_native_chart(book, data_sheet, data_range, frame, spec)
	_write_analysis(book, frame, spec, chart_name)
	_save_chart_state(
		book,
		chart_name,
		data_sheet.name,
		str(data_range.address),
		spec,
		{
			"show_points": spec.show_points,
			"connect_points": spec.connect_points,
			"legend": False,
			"gridlines_enabled": True,
			"title": f"{spec.y_label} in Abhängigkeit von {spec.x_label}",
		},
	)
	data_sheet.activate()
	data_sheet.api.ChartObjects(chart_name).Activate()
	message = f"Diagramm und Analyse erstellt: {spec.x_label} gegen {spec.y_label}."
	_write_status(book, message)
	_message(book, message)


def _analyze_chart_source(
	book: xw.Book,
	data_sheet: xw.Sheet,
	data_range: xw.Range,
	frame: pd.DataFrame,
	columns: list[str],
) -> None:
	if not columns:
		raise ValueError(
			"Der ausgewählte Bereich benötigt mindestens eine numerische y-Messspalte."
		)
	spec = analyze_with_ai(frame, columns, _get_ai_config())
	if spec.needs_user_input:
		_queue_ai_question(book, spec, data_sheet, data_range)
		return
	_create_chart_from_spec(book, data_sheet, data_range, frame, spec)


def choose_pending_choice(selection_index: int) -> None:
	def action(book: xw.Book) -> None:
		count = int(_read_text_name(book, PENDING_CHOICE_COUNT) or "0")
		if not isinstance(selection_index, int) or isinstance(selection_index, bool):
			raise ValueError("Die Ribbon-Auswahl ist ungültig.")
		if not 0 <= selection_index < count:
			raise ValueError("Die ausgewählte Ribbon-Option ist nicht mehr verfügbar.")
		label_parts = (_read_text_name(book, PENDING_CHOICE_LABEL) or "").split("|", 1)
		state_json = _read_text_name(book, PENDING_CHOICE_STATE)
		if len(label_parts) != 2 or not state_json:
			raise ValueError("Die gespeicherte Ribbon-Auswahl ist unvollständig.")
		kind, _ = label_parts
		state = json.loads(state_json)
		if not isinstance(state, dict) or state.get("kind") != kind:
			raise ValueError("Die gespeicherte Ribbon-Auswahl ist ungültig.")
		if kind == "source":
			sources = state.get("sources", [])
			if selection_index == len(sources):
				_clear_pending_choice(book)
				data_sheet, data_range = _select_manual_source(book)
			else:
				source = sources[selection_index]
				data_sheet = book.sheets[str(source["sheet"])]
				data_range = data_sheet.range(str(source["address"]))
				_store_source(book, data_sheet, data_range, str(source["label"]))
				_clear_pending_choice(book)
			frame = _frame_from_range(data_range)
			_analyze_chart_source(book, data_sheet, data_range, frame, get_numeric_columns(frame))
			return
		if kind != "ai":
			raise ValueError(f"Unbekannter Ribbon-Auswahltyp: {kind}")
		choice = _read_text_name(book, f"{PENDING_CHOICE_ITEM_PREFIX}{selection_index}")
		if choice is None:
			raise ValueError("Die ausgewählte KI-Option ist nicht mehr verfügbar.")
		spec = ChartSpecification(**state["spec"])
		question_index = int(state["question_index"])
		question = spec.user_input_options[question_index]
		decisions = dict(state.get("decisions", {}))
		decisions[str(question["type"])] = choice
		question_index += 1
		if question_index < len(spec.user_input_options):
			state["decisions"] = decisions
			state["question_index"] = question_index
			next_question = spec.user_input_options[question_index]
			next_choices = next_question.get("options")
			next_label = str(next_question.get("question", "")).strip()
			if not next_label or not isinstance(next_choices, list) or not next_choices:
				raise AIServiceError("Die KI hat eine ungültige Rückfrage geliefert.")
			_set_pending_choice(
				book,
				"ai",
				next_label,
				[str(option) for option in next_choices],
				state,
			)
			return
		data_sheet = book.sheets[str(state["sheet"])]
		data_range = data_sheet.range(str(state["address"]))
		frame = _frame_from_range(data_range)
		columns = get_numeric_columns(frame)
		final_spec = reanalyze_with_user_input(
			frame,
			columns,
			_get_ai_config(),
			spec,
			decisions,
		)
		_create_chart_from_spec(book, data_sheet, data_range, frame, final_spec)
		_clear_pending_choice(book)

	_safe_action(action)


def create_chart() -> None:
	def action(book: xw.Book) -> None:
		source = _source_for_chart(book)
		if source is None:
			return
		data_sheet, data_range, frame, columns = source
		_analyze_chart_source(book, data_sheet, data_range, frame, columns)

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
