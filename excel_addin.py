"""Excel-Arbeitsblatt-GUI: Shapes rufen VBA auf, VBA ruft diese xlwings-Aktionen."""

from __future__ import annotations

from config import load_config, save_config
import math
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import xlwings as xw

from ai import AIServiceError, analyze_with_ai, test_ai_connection
from config import load_config
from excel import get_numeric_columns
from models import ChartSpecification
from physics import identify_column, infer_chart_candidates, linear_fit


PRODUCT_NAME = "PLVS ULTRA Graphs"
ASSISTANT_SHEET = PRODUCT_NAME
ANALYSIS_SHEET = PRODUCT_NAME
BUTTON_ACTIONS = (
	("PLVS_ULTRA_Analyze", "ANALYSIEREN", 3, 42),
	("PLVS_ULTRA_CreateChart", "DIAGRAMM ERSTELLEN", 5, 50),
	("PLVS_ULTRA_AISettings", "KI-EINSTELLUNGEN", 8, 42),
	("PLVS_ULTRA_ShowAnalysis", "ANALYSE", 10, 42),
	("PLVS_ULTRA_Help", "HILFE", 12, 42),
)


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
		if book is not None:
			if ASSISTANT_SHEET in [sheet.name for sheet in book.sheets]:
				book.sheets[ASSISTANT_SHEET].range("B12").value = message
			_message(book, message, error=True)
		else:
			raise


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


def _frame_from_sheet(sheet: xw.Sheet) -> pd.DataFrame:
	rows = _matrix(sheet.used_range.value)
	if len(rows) < 2:
		raise ValueError("Das aktive Tabellenblatt enthält keine Messwerttabelle mit mindestens einer Datenzeile.")
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
		raise ValueError("Im aktiven Tabellenblatt wurden keine Messwerte gefunden.")
	return frame


def _ensure_dashboard_buttons(sheet: xw.Sheet) -> None:
	shapes = sheet.api.Shapes
	for action_name, caption, row, height in BUTTON_ACTIONS:
		shape_name = f"PLVS_{action_name.removeprefix('PLVS_ULTRA_')}"
		try:
			shape = shapes.Item(shape_name)
		except Exception:
			left = float(sheet.range(f"G{row}").api.Left)
			top = float(sheet.range(f"G{row}").api.Top)
			width = 230
			shape = shapes.AddShape(1, left, top, width, height)
			shape.Name = shape_name
		shape.TextFrame2.TextRange.Text = caption
		shape.TextFrame2.TextRange.Font.Size = 12
		shape.TextFrame2.TextRange.Font.Bold = True
		shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
		shape.Fill.ForeColor.RGB = 0x1A323E if action_name == "PLVS_ULTRA_CreateChart" else 0x1E6A5F
		shape.Line.Visible = 0
		shape.OnAction = f"'{sheet.book.name}'!{action_name}"


def ensure_dashboard(book: xw.Book | None = None) -> xw.Sheet:
	"""Erstellt oder repariert das Dashboardblatt samt Shape-Buttons."""
	if book is None:
		book = _active_book()
	sheet = _settings_sheet(book)
	try:
		data_sheet = _active_data_sheet(book)
		frame = _frame_from_sheet(data_sheet)
		columns = get_numeric_columns(frame)
		if not sheet.range("B3").value:
			sheet.range("B3").value = data_sheet.name
		for cell, value in (("B4", "automatisch"), ("B5", "automatisch"), ("B6", "XY-Streudiagramm"), ("B7", "auto"), ("B8", "auto"), ("B9", "Aktuelle Arbeitsmappe"), ("B10", "Nein")):
			if not sheet.range(cell).value:
				sheet.range(cell).value = value
		quantities = [identify_column(column)[0] for column in columns if not identify_column(column)[2]]
		sheet.range("B11").value = ", ".join(dict.fromkeys(q for q in quantities if q)) or ", ".join(columns)
		if not sheet.range("B12").value:
			sheet.range("B12").value = "Bereit"
	except ValueError as exc:
		sheet.range("B12").value = str(exc)
	_ensure_dashboard_buttons(sheet)
	sheet.activate()
	return sheet


def _active_data_sheet(book: xw.Book) -> xw.Sheet:
	active = book.sheets.active
	if active.name not in {ASSISTANT_SHEET, ANALYSIS_SHEET} and not active.name.startswith("_PhysikDiagrammDaten"):
		return active
	if ASSISTANT_SHEET in [sheet.name for sheet in book.sheets]:
		selected_name = str(book.sheets[ASSISTANT_SHEET].range("B3").value or "").strip()
		if selected_name in [sheet.name for sheet in book.sheets] and selected_name not in {ASSISTANT_SHEET, ANALYSIS_SHEET}:
			return book.sheets[selected_name]
	for sheet in book.sheets:
		if sheet.name not in {ASSISTANT_SHEET, ANALYSIS_SHEET} and not sheet.name.startswith("_PhysikDiagrammDaten") and sheet.api.Visible == -1:
			return sheet
	raise ValueError("Kein sichtbares Datenblatt gefunden. Wechsle zu einem Tabellenblatt mit Messwerten.")


def _settings_sheet(book: xw.Book, refresh_columns: list[str] | None = None) -> xw.Sheet:
	if ASSISTANT_SHEET in [sheet.name for sheet in book.sheets]:
		sheet = book.sheets[ASSISTANT_SHEET]
	else:
		initial_sheet_name = book.sheets.active.name
		sheet = book.sheets.add(ASSISTANT_SHEET, before=book.sheets[0])
		sheet.range("A1").value = PRODUCT_NAME
		sheet.range("A2").value = "Excel-Dashboard für Messreihen"
		sheet.range("A3:A12").value = [
			["Aktuelles Arbeitsblatt"], ["Unabhängige Größe"], ["Abhängige Größe"], ["Diagrammtyp"],
			["Ausgleich"], ["Ursprung bei 0"], ["Arbeitsmodus"], ["KI bei Unsicherheit"],
			["Erkannte Messgrößen"], ["Status"],
		]
		sheet.range("D3:D6").value = [["KI-EINSTELLUNGEN"], ["API-Endpunkt"], ["Modell"], ["API-Key"]]
		sheet.range("E6").value = "Nur config.json; nicht im Workbook"
		sheet.range("A15").value = "ANALYSE"
		sheet.range("A17:A24").value = [
			["Unabhängige Größe"], ["Abhängige Größe"], ["Diagrammtyp"], ["Nullpunkt"],
			["Ausgleich"], ["Fehlerbalken"], ["Confidence"], ["Begründung"],
		]
		sheet.range("A1:G1").merge()
		sheet.range("A2:G2").merge()
		sheet.range("A1").api.Font.Bold = True
		sheet.range("A1").api.Font.Size = 20
		sheet.range("A2").api.Font.Size = 11
		sheet.range("A3:A12").api.Font.Bold = True
		sheet.range("D3:D6").api.Font.Bold = True
		sheet.range("A15:B15").merge()
		sheet.range("A15").api.Font.Bold = True
		sheet.range("A15").api.Font.Size = 14
		sheet.range("A17:A24").api.Font.Bold = True
		sheet.range("A1:G1").api.Interior.Color = 26 + 50 * 256 + 62 * 65536
		sheet.range("A1").api.Font.Color = 255 * 65536 + 255 * 256 + 255
		sheet.range("A3:A12").api.Interior.Color = 235 + 239 * 256 + 239 * 65536
		sheet.range("B3:B12").api.Interior.Color = 255 + 248 * 256 + 224 * 65536
		sheet.range("D3:E3").api.Interior.Color = 30 + 105 * 256 + 95 * 65536
		sheet.range("D3:E3").api.Font.Color = 255 * 65536 + 255 * 256 + 255
		sheet.range("D4:D6").api.Interior.Color = 235 + 239 * 256 + 239 * 65536
		sheet.range("A:A").api.ColumnWidth = 27
		sheet.range("B:B").api.ColumnWidth = 34
		sheet.range("C:C").api.ColumnWidth = 3
		sheet.range("D:D").api.ColumnWidth = 18
		sheet.range("E:E").api.ColumnWidth = 43
		sheet.range("F:F").api.ColumnWidth = 3
		sheet.range("G:G").api.ColumnWidth = 32
		sheet.range("B4:B10").value = [
			["automatisch"], ["automatisch"], ["XY-Streudiagramm"], ["auto"],
			["auto"], ["Aktuelle Arbeitsmappe"], ["Nein"],
		]
		sheet.range("B3").value = initial_sheet_name
		sheet.range("A26").value = "Messwerttabelle: Überschriften in der ersten Tabellenzeile, danach numerische Messwerte."
		sheet.range("A26:G26").merge()
		try:
			config = load_config()
		except (FileNotFoundError, ValueError):
			config = {"endpoint": "", "model": "", "api_key": ""}
		sheet.range("E4").value = config["endpoint"]
		sheet.range("E5").value = config["model"]
		sheet.range("E6").value = "Vorhanden" if config["api_key"] else "Nur lokal in config.json"
	_ensure_dashboard_buttons(sheet)
	if sheet.range("A1").value != PRODUCT_NAME:
		sheet.range("A1").value = PRODUCT_NAME
	if not sheet.range("A2").value:
			sheet.range("A2").value = "Excel-Dashboard für Messreihen"
	if refresh_columns is not None:
		sheet.range("J:J").api.EntireColumn.Hidden = True
		options = ["automatisch", *refresh_columns]
		sheet.range("J1").value = [[item] for item in options]
		last = len(options)
		for cell in ("B4", "B5"):
			validation = sheet.range(cell).api.Validation
			try:
				validation.Delete()
			except Exception:
				pass
			validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f"=$J$1:$J${last}")
		for cell, values in (
			("B6", "XY-Streudiagramm,XY mit Verbindungslinie"),
			("B7", "auto,none,linear,quadratic,cubic"),
			("B8", "auto,ja,nein"),
			("B9", "Aktuelle Arbeitsmappe"),
			("B10", "Nein,Ja"),
		):
			validation = sheet.range(cell).api.Validation
			try:
				validation.Delete()
			except Exception:
				pass
			validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f'"{values}"')
	return sheet


def open_settings() -> None:
	def action(book: xw.Book) -> None:
		data_sheet = _active_data_sheet(book)
		frame = _frame_from_sheet(data_sheet)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Keine zwei numerischen Messspalten im aktiven Blatt gefunden.")
		sheet = _settings_sheet(book, columns)
		sheet.range("B3").value = data_sheet.name
		sheet.range("I:I").api.EntireColumn.Hidden = True
		data_sheets = [
			item.name for item in book.sheets
			if item.name not in {ASSISTANT_SHEET, ANALYSIS_SHEET}
			and not item.name.startswith("_PhysikDiagrammDaten")
			and item.api.Visible == -1
		]
		sheet.range("I1").value = [[name] for name in data_sheets]
		validation = sheet.range("B3").api.Validation
		try:
			validation.Delete()
		except Exception:
			pass
		validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f"=$I$1:$I${len(data_sheets)}")
		sheet.activate()
		sheet.range("B4").select()
		book.app.api.StatusBar = f"{PRODUCT_NAME}: Messspalten, Fit, Ursprung und Ausgabeoptionen einstellen."
	_safe_action(action)


def open_ai_settings() -> None:
	def action(book: xw.Book) -> None:
		sheet = _settings_sheet(book)
		sheet.activate()
		sheet.range("E4").select()
		config = _get_ai_config(book)
		key_entry = book.app.api.InputBox(
			"API-Key optional. Der Wert wird nur lokal in config.json gespeichert, nicht im Workbook.\nLeer lassen, um einen vorhandenen Schlüssel beizubehalten.",
			PRODUCT_NAME,
			"",
			Type=2,
		)
		if key_entry is False:
			return
		config["endpoint"] = str(sheet.range("E4").value or "").strip()
		config["model"] = str(sheet.range("E5").value or "").strip()
		if str(key_entry).strip():
			config["api_key"] = str(key_entry).strip()
		path = save_config(config)
		sheet.range("E6").value = "Vorhanden (lokal gespeichert)" if config["api_key"] else "Nicht gesetzt"
		_write_status(book, "KI-Konfiguration lokal gespeichert.")
		_message(book, f"Endpoint und Modell gespeichert. API-Key wird ausschließlich lokal gespeichert: {path}")
		book.app.api.StatusBar = f"{PRODUCT_NAME}: KI-Konfiguration lokal gespeichert."
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
	settings = _settings_sheet(book)
	candidates = infer_chart_candidates(frame, columns)
	selected_x = str(settings.range("B4").value or "automatisch").strip()
	selected_y = str(settings.range("B5").value or "automatisch").strip()
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
		if ambiguous and allow_ai and str(settings.range("B10").value).strip().lower() in {"ja", "yes"}:
			config = _get_ai_config(book)
			config["endpoint"] = str(settings.range("E4").value or config["endpoint"]).strip()
			config["model"] = str(settings.range("E5").value or config["model"]).strip()
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
	fit = str(settings.range("B7").value or "auto").strip().lower()
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
		chart_type="line_scatter" if str(settings.range("B6").value).startswith("XY mit") else "scatter",
		origin={"ja": True, "nein": False, "auto": "auto"}.get(str(settings.range("B8").value or "auto").strip().lower(), "auto"),
		trendline=fit,
		x_error_column=error_for.get(x_quantity),
		y_error_column=error_for.get(y_quantity),
		x_quantity=x_quantity,
		y_quantity=y_quantity,
		independent_variable=x_quantity,
		dependent_variable=y_quantity,
		confidence=selected_candidate.confidence if selected_candidate else 1.0,
		reason=(f"KI mit Confidence {ai_spec.confidence:.0%} verworfen; Auswahl durch Benutzer." if ai_spec is not None else selected_candidate.reason if selected_candidate else "Benutzer hat x- und y-Spalte ausgewählt."),
	)
	return spec, candidates


def _analysis_rows(spec: ChartSpecification, candidates: list) -> list[list[Any]]:
	confidence = spec.confidence or 0.0
	confidence_label = "Entscheidung" if confidence >= 0.9 else "Mit Vorbehalt" if confidence >= 0.7 else "Manuell bestätigt"
	return [
		["Unabhängige Größe", spec.x_label],
		["Abhängige Größe", spec.y_label],
		["Diagrammtyp", "XY mit Verbindungslinie" if spec.connect_points else "XY-Streudiagramm"],
		["Ursprung", "Ja" if spec.origin is True or spec.origin == "ja" else "Nein" if spec.origin is False or spec.origin == "nein" else "Automatisch"],
		["Ausgleich", spec.trendline],
		["Fehlerbalken", ", ".join(name for name in (spec.x_error_column, spec.y_error_column) if name) or "Keine erkannt"],
		["KI", "Verwendet" if spec.reason.startswith("KI:") else "Lokale Regeln / Auswahl"],
		["Confidence", f"{confidence:.0%} ({confidence_label})"],
		["Begründung", spec.reason],
		["Mögliche Beziehungen", "; ".join(f"{item.relationship}: {item.x_column} → {item.y_column} ({item.confidence:.0%})" for item in candidates[:5])],
	]


def _write_analysis(book: xw.Book, spec: ChartSpecification, candidates: list) -> xw.Sheet:
	sheet = _settings_sheet(book)
	for row_index, row in enumerate(_analysis_rows(spec, candidates), start=17):
		sheet.range(f"A{row_index}").value = row[0]
		sheet.range(f"B{row_index}").value = row[1]
	sheet.range("B17:B25").api.WrapText = True
	sheet.range("A17:A25").api.Font.Bold = True
	sheet.activate()
	return sheet


def show_help() -> None:
	def action(book: xw.Book) -> None:
		help_text = (
			f"{PRODUCT_NAME}\n\n"
			"- Analysiert Messwerttabellen im aktiven Excel-Blatt.\n"
			"- Erkennt physikalische Größen, Einheiten und potenzielle Beziehungen.\n"
			"- Erstellt ein natives Excel-Diagramm im geöffneten Workbook.\n"
			"- Die KI wird nur bei mehrdeutigen Fällen verwendet.\n"
			"- Wenn die Confidence niedrig ist, bleibt die Auswahl nachvollziehbar und manuell prüfbar."
		)
		_message(book, help_text)
		_write_status(book, "Hilfe angezeigt.")
	_safe_action(action)


def _get_ai_config(book: xw.Book) -> dict[str, str]:
	try:
		config = load_config()
	except (FileNotFoundError, ValueError):
		config = {"endpoint": "", "api_key": "", "model": ""}
	if ASSISTANT_SHEET in [sheet.name for sheet in book.sheets]:
		settings = book.sheets[ASSISTANT_SHEET]
		config["endpoint"] = str(settings.range("E4").value or config["endpoint"]).strip()
		config["model"] = str(settings.range("E5").value or config["model"]).strip()
	return config


def _write_status(book: xw.Book, message: str) -> None:
	if ASSISTANT_SHEET in [sheet.name for sheet in book.sheets]:
		book.sheets[ASSISTANT_SHEET].range("B12").value = message
	book.app.api.StatusBar = message


def analyze_table() -> None:
	def action(book: xw.Book) -> None:
		data_sheet = _active_data_sheet(book)
		frame = _frame_from_sheet(data_sheet)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Keine geeigneten Messwertspalten gefunden. Benötigt werden mindestens zwei numerische Spalten.")
		settings = _settings_sheet(book, columns)
		settings.range("B3").value = data_sheet.name
		spec, candidates = _resolve_spec(book, frame, columns, allow_ai=True)
		_write_analysis(book, spec, candidates)
		_write_status(book, "Analyse abgeschlossen.")
		_message(book, f"Unabhängig: {spec.x_label}\nAbhängig: {spec.y_label}\nBeziehung: {spec.reason}")
	_safe_action(action)


def show_analysis() -> None:
	def action(book: xw.Book) -> None:
		data_sheet = _active_data_sheet(book)
		frame = _frame_from_sheet(data_sheet)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Keine geeigneten Messwertspalten gefunden.")
		spec, candidates = _resolve_spec(book, frame, columns, allow_ai=False)
		_write_analysis(book, spec, candidates)
		_write_status(book, "Analyse auf dem PLVS ULTRA Graphs-Dashboard aktualisiert.")
	_safe_action(action)


def _excel_column(index: int) -> str:
	letters = ""
	while index:
		index, remainder = divmod(index - 1, 26)
		letters = chr(65 + remainder) + letters
	return letters


def _chart_data(
	book: xw.Book,
	frame: pd.DataFrame,
	spec: ChartSpecification,
) -> tuple[xw.Sheet, int, list[tuple[str, int, int, int]]]:
	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	if int(valid.sum()) < 2 or np.unique(x_values[valid]).size < 2:
		raise ValueError("Mindestens zwei gültige Messwertpaare mit verschiedenen x-Werten sind erforderlich.")
	chart_frame = pd.DataFrame({spec.x_column: x_values[valid], spec.y_column: y_values[valid]})
	error_series: list[tuple[str, int, int, int]] = []
	for error_column in (spec.x_error_column, spec.y_error_column):
		if error_column and error_column not in frame.columns:
			raise ValueError(f"Fehlerwert-Spalte {error_column!r} existiert nicht.")
	name = "_PhysikDiagrammDaten"
	index = 2
	while name in [sheet.name for sheet in book.sheets]:
		name = f"_PhysikDiagrammDaten_{index}"
		index += 1
	helper = book.sheets.add(name, before=book.sheets[0])
	chart_rows: list[list[Any]] = [chart_frame.columns.tolist(), *chart_frame.values.tolist()]
	for direction, error_column in (("y", spec.y_error_column), ("x", spec.x_error_column)):
		if not error_column:
			continue
		errors = pd.to_numeric(frame[error_column], errors="coerce").to_numpy(dtype=float)[valid]
		if not np.any(np.isfinite(errors) & (errors >= 0)):
			continue
		x_column = len(chart_rows[0]) + 1
		y_column = x_column + 1
		chart_rows[0].extend([f"{direction}-Fehlerbalken x", f"{direction}-Fehlerbalken y"])
		for index, (x_value, y_value, error) in enumerate(zip(x_values[valid], y_values[valid], errors), start=1):
			chart_rows[index].extend([None, None])
			if not np.isfinite(error) or error < 0:
				continue
			if direction == "y":
				chart_rows.extend([
					[None] * (x_column - 1) + [float(x_value), float(y_value - error)],
					[None] * (x_column - 1) + [float(x_value), float(y_value + error)],
					[None] * (x_column - 1) + [None, None],
				])
			else:
				chart_rows.extend([
					[None] * (x_column - 1) + [float(x_value - error), float(y_value)],
					[None] * (x_column - 1) + [float(x_value + error), float(y_value)],
					[None] * (x_column - 1) + [None, None],
				])
		last_error_row = len(chart_rows)
		error_series.append((direction, x_column, y_column, last_error_row))
	width = len(chart_rows[0])
	chart_rows = [row + [None] * (width - len(row)) for row in chart_rows]
	helper.range("A1").value = chart_rows
	helper.api.Visible = 2
	return helper, len(chart_frame) + 1, error_series


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


def _add_native_chart(book: xw.Book, sheet: xw.Sheet, frame: pd.DataFrame, spec: ChartSpecification) -> list[str]:
	helper, last_row, error_series = _chart_data(book, frame, spec)
	used = sheet.used_range.api
	left = float(used.Left) + float(used.Width) + 24
	top = float(used.Top)
	chart_object = sheet.api.ChartObjects().Add(left, top, 500, 300)
	chart = chart_object.Chart
	chart.ChartType = 74 if spec.connect_points or spec.chart_type in {"line", "line_scatter"} else -4169
	chart.HasTitle = True
	chart.ChartTitle.Text = f"{spec.y_label} in Abhängigkeit von {spec.x_label}"
	chart.HasLegend = False
	series = chart.SeriesCollection().NewSeries()
	series.Name = "Messwerte"
	x_letter, y_letter = _excel_column(1), _excel_column(2)
	series.XValues = helper.range(f"{x_letter}2:{x_letter}{last_row}").api
	series.Values = helper.range(f"{y_letter}2:{y_letter}{last_row}").api
	series.MarkerStyle = 8 if spec.show_points else -4142
	chart.DisplayBlanksAs = 1
	for axis_type, label, values, minimum, maximum, quantity in (
		(1, spec.x_label, pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float), spec.x_min, spec.x_max, spec.x_quantity),
		(2, spec.y_label, pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float), spec.y_min, spec.y_max, spec.y_quantity),
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
	for direction, x_column, y_column, error_last_row in error_series:
		error_line = chart.SeriesCollection().NewSeries()
		error_line.Name = f"{direction}-Fehler"
		error_line.XValues = helper.range(f"{_excel_column(x_column)}2:{_excel_column(x_column)}{error_last_row}").api
		error_line.Values = helper.range(f"{_excel_column(y_column)}2:{_excel_column(y_column)}{error_last_row}").api
		error_line.ChartType = 74
		error_line.MarkerStyle = -4142
		error_line.Format.Line.ForeColor.RGB = 0x777777
		error_line.Format.Line.Weight = 1
	if spec.trendline != "none":
		degree = {"linear": 1, "quadratic": 2, "cubic": 3}.get(spec.trendline)
		if degree is None:
			raise ValueError(f"Unbekannte Trendlinie: {spec.trendline}")
		unique_x = np.unique(pd.to_numeric(frame[spec.x_column], errors="coerce").dropna())
		if unique_x.size <= degree:
			raise ValueError(f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} unterschiedliche x-Werte.")
		trendlines = series.Trendlines()
		if degree == 1:
			trendlines.Add(Type=-4132, Name="Lineare Ausgleichsgerade")
		else:
			trendlines.Add(Type=3, Order=degree, Name=f"Polynomfit Grad {degree}")
	return []


def _target_for(book: xw.Book, data_sheet: xw.Sheet) -> tuple[xw.Sheet, None]:
	return data_sheet, None


def create_chart() -> None:
	def action(book: xw.Book) -> None:
		data_sheet = _active_data_sheet(book)
		frame = _frame_from_sheet(data_sheet)
		columns = get_numeric_columns(frame)
		if len(columns) < 2:
			raise ValueError("Keine geeigneten Messwertspalten gefunden. Mindestens zwei numerische Spalten sind erforderlich.")
		settings = _settings_sheet(book, columns)
		spec, candidates = _resolve_spec(book, frame, columns, allow_ai=True)
		target_sheet, _ = _target_for(book, data_sheet)
		warnings = _add_native_chart(book, target_sheet, frame, spec)
		book.save()
		_write_analysis(book, spec, candidates)
		book.save()
		message = f"Diagramm erstellt: {spec.x_label} gegen {spec.y_label}."
		if warnings:
			message += "\n" + " ".join(warnings)
		message += "\nDiagramm direkt in der geöffneten Arbeitsmappe erstellt."
		_write_status(book, message)
		_message(book, message)
	_safe_action(action)


def test_ai() -> None:
	def action(book: xw.Book) -> None:
		_settings_sheet(book)
		config = _get_ai_config(book)
		if not config["endpoint"] or not config["model"]:
			raise AIServiceError("KI-Einstellungen unvollständig. Endpoint und Modell im Dashboard eintragen; den API-Key in config.json.")
		test_ai_connection(config)
		_write_status(book, "KI-Verbindung erfolgreich.")
		_message(book, "KI-Verbindung erfolgreich.")
	_safe_action(action)
