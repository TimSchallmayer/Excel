"""Einlesen und grundlegende Prüfung von Excel-Messwerttabellen."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Literal

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook as load_openpyxl_workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.series import Series as ChartSeries
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from models import ChartSpecification


SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}


def load_workbook(path: Path, sheet: int | str = 0) -> pd.DataFrame:
	if path.suffix.lower() not in SUPPORTED_SUFFIXES:
		supported = ", ".join(sorted(SUPPORTED_SUFFIXES))
		raise ValueError(f"Nicht unterstütztes Excel-Format {path.suffix!r}. Erlaubt: {supported}.")
	frame = pd.read_excel(path, sheet_name=sheet)
	column_names: list[str] = []
	counts: dict[str, int] = {}
	for column in frame.columns:
		name = str(column).strip() or "Spalte"
		counts[name] = counts.get(name, 0) + 1
		column_names.append(name if counts[name] == 1 else f"{name} ({counts[name]})")
	frame.columns = column_names
	return frame


def get_numeric_columns(frame: pd.DataFrame) -> list[str]:
	columns: list[str] = []
	for column in frame.columns:
		values = pd.to_numeric(frame[column], errors="coerce").to_numpy(dtype=float)
		finite_values = values[np.isfinite(values)]
		if finite_values.size >= 2 and np.unique(finite_values).size > 1:
			columns.append(str(column))
	return columns


def _axis_label(label: str, unit: str) -> str:
	if unit and f"[{unit}]" not in label:
		return f"{label} [{unit}]"
	return label


def _axis_bounds(
	values: np.ndarray,
	minimum: float | None,
	maximum: float | None,
	use_zero: bool,
) -> tuple[float, float]:
	data_min = float(np.min(values))
	data_max = float(np.max(values))
	span = data_max - data_min
	padding = span * 0.05 if span else max(abs(data_min) * 0.05, 1.0)
	low = minimum if minimum is not None else data_min - padding
	high = maximum if maximum is not None else data_max + padding
	if use_zero and data_min >= 0 and (minimum is None or minimum <= 0):
		low = 0.0
	if use_zero and data_max <= 0 and (maximum is None or maximum >= 0):
		high = 0.0
	if not math.isfinite(low) or not math.isfinite(high) or low >= high:
		raise ValueError("Die Diagramm-Achsengrenzen sind ungültig.")
	return low, high


def _use_zero_origin(origin: bool | str | None, quantity: str | None, values: np.ndarray) -> bool:
	if origin is True or origin == "yes":
		return float(np.min(values)) >= 0
	if origin is False or origin == "no":
		return False
	return quantity != "Temperatur" and float(np.min(values)) >= 0


def _append_error_bars(
	series: ChartSeries,
	worksheet: Worksheet,
	column: int,
	first_row: int,
	last_row: int,
	direction: Literal["x", "y"],
) -> None:
	column_letter = get_column_letter(column)
	reference = NumDataSource(
		numRef=NumRef(f=f"'{worksheet.title}'!${column_letter}${first_row}:${column_letter}${last_row}")
	)
	series.errBars = ErrorBars(
		errDir=direction,
		errValType="cust",
		plus=reference,
		minus=reference,
	)


def create_excel_chart(
	workbook: Workbook,
	data_sheet: Worksheet,
	frame: pd.DataFrame,
	spec: ChartSpecification,
	fit: str = "none",
) -> ScatterChart:
	"""Fügt ein natives Excel-XY-Diagramm rechts neben den Messdaten ein."""
	if fit not in {"none", "linear", "quadratic", "cubic"}:
		raise ValueError(f"Nicht unterstützte Trendlinie: {fit}")
	for column in (spec.x_column, spec.y_column):
		if column not in frame.columns:
			raise ValueError(f"Diagrammspalte {column!r} fehlt in der Tabelle.")

	x_values = pd.to_numeric(frame[spec.x_column], errors="coerce").to_numpy(dtype=float)
	y_values = pd.to_numeric(frame[spec.y_column], errors="coerce").to_numpy(dtype=float)
	valid = np.isfinite(x_values) & np.isfinite(y_values)
	if valid.sum() < 2 or np.unique(x_values[valid]).size < 2:
		raise ValueError("Für ein XY-Diagramm werden mindestens zwei gültige Messwertpaare mit unterschiedlichen x-Werten benötigt.")
	x_data = x_values[valid]
	y_data = y_values[valid]
	if fit != "none":
		degree = {"linear": 1, "quadratic": 2, "cubic": 3}[fit]
		if np.unique(x_data).size <= degree:
			raise ValueError(f"Ein Fit vom Grad {degree} benötigt mindestens {degree + 1} unterschiedliche x-Werte.")

	data_title = "_PhysikDiagrammDaten"
	suffix = 2
	while data_title in workbook.sheetnames:
		data_title = f"_PhysikDiagrammDaten_{suffix}"
		suffix += 1
	chart_data = workbook.create_sheet(data_title)
	chart_data.sheet_state = "hidden"
	chart_data.append([spec.x_column, spec.y_column])
	for x_value, y_value in zip(x_data, y_data):
		chart_data.append([float(x_value), float(y_value)])
	data_last_row = chart_data.max_row
	chart = ScatterChart()
	if spec.chart_type == "scatter":
		chart.scatterStyle = "marker" if spec.show_points else "line"
	else:
		chart.scatterStyle = "lineMarker" if spec.show_points else "line"
	chart.title = f"{spec.y_label} in Abhängigkeit von {spec.x_label}"
	chart.x_axis.title = _axis_label(spec.x_label, spec.x_unit)
	chart.y_axis.title = _axis_label(spec.y_label, spec.y_unit)
	chart.style = 13
	chart.width = 18
	chart.height = 10
	chart.legend = None
	chart.display_blanks = "gap"
	chart.visible_cells_only = False

	x_reference = Reference(chart_data, min_col=1, min_row=2, max_row=data_last_row)
	y_reference = Reference(chart_data, min_col=2, min_row=1, max_row=data_last_row)
	series = Series(y_reference, x_reference, title="Messwerte")
	series.marker.symbol = "circle"
	series.marker.size = 7
	if chart.scatterStyle == "marker":
		series.graphicalProperties.line.noFill = True
	if fit != "none":
		degree = {"linear": 1, "quadratic": 2, "cubic": 3}[fit]
		series.trendline = Trendline(
			name=f"{fit.title()}er Fit",
			trendlineType="linear" if degree == 1 else "poly",
			order=None if degree == 1 else degree,
			dispRSqr=True,
		)
	chart.series.append(series)

	last_row = data_last_row
	error_bar_columns: list[tuple[int, Literal["x", "y"]]] = []
	for error_column, direction in (
		(spec.x_error_column, "x"),
		(spec.y_error_column, "y"),
	):
		if not error_column:
			continue
		if error_column not in frame.columns:
			raise ValueError(f"Fehlerwert-Spalte {error_column!r} existiert nicht.")
		error_values = pd.to_numeric(frame[error_column], errors="coerce").to_numpy(dtype=float)[valid]
		if not np.any(np.isfinite(error_values) & (error_values >= 0)):
			continue
		column_index = chart_data.max_column + 1
		chart_data.cell(row=1, column=column_index, value=error_column)
		for row_index, error_value in enumerate(error_values, start=2):
			chart_data.cell(
				row=row_index,
				column=column_index,
				value=float(error_value) if math.isfinite(error_value) and error_value >= 0 else 0.0,
			)
		error_bar_columns.append((column_index, direction))
	if error_bar_columns:
		selected_error_bars = next(
			(error for error in error_bar_columns if error[1] == "y"),
			error_bar_columns[0],
		)
		_append_error_bars(series, chart_data, selected_error_bars[0], 2, last_row, selected_error_bars[1])

	x_zero = _use_zero_origin(spec.origin, spec.x_quantity, x_data)
	y_zero = _use_zero_origin(spec.origin, spec.y_quantity, y_data)
	x_min, x_max = _axis_bounds(x_data, spec.x_min, spec.x_max, x_zero)
	y_min, y_max = _axis_bounds(y_data, spec.y_min, spec.y_max, y_zero)
	chart.x_axis.scaling.min = x_min
	chart.x_axis.scaling.max = x_max
	chart.y_axis.scaling.min = y_min
	chart.y_axis.scaling.max = y_max

	data_sheet.add_chart(chart, f"{get_column_letter(data_sheet.max_column + 2)}2")
	return chart


def output_path_for(input_path: Path) -> Path:
	"""Wählt einen neuen Namen, ohne vorhandene Dateien zu überschreiben."""
	base = input_path.with_name(f"{input_path.stem}_physik{input_path.suffix.lower()}")
	return _non_existing_path(base)


def _non_existing_path(path: Path) -> Path:
	base = path
	candidate = base
	index = 2
	while candidate.exists():
		candidate = base.with_name(f"{base.stem}_{index}{base.suffix}")
		index += 1
	return candidate


def save_workbook_with_chart(
	input_path: Path,
	frame: pd.DataFrame,
	sheet: int | str,
	spec: ChartSpecification,
	fit: str = "none",
	output_path: Path | None = None,
) -> Path:
	"""Speichert eine Arbeitsmappenkopie mit nativem Excel-Diagramm."""
	input_path = input_path.resolve()
	if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
		raise ValueError("Es werden nur .xlsx- und .xlsm-Dateien unterstützt.")
	output = output_path or output_path_for(input_path)
	output = output.resolve()
	if output == input_path:
		raise ValueError("Die Ausgabedatei darf nicht die Eingabedatei überschreiben.")
	if output.suffix.lower() != input_path.suffix.lower():
		raise ValueError("Eingabe- und Ausgabeformat müssen übereinstimmen.")
	if output_path is not None:
		output = _non_existing_path(output)
	output.parent.mkdir(parents=True, exist_ok=True)
	workbook = load_openpyxl_workbook(input_path, keep_vba=input_path.suffix.lower() == ".xlsm")
	try:
		if isinstance(sheet, int):
			if sheet < 0 or sheet >= len(workbook.worksheets):
				raise ValueError(f"Tabellenblatt-Index {sheet} ist nicht vorhanden.")
			data_sheet = workbook.worksheets[sheet]
		else:
			if sheet not in workbook.sheetnames:
				raise ValueError(f"Tabellenblatt {sheet!r} ist nicht vorhanden.")
			data_sheet = workbook[sheet]
		create_excel_chart(workbook, data_sheet, frame, spec, fit)
		workbook.save(output)
	finally:
		workbook.close()
		vba_archive = getattr(workbook, "vba_archive", None)
		if vba_archive is not None:
			vba_archive.close()
	return output