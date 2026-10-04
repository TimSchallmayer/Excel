"""Einlesen und grundlegende Prüfung von Excel-Messwerttabellen."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


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