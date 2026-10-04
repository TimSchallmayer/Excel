"""Gemeinsame Datenmodelle für Analyse und Diagrammausgabe."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ChartSpecification:
	x_column: str
	y_column: str
	x_label: str
	y_label: str
	x_unit: str = ""
	y_unit: str = ""
	chart_type: str = "scatter"
	x_min: float | None = None
	x_max: float | None = None
	y_min: float | None = None
	y_max: float | None = None
	origin: str = "auto"
	trendline: str = "none"
	x_error_column: str | None = None
	y_error_column: str | None = None
	x_quantity: str | None = None
	y_quantity: str | None = None