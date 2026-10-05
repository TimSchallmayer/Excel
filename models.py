"""Gemeinsame Datenmodelle für Analyse und Diagrammausgabe."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ChartSpecification:
	x_column: str = ""
	y_column: str = ""
	x_label: str = ""
	y_label: str = ""
	x_unit: str = ""
	y_unit: str = ""
	chart_type: str = "scatter"
	x_min: float | None = None
	x_max: float | None = None
	y_min: float | None = None
	y_max: float | None = None
	origin: bool | None = False
	trendline: str = "none"
	x_error_column: str | None = None
	y_error_column: str | None = None
	independent_variable: str | None = None
	dependent_variable: str | None = None
	show_points: bool = True
	connect_points: bool = False
	confidence: float | None = None
	reason: str = ""
	needs_user_input: bool = False
	uncertainties: list[str] = field(default_factory=list)