"""Gemeinsame Datenmodelle für Analyse und Diagrammausgabe."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal


ChartType = Literal["scatter", "scatter_lines", "line", "line_scatter", "bar", "column"]


@dataclass
class ChartSpecification:
	x_column: str = ""
	y_column: str = ""
	x_label: str = ""
	y_label: str = ""
	x_unit: str = ""
	y_unit: str = ""
	chart_type: ChartType = "scatter"
	x_min: float | None = None
	x_max: float | None = None
	y_min: float | None = None
	y_max: float | None = None
	origin: bool | None = False
	trendline: str = "none"
	polynomial_degree: int = 2
	moving_average_period: int = 3
	x_error_column: str | None = None
	y_error_column: str | None = None
	independent_variable: str | None = None
	dependent_variable: str | None = None
	show_points: bool = True
	connect_points: bool = False
	confidence: float | None = None
	reason: str = ""
	needs_user_input: bool = False
	user_input_reason: str = ""
	user_input_options: list[dict] = field(default_factory=list)
	ai_response: dict[str, Any] = field(default_factory=dict, repr=False)
	uncertainties: list[str] = field(default_factory=list)