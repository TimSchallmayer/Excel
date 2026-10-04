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
	origin: bool | str | None = "auto"
	trendline: str = "none"
	x_error_column: str | None = None
	y_error_column: str | None = None
	x_quantity: str | None = None
	y_quantity: str | None = None
	independent_variable: str | None = None
	dependent_variable: str | None = None
	show_points: bool = True
	connect_points: bool = False
	confidence: float | None = None
	reason: str = ""


@dataclass(frozen=True)
class ColumnInfo:
	name: str
	quantity: str | None
	unit: str
	numeric: bool
	is_error: bool
	finite_count: int
	unique_count: int


@dataclass(frozen=True)
class RelationshipHint:
	x_quantity: str
	y_quantity: str
	formula: str
	context: str
	confidence: float


@dataclass(frozen=True)
class ChartCandidate:
	x_column: str
	y_column: str
	relationship: str
	reason: str
	confidence: float