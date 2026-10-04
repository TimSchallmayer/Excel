"""OpenAI-kompatible KI-Anbindung mit strikt validierter JSON-Ausgabe."""

from __future__ import annotations

import json
import math
import re
from typing import Any, Mapping
from urllib.parse import urlsplit

import numpy as np
import pandas as pd
import requests

from models import ChartSpecification
from physics import identify_column, relationship_hints
from prompts import ANALYSIS_SYSTEM_PROMPT

AI_TIMEOUT_SECONDS = 30
MAX_PROMPT_COLUMNS = 40
MAX_PROMPT_ROWS = 30
MAX_CELL_CHARS = 120


class AIServiceError(ValueError):
	"""Verständlicher, schlüsselinhaltsfreier KI-Konfigurations-/API-Fehler."""


def _validate_config(config: Mapping[str, str]) -> tuple[str, str, str]:
	endpoint = str(config.get("endpoint", "")).strip()
	api_key = str(config.get("api_key", "")).strip()
	model = str(config.get("model", "")).strip()
	if not endpoint:
		raise AIServiceError("In config.json fehlt der Endpoint.")
	if not model:
		raise AIServiceError("In config.json fehlt das Modell.")
	parsed = urlsplit(endpoint)
	if parsed.scheme not in {"http", "https"} or not parsed.netloc:
		raise AIServiceError("Der Endpoint muss eine gültige HTTP- oder HTTPS-URL sein.")
	if parsed.username is not None or parsed.password is not None:
		raise AIServiceError("Zugangsdaten dürfen nicht in der Endpoint-URL stehen.")
	return endpoint, api_key, model


def _error_for_status(status_code: int) -> str:
	messages = {
		400: "HTTP 400 – Anfrage oder Modellparameter werden abgelehnt.",
		401: "HTTP 401 – API-Key fehlt oder ist ungültig.",
		403: "HTTP 403 – Zugriff auf Endpoint oder Modell verweigert.",
		404: "HTTP 404 – Endpoint oder Modell nicht gefunden.",
		429: "HTTP 429 – Rate-Limit oder Kontingent erreicht.",
	}
	if status_code in messages:
		return messages[status_code]
	if status_code >= 500:
		return f"HTTP {status_code} – der KI-Server meldet einen internen Fehler."
	return f"HTTP {status_code} – die KI-Anfrage wurde abgelehnt."


def _post_chat_completion(
	config: Mapping[str, str],
	messages: list[dict[str, str]],
) -> Mapping[str, Any]:
	endpoint, api_key, model = _validate_config(config)
	headers = {"Content-Type": "application/json"}
	if api_key:
		headers["Authorization"] = f"Bearer {api_key}"
	payload = {"model": model, "messages": messages, "temperature": 0}
	try:
		response = requests.post(
			endpoint,
			json=payload,
			headers=headers,
			timeout=AI_TIMEOUT_SECONDS,
			allow_redirects=False,
		)
	except requests.Timeout as exc:
		raise AIServiceError(f"Zeitüberschreitung nach {AI_TIMEOUT_SECONDS} Sekunden bei der KI-Anfrage.") from exc
	except requests.ConnectionError as exc:
		raise AIServiceError("Keine Verbindung zum KI-Endpoint möglich (DNS- oder Netzwerkfehler).") from exc
	except requests.RequestException as exc:
		raise AIServiceError("Die KI-Anfrage konnte wegen eines HTTP-Transportfehlers nicht gesendet werden.") from exc

	if not 200 <= response.status_code < 300:
		raise AIServiceError(_error_for_status(response.status_code))
	if not response.content:
		raise AIServiceError("Der KI-Server hat eine leere Antwort zurückgegeben.")
	try:
		body = response.json()
	except (ValueError, requests.exceptions.JSONDecodeError) as exc:
		raise AIServiceError("Der KI-Server hat keine gültige JSON-HTTP-Antwort zurückgegeben.") from exc
	if not isinstance(body, dict):
		raise AIServiceError("Die HTTP-Antwort des KI-Servers hat ein unerwartetes Format.")
	return body


def _response_text(body: Mapping[str, Any]) -> str:
	choices = body.get("choices")
	if not isinstance(choices, list) or not choices or not isinstance(choices[0], dict):
		raise AIServiceError("In der KI-Antwort fehlt das erwartete Feld choices[0].")
	message = choices[0].get("message")
	if not isinstance(message, dict):
		raise AIServiceError("In der KI-Antwort fehlt das erwartete Feld message.")
	content = message.get("content")
	if isinstance(content, str) and content.strip():
		return content.strip()
	if isinstance(content, list):
		parts = [part.get("text", "") for part in content if isinstance(part, dict)]
		text = "".join(str(part) for part in parts).strip()
		if text:
			return text
	raise AIServiceError("Die KI hat keinen Antworttext zurückgegeben.")


def _extract_json(text: str) -> dict[str, Any]:
	content = text.strip()
	if content.startswith("```"):
		content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE).strip()
	try:
		parsed = json.loads(content)
	except json.JSONDecodeError:
		parsed = None
	if isinstance(parsed, dict):
		return parsed

	start = content.find("{")
	if start < 0:
		raise AIServiceError("Die KI-Antwort enthält kein JSON-Objekt.")
	depth = 0
	in_string = False
	escaped = False
	for index in range(start, len(content)):
		character = content[index]
		if in_string:
			if escaped:
				escaped = False
			elif character == "\\":
				escaped = True
			elif character == '"':
				in_string = False
		elif character == '"':
			in_string = True
		elif character == "{":
			depth += 1
		elif character == "}":
			depth -= 1
			if depth == 0:
				try:
					parsed = json.loads(content[start:index + 1])
				except json.JSONDecodeError as exc:
					raise AIServiceError("Die KI-Antwort enthält ungültiges JSON.") from exc
				if isinstance(parsed, dict):
					return parsed
				break
	raise AIServiceError("Die KI-Antwort enthält kein vollständiges JSON-Objekt.")


def _finite_number(value: Any, field: str) -> float | None:
	if value is None:
		return None
	if isinstance(value, bool) or not isinstance(value, (int, float)):
		raise AIServiceError(f"Das Feld {field} muss eine endliche Zahl oder null sein.")
	number = float(value)
	if not math.isfinite(number):
		raise AIServiceError(f"Das Feld {field} muss eine endliche Zahl oder null sein.")
	return number


def validate_ai_response(
	response: Mapping[str, Any],
	data: pd.DataFrame,
	columns: list[str],
) -> ChartSpecification:
	"""Validiert erlaubte Felder und erzeugt daraus eine ChartSpecification."""
	required = {
		"independent_variable", "dependent_variable", "x_column", "y_column",
		"chart_type", "x_axis_label", "y_axis_label", "x_unit", "y_unit",
		"x_min", "x_max", "y_min", "y_max", "origin", "show_points",
		"trendline", "x_error_column", "y_error_column", "confidence", "reason",
	}
	missing = required.difference(response)
	if missing:
		raise AIServiceError("In der KI-Diagnose fehlen Pflichtfelder: " + ", ".join(sorted(missing)))

	allowed_columns = {str(column) for column in data.columns}
	if not isinstance(response["x_column"], str) or response["x_column"] not in allowed_columns:
		raise AIServiceError("Die von der KI gewählte x-Spalte existiert nicht in der Tabelle.")
	if not isinstance(response["y_column"], str) or response["y_column"] not in allowed_columns:
		raise AIServiceError("Die von der KI gewählte y-Spalte existiert nicht in der Tabelle.")
	x_column, y_column = response["x_column"], response["y_column"]
	if x_column == y_column:
		raise AIServiceError("Die x- und y-Spalte müssen unterschiedlich sein.")
	if x_column not in columns or y_column not in columns:
		raise AIServiceError("Die von der KI gewählte x- oder y-Spalte enthält keine numerischen Messwerte.")
	if identify_column(x_column)[2] or identify_column(y_column)[2]:
		raise AIServiceError("Eine Fehlerwert-Spalte darf nicht als Messachse verwendet werden.")

	values: dict[str, np.ndarray] = {}
	for column in (x_column, y_column):
		numeric = pd.to_numeric(data[column], errors="coerce").to_numpy(dtype=float)
		finite = numeric[np.isfinite(numeric)]
		if finite.size < 2:
			raise AIServiceError(f"Die Spalte {column!r} hat weniger als zwei endliche Messwerte.")
		values[column] = finite
	if np.unique(values[x_column]).size < 2:
		raise AIServiceError("Die x-Spalte benötigt mindestens zwei unterschiedliche Messwerte.")

	chart_type = response["chart_type"]
	if not isinstance(chart_type, str) or chart_type not in {"scatter", "line", "line_scatter"}:
		raise AIServiceError("Die Diagrammart muss scatter, line oder line_scatter sein.")
	trendline = response["trendline"]
	if not isinstance(trendline, str) or trendline not in {"none", "linear", "quadratic", "cubic"}:
		raise AIServiceError("Die Trendlinie muss none, linear, quadratic oder cubic sein.")
	if trendline == "linear" and np.unique(values[x_column]).size < 2:
		raise AIServiceError("Ein linearer Fit benötigt mindestens zwei unterschiedliche x-Werte.")
	if trendline == "quadratic" and np.unique(values[x_column]).size < 3:
		raise AIServiceError("Ein quadratischer Fit benötigt mindestens drei unterschiedliche x-Werte.")
	if trendline == "cubic" and np.unique(values[x_column]).size < 4:
		raise AIServiceError("Ein kubischer Fit benötigt mindestens vier unterschiedliche x-Werte.")

	string_fields = (
		"independent_variable", "dependent_variable", "x_axis_label", "y_axis_label", "x_unit", "y_unit", "reason",
	)
	for field in string_fields:
		if not isinstance(response[field], str) or len(response[field]) > 300:
			raise AIServiceError(f"Das Feld {field} muss ein Text mit höchstens 300 Zeichen sein.")
	if not isinstance(response["show_points"], bool):
		raise AIServiceError("Das Feld show_points muss true oder false sein.")
	origin = response["origin"]
	if origin is not None and not isinstance(origin, bool):
		raise AIServiceError("Das Feld origin muss true, false oder null sein.")
	confidence = _finite_number(response["confidence"], "confidence")
	if confidence is None or not 0 <= confidence <= 1:
		raise AIServiceError("confidence muss zwischen 0 und 1 liegen.")

	axis_limits = {
		field: _finite_number(response[field], field)
		for field in ("x_min", "x_max", "y_min", "y_max")
	}
	for minimum, maximum in (("x_min", "x_max"), ("y_min", "y_max")):
		if axis_limits[minimum] is not None and axis_limits[maximum] is not None:
			if axis_limits[minimum] >= axis_limits[maximum]:
				raise AIServiceError(f"{minimum} muss kleiner als {maximum} sein.")

	error_columns: dict[str, str | None] = {}
	for field in ("x_error_column", "y_error_column"):
		column = response[field]
		if column is not None:
			if not isinstance(column, str) or column not in allowed_columns:
				raise AIServiceError(f"Die Fehlerwert-Spalte aus {field} existiert nicht.")
			if column in {x_column, y_column} or not identify_column(column)[2]:
				raise AIServiceError(f"{field} muss auf eine erkannte Fehlerwert-Spalte zeigen.")
			error_values = pd.to_numeric(data[column], errors="coerce").to_numpy(dtype=float)
			if not np.any(np.isfinite(error_values) & (error_values >= 0)):
				raise AIServiceError(f"Die Fehlerwert-Spalte {column!r} enthält keine gültigen Werte.")
		error_columns[field] = column

	x_quantity, x_unit, _ = identify_column(x_column)
	y_quantity, y_unit, _ = identify_column(y_column)
	x_label = response["x_axis_label"]
	y_label = response["y_axis_label"]
	if x_quantity and x_unit:
		x_label, x_unit = f"{x_quantity} [{x_unit}]", x_unit
	if y_quantity and y_unit:
		y_label, y_unit = f"{y_quantity} [{y_unit}]", y_unit

	return ChartSpecification(
		x_column=x_column,
		y_column=y_column,
		x_label=x_label or x_column,
		y_label=y_label or y_column,
		x_unit=x_unit or response["x_unit"],
		y_unit=y_unit or response["y_unit"],
		chart_type=chart_type,
		x_min=axis_limits["x_min"],
		x_max=axis_limits["x_max"],
		y_min=axis_limits["y_min"],
		y_max=axis_limits["y_max"],
		origin="auto" if origin is None else origin,
		trendline=trendline,
		x_error_column=error_columns["x_error_column"],
		y_error_column=error_columns["y_error_column"],
		x_quantity=x_quantity or response["independent_variable"] or None,
		y_quantity=y_quantity or response["dependent_variable"] or None,
		independent_variable=response["independent_variable"] or None,
		dependent_variable=response["dependent_variable"] or None,
		show_points=response["show_points"],
		confidence=confidence,
		reason=response["reason"],
	)


def _cell_value(value: Any) -> Any:
	if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
		return None
	if isinstance(value, (int, float, np.integer, np.floating)):
		number = float(value)
		return number if math.isfinite(number) else None
	if hasattr(value, "isoformat"):
		return value.isoformat()
	return str(value)[:MAX_CELL_CHARS]


def _table_context(data: pd.DataFrame, columns: list[str]) -> dict[str, Any]:
	all_columns = [str(column) for column in data.columns]
	eligible_columns = set(columns)
	numeric_column_set: set[str] = set()
	error_columns: set[str] = set()
	column_info = []
	for column in all_columns:
		quantity, unit, is_error = identify_column(column)
		values = pd.to_numeric(data[column], errors="coerce").to_numpy(dtype=float)
		if np.isfinite(values).any():
			numeric_column_set.add(column)
		if is_error and np.isfinite(values).any():
			error_columns.add(column)
		column_info.append({
			"name": column,
			"numeric": column in numeric_column_set,
			"eligible_measurement_axis": column in eligible_columns,
			"quantity": quantity,
			"unit": unit or None,
			"possible_error_values": is_error,
			"locally_confident": bool(quantity and not is_error),
		})
	data_columns = [
		column for column in all_columns
		if column in numeric_column_set and (column in eligible_columns or column in error_columns)
	][:MAX_PROMPT_COLUMNS]
	column_info = column_info[:MAX_PROMPT_COLUMNS]

	if len(data) <= MAX_PROMPT_ROWS:
		indexes = list(range(len(data)))
	else:
		indexes = np.linspace(0, len(data) - 1, MAX_PROMPT_ROWS, dtype=int).tolist()
	rows = [
		{column: _cell_value(data.iloc[index][column]) for column in data_columns}
		for index in indexes
	]
	hints = relationship_hints([info["quantity"] for info in column_info])
	return {
		"columns": column_info,
		"local_relationship_hints": [
			{
				"x_quantity": hint.x_quantity,
				"y_quantity": hint.y_quantity,
				"formula": hint.formula,
				"context": hint.context,
				"confidence": hint.confidence,
			}
			for hint in hints
		],
		"sampled_data_columns": data_columns,
		"table_row_count": int(len(data)),
		"sampled_rows_count": len(rows),
		"rows_are_sampled": len(data) > MAX_PROMPT_ROWS,
		"rows": rows,
	}


def analyze_with_ai(
	data: pd.DataFrame,
	columns: list[str],
	config: Mapping[str, str],
) -> ChartSpecification:
	"""Analysiert eine kompakte Messwertvorschau und validiert die JSON-Diagnose."""
	context = _table_context(data, columns)
	body = _post_chat_completion(config, [
		{"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
		{"role": "user", "content": json.dumps(context, ensure_ascii=False, allow_nan=False)},
	])
	response = _extract_json(_response_text(body))
	return validate_ai_response(response, data, columns)


def test_ai_connection(config: Mapping[str, str]) -> None:
	"""Sendet eine minimale Anfrage; löst bei Fehlern eine AIServiceError aus."""
	body = _post_chat_completion(config, [
		{"role": "system", "content": "Reply with a short connection confirmation."},
		{"role": "user", "content": "Connection test."},
	])
	_response_text(body)