"""OpenAI-kompatible KI-Anbindung mit strikt validierter JSON-Ausgabe."""

from __future__ import annotations

import json
import logging
import math
import os
import re
from typing import Any, Mapping
from urllib.parse import urlsplit

import numpy as np
import pandas as pd
import requests

from models import ChartSpecification
from prompts import ANALYSIS_SYSTEM_PROMPT


AI_TIMEOUT_SECONDS = 30
MAX_PROMPT_COLUMNS = 40
MAX_PROMPT_ROWS = 30
MAX_CELL_CHARS = 120
MAX_TEXT_FIELD_CHARS = 300
MAX_REASON_CHARS = 200

VALID_CHART_TYPES = {"scatter", "scatter_lines", "line", "line_scatter", "bar", "column"}

VALID_TRENDLINES = {
	"none",
	"linear",
	"quadratic",
	"cubic",
	"polynomial",
	"exponential",
	"logarithmic",
	"power",
	"moving_average",
}

VALID_USER_INPUT_TYPES = {
	"x_column",
	"y_column",
	"chart_type",
	"trendline",
}

_UNIT_SYMBOLS = {
	"A", "bar", "CHF", "cm", "cm²", "cm³", "cm/s", "d", "deg", "EUR",
	"g", "g/L", "GBP", "h", "Hz", "J", "K", "kg", "kg/m³", "km",
	"km/h", "kN", "kPa", "kW", "L", "m", "m²", "m³", "m/s", "m/s²",
	"m/s^2", "mg/L", "min", "mm", "mol", "mol/L", "ms", "N", "Pa",
	"rad", "rpm", "s", "USD", "V", "W", "°C", "°F", "µm", "μm", "Ω", "%",
}
_BRACKETED_UNIT = re.compile(r"\s*\[\s*([^\[\]]+?)\s*\]\s*$")
_PARENTHESIZED_UNIT = re.compile(r"\s*\(\s*([^()\s]+)\s*\)\s*$")
_UNIT_IN_NAME = re.compile(r"\s+in\s+([^\s]+)\s*$", re.IGNORECASE)


class AIServiceError(ValueError):
	"""Verständlicher, schlüsselinhaltsfreier KI-Konfigurations-/API-Fehler."""


def _response_debug_metadata(
	body: Mapping[str, Any],
	request_metadata: Mapping[str, Any],
) -> dict[str, Any]:
	choices = body.get("choices")
	choice_count = len(choices) if isinstance(choices, list) else None
	first_choice = (
		choices[0]
		if isinstance(choices, list)
		and choices
		and isinstance(choices[0], Mapping)
		else None
	)
	message = (
		first_choice.get("message")
		if isinstance(first_choice, Mapping)
		else None
	)
	message = message if isinstance(message, Mapping) else None
	content = message.get("content") if message is not None else None
	refusal = message.get("refusal") if message is not None else None
	finish_reason = (
		first_choice.get("finish_reason")
		if first_choice is not None
		else None
	)
	if not isinstance(finish_reason, str) or not re.fullmatch(
		r"[A-Za-z0-9_-]{1,64}",
		finish_reason,
	):
		finish_reason = "other" if finish_reason is not None else None

	try:
		content_length = len(content) if content is not None else None
	except TypeError:
		content_length = None

	return {
		"http_status": request_metadata.get("http_status"),
		"model": request_metadata.get("model"),
		"api_model": (
			body.get("model")
			if isinstance(body.get("model"), str)
			and re.fullmatch(r"[A-Za-z0-9._:/-]{1,128}", body["model"])
			else "other" if body.get("model") is not None else None
		),
		"response_format_set": request_metadata.get(
			"response_format_set",
			False,
		),
		"body_type": type(body).__name__,
		"top_level_keys": sorted(str(key) for key in body),
		"choices_type": type(choices).__name__ if choices is not None else None,
		"choices_count": choice_count,
		"first_choice_keys": (
			sorted(str(key) for key in first_choice)
			if first_choice is not None
			else None
		),
		"message_type": (
			type(message).__name__ if message is not None else None
		),
		"message_keys": (
			sorted(str(key) for key in message)
			if message is not None
			else None
		),
		"content_type": type(content).__name__ if content is not None else None,
		"content_length": content_length,
		"content_empty_or_none": (
			content is None
			or content == ""
			or content == []
		),
		"content_contains_opening_brace": (
			"{" in content if isinstance(content, str) else None
		),
		"refusal_present": (
			message is not None and "refusal" in message
		),
		"refusal_type": (
			type(refusal).__name__
			if message is not None and "refusal" in message
			else None
		),
		"finish_reason": finish_reason,
		"structured_content_shape": (
			{
				"type": type(content).__name__,
				"item_types": (
					[type(item).__name__ for item in content[:20]]
					if isinstance(content, list)
					else None
				),
				"item_count": (
					len(content) if isinstance(content, list) else None
				),
			}
			if isinstance(content, (Mapping, list))
			else None
		),
	}


def _log_response_debug(
	body: Mapping[str, Any],
	request_metadata: Mapping[str, Any],
	stage: str,
	failure_kind: str,
) -> None:
	if os.environ.get("PHYSIK_AI_RESPONSE_DEBUG") != "1":
		return

	metadata = _response_debug_metadata(body, request_metadata)
	metadata["failure_stage"] = stage
	metadata["failure_kind"] = failure_kind
	logging.getLogger(__name__).warning(
		"KI-Antwort-Diagnose (nur Metadaten): %s",
		json.dumps(metadata, ensure_ascii=True, sort_keys=True),
	)


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
		raise AIServiceError(
			"Der Endpoint muss eine gültige HTTP- oder HTTPS-URL sein."
		)

	if parsed.username is not None or parsed.password is not None:
		raise AIServiceError(
			"Zugangsdaten dürfen nicht in der Endpoint-URL stehen."
		)

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
		return (
			f"HTTP {status_code} – der KI-Server meldet einen internen Fehler."
		)

	return f"HTTP {status_code} – die KI-Anfrage wurde abgelehnt."


def _post_chat_completion(
	config: Mapping[str, str],
	messages: list[dict[str, str]],
	*,
	json_mode: bool = False,
	debug_metadata: dict[str, Any] | None = None,
) -> Mapping[str, Any]:
	endpoint, api_key, model = _validate_config(config)
	if debug_metadata is not None:
		debug_metadata.update({
			"model": model,
			"response_format_set": json_mode,
		})

	headers = {"Content-Type": "application/json"}

	if api_key:
		headers["Authorization"] = f"Bearer {api_key}"

	payload = {
		"model": model,
		"messages": messages,
		"temperature": 0,
	}

	if json_mode:
		payload["response_format"] = {"type": "json_object"}

	try:
		response = requests.post(
			endpoint,
			json=payload,
			headers=headers,
			timeout=AI_TIMEOUT_SECONDS,
			allow_redirects=False,
		)
	except requests.Timeout as exc:
		raise AIServiceError(
			f"Zeitüberschreitung nach {AI_TIMEOUT_SECONDS} Sekunden "
			"bei der KI-Anfrage."
		) from exc
	except requests.ConnectionError as exc:
		raise AIServiceError(
			"Keine Verbindung zum KI-Endpoint möglich "
			"(DNS- oder Netzwerkfehler)."
		) from exc
	except requests.RequestException as exc:
		raise AIServiceError(
			"Die KI-Anfrage konnte wegen eines HTTP-Transportfehlers "
			"nicht gesendet werden."
		) from exc

	if not 200 <= response.status_code < 300:
		raise AIServiceError(_error_for_status(response.status_code))
	if debug_metadata is not None:
		debug_metadata["http_status"] = response.status_code

	if not response.content:
		raise AIServiceError(
			"Der KI-Server hat eine leere Antwort zurückgegeben."
		)

	try:
		body = response.json()
	except (ValueError, requests.exceptions.JSONDecodeError) as exc:
		raise AIServiceError(
			"Der KI-Server hat keine gültige JSON-HTTP-Antwort zurückgegeben."
		) from exc

	if not isinstance(body, dict):
		raise AIServiceError(
			"Die HTTP-Antwort des KI-Servers hat ein unerwartetes Format."
		)

	return body


def _response_text(
	body: Mapping[str, Any],
) -> str | Mapping[str, Any]:
	choices = body.get("choices")

	if (
		not isinstance(choices, list)
		or not choices
		or not isinstance(choices[0], dict)
	):
		if {"x_column", "y_column", "needs_user_input"}.intersection(body):
			return body
		raise AIServiceError(
			"In der KI-Antwort fehlt das erwartete Feld choices[0]."
		)

	message = choices[0].get("message")

	if not isinstance(message, dict):
		raise AIServiceError(
			"In der KI-Antwort fehlt das erwartete Feld message."
		)

	parsed = message.get("parsed")
	if isinstance(parsed, Mapping):
		return parsed

	content = message.get("content")

	if isinstance(content, str) and content.strip():
		return content.strip()

	if isinstance(content, Mapping):
		return content

	if isinstance(content, list):
		for part in content:
			if isinstance(part, dict):
				structured = part.get("json")
				if isinstance(structured, Mapping):
					return structured

		parts = [
			part.get("text", "")
			for part in content
			if isinstance(part, dict)
		]

		text = "".join(str(part) for part in parts).strip()

		if text:
			return text

	raise AIServiceError(
		"Die KI hat keinen Antworttext zurückgegeben."
	)


def _extract_json(
	text: str | Mapping[str, Any],
) -> dict[str, Any]:
	if isinstance(text, Mapping):
		return dict(text)

	content = text.strip()

	if content.startswith("```"):
		content = re.sub(
			r"^```(?:json)?\s*|\s*```$",
			"",
			content,
			flags=re.IGNORECASE,
		).strip()

	try:
		parsed = json.loads(content)
	except json.JSONDecodeError:
		parsed = None

	if isinstance(parsed, dict):
		return parsed

	starts = [
		match.start()
		for match in re.finditer(r"\{", content)
	]
	if not starts:
		raise AIServiceError(
			"Die KI-Antwort enthält kein JSON-Objekt."
		)

	decoder = json.JSONDecoder()
	for start in starts:
		try:
			parsed, _ = decoder.raw_decode(content, start)
		except json.JSONDecodeError:
			continue
		if isinstance(parsed, dict):
			return parsed

	raise AIServiceError(
		"Die KI-Antwort enthält kein gültiges JSON-Objekt."
	)


def _extract_response_json(
	body: Mapping[str, Any],
	request_metadata: Mapping[str, Any],
	attempt: int = 1,
) -> dict[str, Any]:
	try:
		response_text = _response_text(body)
	except AIServiceError:
		_log_response_debug(
			body,
			request_metadata,
			f"response_text_attempt_{attempt}",
			"no_response_text",
		)
		raise

	try:
		return _extract_json(response_text)
	except AIServiceError as exc:
		failure_kind = (
			"no_json_object"
			if str(exc) == "Die KI-Antwort enthält kein JSON-Objekt."
			else "invalid_or_incomplete_json"
		)
		_log_response_debug(
			body,
			request_metadata,
			f"extract_json_attempt_{attempt}",
			failure_kind,
		)
		raise


def _request_json_response(
	config: Mapping[str, str],
	messages: list[dict[str, str]],
) -> dict[str, Any]:
	debug_metadata: dict[str, Any] = {}
	body = _post_chat_completion(
		config,
		messages,
		json_mode=True,
		debug_metadata=debug_metadata,
	)
	try:
		return _extract_response_json(body, debug_metadata)
	except AIServiceError as exc:
		if str(exc) != "Die KI-Antwort enthält kein JSON-Objekt.":
			raise

	retry_messages = [
		dict(message)
		for message in messages
	]
	retry_instruction = (
		"Deine Antwort muss ausschließlich ein gültiges JSON-Objekt sein.\n"
		"Kein erklärender Text.\n"
		"Kein Markdown.\n"
		"Keine Code-Fences.\n"
		"Keine Einleitung.\n"
		"Keine Ausgabe außerhalb des JSON-Objekts."
	)
	system_message = next(
		(
			message
			for message in retry_messages
			if message.get("role") == "system"
		),
		None,
	)
	if system_message is not None:
		system_message["content"] += "\n\n" + retry_instruction
	else:
		retry_messages.append({
			"role": "system",
			"content": retry_instruction,
		})

	retry_metadata: dict[str, Any] = {}
	retry_body = _post_chat_completion(
		config,
		retry_messages,
		json_mode=True,
		debug_metadata=retry_metadata,
	)
	return _extract_response_json(retry_body, retry_metadata, attempt=2)


def _finite_number(
	value: Any,
	field: str,
) -> float | None:
	if value is None:
		return None

	if isinstance(value, bool) or not isinstance(
		value,
		(int, float),
	):
		raise AIServiceError(
			f"Das Feld {field} muss eine endliche Zahl oder null sein."
		)

	number = float(value)

	if not math.isfinite(number):
		raise AIServiceError(
			f"Das Feld {field} muss eine endliche Zahl oder null sein."
		)

	return number


def _label_base_and_unit(label: str) -> tuple[str, str]:
	for pattern in (_BRACKETED_UNIT, _PARENTHESIZED_UNIT, _UNIT_IN_NAME):
		match = pattern.search(label)
		if match is None:
			continue
		unit = match.group(1).strip()
		if pattern in {_PARENTHESIZED_UNIT, _UNIT_IN_NAME} and unit not in _UNIT_SYMBOLS:
			continue
		if unit:
			return label[:match.start()].strip(), unit
	return label.strip(), ""


def _axis_label_and_unit(
	column: str,
	label: str,
	unit: str,
) -> tuple[str, str]:
	column_base, column_unit = _label_base_and_unit(column)
	label_base, label_unit = _label_base_and_unit(label)
	selected_unit = column_unit or label_unit or unit.strip()
	selected_base = column_base if column_unit else label_base
	if not selected_base:
		selected_base = column_base or column
	if selected_unit:
		return f"{selected_base} [{selected_unit}]", selected_unit
	return selected_base, ""


def _validate_user_input_options(
	response: Mapping[str, Any],
	allowed_columns: set[str],
	numeric_columns: set[str],
) -> tuple[
	bool,
	str,
	list[dict[str, Any]],
]:
	needs_user_input = response["needs_user_input"]

	if not isinstance(needs_user_input, bool):
		raise AIServiceError(
			"needs_user_input muss true oder false sein."
		)

	user_input_reason = response["user_input_reason"]

	if not isinstance(user_input_reason, str):
		raise AIServiceError(
			"user_input_reason muss ein Text sein."
		)

	user_input_options = response["user_input_options"]

	if not isinstance(user_input_options, list):
		raise AIServiceError(
			"user_input_options muss eine Liste sein."
		)

	seen_types: set[str] = set()
	for option in user_input_options:
		if not isinstance(option, dict):
			raise AIServiceError(
				"Jede Benutzeroption muss ein Objekt sein."
			)

		option_type = option.get("type")
		question = option.get("question")
		options = option.get("options")

		if (
			not isinstance(option_type, str)
			or option_type not in VALID_USER_INPUT_TYPES
		):
			raise AIServiceError(
				"Unbekannter Typ in user_input_options."
			)

		if option_type in seen_types:
			raise AIServiceError(
				"Jeder Typ darf in user_input_options nur "
				"einmal vorkommen."
			)
		seen_types.add(option_type)

		if not isinstance(question, str) or not question.strip():
			raise AIServiceError(
				"Jede Benutzeroption benötigt eine Frage."
			)

		if not isinstance(options, list) or not options:
			raise AIServiceError(
				"Jede Benutzeroption benötigt mindestens "
				"eine Auswahlmöglichkeit."
			)

		if not all(
			isinstance(value, str)
			for value in options
		):
			raise AIServiceError(
				"Alle Benutzeroptionen müssen Texte sein."
			)

		if option_type in {"x_column", "y_column"}:
			invalid = [
				value
				for value in options
				if value not in allowed_columns
			]

			if invalid:
				raise AIServiceError(
					"Die KI hat eine nicht vorhandene Spalte "
					"als Benutzeroption vorgeschlagen."
				)
			non_numeric = [
				value
				for value in options
				if value not in numeric_columns
			]

			if option_type == "y_column" and non_numeric:
				raise AIServiceError(
					"Die KI hat eine nicht numerische Spalte "
					"als y-Benutzeroption vorgeschlagen."
				)

		elif option_type == "chart_type":
			invalid = [
				value
				for value in options
				if value not in VALID_CHART_TYPES
			]

			if invalid:
				raise AIServiceError(
					"Die KI hat einen ungültigen Diagrammtyp "
					"als Benutzeroption vorgeschlagen."
				)

		elif option_type == "trendline":
			invalid = [
				value
				for value in options
				if value not in VALID_TRENDLINES
			]

			if invalid:
				raise AIServiceError(
					"Die KI hat eine ungültige Trendlinie "
					"als Benutzeroption vorgeschlagen."
				)

	if needs_user_input and not user_input_options:
		raise AIServiceError(
			"needs_user_input ist true, aber es wurden "
			"keine Benutzeroptionen angegeben."
		)

	if needs_user_input and not user_input_reason.strip():
		raise AIServiceError(
			"needs_user_input ist true, aber user_input_reason "
			"erklärt die Rückfrage nicht."
		)

	if not needs_user_input and user_input_options:
		raise AIServiceError(
			"user_input_options müssen leer sein, wenn "
			"keine Benutzerentscheidung notwendig ist."
		)

	if not needs_user_input and user_input_reason.strip():
		raise AIServiceError(
			"user_input_reason muss leer sein, wenn keine "
			"Benutzerentscheidung notwendig ist."
		)

	return (
		needs_user_input,
		user_input_reason,
		user_input_options,
	)


def validate_ai_response(
	response: Mapping[str, Any],
	data: pd.DataFrame,
	columns: list[str],
	*,
	allow_pending_user_input: bool = False,
) -> ChartSpecification:
	"""Validiert die KI-Antwort und erzeugt daraus eine ChartSpecification."""

	allowed_columns = {
		str(column)
		for column in data.columns
	}

	required = {
		"independent_variable",
		"dependent_variable",
		"x_column",
		"y_column",
		"x_label",
		"y_label",
		"x_unit",
		"y_unit",
		"chart_type",
		"x_min",
		"x_max",
		"y_min",
		"y_max",
		"origin",
		"show_points",
		"connect_points",
		"trendline",
		"x_error_column",
		"y_error_column",
		"confidence",
		"reason",
		"needs_user_input",
		"user_input_reason",
		"user_input_options",
	}

	optional = {"polynomial_degree", "moving_average_period"}
	missing = required.difference(response)
	unexpected = set(response).difference(required | optional)

	if unexpected:
		raise AIServiceError(
			"Im KI-Ergebnis sind unbekannte Felder enthalten: "
			+ ", ".join(sorted(unexpected))
		)

	if (
		"needs_user_input" in response
		and isinstance(response["needs_user_input"], bool)
	):
		needs_user_input = response["needs_user_input"]
		if (
			needs_user_input
			and allow_pending_user_input
			and {
				"needs_user_input",
				"user_input_reason",
				"user_input_options",
			}.issubset(response)
		):
			(
				_,
				_,
				user_input_options,
			) = _validate_user_input_options(
				response,
				allowed_columns,
				set(columns),
			)
			pending_fields = {
				"x_column": {
					"x_column",
					"independent_variable",
				},
				"y_column": {
					"y_column",
					"dependent_variable",
				},
				"chart_type": {"chart_type"},
				"trendline": {"trendline"},
			}
			for option in user_input_options:
				missing.difference_update(
					pending_fields[option["type"]]
				)
	else:
		needs_user_input = False

	if missing:
		raise AIServiceError(
			"Im KI-Ergebnis fehlen Pflichtfelder: "
			+ ", ".join(sorted(missing))
		)

	(
		needs_user_input,
		user_input_reason,
		user_input_options,
	) = _validate_user_input_options(
		response,
		allowed_columns,
		set(columns),
	)

	pending_types = (
		{
			option["type"]
			for option in user_input_options
		}
		if needs_user_input and allow_pending_user_input
		else set()
	)

	x_column = response.get("x_column")
	if "x_column" in pending_types and x_column in (None, ""):
		x_column = ""
	elif not isinstance(x_column, str) or x_column not in allowed_columns:
		raise AIServiceError(
			"Die von der KI gewählte x-Spalte existiert "
			"nicht in der Tabelle."
		)

	y_column = response.get("y_column")
	if "y_column" in pending_types and y_column in (None, ""):
		y_column = ""
	elif not isinstance(y_column, str) or y_column not in allowed_columns:
		raise AIServiceError(
			"Die von der KI gewählte y-Spalte existiert "
			"nicht in der Tabelle."
		)

	if x_column and y_column and x_column == y_column:
		raise AIServiceError(
			"Die x- und y-Spalte müssen unterschiedlich sein."
		)

	chart_type = response.get("chart_type")
	if chart_type is None and "chart_type" not in pending_types:
		raise AIServiceError("Das Feld chart_type fehlt.")
	if chart_type is not None and (
		not isinstance(chart_type, str)
		or chart_type not in VALID_CHART_TYPES
	):
		raise AIServiceError(
			"Die Diagrammart muss scatter, scatter_lines, line, line_scatter, "
			"bar oder column sein."
		)
	if y_column and y_column not in columns:
		raise AIServiceError(
			"Die von der KI gewählte y-Spalte "
			"enthält keine numerischen Messwerte."
		)

	values: dict[str, np.ndarray] = {}

	for column in (x_column, y_column):
		if not column:
			continue
		numeric = pd.to_numeric(data[column], errors="coerce").to_numpy(dtype=float)
		finite = numeric[np.isfinite(numeric)]
		if finite.size >= 2:
			values[column] = finite

	if y_column and y_column not in values:
		raise AIServiceError(
			f"Die Spalte {y_column!r} hat weniger als "
			"zwei endliche Messwerte."
		)

	x_is_numeric = bool(x_column and x_column in columns)
	if x_column and chart_type == "scatter" and not x_is_numeric:
		raise AIServiceError(
			"Ein Scatter-/XY-Diagramm benötigt eine numerische x-Spalte."
		)
	if x_column and chart_type is None and "chart_type" in pending_types:
		x_is_numeric = x_is_numeric or x_column in values
	if x_column and not x_is_numeric:
		categories = data[x_column].dropna()
		if len(categories) < 2:
			raise AIServiceError(
				f"Die Kategorie-Spalte {x_column!r} hat weniger "
				"als zwei Werte."
			)
		values[x_column] = pd.factorize(data[x_column], sort=False)[0].astype(float) + 1

	if x_column and x_is_numeric and np.unique(values[x_column]).size < 2:
		raise AIServiceError(
			"Die x-Spalte benötigt mindestens zwei "
			"unterschiedliche Messwerte."
		)

	trendline = response.get("trendline")
	if trendline is None and "trendline" in pending_types:
		trendline = "none"
	elif (
		not isinstance(trendline, str)
		or trendline not in VALID_TRENDLINES
	):
		raise AIServiceError(
			"Die Trendlinie muss einer der unterstützten "
			"Werte sein: none, linear, quadratic, cubic, "
			"exponential, logarithmic, power, polynomial oder moving_average."
		)

	polynomial_degree = response.get("polynomial_degree", 2)
	if trendline == "quadratic":
		polynomial_degree = 2
	elif trendline == "cubic":
		polynomial_degree = 3
	if (
		not isinstance(polynomial_degree, int)
		or isinstance(polynomial_degree, bool)
		or not 2 <= polynomial_degree <= 6
	):
		raise AIServiceError("Der Polynomgrad muss zwischen 2 und 6 liegen.")
	moving_average_period = response.get("moving_average_period", 3)
	if (
		trendline == "moving_average"
		and "moving_average_period" not in response
		and x_column
		and y_column
	):
		valid_pairs = np.isfinite(values[x_column]) & np.isfinite(values[y_column])
		available_periods = [
			period
			for period in (2, 3, 4, 5, 6, 7, 10, 12, 20)
			if period < int(valid_pairs.sum())
		]
		if not available_periods:
			raise AIServiceError(
				"Der gleitende Durchschnitt benötigt mehr als zwei gültige Messpunkte."
			)
		moving_average_period = 3 if 3 in available_periods else available_periods[0]
	if (
		not isinstance(moving_average_period, int)
		or isinstance(moving_average_period, bool)
		or moving_average_period not in {2, 3, 4, 5, 6, 7, 10, 12, 20}
	):
		raise AIServiceError("Die Periode des gleitenden Durchschnitts ist ungültig.")

	if trendline == "linear" and x_column and y_column:
		if np.unique(values[x_column]).size < 2:
			raise AIServiceError(
				"Ein linearer Fit benötigt mindestens zwei "
				"unterschiedliche x-Werte."
			)

	if trendline == "quadratic" and x_column and y_column:
		if np.unique(values[x_column]).size < 3:
			raise AIServiceError(
				"Ein quadratischer Fit benötigt mindestens "
				"drei unterschiedliche x-Werte."
			)

	if trendline == "cubic" and x_column and y_column:
		if np.unique(values[x_column]).size < 4:
			raise AIServiceError(
				"Ein kubischer Fit benötigt mindestens "
				"vier unterschiedliche x-Werte."
			)

	if trendline == "polynomial" and x_column and y_column:
		if np.unique(values[x_column]).size <= polynomial_degree:
			raise AIServiceError(
				f"Ein Polynomgrad {polynomial_degree} benötigt mindestens "
				f"{polynomial_degree + 1} unterschiedliche x-Werte."
			)

	if trendline == "exponential" and x_column and y_column:
		if np.unique(values[x_column]).size < 3:
			raise AIServiceError(
				"Ein exponentieller Fit benötigt mindestens "
				"drei unterschiedliche x-Werte."
			)

		if np.any(values[y_column] <= 0):
			raise AIServiceError(
				"Ein exponentieller Fit erfordert positive y-Werte."
			)

	if trendline == "logarithmic" and x_column and y_column:
		if np.unique(values[x_column]).size < 2:
			raise AIServiceError(
				"Eine logarithmische Trendlinie benötigt mindestens "
				"zwei unterschiedliche x-Werte."
			)

		if np.any(values[x_column] <= 0):
			raise AIServiceError(
				"Eine logarithmische Trendlinie erfordert "
				"positive x-Werte."
			)

	if trendline == "power" and x_column and y_column:
		if np.unique(values[x_column]).size < 2:
			raise AIServiceError(
				"Eine Potenztrendlinie benötigt mindestens "
				"zwei unterschiedliche x-Werte."
			)

		if (
			np.any(values[x_column] <= 0)
			or np.any(values[y_column] <= 0)
		):
			raise AIServiceError(
				"Eine Potenztrendlinie erfordert positive "
				"x- und y-Werte."
			)

	if trendline == "moving_average" and x_column and y_column:
		valid_pairs = np.isfinite(values[x_column]) & np.isfinite(values[y_column])
		if int(valid_pairs.sum()) <= moving_average_period:
			raise AIServiceError(
				"Der gleitende Durchschnitt benötigt mehr gültige Messpunkte "
				"als die gewählte Periode."
			)

	string_fields = (
		"x_label",
		"y_label",
		"x_unit",
		"y_unit",
	)

	for field in string_fields:
		if (
			not isinstance(response[field], str)
			or len(response[field]) > MAX_TEXT_FIELD_CHARS
		):
			raise AIServiceError(
				f"Das Feld {field} muss ein Text mit höchstens "
				f"{MAX_TEXT_FIELD_CHARS} Zeichen sein."
			)

	independent_variable = response.get("independent_variable", "")
	dependent_variable = response.get("dependent_variable", "")

	if (
		"x_column" not in pending_types
		and independent_variable != x_column
	):
		raise AIServiceError(
			"Die unabhängige Größe muss exakt der "
			"gewählten x-Spalte entsprechen."
		)

	if (
		"y_column" not in pending_types
		and dependent_variable != y_column
	):
		raise AIServiceError(
			"Die abhängige Größe muss exakt der "
			"gewählten y-Spalte entsprechen."
		)

	reason = response["reason"]

	if not isinstance(reason, str):
		raise AIServiceError(
			"Das Feld reason muss ein Text sein."
		)

	if len(reason) > MAX_REASON_CHARS:
		reason = (
			reason[:MAX_REASON_CHARS - 3].rstrip()
			+ "..."
		)

	if not isinstance(response["show_points"], bool):
		raise AIServiceError(
			"Das Feld show_points muss true oder false sein."
		)

	if not isinstance(response["connect_points"], bool):
		raise AIServiceError(
			"Das Feld connect_points muss true oder false sein."
		)

	origin = response["origin"]

	if origin is not None and not isinstance(origin, bool):
		raise AIServiceError(
			"Das Feld origin muss true, false oder null sein."
		)

	confidence = _finite_number(
		response["confidence"],
		"confidence",
	)

	if confidence is None or not 0 <= confidence <= 1:
		raise AIServiceError(
			"confidence muss zwischen 0 und 1 liegen."
		)

	axis_limits = {
		field: _finite_number(
			response[field],
			field,
		)
		for field in (
			"x_min",
			"x_max",
			"y_min",
			"y_max",
		)
	}

	for minimum, maximum in (
		("x_min", "x_max"),
		("y_min", "y_max"),
	):
		if (
			axis_limits[minimum] is not None
			and axis_limits[maximum] is not None
			and axis_limits[minimum] >= axis_limits[maximum]
		):
			raise AIServiceError(
				f"{minimum} muss kleiner als {maximum} sein."
			)

	error_columns: dict[str, str | None] = {}

	for field in (
		"x_error_column",
		"y_error_column",
	):
		column = response[field]

		if column is not None:
			if (
				not isinstance(column, str)
				or column not in allowed_columns
			):
				raise AIServiceError(
					f"Die Fehlerwert-Spalte aus {field} "
					"existiert nicht."
				)

			if column and column in {x_column, y_column}:
				raise AIServiceError(
					f"{field} darf nicht mit einer "
					"Diagrammachse übereinstimmen."
				)
			if field == "x_error_column" and x_column and not x_is_numeric:
				raise AIServiceError(
					"x_error_column ist für eine kategoriale x-Achse "
					"nicht zulässig."
				)

			error_values = pd.to_numeric(
				data[column],
				errors="coerce",
			).to_numpy(dtype=float)

			if not np.any(
				np.isfinite(error_values)
				& (error_values >= 0)
			):
				raise AIServiceError(
					f"Die Fehlerwert-Spalte {column!r} enthält "
					"keine gültigen Werte."
				)

		error_columns[field] = column

	x_label, x_unit = _axis_label_and_unit(
		x_column or "",
		response["x_label"],
		response["x_unit"],
	)
	y_label, y_unit = _axis_label_and_unit(
		y_column or "",
		response["y_label"],
		response["y_unit"],
	)
	return ChartSpecification(
		x_column=x_column,
		y_column=y_column,
		x_label=x_label or x_column,
		y_label=y_label or y_column,
		x_unit=x_unit,
		y_unit=y_unit,
		chart_type=chart_type or "scatter",
		x_min=axis_limits["x_min"],
		x_max=axis_limits["x_max"],
		y_min=axis_limits["y_min"],
		y_max=axis_limits["y_max"],
		origin=False if origin is None else origin,
		trendline=trendline,
		polynomial_degree=polynomial_degree,
		moving_average_period=moving_average_period,
		x_error_column=error_columns["x_error_column"],
		y_error_column=error_columns["y_error_column"],
		independent_variable=(
			independent_variable or None
		),
		dependent_variable=(
			dependent_variable or None
		),
		show_points=response["show_points"],
		connect_points=response["connect_points"],
		confidence=confidence,
		reason=reason,
		needs_user_input=needs_user_input,
		user_input_reason=user_input_reason,
		user_input_options=user_input_options,
		ai_response=dict(response),
	)


def _cell_value(value: Any) -> Any:
	if value is None or (
		not isinstance(value, (list, dict))
		and pd.isna(value)
	):
		return None

	if isinstance(
		value,
		(int, float, np.integer, np.floating),
	):
		number = float(value)
		return number if math.isfinite(number) else None

	if hasattr(value, "isoformat"):
		return value.isoformat()

	return str(value)[:MAX_CELL_CHARS]


def _trendline_candidate_fits(x_values: np.ndarray, y_values: np.ndarray) -> list[dict[str, float | str]]:
	"""Compare supported curve families against y on its original scale."""
	x = np.asarray(x_values, dtype=float)
	y = np.asarray(y_values, dtype=float)
	valid = np.isfinite(x) & np.isfinite(y)
	x = x[valid]
	y = y[valid]
	if x.size < 4 or np.unique(x).size < 2 or np.ptp(y) == 0:
		return []
	x_scale = float(np.ptp(x))
	x_normalized = (x - float(np.min(x))) / x_scale
	n = int(x.size)
	models: list[tuple[str, np.ndarray, int]] = []
	for degree, name in ((1, "linear"), (2, "quadratic"), (3, "cubic")):
		parameter_count = degree + 1
		if n <= parameter_count + 1 or np.unique(x_normalized).size <= degree:
			continue
		coefficients = np.polynomial.polynomial.polyfit(x_normalized, y, degree)
		prediction = np.polynomial.polynomial.polyval(x_normalized, coefficients)
		models.append((name, prediction, parameter_count))
	if np.all(y > 0) and n > 3:
		coefficients = np.polynomial.polynomial.polyfit(x_normalized, np.log(y), 1)
		models.append(("exponential", np.exp(np.polynomial.polynomial.polyval(x_normalized, coefficients)), 2))
	if np.all(x > 0):
		log_x = np.log(x)
		log_x = (log_x - float(np.min(log_x))) / float(np.ptp(log_x))
		if np.ptp(log_x) > 0:
			coefficients = np.polynomial.polynomial.polyfit(log_x, y, 1)
			models.append(("logarithmic", np.polynomial.polynomial.polyval(log_x, coefficients), 2))
			if np.all(y > 0):
				power_coefficients = np.polynomial.polynomial.polyfit(log_x, np.log(y), 1)
				models.append(("power", np.exp(
					np.polynomial.polynomial.polyval(log_x, power_coefficients)
				), 2))

	total_sum_squares = float(np.sum((y - float(np.mean(y))) ** 2))
	results: list[dict[str, float | str]] = []
	for name, prediction, parameter_count in models:
		residual_sum_squares = float(np.sum((y - prediction) ** 2))
		r_squared = 1.0 - residual_sum_squares / total_sum_squares
		adjusted_r_squared = 1.0 - (
			(1.0 - r_squared) * (n - 1) / (n - parameter_count - 1)
		)
		results.append({
			"model": name,
			"r_squared": round(r_squared, 6),
			"adjusted_r_squared": round(adjusted_r_squared, 6),
		})
	return sorted(
		results,
		key=lambda result: (
			-float(result["adjusted_r_squared"]),
			("linear", "quadratic", "cubic", "exponential", "logarithmic", "power").index(
				str(result["model"])
			),
		),
	)


def _numeric_trendline_evidence(
	data: pd.DataFrame,
	columns: list[str],
) -> dict[str, Any]:
	"""Provide compact curve-fit evidence to the AI without making its decision."""
	numeric_columns = [
		column
		for column in dict.fromkeys(columns)
		if column in data.columns
		and pd.to_numeric(data[column], errors="coerce").notna().sum() >= 4
	]
	numeric_values = {
		column: pd.to_numeric(data[column], errors="coerce").to_numpy(dtype=float)
		for column in numeric_columns
	}
	pairs: list[dict[str, Any]] = []
	for x_column in numeric_columns:
		for y_column in numeric_columns:
			if x_column == y_column:
				continue
			x_values = numeric_values[x_column]
			y_values = numeric_values[y_column]
			fits = _trendline_candidate_fits(x_values, y_values)
			if fits:
				valid_pair_count = int(
					(np.isfinite(x_values) & np.isfinite(y_values)).sum()
				)
				pairs.append({
					"x": x_column,
					"y": y_column,
					"valid_pair_count": valid_pair_count,
					"best_fit_candidates": fits[:3],
				})
	pairs.sort(
		key=lambda pair: -float(pair["best_fit_candidates"][0]["adjusted_r_squared"])
	)
	return {
		"evaluated_numeric_pairs": len(pairs),
		"leading_pairs": pairs[:12],
	}


def _table_context(
	data: pd.DataFrame,
	columns: list[str],
) -> dict[str, Any]:
	all_columns = [
		str(column)
		for column in data.columns
	]

	eligible_columns = set(columns)

	column_info = []

	context_columns = list(
		dict.fromkeys(
			[*columns, *all_columns]
		)
	)[:MAX_PROMPT_COLUMNS]

	for column in context_columns:
		values = pd.to_numeric(
			data[column],
			errors="coerce",
		).to_numpy(dtype=float)

		finite_values = values[
			np.isfinite(values)
		]

		non_empty_count = int(
			data[column].notna().sum()
		)

		column_info.append({
			"name": column,
			"data_type": str(data[column].dtype),
			"numeric_measurement_candidate": (
				column in eligible_columns
			),
			"non_empty_count": non_empty_count,
			"missing_count": (
				int(len(data) - non_empty_count)
			),
			"unique_count": int(
				data[column].nunique(
					dropna=True
				)
			),
			"numeric_value_count": int(
				finite_values.size
			),
			"numeric_min": (
				float(finite_values.min())
				if finite_values.size
				else None
			),
			"numeric_max": (
				float(finite_values.max())
				if finite_values.size
				else None
			),
		})

	data_columns = list(
		dict.fromkeys(
			[*columns, *all_columns]
		)
	)[:MAX_PROMPT_COLUMNS]

	if len(data) <= MAX_PROMPT_ROWS:
		indexes = list(range(len(data)))
	else:
		indexes = np.linspace(
			0,
			len(data) - 1,
			MAX_PROMPT_ROWS,
			dtype=int,
		).tolist()

	rows = [
		{
			column: _cell_value(
				data.iloc[index][column]
			)
			for column in data_columns
		}
		for index in indexes
	]

	return {
		"columns": column_info,
		"sampled_data_columns": data_columns,
		"table_row_count": int(len(data)),
		"sampled_rows_count": len(rows),
		"rows_are_sampled": (
			len(data) > MAX_PROMPT_ROWS
		),
		"rows": rows,
		"numeric_trendline_evidence": _numeric_trendline_evidence(
			data.iloc[indexes],
			columns,
		),
	}


def analyze_with_ai(
	data: pd.DataFrame,
	columns: list[str],
	config: Mapping[str, str],
) -> ChartSpecification:
	"""Analysiert die Tabelle mit der KI und validiert die JSON-Antwort."""

	context = _table_context(
		data,
		columns,
	)
	response = _request_json_response(
		config,
		[
			{
				"role": "system",
				"content": ANALYSIS_SYSTEM_PROMPT,
			},
			{
				"role": "user",
				"content": json.dumps(
					context,
					ensure_ascii=False,
					allow_nan=False,
				),
			},
		],
	)

	return validate_ai_response(
		response,
		data,
		columns,
		allow_pending_user_input=True,
	)

def reanalyze_with_user_input(
	data: pd.DataFrame,
	columns: list[str],
	config: Mapping[str, str],
	previous_spec: ChartSpecification,
	user_decisions: Mapping[str, str],
) -> ChartSpecification:
	"""Führt die Analyse nach einer Benutzerentscheidung erneut durch."""

	context = _table_context(
		data,
		columns,
	)

	previous_analysis = previous_spec.ai_response or {
		"x_column": previous_spec.x_column,
		"y_column": previous_spec.y_column,
		"x_label": previous_spec.x_label,
		"y_label": previous_spec.y_label,
		"x_unit": previous_spec.x_unit,
		"y_unit": previous_spec.y_unit,
		"chart_type": previous_spec.chart_type,
		"x_min": previous_spec.x_min,
		"x_max": previous_spec.x_max,
		"y_min": previous_spec.y_min,
		"y_max": previous_spec.y_max,
		"origin": previous_spec.origin,
		"show_points": previous_spec.show_points,
		"connect_points": previous_spec.connect_points,
		"trendline": previous_spec.trendline,
		"polynomial_degree": previous_spec.polynomial_degree,
		"moving_average_period": previous_spec.moving_average_period,
		"x_error_column": previous_spec.x_error_column,
		"y_error_column": previous_spec.y_error_column,
		"independent_variable": previous_spec.independent_variable,
		"dependent_variable": previous_spec.dependent_variable,
		"confidence": previous_spec.confidence,
		"reason": previous_spec.reason,
		"needs_user_input": previous_spec.needs_user_input,
		"user_input_reason": previous_spec.user_input_reason,
		"user_input_options": previous_spec.user_input_options,
	}

	requested_types = {
		option["type"]: option["options"]
		for option in previous_spec.user_input_options
	}
	if (
		not previous_spec.needs_user_input
		or set(user_decisions) != set(requested_types)
		or any(
			not isinstance(value, str)
			or value not in requested_types[key]
			for key, value in user_decisions.items()
		)
	):
		raise AIServiceError(
			"Die Benutzerentscheidungen passen nicht zu den "
			"Rückfragen der vorherigen Analyse."
		)

	request_context = {
		"table": context,
		"previous_analysis": previous_analysis,
		"user_decision": dict(user_decisions),
	}
	response = _request_json_response(
		config,
		[
			{
				"role": "system",
				"content": ANALYSIS_SYSTEM_PROMPT,
			},
			{
				"role": "user",
				"content": json.dumps(
					request_context,
					ensure_ascii=False,
					allow_nan=False,
				),
			},
		],
	)

	spec = validate_ai_response(
		response,
		data,
		columns,
	)

	if spec.needs_user_input:
		raise AIServiceError(
			"Die KI benötigt trotz der Benutzerentscheidung "
			"noch weitere Angaben."
		)

	for decision_type, decision in user_decisions.items():
		if getattr(spec, decision_type) != decision:
			raise AIServiceError(
				f"Die KI hat die verbindliche Benutzerentscheidung "
				f"für {decision_type} nicht übernommen."
			)

	return spec

def test_ai_connection(
	config: Mapping[str, str],
) -> None:
	"""Sendet eine minimale Anfrage."""
	body = _post_chat_completion(
		config,
		[
			{
				"role": "system",
				"content": (
					"Reply with a short connection confirmation."
				),
			},
			{
				"role": "user",
				"content": "Connection test.",
			},
		],
	)

	_response_text(body)