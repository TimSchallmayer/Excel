from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import pandas as pd

from ai import (
	AIServiceError,
	_extract_json,
	_extract_response_json,
	_post_chat_completion,
	_response_text,
	analyze_with_ai,
	reanalyze_with_user_input,
	validate_ai_response,
)
from config import load_config


class AIValidationTests(unittest.TestCase):
	def setUp(self) -> None:
		self.data = pd.DataFrame({
			"Zeit [s]": [0.0, 1.0, 2.0],
			"Strecke [m]": [0.0, 2.0, 4.0],
			"Strecke Fehler [m]": [0.1, 0.1, 0.2],
		})

		self.columns = [
			"Zeit [s]",
			"Strecke [m]",
		]

		self.answer = {
			"independent_variable": "Zeit [s]",
			"dependent_variable": "Strecke [m]",
			"x_column": "Zeit [s]",
			"y_column": "Strecke [m]",
			"x_label": "Zeit [s]",
			"y_label": "Strecke [m]",
			"x_unit": "s",
			"y_unit": "m",
			"chart_type": "scatter",
			"x_min": 0,
			"x_max": 2,
			"y_min": 0,
			"y_max": 4,
			"origin": True,
			"show_points": True,
			"connect_points": False,
			"trendline": "linear",
			"x_error_column": None,
			"y_error_column": "Strecke Fehler [m]",
			"confidence": 0.97,
			"reason": "Zeit ist die unabhängige Größe.",
			"needs_user_input": False,
			"user_input_reason": "",
			"user_input_options": [],
		}

	def test_valid_response_creates_chart_specification(self) -> None:
		spec = validate_ai_response(
			self.answer,
			self.data,
			self.columns,
		)

		self.assertEqual(
			spec.x_column,
			"Zeit [s]",
		)

		self.assertEqual(
			spec.y_label,
			"Strecke [m]",
		)

		self.assertEqual(
			spec.trendline,
			"linear",
		)

		self.assertEqual(
			spec.confidence,
			0.97,
		)

		self.assertEqual(
			spec.y_error_column,
			"Strecke Fehler [m]",
		)

		self.assertFalse(
			spec.needs_user_input
		)

		self.assertEqual(
			spec.user_input_options,
			[],
		)

	def test_invalid_response_missing_required_field(self) -> None:
		answer = self.answer.copy()
		del answer["reason"]

		with self.assertRaisesRegex(
			AIServiceError,
			"Pflichtfelder",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_unexpected_response_fields_are_rejected(self) -> None:
		with self.assertRaisesRegex(AIServiceError, "unbekannte Felder"):
			validate_ai_response(
				self.answer | {"unexpected": "value"},
				self.data,
				self.columns,
			)

	def test_missing_column_is_rejected(self) -> None:
		answer = self.answer | {
			"x_column": "Zeitpunkt"
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"x-Spalte existiert nicht",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_unsupported_chart_type_is_rejected(self) -> None:
		answer = self.answer | {
			"chart_type": "pie"
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"Diagrammart",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_invalid_axis_ranges_are_rejected(self) -> None:
		answer = self.answer | {
			"x_min": 3,
			"x_max": 2,
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"x_min muss kleiner",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_non_finite_axis_ranges_are_rejected(self) -> None:
		answer = self.answer | {
			"y_max": float("inf")
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"endliche Zahl",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_confidence_out_of_range_is_rejected(self) -> None:
		answer = self.answer | {
			"confidence": 1.01
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"zwischen 0 und 1",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_cubic_trendline_requires_four_distinct_x_values(self) -> None:
		answer = self.answer | {
			"trendline": "cubic"
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"mindestens vier",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_user_input_requires_options(self) -> None:
		answer = self.answer | {
			"needs_user_input": True,
			"user_input_reason": "Diagrammart ist unklar.",
			"user_input_options": [],
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"keine Benutzeroptionen",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	def test_user_input_reason_must_match_user_input_flag(self) -> None:
		for needs_user_input, reason in (
			(True, ""),
			(False, "Keine Rückfrage erforderlich."),
		):
			answer = self.answer | {
				"needs_user_input": needs_user_input,
				"user_input_reason": reason,
				"user_input_options": (
					[
						{
							"type": "chart_type",
							"question": "Welche Diagrammart?",
							"options": ["scatter", "line"],
						}
					]
					if needs_user_input
					else []
				),
			}

			with self.subTest(needs_user_input=needs_user_input):
				with self.assertRaisesRegex(AIServiceError, "user_input_reason"):
					validate_ai_response(
						answer,
						self.data,
						self.columns,
					)

	def test_needs_user_input_boolean_does_not_allow_schema_omissions(self) -> None:
		for needs_user_input in (True, False):
			with self.subTest(needs_user_input=needs_user_input):
				with self.assertRaisesRegex(AIServiceError, "Pflichtfelder"):
					validate_ai_response(
						{"needs_user_input": needs_user_input},
						self.data,
						self.columns,
						allow_pending_user_input=True,
					)

	def test_user_input_options_are_validated(self) -> None:
		answer = self.answer | {
			"needs_user_input": True,
			"user_input_reason": (
				"Mehrere Diagrammarten sind plausibel."
			),
			"user_input_options": [
				{
					"type": "chart_type",
					"question": (
						"Welche Diagrammart soll verwendet werden?"
					),
					"options": [
						"scatter",
						"line",
					],
				}
			],
		}

		spec = validate_ai_response(
			answer,
			self.data,
			self.columns,
		)

		self.assertTrue(
			spec.needs_user_input
		)

		self.assertEqual(
			spec.user_input_reason,
			"Mehrere Diagrammarten sind plausibel.",
		)

		self.assertEqual(
			spec.user_input_options[0]["type"],
			"chart_type",
		)

	def test_user_input_option_type_must_be_a_valid_string(self) -> None:
		for option_type in ([], {}, 123):
			answer = self.answer | {
				"needs_user_input": True,
				"user_input_reason": "Die Diagrammart ist unklar.",
				"user_input_options": [
					{
						"type": option_type,
						"question": "Welche Diagrammart?",
						"options": ["scatter", "line"],
					}
				],
			}
			with self.subTest(option_type=option_type):
				with self.assertRaisesRegex(AIServiceError, "Unbekannter Typ"):
					validate_ai_response(answer, self.data, self.columns)

		answer = self.answer | {
			"needs_user_input": True,
			"user_input_reason": "Die Diagrammart ist unklar.",
			"user_input_options": [
				{
					"type": "chart_type",
					"question": "Welche Diagrammart?",
					"options": ["scatter", "line"],
				}
			],
		}
		spec = validate_ai_response(answer, self.data, self.columns)
		self.assertEqual(spec.user_input_options[0]["type"], "chart_type")

	def test_provisional_responses_may_omit_only_queried_fields(self) -> None:
		cases = (
			(
				"x_column",
				("x_column", "independent_variable"),
				["Zeit [s]", "Strecke [m]"],
			),
			(
				"y_column",
				("y_column", "dependent_variable"),
				["Zeit [s]", "Strecke [m]"],
			),
			("chart_type", ("chart_type",), ["scatter", "line"]),
			("trendline", ("trendline",), ["none", "linear"]),
		)

		for question_type, omitted_fields, options in cases:
			with self.subTest(question_type=question_type):
				answer = self.answer | {
					"needs_user_input": True,
					"user_input_reason": "Diese Entscheidung ist unklar.",
					"user_input_options": [
						{
							"type": question_type,
							"question": "Bitte auswählen.",
							"options": options,
						}
					],
				}
				for field in omitted_fields:
					answer.pop(field)

				spec = validate_ai_response(
					answer,
					self.data,
					self.columns,
					allow_pending_user_input=True,
				)

				self.assertTrue(spec.needs_user_input)
				self.assertEqual(spec.user_input_options[0]["type"], question_type)

				with self.assertRaises(AIServiceError):
					validate_ai_response(
						answer,
						self.data,
						self.columns,
					)

	def test_invalid_user_input_column_is_rejected(self) -> None:
		answer = self.answer | {
			"needs_user_input": True,
			"user_input_reason": "X ist unklar.",
			"user_input_options": [
				{
					"type": "x_column",
					"question": (
						"Welche Spalte soll X sein?"
					),
					"options": [
						"Zeitpunkt"
					],
				}
			],
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"nicht vorhandene Spalte",
		):
			validate_ai_response(
				answer,
				self.data,
				self.columns,
			)

	@patch("ai._post_chat_completion")
	def test_api_json_is_parsed_and_validated(
		self,
		post: Mock,
	) -> None:
		post.return_value = {
			"choices": [
				{
					"message": {
						"content": (
							"```json\n"
							+ json.dumps(
								self.answer,
								ensure_ascii=False,
							)
							+ "\n```"
						)
					}
				}
			]
		}

		spec = analyze_with_ai(
			self.data,
			self.columns,
			{
				"endpoint": (
					"https://example.invalid/v1/chat/completions"
				),
				"api_key": "",
				"model": "test-model",
			},
		)

		self.assertEqual(
			spec.x_column,
			"Zeit [s]",
		)

		self.assertFalse(
			spec.needs_user_input
		)

		post.assert_called_once()
		self.assertTrue(post.call_args.kwargs["json_mode"])

	def test_json_extraction_accepts_plain_markdown_and_prefixed_json(self) -> None:
		encoded = json.dumps(self.answer, ensure_ascii=False)

		for text in (
			encoded,
			f"```json\n{encoded}\n```",
			f"```text\n{encoded}\n```",
			f"Hier ist die Analyse:\n{encoded}",
		):
			with self.subTest(text=text[:32]):
				self.assertEqual(_extract_json(text), self.answer)

	def test_json_extraction_rejects_plain_text(self) -> None:
		with self.assertRaisesRegex(
			AIServiceError,
			"enthält kein JSON-Objekt",
		):
			_extract_json("Ich kann die Tabelle analysieren.")

	@patch.dict("os.environ", {"PHYSIK_AI_RESPONSE_DEBUG": "1"})
	@patch("ai.logging.getLogger")
	def test_failed_extraction_logs_only_safe_response_metadata(
		self,
		get_logger: Mock,
	) -> None:
		body = {
			"choices": [{
				"message": {
					"content": "private table-derived response",
					"refusal": None,
				},
				"finish_reason": "stop",
			}]
		}
		logger = Mock()
		get_logger.return_value = logger

		with self.assertRaises(AIServiceError):
			_extract_response_json(
				body,
				{
					"http_status": 200,
					"model": "test-model",
					"response_format_set": True,
				},
			)

		logged = logger.warning.call_args.args[1]
		self.assertNotIn("private table-derived response", logged)
		self.assertIn('"http_status": 200', logged)
		self.assertIn('"model": "test-model"', logged)
		self.assertIn('"response_format_set": true', logged)
		self.assertIn('"content_type": "str"', logged)
		self.assertIn('"content_length": 30', logged)
		self.assertIn('"content_contains_opening_brace": false', logged)
		self.assertIn('"finish_reason": "stop"', logged)
		self.assertIn('"failure_stage": "extract_json_attempt_1"', logged)
		self.assertIn('"failure_kind": "no_json_object"', logged)

	@patch("ai._post_chat_completion")
	def test_analyze_reports_plain_text_response_as_ai_service_error(
		self,
		post: Mock,
	) -> None:
		post.return_value = {
			"choices": [
				{"message": {"content": "Ich kann die Tabelle analysieren."}}
			]
		}

		with self.assertRaisesRegex(
			AIServiceError,
			"enthält kein JSON-Objekt",
		):
			analyze_with_ai(
				self.data,
				self.columns,
				{
					"endpoint": "https://example.invalid/v1/chat/completions",
					"api_key": "",
					"model": "test-model",
				},
			)

	def test_response_text_accepts_structured_api_json(self) -> None:
		bodies = (
			self.answer,
			{"choices": [{"message": {"parsed": self.answer}}]},
			{"choices": [{"message": {"content": self.answer}}]},
			{
				"choices": [{
					"message": {
						"content": [{"type": "json", "json": self.answer}]
					}
				}]
			},
		)

		for body in bodies:
			with self.subTest(body_type=type(body).__name__):
				self.assertEqual(
					_extract_json(_response_text(body)),
					self.answer,
				)

	@patch("ai.requests.post")
	def test_analysis_request_enables_json_object_mode(
		self,
		post: Mock,
	) -> None:
		response = Mock(status_code=200, content=b"{}")
		response.json.return_value = {}
		post.return_value = response

		_post_chat_completion(
			{
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "",
				"model": "test-model",
			},
			[{"role": "user", "content": "test"}],
			json_mode=True,
		)

		self.assertEqual(
			post.call_args.kwargs["json"]["response_format"],
			{"type": "json_object"},
		)

	@patch("ai._post_chat_completion")
	def test_x_column_user_input_workflow_returns_final_chart_specification(
		self,
		post: Mock,
	) -> None:
		data = pd.DataFrame({
			"Zeit [s]": [0.0, 1.0, 2.0],
			"Temperatur [°C]": [20.0, 22.0, 24.0],
			"Messwert": [1.0, 2.0, 4.0],
		})
		columns = list(data.columns)
		first_answer = self.answer | {
			"independent_variable": "",
			"dependent_variable": "Messwert",
			"x_column": None,
			"y_column": "Messwert",
			"x_label": "",
			"y_label": "Messwert",
			"y_unit": "m",
			"x_error_column": None,
			"y_error_column": None,
			"needs_user_input": True,
			"user_input_reason": "Die X-Achse ist nicht eindeutig.",
			"user_input_options": [{
				"type": "x_column",
				"question": (
					"Welche Spalte soll als X-Achse verwendet werden?"
				),
				"options": ["Zeit [s]", "Temperatur [°C]"],
			}],
		}
		final_answer = first_answer | {
			"independent_variable": "Zeit [s]",
			"dependent_variable": "Messwert",
			"x_column": "Zeit [s]",
			"x_label": "Zeit [s]",
			"chart_type": "scatter",
			"needs_user_input": False,
			"user_input_reason": "",
			"user_input_options": [],
		}
		post.side_effect = [
			{
				"choices": [{
					"message": {
						"content": json.dumps(first_answer, ensure_ascii=False)
					}
				}]
			},
			{
				"choices": [{
					"message": {
						"content": json.dumps(final_answer, ensure_ascii=False)
					}
				}]
			},
		]
		config = {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		}

		provisional_spec = analyze_with_ai(data, columns, config)

		self.assertTrue(provisional_spec.needs_user_input)
		self.assertEqual(provisional_spec.x_column, "")
		self.assertEqual(provisional_spec.y_column, "Messwert")
		self.assertEqual(
			provisional_spec.user_input_options,
			first_answer["user_input_options"],
		)

		selected_option = provisional_spec.user_input_options[0]
		user_decisions = {
			selected_option["type"]: selected_option["options"][0]
		}
		self.assertEqual(user_decisions, {"x_column": "Zeit [s]"})

		final_spec = reanalyze_with_user_input(
			data,
			columns,
			config,
			provisional_spec,
			user_decisions,
		)

		self.assertEqual(post.call_count, 2)
		second_request = json.loads(
			post.call_args_list[1].args[1][1]["content"]
		)
		self.assertEqual(second_request["table"]["rows"], [
			{
				"Zeit [s]": 0.0,
				"Temperatur [°C]": 20.0,
				"Messwert": 1.0,
			},
			{
				"Zeit [s]": 1.0,
				"Temperatur [°C]": 22.0,
				"Messwert": 2.0,
			},
			{
				"Zeit [s]": 2.0,
				"Temperatur [°C]": 24.0,
				"Messwert": 4.0,
			},
		])
		self.assertEqual(second_request["previous_analysis"], first_answer)
		self.assertEqual(second_request["user_decision"], user_decisions)
		self.assertFalse(final_spec.needs_user_input)
		self.assertEqual(final_spec.x_column, "Zeit [s]")
		self.assertEqual(final_spec.y_column, "Messwert")
		self.assertEqual(final_spec.chart_type, "scatter")
		self.assertEqual(final_spec.ai_response, final_answer)

	@patch("ai._post_chat_completion")
	def test_x_column_reanalysis_rejects_a_second_user_input_request(
		self,
		post: Mock,
	) -> None:
		data = pd.DataFrame({
			"Zeit [s]": [0.0, 1.0, 2.0],
			"Temperatur [°C]": [20.0, 22.0, 24.0],
			"Messwert": [1.0, 2.0, 4.0],
		})
		columns = list(data.columns)
		pending_answer = self.answer | {
			"independent_variable": "",
			"dependent_variable": "Messwert",
			"x_column": None,
			"y_column": "Messwert",
			"x_label": "",
			"y_label": "Messwert",
			"x_error_column": None,
			"y_error_column": None,
			"needs_user_input": True,
			"user_input_reason": "Die X-Achse ist nicht eindeutig.",
			"user_input_options": [{
				"type": "x_column",
				"question": "Welche X-Achse?",
				"options": ["Zeit [s]", "Temperatur [°C]"],
			}],
		}
		second_pending_answer = pending_answer | {
			"independent_variable": "Zeit [s]",
			"x_column": "Zeit [s]",
			"chart_type": "scatter",
			"user_input_reason": "Die Diagrammart bleibt unklar.",
			"user_input_options": [{
				"type": "chart_type",
				"question": "Welche Diagrammart?",
				"options": ["scatter", "line"],
			}],
		}
		post.side_effect = [
			{
				"choices": [{
					"message": {
						"content": json.dumps(pending_answer, ensure_ascii=False)
					}
				}]
			},
			{
				"choices": [{
					"message": {
						"content": json.dumps(
							second_pending_answer,
							ensure_ascii=False,
						)
					}
				}]
			},
		]
		config = {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		}

		provisional_spec = analyze_with_ai(data, columns, config)

		with self.assertRaisesRegex(
			AIServiceError,
			"noch weitere Angaben",
		):
			reanalyze_with_user_input(
				data,
				columns,
				config,
				provisional_spec,
				{"x_column": "Zeit [s]"},
			)

		self.assertEqual(post.call_count, 2)

	@patch("ai._post_chat_completion")
	def test_json_retry_recovers_after_plain_text_and_keeps_request_context(
		self,
		post: Mock,
	) -> None:
		post.side_effect = [
			{
				"choices": [{
					"message": {
						"content": "Ich kann die Tabelle analysieren."
					}
				}]
			},
			{
				"choices": [{
					"message": {
						"content": json.dumps(self.answer, ensure_ascii=False)
					}
				}]
			},
		]

		spec = analyze_with_ai(
			self.data,
			self.columns,
			{
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "",
				"model": "test-model",
			},
		)

		self.assertEqual(spec.x_column, "Zeit [s]")
		self.assertEqual(post.call_count, 2)
		for call in post.call_args_list:
			self.assertTrue(call.kwargs["json_mode"])
		first_messages = post.call_args_list[0].args[1]
		retry_messages = post.call_args_list[1].args[1]
		self.assertEqual(
			retry_messages[0]["content"][:len(first_messages[0]["content"])],
			first_messages[0]["content"],
		)
		for instruction in (
			"Deine Antwort muss ausschließlich ein gültiges JSON-Objekt sein.",
			"Kein erklärender Text.",
			"Kein Markdown.",
			"Keine Code-Fences.",
			"Keine Einleitung.",
			"Keine Ausgabe außerhalb des JSON-Objekts.",
		):
			self.assertIn(instruction, retry_messages[0]["content"])
		self.assertEqual(retry_messages[1], first_messages[1])

	@patch("ai._post_chat_completion")
	def test_json_retry_fails_after_exactly_two_plain_text_responses(
		self,
		post: Mock,
	) -> None:
		plain_text = {
			"choices": [{
				"message": {"content": "Ich kann die Tabelle analysieren."}
			}]
		}
		post.return_value = plain_text

		with self.assertRaisesRegex(
			AIServiceError,
			"enthält kein JSON-Objekt",
		):
			analyze_with_ai(
				self.data,
				self.columns,
				{
					"endpoint": "https://example.invalid/v1/chat/completions",
					"api_key": "",
					"model": "test-model",
				},
			)

		self.assertEqual(post.call_count, 2)

	@patch("ai._post_chat_completion")
	def test_reanalysis_uses_the_shared_json_retry(
		self,
		post: Mock,
	) -> None:
		previous_spec = validate_ai_response(
			self.answer | {
				"needs_user_input": True,
				"user_input_reason": "Diagrammart ist unklar.",
				"user_input_options": [{
					"type": "chart_type",
					"question": "Welche Diagrammart?",
					"options": ["scatter", "line"],
				}],
			},
			self.data,
			self.columns,
		)
		post.side_effect = [
			{
				"choices": [{
					"message": {
						"content": "Ich kann die Tabelle analysieren."
					}
				}]
			},
			{
				"choices": [{
					"message": {
						"content": json.dumps(
							self.answer | {"chart_type": "line"},
							ensure_ascii=False,
						)
					}
				}]
			},
		]

		final_spec = reanalyze_with_user_input(
			self.data,
			self.columns,
			{
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "",
				"model": "test-model",
			},
			previous_spec,
			{"chart_type": "line"},
		)

		self.assertEqual(final_spec.chart_type, "line")
		self.assertFalse(final_spec.needs_user_input)
		self.assertEqual(post.call_count, 2)

	@patch("ai._post_chat_completion")
	def test_reanalysis_with_user_input(
		self,
		post: Mock,
	) -> None:
		first_answer = self.answer | {
			"chart_type": "scatter",
			"needs_user_input": True,
			"user_input_reason": (
				"Scatter und Linie sind beide plausibel."
			),
			"user_input_options": [
				{
					"type": "chart_type",
					"question": (
						"Welche Diagrammart soll verwendet werden?"
					),
					"options": [
						"scatter",
						"line",
					],
				}
			],
		}

		final_answer = self.answer | {
			"chart_type": "line",
			"show_points": True,
			"connect_points": True,
			"needs_user_input": False,
			"user_input_reason": "",
			"user_input_options": [],
			"reason": (
				"Die gewählte Darstellung verbindet "
				"die Messpunkte als Linien."
			),
		}

		post.side_effect = [
			{
				"choices": [
					{
						"message": {
							"content": json.dumps(
								first_answer,
								ensure_ascii=False,
							)
						}
					}
				]
			},
			{
				"choices": [
					{
						"message": {
							"content": json.dumps(
								final_answer,
								ensure_ascii=False,
							)
						}
					}
				]
			},
		]

		first_spec = analyze_with_ai(
			self.data,
			self.columns,
			{
				"endpoint": (
					"https://example.invalid/v1/chat/completions"
				),
				"api_key": "",
				"model": "test-model",
			},
		)

		self.assertTrue(
			first_spec.needs_user_input
		)

		self.assertEqual(
			first_spec.user_input_options[0]["type"],
			"chart_type",
		)

		final_spec = reanalyze_with_user_input(
			self.data,
			self.columns,
			{
				"endpoint": (
					"https://example.invalid/v1/chat/completions"
				),
				"api_key": "",
				"model": "test-model",
			},
			first_spec,
			{
				"chart_type": "line",
			},
		)

		self.assertFalse(
			final_spec.needs_user_input
		)

		self.assertEqual(
			final_spec.chart_type,
			"line",
		)

		self.assertTrue(
			final_spec.connect_points
		)

		self.assertEqual(
			final_spec.x_column,
			"Zeit [s]",
		)

		self.assertEqual(
			final_spec.y_column,
			"Strecke [m]",
		)

		self.assertEqual(
			post.call_count,
			2,
		)
		self.assertTrue(
			all(call.kwargs["json_mode"] for call in post.call_args_list)
		)

		request = json.loads(
			post.call_args_list[1].args[1][1]["content"]
		)
		self.assertEqual(request["table"]["rows"], [
			{
				"Zeit [s]": 0.0,
				"Strecke [m]": 0.0,
				"Strecke Fehler [m]": 0.1,
			},
			{
				"Zeit [s]": 1.0,
				"Strecke [m]": 2.0,
				"Strecke Fehler [m]": 0.1,
			},
			{
				"Zeit [s]": 2.0,
				"Strecke [m]": 4.0,
				"Strecke Fehler [m]": 0.2,
			},
		])
		self.assertEqual(request["previous_analysis"], first_answer)
		self.assertEqual(
			request["user_decision"],
			{"chart_type": "line"},
		)

	@patch("ai._post_chat_completion")
	def test_reanalysis_rejects_a_response_that_ignores_user_decision(
		self,
		post: Mock,
	) -> None:
		previous_spec = validate_ai_response(
			self.answer | {
				"needs_user_input": True,
				"user_input_reason": "Diagrammart ist unklar.",
				"user_input_options": [
					{
						"type": "chart_type",
						"question": "Welche Diagrammart?",
						"options": ["scatter", "line"],
					}
				],
			},
			self.data,
			self.columns,
		)
		post.return_value = {
			"choices": [
				{
					"message": {
						"content": json.dumps(self.answer),
					}
				}
			]
		}

		with self.assertRaisesRegex(AIServiceError, "verbindliche Benutzerentscheidung"):
			reanalyze_with_user_input(
				self.data,
				self.columns,
				{
					"endpoint": "https://example.invalid/v1/chat/completions",
					"api_key": "",
					"model": "test-model",
				},
				previous_spec,
				{"chart_type": "line"},
			)

	@patch("ai._post_chat_completion")
	def test_reanalysis_rejects_another_user_input_request(
		self,
		post: Mock,
	) -> None:
		previous_spec = validate_ai_response(
			self.answer | {
				"needs_user_input": True,
				"user_input_reason": "Diagrammart ist unklar.",
				"user_input_options": [
					{
						"type": "chart_type",
						"question": "Welche Diagrammart?",
						"options": ["scatter", "line"],
					}
				],
			},
			self.data,
			self.columns,
		)
		second_answer = self.answer | {
			"chart_type": "line",
			"needs_user_input": True,
			"user_input_reason": "Trendlinie ist unklar.",
			"user_input_options": [
				{
					"type": "trendline",
					"question": "Welche Trendlinie?",
					"options": ["none", "linear"],
				}
			],
		}
		post.return_value = {
			"choices": [
				{
					"message": {
						"content": json.dumps(second_answer),
					}
				}
			]
		}

		with self.assertRaisesRegex(AIServiceError, "noch weitere Angaben"):
			reanalyze_with_user_input(
				self.data,
				self.columns,
				{
					"endpoint": "https://example.invalid/v1/chat/completions",
					"api_key": "",
					"model": "test-model",
				},
				previous_spec,
				{"chart_type": "line"},
			)

	@patch("ai.requests.post")
	def test_http_error_is_safe_and_actionable(
		self,
		post: Mock,
	) -> None:
		response = Mock(
			status_code=401,
			content=b"server echoed details",
		)

		post.return_value = response

		secret = "test-secret-not-for-errors"

		with self.assertRaises(AIServiceError) as context:
			analyze_with_ai(
				self.data,
				self.columns,
				{
					"endpoint": (
						"https://example.invalid/"
						"v1/chat/completions"
					),
					"api_key": secret,
					"model": "test-model",
				},
			)

		self.assertIn(
			"HTTP 401",
			str(context.exception),
		)

		self.assertNotIn(
			secret,
			str(context.exception),
		)

		self.assertNotIn(
			"server echoed details",
			str(context.exception),
		)

	@patch("ai.requests.post")
	def test_api_key_is_only_sent_in_authorization_header(
		self,
		post: Mock,
	) -> None:
		response = Mock(
			status_code=200,
			content=(
				b'{"choices": '
				b'[{"message": {"content": "OK"}}]}'
			),
		)

		response.json.return_value = {
			"choices": [
				{
					"message": {
						"content": "OK"
					}
				}
			]
		}

		post.return_value = response

		secret = "header-only-test-key"

		from ai import test_ai_connection

		test_ai_connection({
			"endpoint": (
				"https://example.invalid/"
				"v1/chat/completions"
			),
			"api_key": secret,
			"model": "test-model",
		})

		kwargs = post.call_args.kwargs

		self.assertEqual(
			kwargs["headers"]["Authorization"],
			f"Bearer {secret}",
		)

		self.assertEqual(
			kwargs["timeout"],
			30,
		)

		self.assertNotIn(
			secret,
			json.dumps(kwargs["json"]),
		)

	def test_missing_config_file_is_reported(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			with self.assertRaisesRegex(
				FileNotFoundError,
				"Keine config.json",
			):
				load_config(
					Path(folder) / "config.json"
				)

	def test_long_ai_reason_is_capped_without_aborting_analysis(
		self,
	) -> None:
		answer = self.answer | {
			"reason": "Begründung " * 40
		}

		spec = validate_ai_response(
			answer,
			self.data,
			self.columns,
		)

		self.assertLessEqual(
			len(spec.reason),
			200,
		)

		self.assertTrue(
			spec.reason.endswith("...")
		)


if __name__ == "__main__":
	unittest.main()