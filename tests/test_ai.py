from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import pandas as pd

from ai import AIServiceError, analyze_with_ai, validate_ai_response
from config import load_config


class AIValidationTests(unittest.TestCase):
	def setUp(self) -> None:
		self.data = pd.DataFrame({
			"Zeit [s]": [0.0, 1.0, 2.0],
			"Strecke [m]": [0.0, 2.0, 4.0],
			"Strecke Fehler [m]": [0.1, 0.1, 0.2],
		})
		self.columns = ["Zeit [s]", "Strecke [m]"]
		self.answer = {
			"independent_variable": "Zeit",
			"dependent_variable": "Strecke",
			"x_column": "Zeit [s]",
			"y_column": "Strecke [m]",
			"chart_type": "scatter",
			"x_axis_label": "Zeit [s]",
			"y_axis_label": "Strecke [m]",
			"x_unit": "s",
			"y_unit": "m",
			"x_min": 0,
			"x_max": 2,
			"y_min": 0,
			"y_max": 4,
			"origin": True,
			"show_points": True,
			"trendline": "linear",
			"x_error_column": None,
			"y_error_column": "Strecke Fehler [m]",
			"confidence": 0.97,
			"reason": "Zeit ist die unabhängige Größe.",
		}

	def test_valid_response_creates_chart_specification(self) -> None:
		spec = validate_ai_response(self.answer, self.data, self.columns)
		self.assertEqual(spec.x_column, "Zeit [s]")
		self.assertEqual(spec.y_label, "Strecke [m]")
		self.assertEqual(spec.trendline, "linear")
		self.assertEqual(spec.confidence, 0.97)
		self.assertEqual(spec.y_error_column, "Strecke Fehler [m]")

	def test_invalid_response_missing_required_field(self) -> None:
		answer = self.answer.copy()
		del answer["reason"]
		with self.assertRaisesRegex(AIServiceError, "Pflichtfelder"):
			validate_ai_response(answer, self.data, self.columns)

	def test_missing_column_is_rejected(self) -> None:
		answer = self.answer | {"x_column": "Zeitpunkt"}
		with self.assertRaisesRegex(AIServiceError, "x-Spalte existiert nicht"):
			validate_ai_response(answer, self.data, self.columns)

	def test_unsupported_chart_type_is_rejected(self) -> None:
		answer = self.answer | {"chart_type": "pie"}
		with self.assertRaisesRegex(AIServiceError, "Diagrammart"):
			validate_ai_response(answer, self.data, self.columns)

	def test_invalid_axis_ranges_are_rejected(self) -> None:
		answer = self.answer | {"x_min": 3, "x_max": 2}
		with self.assertRaisesRegex(AIServiceError, "x_min muss kleiner"):
			validate_ai_response(answer, self.data, self.columns)

	def test_non_finite_axis_ranges_are_rejected(self) -> None:
		answer = self.answer | {"y_max": float("inf")}
		with self.assertRaisesRegex(AIServiceError, "endliche Zahl"):
			validate_ai_response(answer, self.data, self.columns)

	def test_confidence_out_of_range_is_rejected(self) -> None:
		answer = self.answer | {"confidence": 1.01}
		with self.assertRaisesRegex(AIServiceError, "zwischen 0 und 1"):
			validate_ai_response(answer, self.data, self.columns)

	def test_cubic_trendline_requires_four_distinct_x_values(self) -> None:
		answer = self.answer | {"trendline": "cubic"}
		with self.assertRaisesRegex(AIServiceError, "mindestens vier"):
			validate_ai_response(answer, self.data, self.columns)

	@patch("ai._post_chat_completion")
	def test_api_json_is_parsed_and_validated(self, post: Mock) -> None:
		post.return_value = {
			"choices": [{"message": {"content": "```json\n" + json.dumps(self.answer) + "\n```"}}],
		}
		spec = analyze_with_ai(self.data, self.columns, {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		})
		self.assertEqual(spec.x_column, "Zeit [s]")
		post.assert_called_once()

	@patch("ai.requests.post")
	def test_http_error_is_safe_and_actionable(self, post: Mock) -> None:
		response = Mock(status_code=401, content=b"server echoed details")
		post.return_value = response
		secret = "test-secret-not-for-errors"
		with self.assertRaises(AIServiceError) as context:
			analyze_with_ai(self.data, self.columns, {
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": secret,
				"model": "test-model",
			})
		self.assertIn("HTTP 401", str(context.exception))
		self.assertNotIn(secret, str(context.exception))
		self.assertNotIn("server echoed details", str(context.exception))

	@patch("ai.requests.post")
	def test_api_key_is_only_sent_in_authorization_header(self, post: Mock) -> None:
		response = Mock(
			status_code=200,
			content=b'{"choices": [{"message": {"content": "OK"}}]}',
		)
		response.json.return_value = {"choices": [{"message": {"content": "OK"}}]}
		post.return_value = response
		secret = "header-only-test-key"
		from ai import test_ai_connection
		test_ai_connection({
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": secret,
			"model": "test-model",
		})
		kwargs = post.call_args.kwargs
		self.assertEqual(kwargs["headers"]["Authorization"], f"Bearer {secret}")
		self.assertEqual(kwargs["timeout"], 30)
		self.assertNotIn(secret, json.dumps(kwargs["json"]))

	def test_missing_config_file_is_reported(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			with self.assertRaisesRegex(FileNotFoundError, "Keine config.json"):
				load_config(Path(folder) / "config.json")


if __name__ == "__main__":
	unittest.main()