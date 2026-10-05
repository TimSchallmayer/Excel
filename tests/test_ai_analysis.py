from __future__ import annotations

import json
import unittest
from unittest.mock import Mock, patch

import pandas as pd

from ai import AIServiceError, _table_context, analyze_with_ai


class GeneralAIAnalysisTests(unittest.TestCase):
	def setUp(self) -> None:
		self.data = pd.DataFrame({
			"Week": [0, 1, 2, 3],
			"Active users": [1, 2, 4, 8],
		})
		self.columns = list(self.data.columns)
		self.response = {
			"independent_variable": "Week",
			"dependent_variable": "Active users",
			"x_column": "Week",
			"y_column": "Active users",
			"chart_type": "scatter",
			"x_axis_label": "Week",
			"y_axis_label": "Active users",
			"x_unit": "",
			"y_unit": "",
			"x_min": 0,
			"x_max": 3,
			"y_min": 1,
			"y_max": 8,
			"origin": False,
			"show_points": True,
			"trendline": "exponential",
			"x_error_column": None,
			"y_error_column": None,
			"confidence": 0.96,
			"reason": "The values double each week.",
		}

	def test_context_contains_generic_table_facts_without_local_relationship_rules(self) -> None:
		context = _table_context(self.data, self.columns)
		self.assertEqual([column["name"] for column in context["columns"]], self.columns)
		self.assertEqual(context["table_row_count"], 4)
		self.assertEqual(context["columns"][0]["numeric_min"], 0)
		self.assertEqual(context["columns"][0]["numeric_max"], 3)
		self.assertEqual(context["columns"][1]["missing_count"], 0)
		self.assertNotIn("local_relationship_hints", context)
		self.assertNotIn("quantity", context["columns"][0])
		self.assertEqual(context["rows"][0], {"Week": 0, "Active users": 1})

	@patch("ai._post_chat_completion")
	def test_ai_can_choose_an_exponential_fit_for_general_data(self, post: Mock) -> None:
		post.return_value = {
			"choices": [{"message": {"content": json.dumps(self.response)}}],
		}
		spec = analyze_with_ai(self.data, self.columns, {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		})
		self.assertEqual(spec.trendline, "exponential")
		self.assertEqual(spec.x_column, "Week")
		self.assertEqual(spec.y_column, "Active users")

	@patch("ai._post_chat_completion")
	def test_dependency_labels_must_match_the_selected_axis_columns(self, post: Mock) -> None:
		response = self.response | {"dependent_variable": "Week"}
		post.return_value = {
			"choices": [{"message": {"content": json.dumps(response)}}],
		}
		with self.assertRaisesRegex(AIServiceError, "abhängige Größe"):
			analyze_with_ai(self.data, self.columns, {
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "",
				"model": "test-model",
			})

	@patch("ai._post_chat_completion")
	def test_low_confidence_does_not_fall_back_to_local_rules(self, post: Mock) -> None:
		response = self.response | {"confidence": 0.5}
		post.return_value = {
			"choices": [{"message": {"content": json.dumps(response)}}],
		}
		with self.assertRaisesRegex(AIServiceError, "zu unsicher"):
			analyze_with_ai(self.data, self.columns, {
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "",
				"model": "test-model",
			})


if __name__ == "__main__":
	unittest.main()
