from __future__ import annotations

import json
import unittest
from unittest.mock import Mock, patch

import pandas as pd

from ai import (
	AIServiceError,
	_numeric_trendline_evidence,
	_table_context,
	_trendline_candidate_fits,
	analyze_with_ai,
	validate_ai_response,
)
from prompts import ANALYSIS_SYSTEM_PROMPT


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
			"x_label": "Week",
			"y_label": "Active users",
			"x_unit": "",
			"y_unit": "",
			"x_min": 0,
			"x_max": 3,
			"y_min": 1,
			"y_max": 8,
			"origin": False,
			"show_points": True,
			"connect_points": False,
			"trendline": "exponential",
			"x_error_column": None,
			"y_error_column": None,
			"confidence": 0.96,
			"reason": "The values double each week.",
			"needs_user_input": False,
			"user_input_reason": "",
			"user_input_options": [],
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

	def test_ai_trendline_state_includes_polynomial_degree_and_moving_average_period(self) -> None:
		polynomial = validate_ai_response(
			self.response | {
				"trendline": "polynomial",
				"polynomial_degree": 3,
			},
			self.data,
			self.columns,
		)
		self.assertEqual(polynomial.trendline, "polynomial")
		self.assertEqual(polynomial.polynomial_degree, 3)

		moving_average = validate_ai_response(
			self.response | {
				"trendline": "moving_average",
				"moving_average_period": 3,
			},
			self.data,
			self.columns,
		)
		self.assertEqual(moving_average.trendline, "moving_average")
		self.assertEqual(moving_average.moving_average_period, 3)

	def test_ai_rejects_unavailable_polynomial_degree_and_average_period(self) -> None:
		with self.assertRaisesRegex(AIServiceError, "Polynomgrad"):
			validate_ai_response(
				self.response | {
					"trendline": "polynomial",
					"polynomial_degree": 4,
				},
				self.data,
				self.columns,
			)
		with self.assertRaisesRegex(AIServiceError, "mehr gültige Messpunkte"):
			validate_ai_response(
				self.response | {
					"trendline": "moving_average",
					"moving_average_period": 7,
				},
				self.data,
				self.columns,
			)

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
	def test_low_confidence_does_not_trigger_a_user_input_threshold(self, post: Mock) -> None:
		response = self.response | {"confidence": 0.5}
		post.return_value = {
			"choices": [{"message": {"content": json.dumps(response)}}],
		}
		spec = analyze_with_ai(self.data, self.columns, {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		})
		self.assertEqual(spec.confidence, 0.5)
		self.assertFalse(spec.needs_user_input)

	def test_axis_units_are_preserved_from_column_names(self) -> None:
		for x_name, y_name in (
			("Zeit [s]", "Strecke [m]"),
			("Zeit (s)", "Strecke in m"),
		):
			with self.subTest(x_name=x_name, y_name=y_name):
				data = pd.DataFrame({x_name: [0, 1, 2], y_name: [0, 2, 4]})
				response = self.response | {
					"independent_variable": x_name,
					"dependent_variable": y_name,
					"x_column": x_name,
					"y_column": y_name,
					"x_label": x_name,
					"y_label": y_name,
					"x_unit": "",
					"y_unit": "",
					"x_min": 0,
					"x_max": 2,
					"y_min": 0,
					"y_max": 4,
					"trendline": "linear",
				}
				spec = validate_ai_response(response, data, list(data.columns))
				self.assertEqual(spec.x_label, "Zeit [s]")
				self.assertEqual(spec.y_label, "Strecke [m]")
				self.assertEqual(spec.x_unit, "s")
				self.assertEqual(spec.y_unit, "m")

	def test_general_columns_do_not_gain_invented_units(self) -> None:
		spec = validate_ai_response(self.response, self.data, self.columns)
		self.assertEqual(spec.x_unit, "")
		self.assertEqual(spec.y_unit, "")
		self.assertEqual(spec.x_label, "Week")
		self.assertEqual(spec.y_label, "Active users")
		europe_data = pd.DataFrame({
			"Region in Europe": [1, 2, 3],
			"Sales": [10, 11, 13],
		})
		europe_response = self.response | {
			"independent_variable": "Region in Europe",
			"dependent_variable": "Sales",
			"x_column": "Region in Europe",
			"y_column": "Sales",
			"x_label": "Region in Europe",
			"y_label": "Sales",
			"x_unit": "",
			"y_unit": "",
			"x_min": 1,
			"x_max": 3,
			"y_min": 10,
			"y_max": 13,
			"trendline": "none",
		}
		europe_spec = validate_ai_response(
			europe_response,
			europe_data,
			list(europe_data.columns),
		)
		self.assertEqual(europe_spec.x_unit, "")
		self.assertEqual(europe_spec.x_label, "Region in Europe")

	@patch("ai._post_chat_completion")
	def test_ai_receives_numeric_values_for_trendline_selection(self, post: Mock) -> None:
		cases = (
			([1, 2, 3, 4, 5], [2, 4, 6, 8, 10], "linear"),
			([1, 2, 3, 4, 5], [1, 4, 9, 16, 25], "quadratic"),
			([0, 1, 2, 3, 4], [1, 2, 4, 8, 16], "exponential"),
			([1, 2, 3, 4, 5], [1, 20, 3, 18, 7], "none"),
		)
		for x_values, y_values, fit in cases:
			with self.subTest(fit=fit):
				data = pd.DataFrame({"Input": x_values, "Output": y_values})
				response = self.response | {
					"independent_variable": "Input",
					"dependent_variable": "Output",
					"x_column": "Input",
					"y_column": "Output",
					"x_label": "Input",
					"y_label": "Output",
					"x_unit": "",
					"y_unit": "",
					"x_min": min(x_values),
					"x_max": max(x_values),
					"y_min": min(y_values),
					"y_max": max(y_values),
					"trendline": fit,
				}
				post.return_value = {
					"choices": [{"message": {"content": json.dumps(response)}}],
				}
				spec = analyze_with_ai(data, list(data.columns), {
					"endpoint": "https://example.invalid/v1/chat/completions",
					"api_key": "",
					"model": "test-model",
				})
				request = json.loads(post.call_args.args[1][1]["content"])
				self.assertEqual(spec.trendline, fit)
				trendline_evidence = request["numeric_trendline_evidence"]
				pair = next(
					relationship
					for relationship in trendline_evidence["leading_pairs"]
					if relationship["x"] == "Input" and relationship["y"] == "Output"
				)
				if fit == "none":
					self.assertLess(
						float(pair["best_fit_candidates"][0]["adjusted_r_squared"]),
						0.0,
					)
				else:
					self.assertIn(
						fit,
						[candidate["model"] for candidate in pair["best_fit_candidates"]],
					)
				self.assertEqual(
					[(row["Input"], row["Output"]) for row in request["rows"]],
					list(zip(x_values, y_values)),
				)
				post.reset_mock()
		self.assertIn("tatsächlichen numerischen Werte", ANALYSIS_SYSTEM_PROMPT)

	def test_numeric_fit_evidence_distinguishes_supported_curve_families(self) -> None:
		cases = (
			("linear", [1, 2, 3, 4, 5], [2, 4, 6, 8, 10]),
			("quadratic", [1, 2, 3, 4, 5], [1, 4, 9, 16, 25]),
			("cubic", [0, 1, 2, 3, 4, 5], [1, 1, 3, 13, 37, 81]),
			("exponential", [0, 1, 2, 3, 4], [1, 2, 4, 8, 16]),
			("logarithmic", [1, 2, 3, 4, 5], [1, 2, 2.5, 3, 3.3]),
			("power", [1, 2, 3, 4, 5], [1, 2**1.5, 3**1.5, 8, 5**1.5]),
		)
		for expected_model, x_values, y_values in cases:
			with self.subTest(expected_model=expected_model):
				fits = _trendline_candidate_fits(
					pd.Series(x_values).to_numpy(dtype=float),
					pd.Series(y_values).to_numpy(dtype=float),
				)
				self.assertEqual(fits[0]["model"], expected_model)
				self.assertGreater(float(fits[0]["adjusted_r_squared"]), 0.95)

	def test_numeric_fit_evidence_does_not_overstate_unstructured_data(self) -> None:
		fits = _trendline_candidate_fits(
			pd.Series([1, 2, 3, 4, 5]).to_numpy(dtype=float),
			pd.Series([1, 20, 3, 18, 7]).to_numpy(dtype=float),
		)
		self.assertTrue(fits)
		self.assertLess(float(fits[0]["adjusted_r_squared"]), 0.0)
		evidence = _numeric_trendline_evidence(
			pd.DataFrame({"Input": [1, 2, 3, 4, 5], "Output": [1, 20, 3, 18, 7]}),
			["Input", "Output"],
		)
		self.assertEqual(evidence["evaluated_numeric_pairs"], 2)

	@patch("ai._post_chat_completion")
	def test_general_categorical_time_series_can_be_line_with_trendline(self, post: Mock) -> None:
		data = pd.DataFrame({
			"Monat": ["Jan", "Feb", "Mär", "Apr", "Mai"],
			"Preis": [12, 13, 14, 15, 17],
		})
		response = self.response | {
			"independent_variable": "Monat",
			"dependent_variable": "Preis",
			"x_column": "Monat",
			"y_column": "Preis",
			"x_label": "Monat",
			"y_label": "Preis",
			"x_unit": "",
			"y_unit": "",
			"chart_type": "line",
			"x_min": None,
			"x_max": None,
			"y_min": 12,
			"y_max": 17,
			"trendline": "linear",
			"connect_points": True,
		}
		post.return_value = {
			"choices": [{"message": {"content": json.dumps(response, ensure_ascii=False)}}],
		}
		spec = analyze_with_ai(data, ["Preis"], {
			"endpoint": "https://example.invalid/v1/chat/completions",
			"api_key": "",
			"model": "test-model",
		})
		self.assertEqual(spec.chart_type, "line")
		self.assertEqual(spec.trendline, "linear")
		self.assertTrue(spec.connect_points)

	def test_category_chart_types_are_validated(self) -> None:
		data = pd.DataFrame({"Produkt": ["A", "B"], "Verkäufe": [120, 90]})
		for chart_type in ("bar", "column"):
			with self.subTest(chart_type=chart_type):
				response = self.response | {
					"independent_variable": "Produkt",
					"dependent_variable": "Verkäufe",
					"x_column": "Produkt",
					"y_column": "Verkäufe",
					"x_label": "Produkt",
					"y_label": "Verkäufe",
					"x_unit": "",
					"y_unit": "",
					"chart_type": chart_type,
					"x_min": None,
					"x_max": None,
					"y_min": 90,
					"y_max": 120,
					"trendline": "none",
				}
				spec = validate_ai_response(response, data, ["Verkäufe"])
				self.assertEqual(spec.chart_type, chart_type)


if __name__ == "__main__":
	unittest.main()
