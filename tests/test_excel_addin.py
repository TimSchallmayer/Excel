from __future__ import annotations

import unittest
import re
import json
from dataclasses import asdict
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd

from excel_addin import _add_native_chart, _analysis_label, _available_trendlines, _chart_type_code, _choose_data_source, _connected_regions, _current_data_range, _display_address, _edit_chart_settings, _frame_from_values, _selected_color, _set_text_name, _source_for_chart, _source_label, _stored_color, _swap_chart_axes, choose_pending_choice, create_chart
from models import ChartSpecification


class DataSourceTests(unittest.TestCase):
	def test_trendline_choices_follow_chart_type_and_numeric_domains(self) -> None:
		all_options, degrees, periods = _available_trendlines(
			pd.Series([1, 2, 3, 4, 5, 6, 7, 8]).to_numpy(dtype=float),
			pd.Series([2, 4, 8, 16, 32, 64, 128, 256]).to_numpy(dtype=float),
			-4169,
		)
		self.assertEqual(
			all_options,
			("none", "linear", "exponential", "logarithmic", "power", "polynomial", "moving_average"),
		)
		self.assertEqual(degrees, (2, 3, 4, 5, 6))
		self.assertEqual(periods, (2, 3, 4, 5, 6, 7))

		negative_values, _, _ = _available_trendlines(
			[1, 2, 3, 4],
			[-2, 4, 8, 16],
			-4169,
		)
		self.assertNotIn("exponential", negative_values)
		self.assertNotIn("power", negative_values)
		self.assertIn("logarithmic", negative_values)
		repeated_x, _, _ = _available_trendlines([1, 1, 1], [2, 4, 8], -4169)
		self.assertNotIn("exponential", repeated_x)
		self.assertEqual(_available_trendlines([1, 2], [2, 3], 3), ((), (), ()))

	def test_failed_trendline_application_does_not_save_a_false_state(self) -> None:
		frame = pd.DataFrame({"x": [1, 2, 3], "y": [2, 4, 6]})
		spec = ChartSpecification(x_column="x", y_column="y", trendline="none")
		state = {"spec": {field: getattr(spec, field) for field in (
			"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
			"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
			"trendline", "polynomial_degree", "moving_average_period",
			"x_error_column", "y_error_column", "independent_variable",
			"dependent_variable", "show_points", "connect_points", "confidence", "reason",
		)}}
		series = SimpleNamespace()
		chart = SimpleNamespace(
			ChartType=-4169,
			SeriesCollection=Mock(return_value=series),
		)
		with (
			patch("excel_addin._selected_chart", return_value=("chart", None, chart, state)),
			patch("excel_addin._frame_for_state", return_value=(None, None, frame)),
			patch("excel_addin._set_trendline", side_effect=RuntimeError("Excel rejected it")),
			patch("excel_addin._write_analysis"),
			patch("excel_addin._write_status"),
			patch("excel_addin._save_chart_state") as save_state,
		):
			with self.assertRaisesRegex(RuntimeError, "Excel rejected it"):
				_edit_chart_settings(SimpleNamespace(), "trendline", 1)
		save_state.assert_not_called()

	def test_native_chart_applies_custom_x_and_y_error_bars(self) -> None:
		frame = pd.DataFrame({
			"Input": [1.0, 2.0, 3.0],
			"Output": [2.0, 4.0, 6.0],
			"Input error": [0.1, 0.1, 0.2],
			"Output error": [0.2, 0.3, 0.4],
		})
		series = SimpleNamespace(ErrorBar=Mock())
		chart = SimpleNamespace(
			ChartType=None,
			HasTitle=None,
			ChartTitle=SimpleNamespace(Text=None),
			HasLegend=None,
			SeriesCollection=Mock(return_value=SimpleNamespace(
				NewSeries=Mock(return_value=series),
			)),
			DisplayBlanksAs=None,
			Axes=Mock(side_effect=[
				SimpleNamespace(AxisTitle=SimpleNamespace(Text=None)),
				SimpleNamespace(AxisTitle=SimpleNamespace(Text=None)),
			]),
		)
		chart_object = SimpleNamespace(Name=None, Chart=chart)
		chart_objects = SimpleNamespace(
			Count=0,
			Add=Mock(return_value=chart_object),
		)
		sheet = SimpleNamespace(
			name="Data",
			used_range=SimpleNamespace(api=SimpleNamespace(Left=0, Width=100, Top=0)),
			api=SimpleNamespace(ChartObjects=Mock(return_value=chart_objects)),
			range=Mock(side_effect=lambda *args: SimpleNamespace(api=object())),
		)
		data_range = SimpleNamespace(
			api=SimpleNamespace(
				Row=1,
				Rows=SimpleNamespace(Count=4),
				Column=1,
			)
		)
		spec = ChartSpecification(
			x_column="Input",
			y_column="Output",
			x_label="Input",
			y_label="Output",
			x_error_column="Input error",
			y_error_column="Output error",
		)

		with (
			patch("excel_addin._next_chart_name", return_value="PLVS_ULTRA_Chart_1"),
			patch("excel_addin._set_text_name"),
		):
			_add_native_chart(SimpleNamespace(), sheet, data_range, frame, spec)

		self.assertEqual(series.ErrorBar.call_count, 2)
		self.assertEqual(
			[call.kwargs["Direction"] for call in series.ErrorBar.call_args_list],
			[-4168, 1],
		)
		for call in series.ErrorBar.call_args_list:
			self.assertEqual(call.kwargs["Include"], 1)
			self.assertEqual(call.kwargs["Type"], -4114)
			self.assertIs(call.kwargs["Amount"], call.kwargs["MinusValues"])

	def test_native_chart_without_error_columns_does_not_add_error_bars(self) -> None:
		frame = pd.DataFrame({"Input": [1.0, 2.0, 3.0], "Output": [2.0, 4.0, 6.0]})
		series = SimpleNamespace(ErrorBar=Mock())
		chart = SimpleNamespace(
			ChartType=None,
			HasTitle=None,
			ChartTitle=SimpleNamespace(Text=None),
			HasLegend=None,
			SeriesCollection=Mock(return_value=SimpleNamespace(
				NewSeries=Mock(return_value=series),
			)),
			DisplayBlanksAs=None,
			Axes=Mock(side_effect=[
				SimpleNamespace(AxisTitle=SimpleNamespace(Text=None)),
				SimpleNamespace(AxisTitle=SimpleNamespace(Text=None)),
			]),
		)
		chart_object = SimpleNamespace(Name=None, Chart=chart)
		chart_objects = SimpleNamespace(Count=0, Add=Mock(return_value=chart_object))
		sheet = SimpleNamespace(
			name="Data",
			used_range=SimpleNamespace(api=SimpleNamespace(Left=0, Width=100, Top=0)),
			api=SimpleNamespace(ChartObjects=Mock(return_value=chart_objects)),
			range=Mock(side_effect=lambda *args: SimpleNamespace(api=object())),
		)
		data_range = SimpleNamespace(
			api=SimpleNamespace(
				Row=1,
				Rows=SimpleNamespace(Count=4),
				Column=1,
			)
		)
		spec = ChartSpecification(
			x_column="Input",
			y_column="Output",
			x_label="Input",
			y_label="Output",
		)

		with (
			patch("excel_addin._next_chart_name", return_value="PLVS_ULTRA_Chart_1"),
			patch("excel_addin._set_text_name"),
		):
			_add_native_chart(SimpleNamespace(), sheet, data_range, frame, spec)

		series.ErrorBar.assert_not_called()

	def test_ai_question_is_queued_as_labeled_ribbon_options(self) -> None:
		frame = pd.DataFrame({"Zeit": [0, 1, 2], "Wert": [1, 2, 4]})
		spec = ChartSpecification(
			needs_user_input=True,
			user_input_options=[{
				"type": "chart_type",
				"question": "Welche Diagrammart passt?",
				"options": ["Punktdiagramm", "Liniendiagramm"],
			}],
		)
		sheet = SimpleNamespace(name="Data")
		data_range = SimpleNamespace(address="A1:B4")
		book = SimpleNamespace()
		with (
			patch("excel_addin._active_book", return_value=book),
			patch("excel_addin._safe_action", side_effect=lambda action: action(book)),
			patch("excel_addin._source_for_chart", return_value=(sheet, data_range, frame, ["Zeit", "Wert"])),
			patch("excel_addin._get_ai_config", return_value={}),
			patch("excel_addin.analyze_with_ai", return_value=spec),
			patch("excel_addin._set_pending_choice") as set_pending,
			patch("excel_addin._add_native_chart") as add_chart,
		):
			create_chart()
		set_pending.assert_called_once()
		self.assertEqual(set_pending.call_args.args[1:4], (
			"ai",
			"Welche Diagrammart passt?",
			["Punktdiagramm", "Liniendiagramm"],
		))
		self.assertEqual(set_pending.call_args.args[4]["kind"], "ai")
		add_chart.assert_not_called()

	def test_clicking_ai_ribbon_option_reanalyzes_and_creates_chart(self) -> None:
		frame = pd.DataFrame({"Woche": ["Mo", "Di"], "Nutzer": [12, 15]})
		provisional = ChartSpecification(
			needs_user_input=True,
			user_input_options=[{
				"type": "chart_type",
				"question": "Welche Darstellung?",
				"options": ["scatter", "line"],
			}],
		)
		final = ChartSpecification(x_column="Woche", y_column="Nutzer", chart_type="line")
		data_range = "A1:B3"
		sheet = SimpleNamespace(name="Data", range=Mock(return_value=data_range))
		book = SimpleNamespace(sheets={"Data": sheet})
		state = {
			"kind": "ai",
			"sheet": "Data",
			"address": data_range,
			"spec": asdict(provisional),
			"decisions": {},
			"question_index": 0,
		}
		stored = {
			"_PLVS_ULTRA_PendingChoice_Count": "2",
			"_PLVS_ULTRA_PendingChoice_Label": "ai|Welche Darstellung?",
			"_PLVS_ULTRA_PendingChoice_State": json.dumps(state),
			"_PLVS_ULTRA_PendingChoice_Item_1": "line",
		}
		with (
			patch("excel_addin._active_book", return_value=book),
			patch("excel_addin._safe_action", side_effect=lambda action: action(book)),
			patch("excel_addin._read_text_name", side_effect=lambda _book, name: stored.get(name)),
			patch("excel_addin._frame_from_range", return_value=frame),
			patch("excel_addin.get_numeric_columns", return_value=["Nutzer"]),
			patch("excel_addin._get_ai_config", return_value={"model": "test"}),
			patch("excel_addin.reanalyze_with_user_input", return_value=final) as reanalyze,
			patch("excel_addin._create_chart_from_spec") as create_chart_from_spec,
			patch("excel_addin._clear_pending_choice") as clear_pending,
		):
			choose_pending_choice(1)
		reanalyze.assert_called_once()
		self.assertEqual(reanalyze.call_args.args[4], {"chart_type": "line"})
		create_chart_from_spec.assert_called_once_with(book, sheet, data_range, frame, final)
		clear_pending.assert_called_once_with(book)

	def test_clicking_data_source_ribbon_option_continues_analysis(self) -> None:
		data_range = "A1:B4"
		sheet = SimpleNamespace(name="Messwerte", range=Mock(return_value=data_range))
		book = SimpleNamespace(sheets={"Messwerte": sheet})
		state = {
			"kind": "source",
			"sources": [{
				"sheet": "Messwerte",
				"address": data_range,
				"label": "Messwerte: A1:B4",
			}],
		}
		stored = {
			"_PLVS_ULTRA_PendingChoice_Count": "2",
			"_PLVS_ULTRA_PendingChoice_Label": "source|Welche Datenquelle?",
			"_PLVS_ULTRA_PendingChoice_State": json.dumps(state),
			"_PLVS_ULTRA_PendingChoice_Item_0": "Messwerte: A1:B4",
		}
		frame = pd.DataFrame({"Zeit": [0, 1, 2], "Wert": [2, 3, 4]})
		with (
			patch("excel_addin._active_book", return_value=book),
			patch("excel_addin._safe_action", side_effect=lambda action: action(book)),
			patch("excel_addin._read_text_name", side_effect=lambda _book, name: stored.get(name)),
			patch("excel_addin._store_source") as store_source,
			patch("excel_addin._clear_pending_choice") as clear_pending,
			patch("excel_addin._frame_from_range", return_value=frame),
			patch("excel_addin.get_numeric_columns", return_value=["Wert"]),
			patch("excel_addin._analyze_chart_source") as analyze_source,
		):
			choose_pending_choice(0)
		store_source.assert_called_once_with(book, sheet, data_range, "Messwerte: A1:B4")
		clear_pending.assert_called_once_with(book)
		analyze_source.assert_called_once_with(book, sheet, data_range, frame, ["Wert"])

	def test_create_chart_reanalyzes_before_creating_native_chart(self) -> None:
		frame = pd.DataFrame({"Zeit": [0, 1, 2], "Wert": [1, 2, 4]})
		provisional = ChartSpecification(
			needs_user_input=True,
			user_input_options=[
				{
					"type": "chart_type",
					"question": "Welche Diagrammart?",
					"options": ["scatter", "line"],
				}
			],
		)
		final = ChartSpecification(
			x_column="Zeit",
			y_column="Wert",
			x_label="Zeit",
			y_label="Wert",
		)
		calls = []
		sheet = SimpleNamespace(
			name="Data",
			activate=Mock(),
			api=SimpleNamespace(ChartObjects=Mock()),
		)
		data_range = SimpleNamespace(address="A1:B4")
		frame = pd.DataFrame({"Zeit": [0, 1, 2], "Wert": [1, 2, 4]})
		book = SimpleNamespace()

		with (
			patch("excel_addin._active_book", return_value=book),
			patch("excel_addin._safe_action", side_effect=lambda action: action(book)),
			patch(
				"excel_addin._source_for_chart",
				return_value=(sheet, data_range, frame, ["Zeit", "Wert"]),
			),
			patch("excel_addin._frame_from_range") as frame_from_range,
			patch("excel_addin.get_numeric_columns", return_value=["Zeit", "Wert"]),
			patch("excel_addin._get_ai_config", return_value={}),
			patch(
				"excel_addin.analyze_with_ai",
				side_effect=lambda *args: (calls.append("analyze"), provisional)[1],
			),
			patch("excel_addin._queue_ai_question") as queue_question,
			patch(
				"excel_addin.reanalyze_with_user_input",
				side_effect=lambda *args: (calls.append("reanalyze"), final)[1],
			) as reanalyze,
			patch(
				"excel_addin._add_native_chart",
				side_effect=lambda *args: (calls.append("chart"), "chart")[1],
			),
			patch("excel_addin._write_analysis"),
			patch("excel_addin._save_chart_state") as save_state,
			patch("excel_addin._write_status"),
			patch("excel_addin._message"),
		):
			create_chart()

		frame_from_range.assert_not_called()
		self.assertEqual(calls, ["analyze"])
		queue_question.assert_called_once_with(book, provisional, sheet, data_range)
		reanalyze.assert_not_called()
		save_state.assert_not_called()

	def test_axis_swap_updates_spec_sources_labels_units_and_errors(self) -> None:
		frame = pd.DataFrame({
			"Zeit [s]": [0, 1, 2],
			"Strecke [m]": [0, 2, 4],
			"X Fehler": [0.1, 0.1, 0.2],
			"Y Fehler": [0.2, 0.3, 0.4],
		})
		spec = ChartSpecification(
			x_column="Zeit [s]",
			y_column="Strecke [m]",
			x_label="Zeit [s]",
			y_label="Strecke [m]",
			x_unit="s",
			y_unit="m",
			x_error_column="X Fehler",
			y_error_column="Y Fehler",
			trendline="linear",
			connect_points=True,
		)
		state = {
			"sheet": "Data",
			"address": "A1:D4",
			"spec": {
				field: getattr(spec, field)
				for field in (
					"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
					"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
					"trendline", "polynomial_degree", "moving_average_period",
					"x_error_column", "y_error_column",
					"independent_variable", "dependent_variable", "show_points",
					"connect_points", "confidence", "reason",
				)
			},
		}
		series = SimpleNamespace(
			XValues=None,
			Values=None,
			Format=SimpleNamespace(
				Fill=SimpleNamespace(ForeColor=SimpleNamespace(RGB=0)),
				Line=SimpleNamespace(ForeColor=SimpleNamespace(RGB=0), Weight=1, Visible=True),
			),
			MarkerForegroundColor=0,
			MarkerBackgroundColor=0,
			MarkerSize=5,
			MarkerStyle=8,
		)
		axes = {
			index: SimpleNamespace(HasTitle=False, AxisTitle=SimpleNamespace(Text=""))
			for index in (1, 2)
		}
		chart = SimpleNamespace(
			SeriesCollection=Mock(return_value=series),
			Axes=Mock(side_effect=lambda index: axes[index]),
			HasTitle=False,
			ChartTitle=SimpleNamespace(Text=""),
		)
		sheet = SimpleNamespace(name="Data")
		data_range = SimpleNamespace()
		with (
			patch("excel_addin._frame_for_state", return_value=(sheet, data_range, frame)),
			patch("excel_addin._source_reference", side_effect=lambda _sheet, _range, column: f"ref:{column}"),
			patch("excel_addin._clear_error_bars") as clear_errors,
			patch("excel_addin._apply_error_bars") as apply_errors,
			patch("excel_addin._remove_trendlines"),
			patch("excel_addin._set_trendline") as set_trendline,
			patch("excel_addin._write_analysis") as write_analysis,
			patch("excel_addin._save_chart_state") as save_state,
			patch("excel_addin._write_status") as write_status,
		):
			result = _swap_chart_axes(SimpleNamespace(), "PLVS_ULTRA_Chart_1", chart, state)
		self.assertEqual(result.x_column, "Strecke [m]")
		self.assertEqual(result.y_column, "Zeit [s]")
		self.assertEqual(result.x_label, "Strecke [m]")
		self.assertEqual(result.y_label, "Zeit [s]")
		self.assertEqual((result.x_unit, result.y_unit), ("m", "s"))
		self.assertEqual((result.x_error_column, result.y_error_column), ("Y Fehler", "X Fehler"))
		self.assertEqual((series.XValues, series.Values), ("ref:Strecke [m]", "ref:Zeit [s]"))
		self.assertEqual(chart.ChartType, 74)
		clear_errors.assert_called_once_with(series)
		apply_errors.assert_called_once_with(series, sheet, data_range, result)
		set_trendline.assert_not_called()
		write_analysis.assert_called_once_with(SimpleNamespace(), frame, result, "PLVS_ULTRA_Chart_1")
		self.assertEqual(save_state.call_args.args[4], result)
		self.assertIn("Achsen getauscht", write_status.call_args.args[1])

	def test_categorical_axis_swap_changes_bar_orientation_without_replacing_sources(self) -> None:
		frame = pd.DataFrame({"Produkt": ["A", "B"], "Verkäufe": [120, 90]})
		for chart_type, expected_type in (("bar", "column"), ("column", "bar"), ("line", "bar")):
			with self.subTest(chart_type=chart_type):
				spec = ChartSpecification(
					x_column="Produkt",
					y_column="Verkäufe",
					x_label="Produkt",
					y_label="Verkäufe",
					chart_type=chart_type,
				)
				state = {
					"sheet": "Data",
					"address": "A1:B3",
					"spec": {
						field: getattr(spec, field)
						for field in (
							"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
							"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
							"trendline", "polynomial_degree", "moving_average_period",
							"x_error_column", "y_error_column",
							"independent_variable", "dependent_variable", "show_points",
							"connect_points", "confidence", "reason",
						)
					},
				}
				series = SimpleNamespace(
					XValues="categories",
					Values="sales",
					Format=SimpleNamespace(
						Fill=SimpleNamespace(ForeColor=SimpleNamespace(RGB=0)),
						Line=SimpleNamespace(ForeColor=SimpleNamespace(RGB=0), Weight=1, Visible=True),
					),
				)
				if chart_type == "line":
					series.MarkerForegroundColor = 0
					series.MarkerBackgroundColor = 0
					series.MarkerSize = 5
					series.MarkerStyle = 8
				chart = SimpleNamespace(
					SeriesCollection=Mock(return_value=series),
					Axes=Mock(side_effect=lambda _: SimpleNamespace(
						HasTitle=False,
						AxisTitle=SimpleNamespace(Text=""),
					)),
					HasTitle=False,
					ChartTitle=SimpleNamespace(Text=""),
				)
				with (
					patch("excel_addin._frame_for_state", return_value=(SimpleNamespace(name="Data"), SimpleNamespace(), frame)),
					patch("excel_addin._clear_error_bars"),
					patch("excel_addin._apply_error_bars"),
					patch("excel_addin._remove_trendlines"),
					patch("excel_addin._set_trendline"),
					patch("excel_addin._write_analysis"),
					patch("excel_addin._save_chart_state"),
					patch("excel_addin._write_status"),
				):
					result = _swap_chart_axes(SimpleNamespace(), "chart", chart, state)
				self.assertEqual(result.chart_type, expected_type)
				self.assertEqual(chart.ChartType, 51 if expected_type == "column" else 57)
				self.assertEqual((series.XValues, series.Values), ("categories", "sales"))
				self.assertEqual((result.x_column, result.y_column), ("Verkäufe", "Produkt"))

	def test_chart_style_edits_update_actual_properties_and_saved_state(self) -> None:
		frame = pd.DataFrame({"X": [1, 2, 3], "Y": [2, 4, 6]})
		spec = ChartSpecification(x_column="X", y_column="Y", x_label="X", y_label="Y")
		state = {
			"sheet": "Data",
			"address": "A1:B4",
			"spec": {
				field: getattr(spec, field)
				for field in (
					"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
					"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
					"trendline", "polynomial_degree", "moving_average_period",
					"x_error_column", "y_error_column",
					"independent_variable", "dependent_variable", "show_points",
					"connect_points", "confidence", "reason",
				)
			},
			"settings": {},
		}
		line = SimpleNamespace(Weight=None, Visible=True, ForeColor=SimpleNamespace(RGB=None))
		axes = {
			index: SimpleNamespace(
				HasTitle=False,
				AxisTitle=SimpleNamespace(Text=""),
				HasMajorGridlines=False,
			)
			for index in (1, 2)
		}
		series = SimpleNamespace(
			Format=SimpleNamespace(
				Fill=SimpleNamespace(ForeColor=SimpleNamespace(RGB=None)),
				Line=line,
			),
			Border=SimpleNamespace(Color=None),
			Trendlines=Mock(return_value=SimpleNamespace(Count=0)),
			MarkerForegroundColor=None,
			MarkerBackgroundColor=None,
			MarkerSize=7,
			MarkerStyle=8,
		)
		chart = SimpleNamespace(
			Parent=SimpleNamespace(Activate=Mock()),
			SeriesCollection=Mock(return_value=series),
			ChartTitle=SimpleNamespace(
				Text="Title",
				Font=SimpleNamespace(Name="Calibri", Size=11, Color=0, Bold=False, Italic=False),
			),
			HasTitle=True,
			HasLegend=False,
			Axes=Mock(side_effect=lambda index: axes[index]),
			ChartType=-4169,
		)
		actions = (
			("series_color", 0x563412),
			("line_color", 0xEFCDAB),
			("line_width", 4),
			("point_size", 3),
			("points", None),
			("connections", None),
			("trendline", 1),
			("title", "Updated title"),
			("x_label", "New X"),
			("y_label", "New Y"),
			("legend", None),
			("legend_state", True),
			("connections_state", True),
			("connections_state", False),
			("gridlines", None),
			("text_target", 0),
			("font_name", 1),
			("font_size", 5),
			("font_color", 0x123456),
			("font_bold", None),
			("font_italic", None),
			("title_toggle", None),
		)
		spec_fields = (
			"x_column", "y_column", "x_label", "y_label", "x_unit", "y_unit",
			"chart_type", "x_min", "x_max", "y_min", "y_max", "origin",
			"trendline", "polynomial_degree", "moving_average_period",
			"x_error_column", "y_error_column",
			"independent_variable", "dependent_variable", "show_points",
			"connect_points", "confidence", "reason",
		)
		def save_state(_book, _name, _sheet, _address, updated_spec, updated_settings):
			state["spec"] = {field: getattr(updated_spec, field) for field in spec_fields}
			state["settings"] = updated_settings
		with (
			patch("excel_addin._selected_chart", return_value=("chart", SimpleNamespace(), chart, state)),
			patch("excel_addin._frame_for_state", return_value=(SimpleNamespace(), SimpleNamespace(), frame)),
			patch("excel_addin._ask_text", side_effect=["Updated title", "New X", "New Y"]),
			patch("excel_addin._set_trendline") as set_fit,
			patch("excel_addin._write_analysis"),
			patch("excel_addin._save_chart_state", side_effect=save_state) as save_state,
			patch("excel_addin._write_status"),
		):
			for action, value in actions:
				_edit_chart_settings(SimpleNamespace(), action, value)
				if action == "series_color":
					self.assertEqual(series.Format.Line.ForeColor.RGB, 0x563412)
					self.assertEqual(series.MarkerForegroundColor, 0x563412)
					self.assertEqual(series.MarkerBackgroundColor, 0x563412)
		self.assertIsNone(series.Format.Fill.ForeColor.RGB)
		self.assertEqual(series.Format.Line.ForeColor.RGB, 0xEFCDAB)
		self.assertEqual((series.MarkerForegroundColor, series.MarkerBackgroundColor), (0x563412, 0x563412))
		self.assertEqual(series.Format.Line.Weight, 3)
		self.assertEqual(series.MarkerSize, 9)
		self.assertFalse(state["spec"]["show_points"])
		self.assertFalse(state["spec"]["connect_points"])
		self.assertEqual(chart.ChartType, -4169)
		self.assertEqual(set_fit.call_args.args[:2], (series, "linear"))
		self.assertEqual(set_fit.call_args.args[-1], -4169)
		self.assertEqual(chart.ChartTitle.Text, "Updated title")
		self.assertEqual(chart.Axes(1).AxisTitle.Text, "New X")
		self.assertEqual(chart.Axes(2).AxisTitle.Text, "New Y")
		self.assertTrue(chart.HasLegend)
		self.assertTrue(chart.Axes(2).HasMajorGridlines)
		self.assertEqual(chart.ChartTitle.Font.Name, "Arial")
		self.assertEqual(chart.ChartTitle.Font.Size, 14)
		self.assertEqual(chart.ChartTitle.Font.Color, 0x123456)
		self.assertTrue(chart.ChartTitle.Font.Bold)
		self.assertTrue(chart.ChartTitle.Font.Italic)
		self.assertFalse(chart.HasTitle)
		self.assertEqual(state["settings"]["text_formats"]["title"]["name"], "Arial")
		self.assertEqual(save_state.call_count, len(actions))

	def test_color_and_chart_type_helpers_are_validated(self) -> None:
		self.assertEqual(_selected_color(0x563412), 0x563412)
		with self.assertRaisesRegex(ValueError, "Farbwert"):
			_selected_color(-1)
		self.assertEqual(_stored_color("#123456"), 0x563412)
		self.assertEqual(_stored_color(0x563412), 0x563412)
		self.assertEqual(_chart_type_code("scatter", True, False), -4169)
		self.assertEqual(_chart_type_code("scatter", True, True), -4169)
		self.assertEqual(_chart_type_code("scatter_lines", True, True), 74)
		self.assertEqual(_chart_type_code("scatter_lines", False, True), 75)
		self.assertEqual(_chart_type_code("line", True, True, categorical=True), 4)
		self.assertEqual(_chart_type_code("line_scatter", True, True), 65)
		self.assertEqual(_chart_type_code("bar", True, False), 57)

	def test_data_values_become_frame_and_duplicate_headers_are_unique(self) -> None:
		frame = _frame_from_values((
			("Zeit [s]", "Spannung [V]", "Zeit [s]"),
			(0, 0, 0),
			(1, 2, 1),
			(None, None, None),
		))
		self.assertEqual(list(frame.columns), ["Zeit [s]", "Spannung [V]", "Zeit [s] (2)"])
		self.assertEqual(len(frame), 2)
		self.assertEqual(frame.iloc[1, 1], 2)

	def test_contiguous_regions_detect_separate_candidate_tables(self) -> None:
		regions = _connected_regions(
			[
				["Zeit [s]", "Strecke [m]", None, "Kraft [N]", "Dehnung [mm]"],
				[0, 0, None, 0, 0],
				[1, 2, None, 10, 1.5],
			]
		)
		self.assertEqual(regions, [(0, 0, 2, 1), (0, 3, 2, 4)])

	def test_data_source_selection_uses_clickable_modal_options(self) -> None:
		first = (
			SimpleNamespace(name="Versuch1"),
			SimpleNamespace(address="$A$1:$B$4"),
			"Versuch1: A1:B4",
		)
		second = (
			SimpleNamespace(name="Versuch2"),
			SimpleNamespace(address="$A$1:$B$4"),
			"Versuch2: A1:B4",
		)
		book = SimpleNamespace()
		with (
			patch("excel_addin._discover_data_sources", return_value=[first, second]),
			patch("excel_addin._show_data_source_dialog", return_value=1) as show_dialog,
			patch("excel_addin._store_source") as store_source,
		):
			selected = _choose_data_source(book)
		self.assertEqual(selected, second[:2])
		show_dialog.assert_called_once_with(book, ["Versuch1: A1:B4", "Versuch2: A1:B4"])
		store_source.assert_called_once_with(book, *second)

	def test_cancelled_data_source_dialog_does_not_select_or_continue(self) -> None:
		sources = [
			(SimpleNamespace(name="Versuch1"), SimpleNamespace(address="A1:B4"), "Tabelle 1"),
			(SimpleNamespace(name="Versuch2"), SimpleNamespace(address="A1:B4"), "Tabelle 2"),
		]
		with (
			patch("excel_addin._discover_data_sources", return_value=sources),
			patch("excel_addin._show_data_source_dialog", return_value=None),
			patch("excel_addin._store_source") as store_source,
		):
			self.assertIsNone(_choose_data_source(SimpleNamespace()))
		store_source.assert_not_called()

	def test_active_table_is_used_without_opening_source_dialog(self) -> None:
		book = SimpleNamespace()
		active = (SimpleNamespace(name="Aktiv"), SimpleNamespace(address="A1:B4"))
		frame = pd.DataFrame({"Zeit": [0, 1], "Wert": [2, 3]})
		columns = ["Zeit", "Wert"]
		with (
			patch("excel_addin._current_data_range", return_value=(*active, frame, columns)),
			patch("excel_addin._store_source") as store_source,
			patch("excel_addin._choose_data_source") as choose_source,
		):
			self.assertEqual(_source_for_chart(book), (*active, frame, columns))
		store_source.assert_called_once_with(book, active[0], active[1], "Aktiv: A1:B4")
		choose_source.assert_not_called()

	def test_current_data_range_reuses_the_validated_frame_and_columns(self) -> None:
		data_range = SimpleNamespace()
		frame = pd.DataFrame({"Zeit": [0, 1], "Wert": [2, 3]})
		columns = ["Zeit", "Wert"]
		sheet = SimpleNamespace(range=Mock(return_value=data_range))
		selected = SimpleNamespace(
			Areas=SimpleNamespace(Count=1),
			CountLarge=6,
			Worksheet=SimpleNamespace(
				Parent=SimpleNamespace(FullName="book.xlsx"),
				Name="Data",
			),
			Address="A1:B3",
		)
		book = SimpleNamespace(
			app=SimpleNamespace(api=SimpleNamespace(Selection=selected)),
			fullname="book.xlsx",
			sheets={"Data": sheet},
		)

		with (
			patch("excel_addin._frame_from_range", return_value=frame) as read_range,
			patch("excel_addin.get_numeric_columns", return_value=columns) as numeric_columns,
		):
			result = _current_data_range(book)

		self.assertEqual(result, (sheet, data_range, frame, columns))
		read_range.assert_called_once_with(data_range)
		numeric_columns.assert_called_once_with(frame)

	def test_single_active_cell_is_not_treated_as_a_marked_data_source(self) -> None:
		book = SimpleNamespace(
			app=SimpleNamespace(api=SimpleNamespace(Selection=SimpleNamespace(
				Areas=SimpleNamespace(Count=1),
				CountLarge=1,
			))),
		)
		self.assertIsNone(_current_data_range(book))

	def test_source_dialog_is_reached_only_when_no_active_table_exists(self) -> None:
		book = SimpleNamespace()
		source = (SimpleNamespace(name="Tabelle"), SimpleNamespace(address="A1:B4"))
		frame = pd.DataFrame({"Zeit": [0, 1], "Wert": [2, 3]})
		columns = ["Zeit", "Wert"]
		with (
			patch("excel_addin._current_data_range", return_value=None),
			patch("excel_addin._choose_data_source", return_value=source) as choose_source,
			patch("excel_addin._frame_from_range", return_value=frame),
			patch("excel_addin.get_numeric_columns", return_value=columns),
		):
			self.assertEqual(_source_for_chart(book), (*source, frame, columns))
		choose_source.assert_called_once_with(book)

	def test_display_address_uses_readable_relative_a1_notation(self) -> None:
		self.assertEqual(_display_address(SimpleNamespace(address="$A$1:$B$15")), "A1:B15")

	def test_source_label_uses_sheet_and_range_without_exclamation_mark(self) -> None:
		label = _source_label(
			SimpleNamespace(name="Tabelle1"),
			SimpleNamespace(address="$A$1:$B$8"),
		)
		self.assertEqual(label, "Tabelle1: A1:B8")

	def test_long_analysis_text_uses_excel_safe_formula_chunks(self) -> None:
		value = ('Analyse "KI" und Zusammenhang; ' * 15).strip()
		with patch("excel_addin._set_name") as set_name:
			_set_text_name(SimpleNamespace(), "_PLVS_ULTRA_Analysis", value)
		formula = set_name.call_args.args[2]
		escaped_chunks = re.findall(r'"((?:""|[^"])*)"', formula)
		self.assertGreater(len(escaped_chunks), 1)
		self.assertTrue(all(len(chunk) <= 240 for chunk in escaped_chunks))
		self.assertEqual("".join(chunk.replace('""', '"') for chunk in escaped_chunks), value)

	def test_analysis_summary_includes_relevant_measurement_statistics(self) -> None:
		frame = pd.DataFrame({"Zeit [s]": [0, 1, 2], "Strecke [m]": [0, 2, 4]})
		spec = ChartSpecification(
			x_column="Zeit [s]",
			y_column="Strecke [m]",
			x_label="Zeit [s]",
			y_label="Strecke [m]",
			x_unit="s",
			y_unit="m",
			trendline="linear",
			confidence=0.95,
		)
		label = _analysis_label(frame, spec)
		for expected in (
			"Unabhängig: Zeit [s]; Abhängig: Strecke [m]",
			"Zusammenhang: Zeit [s] → Strecke [m]",
			"Diagramm: Punktdiagramm (XY); Trendlinie: linear",
			"Messpunkte: 3 von 3; Mittelwert y: 2 m",
			"Statistik: Median y: 2 m; Min–Max y: 0–4 m; Stdabw. y: 2 m; Korrelation: 1",
			"KI-Einschätzung: 95% —",
		):
			self.assertIn(expected, label)
		self.assertNotIn("Veränderung:", label)

	def test_incomplete_rows_are_excluded_from_measurement_statistics(self) -> None:
		frame = pd.DataFrame({"Zeit": [1, 2, 3], "Messwert": [10, None, 30]})
		spec = ChartSpecification(
			x_column="Zeit",
			y_column="Messwert",
			x_label="Zeit",
			y_label="Messwert",
		)
		label = _analysis_label(frame, spec)
		self.assertIn("Messpunkte: 2 von 3; Mittelwert y: 20", label)
		self.assertIn("Median y: 20; Min–Max y: 10–30", label)

	def test_analysis_values_cannot_break_ribbon_field_separators(self) -> None:
		frame = pd.DataFrame({"Zeit | x": [1, 2], "Messwert | y": [2, 4]})
		spec = ChartSpecification(
			x_column="Zeit | x",
			y_column="Messwert | y",
			x_label="Zeit | x",
			y_label="Messwert | y",
			confidence=0.95,
			reason="Begründung | mit Trenner.",
		)
		label = _analysis_label(frame, spec)
		self.assertEqual(label.count(" | "), 5)
		self.assertIn("Unabhängig: Zeit / x; Abhängig: Messwert / y", label)
		self.assertIn("Begründung / mit Trenner.", label)

	def test_categorical_axis_swap_analysis_keeps_pair_count(self) -> None:
		frame = pd.DataFrame({"Verkäufe": [120, 90], "Produkt": ["A", "B"]})
		spec = ChartSpecification(
			x_column="Verkäufe",
			y_column="Produkt",
			x_label="Verkäufe",
			y_label="Produkt",
			chart_type="bar",
		)
		label = _analysis_label(frame, spec)
		self.assertIn("Messpunkte: 2 von 2", label)

	def test_price_analysis_includes_absolute_and_percentage_change(self) -> None:
		frame = pd.DataFrame({"Datum": [3, 1, 2], "Aktienkurs [EUR]": [15, 10, 12]})
		spec = ChartSpecification(
			x_column="Datum",
			y_column="Aktienkurs [EUR]",
			x_label="Datum",
			y_label="Aktienkurs [EUR]",
			y_unit="EUR",
			confidence=0.95,
		)
		label = _analysis_label(frame, spec)
		self.assertIn("Mittelwert y: 12.33 EUR", label)
		self.assertIn("Veränderung: Kurs 10 → 15 EUR; absolut +5 EUR; Prozent: +50%", label)

	def test_price_percentage_change_is_defined_when_start_value_is_nonzero(self) -> None:
		frame = pd.DataFrame({"Woche": [1, 2], "Preis [EUR]": [0, 5]})
		spec = ChartSpecification(
			x_column="Woche",
			y_column="Preis [EUR]",
			x_label="Woche",
			y_label="Preis [EUR]",
			y_unit="EUR",
		)
		self.assertIn(
			"Prozent: nicht definiert (Startwert 0)",
			_analysis_label(frame, spec),
		)

if __name__ == "__main__":
	unittest.main()
