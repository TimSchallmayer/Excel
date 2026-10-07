from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from uuid import uuid4
from xml.etree import ElementTree
from zipfile import ZipFile

import xlwings as xw
import pandas as pd

from excel_addin import _add_native_chart, _analysis_label, _available_trendlines, _chart_type_code, _discover_data_sources, _edit_chart_settings, _read_chart_state, _save_chart_state, _set_trendline, _swap_chart_axes, create_chart
from excel_bridge import _install_vba_bridge, _set_xlwings_config, create_addin
from models import ChartSpecification


@unittest.skipUnless(os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1", "Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed")
class ExcelAddinCOMTests(unittest.TestCase):
	def _analyze(self, frame, columns, config):
		self.assertIsInstance(config, dict)
		if "Zeit [s]" in frame.columns and "Zeit [s]" in columns:
			return ChartSpecification(
				x_column="Zeit [s]",
				y_column="Strecke [m]",
				x_label="Zeit [s]",
				y_label="Strecke [m]",
				confidence=0.96,
				reason="Testentscheidung für Zeit und Strecke.",
			)
		return ChartSpecification(
			x_column="Dehnung [mm]",
			y_column="Kraft [N]",
			x_label="Dehnung [mm]",
			y_label="Kraft [N]",
			confidence=0.96,
			reason="Testentscheidung für Dehnung und Kraft.",
		)

	def test_ribbon_addin_keeps_multiple_native_charts_with_individual_analyses(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			addin_path = create_addin(Path(folder) / f"PLVS Ribbon Test {uuid4().hex}.xlam")
			namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
			with ZipFile(addin_path) as package:
				ribbon = ElementTree.fromstring(package.read("customUI/customUI.xml"))
			axes_group = ribbon.find(".//r:group[@id='PLVSAxesGroup']", namespace)
			self.assertIsNotNone(axes_group)
			swap_button = axes_group.find("r:button[@id='PLVSSwapAxes']", namespace)
			self.assertIsNotNone(swap_button)
			self.assertEqual(swap_button.attrib["onAction"], "PLVS_RibbonSwapAxes")
			app = xw.App(visible=False, add_book=True)
			book = app.books.active
			installed_addins = []
			try:
				app.api.DisplayAlerts = False
				first = book.sheets[0]
				first.name = "Versuch1"
				first.range("A1").value = [
					["Zeit [s]", "Strecke [m]"],
					[0, 0],
					[1, 2.1],
					[2, 4.3],
					[3, 6.2],
				]
				first_table = first.api.ListObjects.Add(1, first.range("A1:B5").api, None, 1)
				first_table.Name = "Versuch1Tabelle"
				second = book.sheets.add("Versuch2")
				second.range("A1").value = [
					["Kraft [N]", "Dehnung [mm]"],
					[0, 0],
					[10, 1.5],
					[20, 3.0],
					[30, 4.5],
				]
				second_table = second.api.ListObjects.Add(1, second.range("A1:B5").api, None, 1)
				second_table.Name = "Versuch2Tabelle"
				for path in (
					Path(xw.__file__).resolve().parent / "addin" / "xlwings.xlam",
					addin_path,
				):
					addin = app.api.AddIns.Add(str(path), False)
					was_installed = bool(addin.Installed)
					if not was_installed:
						addin.Installed = True
					installed_addins.append((addin, was_installed))
				self.assertEqual(len(installed_addins), 2)
				self.assertTrue(all(addin.Installed for addin, _ in installed_addins))
				book.activate()

				initial_sheet_names = tuple(book.sheet_names)
				initial_book_count = len(app.books)
				sources = _discover_data_sources(book)
				self.assertEqual(len(sources), 2)
				self.assertIn("Versuch1: A1:B5", " ".join(source[2] for source in sources))
				self.assertIn("Versuch2: A1:B5", " ".join(source[2] for source in sources))
				first.activate()
				first.range("A1:B5").select()
				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze) as analyze,
				):
					create_chart()
				analyze.assert_called_once()
				self.assertEqual(
					list(analyze.call_args.args[0].columns),
					["Zeit [s]", "Strecke [m]"],
				)
				self.assertEqual(tuple(book.sheet_names), initial_sheet_names)
				self.assertEqual(len(app.books), initial_book_count)
				self.assertEqual(first.api.ChartObjects().Count, 1)
				self.assertEqual(second.api.ChartObjects().Count, 0)
				first_chart_object = first.api.ChartObjects(1)
				self.assertEqual(first_chart_object.Name, "PLVS_ULTRA_Chart_1")
				self.assertTrue(str(app.api.ActiveChart.Name).endswith(first_chart_object.Name))
				first_chart = first.api.ChartObjects(1).Chart
				self.assertEqual(tuple(first_chart.SeriesCollection(1).XValues), (0.0, 1.0, 2.0, 3.0))
				self.assertEqual(tuple(first_chart.SeriesCollection(1).Values), (0.0, 2.1, 4.3, 6.2))

				second.activate()
				second.range("A1:B5").select()
				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze),
				):
					create_chart()
				self.assertEqual(tuple(book.sheet_names), initial_sheet_names)
				self.assertEqual(len(app.books), initial_book_count)
				self.assertEqual(second.api.ChartObjects().Count, 1, str(app.api.StatusBar))
				self.assertEqual(first.api.ChartObjects().Count, 1)
				second_chart_object = second.api.ChartObjects(1)
				self.assertEqual(second_chart_object.Name, "PLVS_ULTRA_Chart_2")
				self.assertTrue(str(app.api.ActiveChart.Name).endswith(second_chart_object.Name))
				chart = second_chart_object.Chart
				self.assertEqual(chart.ChartType, -4169)
				self.assertEqual(tuple(chart.SeriesCollection(1).XValues), (0.0, 1.5, 3.0, 4.5))
				self.assertEqual(tuple(chart.SeriesCollection(1).Values), (0.0, 10.0, 20.0, 30.0))
				analysis_name = book.api.Names.Item("_PLVS_ULTRA_Analysis")
				analysis = str(book.app.api.Evaluate(analysis_name.Name))
				self.assertIn("Unabhängig: Dehnung [mm]", analysis)
				self.assertIn("Abhängig: Kraft [N]", analysis)
				self.assertIn("Diagramm: Punktdiagramm (XY); Trendlinie: kein", analysis)
				self.assertIn("Messpunkte: 4 von 4; Mittelwert y:", analysis)
				self.assertIn("KI-Einschätzung: 96%", analysis)
				first_analysis = str(
					book.app.api.Evaluate(
						book.api.Names.Item("_PLVS_ULTRA_Analysis_PLVS_ULTRA_Chart_1").Name
					)
				)
				self.assertIn("Unabhängig: Zeit [s]", first_analysis)
				self.assertIn("Abhängig: Strecke [m]", first_analysis)

				second.activate()
				second.range("A1:B5").select()
				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze) as analyze,
				):
					create_chart()
				analyze.assert_called_once()
				self.assertEqual(
					list(analyze.call_args.args[0].columns),
					["Kraft [N]", "Dehnung [mm]"],
				)
				self.assertEqual(tuple(book.sheet_names), initial_sheet_names)
				self.assertEqual(second.api.ChartObjects().Count, 2)
				chart_positions = {
					str(second.api.ChartObjects(index).Name): float(second.api.ChartObjects(index).Top)
					for index in range(1, int(second.api.ChartObjects().Count) + 1)
				}
				self.assertEqual(set(chart_positions), {"PLVS_ULTRA_Chart_2", "PLVS_ULTRA_Chart_3"})
				self.assertGreater(
					chart_positions["PLVS_ULTRA_Chart_3"],
					chart_positions["PLVS_ULTRA_Chart_2"],
				)
			finally:
				for addin, was_installed in reversed(installed_addins):
					try:
						addin.Installed = was_installed
					except Exception:
						pass
				try:
					book.close()
				except Exception:
					pass
				app.quit()

	def test_workbook_bridge_runs_without_dashboard_sheet_or_shape_buttons(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			workbook_path = Path(folder) / "messwerte.xlsm"
			app = xw.App(visible=False, add_book=True)
			book = None
			try:
				book = app.books.active
				sheet = book.sheets[0]
				sheet.name = "Messdaten"
				sheet.range("A1").value = [
					["Zeit [s]", "Strecke [m]"],
					[0, 0],
					[1, 2.1],
					[2, 4.3],
					[3, 6.2],
				]
				book.api.SaveAs(str(workbook_path), FileFormat=52)
				_install_vba_bridge(book)
				_set_xlwings_config(book)
				book.save()
				book.close()
				book = app.books.open(str(workbook_path), update_links=False, read_only=False)

				self.assertEqual(book.sheet_names, ["Messdaten", "xlwings.conf"])
				config_sheet = book.sheets["xlwings.conf"]
				config = dict(config_sheet.range("A1:B2").value)
				self.assertEqual(config["INTERPRETER_WIN"], str(Path(sys.executable).resolve()))
				self.assertIn("PYTHONPATH", config)

				data_sheet = book.sheets["Messdaten"]
				initial_sheet_names = tuple(book.sheet_names)
				data_sheet.activate()
				data_sheet.range("A1").select()
				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze),
				):
					create_chart()
				self.assertEqual(tuple(book.sheet_names), initial_sheet_names)
				self.assertEqual(data_sheet.api.ChartObjects().Count, 1)
				chart = data_sheet.api.ChartObjects(1).Chart
				self.assertEqual(chart.ChartType, -4169)
				self.assertEqual(chart.SeriesCollection().Count, 1)
				self.assertEqual(chart.SeriesCollection(1).XValues[0], 0)
				self.assertEqual(chart.SeriesCollection(1).Values[1], 2.1)
				analysis = str(book.app.api.Evaluate(book.api.Names.Item("_PLVS_ULTRA_Analysis").Name))
				self.assertIn("Unabhängig: Zeit [s]", analysis)
				self.assertIn("Abhängig: Strecke [m]", analysis)
			finally:
				if book is not None:
					book.close()
				app.quit()

	def test_native_chart_applies_custom_x_and_y_error_bars(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			sheet = book.sheets[0]
			sheet.range("A1").value = [
				["Input", "Output", "Input error", "Output error"],
				[1, 2, 0.1, 0.2],
				[2, 4, 0.1, 0.3],
				[3, 6, 0.2, 0.4],
			]
			data_range = sheet.range("A1:D4")
			frame = pd.DataFrame(
				data_range.value[1:],
				columns=data_range.value[0],
			)
			spec = ChartSpecification(
				x_column="Input",
				y_column="Output",
				x_label="Input",
				y_label="Output",
				x_error_column="Input error",
				y_error_column="Output error",
			)

			chart_name = _add_native_chart(
				book,
				sheet,
				data_range,
				frame,
				spec,
			)

			series = sheet.api.ChartObjects(chart_name).Chart.SeriesCollection(1)
			self.assertTrue(series.HasErrorBars)
		finally:
			book.close()
			app.quit()

	def test_native_categorical_bar_column_and_line_trend_charts(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			cases = (
				("Products", "bar", [("A", 120), ("B", 90)], "none", 57),
				("ProductsColumn", "column", [("A", 120), ("B", 90)], "none", 51),
				("MonthlyPrices", "line", [("Jan", 12), ("Feb", 13), ("Mar", 15)], "linear", 4),
			)
			for sheet_name, chart_type, rows, fit, expected_excel_type in cases:
				with self.subTest(chart_type=chart_type):
					sheet = book.sheets.add(sheet_name)
					sheet.range("A1").value = [["Category", "Value"], *rows]
					data_range = sheet.range("A1:B" + str(len(rows) + 1))
					frame = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
					spec = ChartSpecification(
						x_column="Category",
						y_column="Value",
						x_label="Category",
						y_label="Value",
						chart_type=chart_type,
						trendline=fit,
					)
					chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
					chart = sheet.api.ChartObjects(chart_name).Chart
					self.assertEqual(chart.ChartType, expected_excel_type)
					self.assertEqual(
						tuple(chart.SeriesCollection(1).XValues),
						tuple(category for category, _ in rows),
					)
					if fit != "none":
						self.assertEqual(chart.SeriesCollection(1).Trendlines().Count, 1)
					analysis = _analysis_label(frame, spec)
					self.assertIn(f"Messpunkte: {len(rows)} von {len(rows)}", analysis)
					self.assertIn("Trendlinie: linear" if fit == "linear" else "Trendlinie: kein", analysis)
		finally:
			book.close()
			app.quit()

	def test_native_trendline_types_degree_period_and_persisted_state(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		workbook_path = None
		try:
			sheet = book.sheets[0]
			sheet.name = "Trendlinien"
			rows = [["x", "y"], *[[x, 2 ** (x / 2)] for x in range(1, 9)]]
			sheet.range("A1").value = rows
			data_range = sheet.range("A1:B9")
			frame = pd.DataFrame(rows[1:], columns=rows[0])
			spec = ChartSpecification(
				x_column="x",
				y_column="y",
				x_label="x",
				y_label="y",
				chart_type="scatter",
			)
			chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
			chart = sheet.api.ChartObjects(chart_name).Chart
			series = chart.SeriesCollection(1)
			x_values = frame["x"].to_numpy(dtype=float)
			y_values = frame["y"].to_numpy(dtype=float)
			for trendline, expected_type in (
				("linear", -4132),
				("exponential", 5),
				("logarithmic", -4133),
				("power", 4),
				("polynomial", 3),
				("moving_average", 6),
			):
				with self.subTest(trendline=trendline):
					_set_trendline(
						series,
						trendline,
						4,
						4,
						x_values,
						y_values,
						int(chart.ChartType),
					)
					self.assertEqual(int(series.Trendlines(1).Type), expected_type)
					if trendline == "polynomial":
						self.assertEqual(int(series.Trendlines(1).Order), 4)
					if trendline == "moving_average":
						self.assertEqual(int(series.Trendlines(1).Period), 4)
			_set_trendline(
				series, "none", 4, 4, x_values, y_values, int(chart.ChartType)
			)
			self.assertEqual(int(series.Trendlines().Count), 0)

			_save_chart_state(book, chart_name, sheet.name, "A1:B9", spec)
			state = _read_chart_state(book, chart_name)
			options, degrees, periods = _available_trendlines(
				x_values, y_values, int(chart.ChartType)
			)
			self.assertIn("polynomial", options)
			self.assertIn("moving_average", options)
			self.assertEqual(degrees, (2, 3, 4, 5, 6))
			self.assertEqual(periods, (2, 3, 4, 5, 6, 7))

			def selected_chart(_book, _target_sheet=None, _target_chart=None):
				return chart_name, sheet, chart, _read_chart_state(book, chart_name)

			with (
				patch("excel_addin._selected_chart", side_effect=selected_chart),
				patch("excel_addin._frame_for_state", return_value=(sheet, data_range, frame)),
				patch("excel_addin._write_analysis"),
				patch("excel_addin._write_status"),
			):
				_edit_chart_settings(book, "trendline", options.index("polynomial"))
				_edit_chart_settings(book, "polynomial_degree", 3)
				self.assertEqual(int(series.Trendlines(1).Order), 5)
				_edit_chart_settings(book, "trendline", options.index("moving_average"))
				_edit_chart_settings(book, "moving_average_period", 3)
				self.assertEqual(int(series.Trendlines(1).Type), 6)
				self.assertEqual(int(series.Trendlines(1).Period), 5)

			saved_spec = _read_chart_state(book, chart_name)["spec"]
			self.assertEqual(saved_spec["trendline"], "moving_average")
			self.assertEqual(saved_spec["polynomial_degree"], 5)
			self.assertEqual(saved_spec["moving_average_period"], 5)

			with tempfile.TemporaryDirectory() as folder:
				workbook_path = Path(folder) / "trendline-state.xlsx"
				book.api.SaveAs(str(workbook_path), FileFormat=51)
				book.close()
				book = app.books.open(str(workbook_path), update_links=False, read_only=False)
				self.assertEqual(
					_read_chart_state(book, chart_name)["spec"]["moving_average_period"],
					5,
				)
				reloaded_chart = book.sheets["Trendlinien"].api.ChartObjects(chart_name).Chart
				self.assertEqual(
					int(reloaded_chart.SeriesCollection(1).Trendlines(1).Period),
					5,
				)
				book.close()
		finally:
			if book is not None:
				try:
					book.close()
				except Exception:
					pass
			app.quit()

	def test_native_axis_swap_preserves_chart_state_and_does_not_change_cells(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			cases = (
				("ScatterSwap", "scatter", "Zeit [s]", "Strecke [m]", [0, 1, 2, 3], [0, 2, 4, 6], -4169),
				("LineSwap", "line", "Zeit [s]", "Strecke [m]", [0, 1, 2, 3], [0, 2, 4, 6], 75),
				("LineScatterSwap", "line_scatter", "Zeit [s]", "Strecke [m]", [0, 1, 2, 3], [0, 2, 4, 6], 74),
			)
			for sheet_name, chart_type, x_name, y_name, x_values, y_values, expected_chart_type in cases:
				with self.subTest(chart_type=chart_type):
					sheet = book.sheets.add(sheet_name)
					rows = [
						[x_name, y_name, "X Fehler", "Y Fehler"],
						*[list(values) for values in zip(x_values, y_values, [0.1] * 4, [0.2] * 4)],
					]
					sheet.range("A1").value = rows
					data_range = sheet.range("A1:D5")
					frame = pd.DataFrame(rows[1:], columns=rows[0])
					spec = ChartSpecification(
						x_column=x_name,
						y_column=y_name,
						x_label=x_name,
						y_label=y_name,
						x_unit="s",
						y_unit="m",
						chart_type=chart_type,
						trendline="linear",
						x_error_column="X Fehler",
						y_error_column="Y Fehler",
						show_points=True,
						connect_points=chart_type != "scatter",
					)
					before = sheet.range("A1:D5").value
					chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
					chart_object = sheet.api.ChartObjects(chart_name)
					chart = chart_object.Chart
					series = chart.SeriesCollection(1)
					series.Format.Line.Weight = 2
					series.MarkerForegroundColor = 0x563412
					state = {
						"sheet": sheet.name,
						"address": "A1:D5",
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
					swapped = _swap_chart_axes(book, chart_name, chart, state)
					self.assertEqual(swapped.x_column, y_name)
					self.assertEqual(swapped.y_column, x_name)
					self.assertEqual((swapped.x_label, swapped.y_label), (y_name, x_name))
					self.assertEqual((swapped.x_unit, swapped.y_unit), ("m", "s"))
					self.assertEqual((swapped.x_error_column, swapped.y_error_column), ("Y Fehler", "X Fehler"))
					if chart_type == "scatter":
						self.assertIn(chart.ChartType, {-4169, 74})
						self.assertFalse(bool(chart.SeriesCollection(1).Format.Line.Visible))
					else:
						self.assertEqual(chart.ChartType, expected_chart_type)
					series = chart.SeriesCollection(1)
					self.assertEqual(tuple(series.XValues), tuple(y_values))
					self.assertEqual(tuple(series.Values), tuple(x_values))
					self.assertEqual(series.Trendlines().Count, 1)
					self.assertTrue(series.HasErrorBars)
					if chart_type != "scatter":
						self.assertEqual(series.Format.Line.Weight, 2)
					if series.MarkerForegroundColor != -1:
						self.assertEqual(series.MarkerForegroundColor, 0x563412)
					self.assertEqual(sheet.range("A1:D5").value, before)

			for sheet_name, chart_type, expected_type in (
				("BarSwap", "bar", 51),
				("ColumnSwap", "column", 57),
				("CategoryLineSwap", "line", 57),
			):
				with self.subTest(chart_type=chart_type):
					sheet = book.sheets.add(sheet_name)
					rows = [["Produkt", "Verkäufe"], ["A", 120], ["B", 90], ["C", 60]]
					sheet.range("A1").value = rows
					data_range = sheet.range("A1:B4")
					frame = pd.DataFrame(rows[1:], columns=rows[0])
					spec = ChartSpecification(
						x_column="Produkt",
						y_column="Verkäufe",
						x_label="Produkt",
						y_label="Verkäufe",
						chart_type=chart_type,
						trendline="linear" if chart_type == "line" else "none",
						connect_points=chart_type == "line",
					)
					before = sheet.range("A1:B4").value
					chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
					chart = sheet.api.ChartObjects(chart_name).Chart
					swapped = _swap_chart_axes(
						book,
						chart_name,
						chart,
						{
							"sheet": sheet.name,
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
						},
					)
					self.assertEqual(chart.ChartType, expected_type)
					self.assertEqual((swapped.x_column, swapped.y_column), ("Verkäufe", "Produkt"))
					self.assertEqual(sheet.range("A1:B4").value, before)
					if spec.trendline != "none":
						self.assertEqual(chart.SeriesCollection(1).Trendlines().Count, 1)
		finally:
			book.close()
			app.quit()

	def test_native_chart_gridlines_follow_chart_type_and_persist_manual_state(self) -> None:
		temporary_folder = tempfile.TemporaryDirectory()
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			cases = (
				("Scatter", "scatter", "Zeit [s]", [0, 1, 2], (True, True)),
				("Bar", "bar", "Monat", ["Jan", "Feb", "Mär"], (True, False)),
				("Column", "column", "Monat", ["Jan", "Feb", "Mär"], (True, False)),
				("CategoryLine", "line", "Monat", ["Jan", "Feb", "Mär"], (False, True)),
				("NumericLine", "line", "Zeit [s]", [0, 1, 2], (True, True)),
			)
			for sheet_name, chart_type, x_name, x_values, expected in cases:
				with self.subTest(chart_type=chart_type):
					sheet = book.sheets.add(sheet_name)
					rows = [
						[x_name, "Umsatz [€]"],
						*[list(row) for row in zip(x_values, [10, 20, 30])],
					]
					sheet.range("A1").value = rows
					data_range = sheet.range(f"A1:B{len(rows)}")
					frame = pd.DataFrame(rows[1:], columns=rows[0])
					spec = ChartSpecification(
						x_column=x_name,
						y_column="Umsatz [€]",
						x_label=x_name,
						y_label="Umsatz [€]",
						chart_type=chart_type,
						connect_points=chart_type == "line",
					)
					chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
					chart_object = sheet.api.ChartObjects(chart_name)
					chart_object.Activate()
					chart = chart_object.Chart
					expected_chart_type = _chart_type_code(
						chart_type,
						spec.show_points,
						spec.connect_points,
					)
					self.assertEqual(int(chart.ChartType), expected_chart_type)
					expected_major_axes = tuple(
						bool(chart.Axes(axis).HasMajorGridlines)
						for axis in (1, 2)
					)
					self.assertEqual(expected_major_axes, expected)
					_save_chart_state(
						book,
						chart_name,
						sheet.name,
						str(data_range.address),
						spec,
						{"gridlines_enabled": True},
					)
					if chart_type == "scatter":
						chart.Axes(1).HasMinorGridlines = True
					_edit_chart_settings(book, "gridlines")
					for axis_type in (1, 2):
						axis = chart.Axes(axis_type)
						self.assertFalse(bool(axis.HasMajorGridlines))
						self.assertFalse(bool(axis.HasMinorGridlines))
					self.assertFalse(_read_chart_state(book, chart_name)["settings"]["gridlines_enabled"])
					_edit_chart_settings(book, "gridlines")
					actual_major_axes = tuple(
						bool(chart.Axes(axis).HasMajorGridlines)
						for axis in (1, 2)
					)
					self.assertEqual(actual_major_axes, expected)
					_edit_chart_settings(book, "gridlines")
					with patch("excel_addin._ask_text", return_value="Gitternetzlinien bleiben aus"):
						_edit_chart_settings(book, "title")
					for axis_type in (1, 2):
						axis = chart.Axes(axis_type)
						self.assertFalse(bool(axis.HasMajorGridlines))
						self.assertFalse(bool(axis.HasMinorGridlines))
					self.assertFalse(_read_chart_state(book, chart_name)["settings"]["gridlines_enabled"])

			workbook_path = Path(temporary_folder.name) / "gridlines.xlsx"
			book.api.SaveAs(str(workbook_path), FileFormat=51)
			book.close()
			book = app.books.open(str(workbook_path))
			for sheet_name, *_ in cases:
				sheet = book.sheets[sheet_name]
				chart = sheet.api.ChartObjects(1).Chart
				chart_name = str(sheet.api.ChartObjects(1).Name)
				self.assertFalse(_read_chart_state(book, chart_name)["settings"]["gridlines_enabled"])
				for axis_type in (1, 2):
					axis = chart.Axes(axis_type)
					self.assertFalse(bool(axis.HasMajorGridlines))
					self.assertFalse(bool(axis.HasMinorGridlines))
		finally:
			book.close()
			app.quit()
			temporary_folder.cleanup()

	def test_native_chart_manual_formatting_and_text_edits(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			sheet = book.sheets[0]
			rows = [["Input", "Output"], [1, 2], [2, 4], [3, 6]]
			sheet.range("A1").value = rows
			data_range = sheet.range("A1:B4")
			frame = pd.DataFrame(rows[1:], columns=rows[0])
			spec = ChartSpecification(
				x_column="Input",
				y_column="Output",
				x_label="Input",
				y_label="Output",
				trendline="linear",
				chart_type="scatter_lines",
				connect_points=True,
			)
			chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
			chart_object = sheet.api.ChartObjects(chart_name)
			chart_object.Activate()
			chart = chart_object.Chart
			_save_chart_state(book, chart_name, sheet.name, "A1:B4", spec)
			gridlines_before = bool(chart.Axes(2).HasMajorGridlines)
			bold_before = bool(chart.ChartTitle.Font.Bold)
			italic_before = bool(chart.ChartTitle.Font.Italic)
			with (
				patch("excel_addin._ask_text", side_effect=["Updated title", "Horizontal", "Vertical"]),
			):
				from excel_addin import _edit_chart_settings
				for action, value in (
					("series_color", 0x563412),
					("line_color", 0xEFCDAB),
					("line_width", 4),
					("point_size", 3),
					("chart_type", 2),
					("trendline", 0),
					("title", None),
					("x_label", None),
					("y_label", None),
					("legend", None),
					("gridlines", None),
					("font_name", 1),
					("font_size", 5),
					("font_color", 0x123456),
					("font_bold", None),
					("font_italic", None),
				):
					_edit_chart_settings(book, action, value)
					series = chart.SeriesCollection(1)
					if action == "series_color":
						self.assertEqual(series.MarkerForegroundColor, 0x563412)
					elif action == "line_color":
						self.assertEqual(series.Format.Line.ForeColor.RGB, 0xEFCDAB)
			series = chart.SeriesCollection(1)
			self.assertEqual(series.MarkerBackgroundColor, 0x563412)
			self.assertEqual(series.Format.Line.ForeColor.RGB, 0xEFCDAB)
			self.assertEqual(series.MarkerForegroundColor, 0xEFCDAB)
			self.assertEqual(series.Format.Line.Weight, 3)
			self.assertEqual(series.MarkerSize, 9)
			self.assertEqual(chart.ChartType, 4)
			self.assertEqual(series.Trendlines().Count, 0)
			self.assertEqual(chart.ChartTitle.Text, "Updated title")
			self.assertEqual(chart.Axes(1).AxisTitle.Text, "Horizontal")
			self.assertEqual(chart.Axes(2).AxisTitle.Text, "Vertical")
			self.assertTrue(chart.HasLegend)
			self.assertNotEqual(bool(chart.Axes(2).HasMajorGridlines), gridlines_before)
			self.assertEqual(chart.ChartTitle.Font.Name, "Arial")
			self.assertEqual(chart.ChartTitle.Font.Size, 14)
			self.assertEqual(chart.ChartTitle.Font.Color, 0x123456)
			self.assertEqual(bool(chart.ChartTitle.Font.Bold), not bold_before)
			self.assertEqual(bool(chart.ChartTitle.Font.Italic), not italic_before)
			for target_index, target_object in (
				(1, chart.Axes(1).AxisTitle),
				(2, chart.Axes(2).AxisTitle),
				(3, chart.Legend),
			):
				_edit_chart_settings(book, "text_target", target_index)
				_edit_chart_settings(book, "font_size", 2)
				self.assertEqual(target_object.Font.Size, 10)
			saved_state = _read_chart_state(book, chart_name)
			self.assertEqual(saved_state["settings"]["series_color"], 0x563412)
			self.assertEqual(saved_state["settings"]["line_width"], 3)
			self.assertEqual(saved_state["settings"]["point_size"], 9)
			self.assertEqual(saved_state["settings"]["text_formats"]["title"]["name"], "Arial")
			self.assertEqual(saved_state["settings"]["text_formats"]["x_axis"]["size"], 10)
			self.assertEqual(saved_state["settings"]["text_formats"]["y_axis"]["size"], 10)
			self.assertEqual(saved_state["settings"]["text_formats"]["legend"]["size"], 10)
		finally:
			book.close()
			app.quit()

	def test_native_manual_edits_preserve_chart_identity_and_chart_type_state(self) -> None:
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		try:
			sheet = book.sheets[0]
			sheet.range("A1").value = [["Input", "Output"], [1, 2], [2, 4], [3, 6]]
			data_range = sheet.range("A1:B4")
			frame = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
			chart_names = []
			for _ in range(2):
				spec = ChartSpecification(
					x_column="Input",
					y_column="Output",
					x_label="Input",
					y_label="Output",
					chart_type="scatter",
				)
				chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
				chart_names.append(chart_name)
				_save_chart_state(book, chart_name, sheet.name, "A1:B4", spec)
			chart_a = sheet.api.ChartObjects(chart_names[0]).Chart
			chart_b = sheet.api.ChartObjects(chart_names[1]).Chart
			chart_b_type = int(chart_b.ChartType)
			chart_b_legend = bool(chart_b.HasLegend)
			with (
				patch("excel_addin._write_analysis"),
				patch("excel_addin._write_status"),
			):
				for index, expected_spec_type in (
					(2, "line"),
					(0, "scatter"),
					(3, "line_scatter"),
					(5, "column"),
					(4, "bar"),
					(0, "scatter"),
				):
					with self.subTest(chart_type=expected_spec_type):
						_edit_chart_settings(
							book,
							"chart_type",
							index,
							target_sheet=sheet.name,
							target_chart=chart_names[0],
						)
						saved_state = _read_chart_state(book, chart_names[0])
						saved_spec = saved_state["spec"]
						saved_settings = saved_state["settings"]
						self.assertEqual(saved_spec["chart_type"], expected_spec_type)
						if "line_width" in saved_settings:
							self.assertGreaterEqual(float(saved_settings["line_width"]), 0.75)
							self.assertLessEqual(float(saved_settings["line_width"]), 6)
						self.assertEqual(
							int(chart_a.ChartType),
							_chart_type_code(
								expected_spec_type,
								saved_spec["show_points"],
								saved_spec["connect_points"],
							),
						)
						self.assertEqual(int(chart_b.ChartType), chart_b_type)
						self.assertTrue(str(app.api.ActiveChart.Name).endswith(chart_names[0]))

				chart_a_line_color = int(chart_a.SeriesCollection(1).Format.Line.ForeColor.RGB)
				_edit_chart_settings(
					book,
					"chart_type",
					2,
					target_sheet=sheet.name,
					target_chart=chart_names[1],
				)
				_edit_chart_settings(
					book,
					"series_color",
					0x34AB12,
					target_sheet=sheet.name,
					target_chart=chart_names[1],
				)
				self.assertEqual(int(chart_b.ChartType), 4)
				self.assertEqual(
					int(chart_b.SeriesCollection(1).Format.Line.ForeColor.RGB),
					0x34AB12,
					f"type={chart_b.ChartType}; line_visible={chart_b.SeriesCollection(1).Format.Line.Visible}; "
					f"marker={chart_b.SeriesCollection(1).MarkerForegroundColor}; "
					f"state={_read_chart_state(book, chart_names[1])['settings']}",
				)
				self.assertEqual(
					int(chart_a.SeriesCollection(1).Format.Line.ForeColor.RGB),
					chart_a_line_color,
				)
				self.assertTrue(str(app.api.ActiveChart.Name).endswith(chart_names[1]))
				_edit_chart_settings(
					book,
					"chart_type",
					5,
					target_sheet=sheet.name,
					target_chart=chart_names[1],
				)
				_edit_chart_settings(
					book,
					"series_color",
					0x0000FF,
					target_sheet=sheet.name,
					target_chart=chart_names[1],
				)
				self.assertEqual(int(chart_b.ChartType), 51)
				chart_b_type = int(chart_b.ChartType)
				self.assertEqual(int(chart_b.SeriesCollection(1).Format.Fill.ForeColor.RGB), 0x0000FF)
				_edit_chart_settings(
					book,
					"font_name",
					11,
					target_sheet=sheet.name,
					target_chart=chart_names[1],
				)
				self.assertEqual(str(chart_b.ChartTitle.Font.Name), "Segoe UI")

				for expected_legend in (True, False):
					_edit_chart_settings(
						book,
						"legend_state",
						expected_legend,
						target_sheet=sheet.name,
						target_chart=chart_names[0],
					)
					self.assertEqual(bool(chart_a.HasLegend), expected_legend)
					self.assertEqual(
						_read_chart_state(book, chart_names[0])["settings"]["legend"],
						expected_legend,
					)
				self.assertEqual(bool(chart_b.HasLegend), chart_b_legend)

				for expected_markers, expected_connections in (
					(False, True),
					(True, True),
					(True, False),
					(False, False),
				):
					chart_a_series = chart_a.SeriesCollection(1)
					actual_markers = int(chart_a_series.MarkerStyle) != -4142
					actual_connections = bool(chart_a_series.Format.Line.Visible)
					if actual_markers != expected_markers:
						_edit_chart_settings(
							book,
							"points",
							target_sheet=sheet.name,
							target_chart=chart_names[0],
						)
					if actual_connections != expected_connections:
						_edit_chart_settings(
							book,
							"connections_state",
							expected_connections,
							target_sheet=sheet.name,
							target_chart=chart_names[0],
						)
					self.assertEqual(
						int(chart_a.SeriesCollection(1).MarkerStyle) != -4142,
						expected_markers,
					)
					self.assertEqual(
						bool(chart_a.SeriesCollection(1).Format.Line.Visible),
						expected_connections,
					)
					state = _read_chart_state(book, chart_names[0])["spec"]
					self.assertEqual(state["show_points"], expected_markers)
					self.assertEqual(state["connect_points"], expected_connections)
					self.assertEqual(int(chart_b.ChartType), chart_b_type)

				chart_b_series = chart_b.SeriesCollection(1)
				chart_b_line_visibility = bool(chart_b_series.Format.Line.Visible)
				chart_b_marker_style = int(chart_b_series.MarkerStyle)
				for connected in (False, True):
					_edit_chart_settings(
						book, "chart_type", 0, target_sheet=sheet.name, target_chart=chart_names[0]
					)
					_edit_chart_settings(
						book, "connections_state", connected,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					series_a = chart_a.SeriesCollection(1)
					if int(series_a.MarkerStyle) == -4142:
						_edit_chart_settings(
							book, "points", None,
							target_sheet=sheet.name, target_chart=chart_names[0],
						)
					series_a = chart_a.SeriesCollection(1)
					_edit_chart_settings(
						book, "series_color", 0x123456,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					_edit_chart_settings(
						book, "point_size", 2,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					series_a = chart_a.SeriesCollection(1)
					self.assertEqual(int(series_a.MarkerStyle) != -4142, True)
					self.assertEqual(bool(series_a.Format.Line.Visible), connected)
					self.assertEqual(
						_read_chart_state(book, chart_names[0])["spec"]["connect_points"],
						connected,
					)

					_edit_chart_settings(
						book, "line_width", 4,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					self.assertEqual(bool(series_a.Format.Line.Visible), connected)
					self.assertEqual(float(series_a.Format.Line.Weight), 3.0)

					_edit_chart_settings(
						book, "chart_type", 4,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					self.assertEqual(int(chart_a.ChartType), 57)
					self.assertEqual(
						_read_chart_state(book, chart_names[0])["spec"]["connect_points"],
						False,
					)
					_edit_chart_settings(
						book, "chart_type", 0,
						target_sheet=sheet.name, target_chart=chart_names[0],
					)
					series_a = chart_a.SeriesCollection(1)
					self.assertEqual(int(chart_a.ChartType), -4169)
					self.assertEqual(int(series_a.MarkerStyle) != -4142, True)
					self.assertEqual(bool(series_a.Format.Line.Visible), False)
					self.assertEqual(int(series_a.MarkerSize), 7)
					self.assertEqual(int(series_a.MarkerForegroundColor), 0x123456)
					self.assertEqual(int(series_a.MarkerBackgroundColor), 0x123456)
					self.assertEqual(float(series_a.Format.Line.Weight), 3.0)
					self.assertEqual(
						_read_chart_state(book, chart_names[0])["spec"]["connect_points"],
						False,
					)
					self.assertEqual(int(chart_b.ChartType), chart_b_type)
					self.assertEqual(bool(chart_b_series.Format.Line.Visible), chart_b_line_visibility)
					self.assertEqual(int(chart_b_series.MarkerStyle), chart_b_marker_style)
		finally:
			book.close()
			app.quit()


if __name__ == "__main__":
	unittest.main()
