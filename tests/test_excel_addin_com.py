from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from uuid import uuid4

import xlwings as xw

from excel_addin import _discover_data_sources, create_chart
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
				first.range("A1").select()
				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze),
				):
					create_chart()
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
				second.range("A1").select()
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

				with (
					patch("excel_addin._active_book", return_value=book),
					patch("excel_addin.analyze_with_ai", side_effect=self._analyze),
				):
					create_chart()
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


if __name__ == "__main__":
	unittest.main()
