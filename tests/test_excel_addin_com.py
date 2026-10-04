from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

import xlwings as xw

from excel_addin import PRODUCT_NAME
from excel_bridge import _install_vba_bridge, _set_xlwings_config


@unittest.skipUnless(os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1", "Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed")
class ExcelAddinCOMTests(unittest.TestCase):
	def test_workbook_open_and_shape_macro_create_native_chart(self) -> None:
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

				self.assertIn(PRODUCT_NAME, book.sheet_names)
				dashboard = book.sheets[PRODUCT_NAME]
				self.assertEqual(dashboard.range("B3").value, "Messdaten")
				self.assertEqual(dashboard.range("B12").value, "Bereit")
				config_sheet = book.sheets["xlwings.conf"]
				config = dict(config_sheet.range("A1:B2").value)
				self.assertEqual(config["INTERPRETER_WIN"], str(Path(sys.executable).resolve()))
				self.assertIn("PYTHONPATH", config)

				button_actions = (
					("PLVS_Analyze", "PLVS_ULTRA_Analyze"),
					("PLVS_CreateChart", "PLVS_ULTRA_CreateChart"),
					("PLVS_AISettings", "PLVS_ULTRA_AISettings"),
					("PLVS_ShowAnalysis", "PLVS_ULTRA_ShowAnalysis"),
					("PLVS_Help", "PLVS_ULTRA_Help"),
				)
				for shape_name, macro_name in button_actions:
					self.assertEqual(
						dashboard.api.Shapes.Item(shape_name).OnAction,
						f"{book.name}!{macro_name}",
					)

				app.api.Run(dashboard.api.Shapes.Item("PLVS_Analyze").OnAction)
				self.assertEqual(dashboard.range("B17").value, "Zeit [s]")
				self.assertEqual(dashboard.range("B18").value, "Strecke [m]")
				self.assertEqual(book.sheets["Messdaten"].api.ChartObjects().Count, 0)

				shape = dashboard.api.Shapes.Item("PLVS_CreateChart")
				app.api.Run(shape.OnAction)
				book.save()

				data_sheet = book.sheets["Messdaten"]
				chart_count = data_sheet.api.ChartObjects().Count
				self.assertEqual(chart_count, 1, str(dashboard.range("B12").value))
				chart = data_sheet.api.ChartObjects(1).Chart
				self.assertEqual(chart.ChartType, -4169)
				self.assertEqual(chart.SeriesCollection().Count, 1)
				self.assertEqual(chart.SeriesCollection(1).XValues[0], 0)
				self.assertEqual(chart.SeriesCollection(1).Values[1], 2.1)
				self.assertEqual(dashboard.range("B17").value, "Zeit [s]")
				self.assertEqual(dashboard.range("B18").value, "Strecke [m]")
			finally:
				if book is not None:
					book.close()
				app.quit()


if __name__ == "__main__":
	unittest.main()
