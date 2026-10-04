from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

import xlwings as xw

from excel import get_numeric_columns
from excel_addin import _add_native_chart, _frame_from_sheet, _resolve_spec, _settings_sheet, _target_for


@unittest.skipUnless(os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1", "Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed")
class ExcelAddinCOMTests(unittest.TestCase):
	def test_active_workbook_is_copied_and_receives_native_chart(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			root = Path(folder)
			source_path = root / "messwerte.xlsx"
			app = xw.App(visible=False, add_book=True)
			book = app.books.active
			try:
				sheet = book.sheets[0]
				sheet.name = "Messwerte"
				sheet.range("A1").value = [
					["Zeit [s]", "Strecke [m]"],
					[0, 0],
					[1, 2.1],
					[2, 4.3],
					[3, 6.2],
				]
				book.save(source_path)
				original_bytes = source_path.read_bytes()
				frame = _frame_from_sheet(sheet)
				columns = get_numeric_columns(frame)
				settings = _settings_sheet(book, columns)
				settings.range("B3").value = sheet.name
				settings.range("B7").value = "linear"
				spec, _ = _resolve_spec(book, frame, columns, allow_ai=False)
				target_sheet, output_path = _target_for(book, sheet, settings)
				warnings = _add_native_chart(book, target_sheet, frame, spec)
				book.save()
				chart = target_sheet.api.ChartObjects(1).Chart
				self.assertTrue(output_path)
				self.assertEqual(chart.ChartType, -4169)
				self.assertEqual(chart.SeriesCollection().Count, 1)
				self.assertEqual(chart.SeriesCollection(1).Trendlines().Count, 1)
				self.assertEqual(source_path.read_bytes(), original_bytes)
				self.assertEqual(warnings, [])
				self.assertTrue(Path(output_path).is_file())
			finally:
				book.close()
				app.quit()


if __name__ == "__main__":
	unittest.main()
