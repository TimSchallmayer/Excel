from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import main


class ExcelChartIntegrationTests(unittest.TestCase):
	def _make_workbook(
		self,
		folder: str,
		headers: list[str],
		rows: list[tuple[object, ...]],
	) -> Path:
		path = Path(folder) / "messwerte.xlsx"
		workbook = Workbook()
		worksheet = workbook.active
		worksheet.title = "Messwerte"
		worksheet.append(headers)
		for row in rows:
			worksheet.append(row)
		workbook.save(path)
		return path

	def _run_and_check_chart(self, path: Path, x_column: str, y_column: str) -> Path:
		result = main.main([
			str(path),
			"--x", x_column,
			"--y", y_column,
			"--fit", "none",
		])
		self.assertEqual(result, 0)
		output = path.with_name("messwerte_physik.xlsx")
		self.assertTrue(output.is_file())
		workbook = load_workbook(output)
		worksheet = workbook["Messwerte"]
		self.assertEqual(worksheet["A1"].value, x_column)
		self.assertEqual(len(worksheet._charts), 1)
		chart = worksheet._charts[0]
		self.assertEqual(chart.__class__.__name__, "ScatterChart")
		self.assertEqual(chart.x_axis.title.tx.rich.p[0].r[0].t, x_column)
		self.assertEqual(chart.y_axis.title.tx.rich.p[0].r[0].t, y_column)
		self.assertIsNone(chart.legend)
		workbook.close()
		return output

	def test_time_distance_xy_chart(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], [(0, 0), (1, 2.1), (2, 4.3)])
			self._run_and_check_chart(path, "Zeit [s]", "Strecke [m]")

	def test_voltage_current_xy_chart(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Spannung [V]", "Stromstärke [A]"], [(1, 0.1), (2, 0.2), (3, 0.3)])
			self._run_and_check_chart(path, "Spannung [V]", "Stromstärke [A]")

	def test_force_deformation_xy_chart(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Kraft [N]", "Dehnung [m]"], [(1, 0.01), (2, 0.02), (3, 0.03)])
			self._run_and_check_chart(path, "Kraft [N]", "Dehnung [m]")

	def test_unrecognized_numeric_columns_can_be_selected(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Messkanal A", "Messkanal B"], [(1, 4), (2, 6), (3, 8)])
			self._run_and_check_chart(path, "Messkanal A", "Messkanal B")

	def test_empty_measurements_are_rejected(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], [])
			self.assertEqual(main.main([str(path)]), 2)
			self.assertFalse(path.with_name("messwerte_physik.xlsx").exists())

	def test_single_numeric_column_is_rejected(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Notiz"], [(0, "a"), (1, "b"), (2, "c")])
			self.assertEqual(main.main([str(path)]), 2)
			self.assertFalse(path.with_name("messwerte_physik.xlsx").exists())

	def test_error_column_is_added_as_native_error_bars(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(
				folder,
				["Zeit [s]", "Strecke [m]", "Strecke Fehler [m]"],
				[(0, 0, 0.1), (1, 2, 0.2), (2, 4, 0.1)],
			)
			self.assertEqual(main.main([str(path), "--fit", "none"]), 0)
			workbook = load_workbook(path.with_name("messwerte_physik.xlsx"))
			chart = workbook["Messwerte"]._charts[0]
			self.assertEqual(chart.series[0].errBars.errDir, "y")
			self.assertEqual(chart.series[0].errBars.errValType, "cust")
			workbook.close()

	def test_temperature_axis_is_not_forced_to_zero(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(
				folder,
				["Temperatur [°C]", "Spannung [V]"],
				[(19, 1), (20, 2), (21, 3)],
			)
			self.assertEqual(main.main([str(path), "--fit", "none"]), 0)
			workbook = load_workbook(path.with_name("messwerte_physik.xlsx"))
			chart = workbook["Messwerte"]._charts[0]
			self.assertGreater(chart.x_axis.scaling.min, 0)
			workbook.close()

	def test_existing_output_gets_a_unique_name_and_source_is_preserved(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], [(0, 0), (1, 2), (2, 4)])
			original = path.read_bytes()
			first = self._run_and_check_chart(path, "Zeit [s]", "Strecke [m]")
			self.assertEqual(main.main([str(path), "--x", "Zeit [s]", "--y", "Strecke [m]", "--fit", "none"]), 0)
			self.assertEqual(first.read_bytes()[:2], b"PK")
			self.assertTrue(path.with_name("messwerte_physik_2.xlsx").is_file())
			self.assertEqual(path.read_bytes(), original)

	def test_existing_explicit_output_gets_a_unique_name(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], [(0, 0), (1, 2), (2, 4)])
			requested_output = Path(folder) / "Ergebnis.xlsx"
			requested_output.write_bytes(b"keep existing")
			self.assertEqual(main.main([
				str(path), "--x", "Zeit [s]", "--y", "Strecke [m]",
				"--fit", "none", "--output", str(requested_output),
			]), 0)
			self.assertEqual(requested_output.read_bytes(), b"keep existing")
			self.assertTrue(Path(folder, "Ergebnis_2.xlsx").is_file())

	def test_native_linear_trendline_is_written(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], [(0, 0), (1, 2), (2, 4)])
			self.assertEqual(main.main([str(path), "--x", "Zeit [s]", "--y", "Strecke [m]", "--fit", "linear"]), 0)
			workbook = load_workbook(path.with_name("messwerte_physik.xlsx"))
			chart = workbook["Messwerte"]._charts[0]
			self.assertEqual(chart.series[0].trendline.trendlineType, "linear")
			workbook.close()

	def test_quadratic_and_cubic_native_trendlines_are_written(self) -> None:
		for fit, order in (("quadratic", 2), ("cubic", 3)):
			with self.subTest(fit=fit), tempfile.TemporaryDirectory() as folder:
				rows = [(index, index ** order) for index in range(order + 2)]
				path = self._make_workbook(folder, ["Zeit [s]", "Strecke [m]"], rows)
				self.assertEqual(main.main([
					str(path), "--x", "Zeit [s]", "--y", "Strecke [m]", "--fit", fit,
				]), 0)
				workbook = load_workbook(path.with_name("messwerte_physik.xlsx"))
				chart = workbook["Messwerte"]._charts[0]
				self.assertEqual(chart.series[0].trendline.trendlineType, "poly")
				self.assertEqual(chart.series[0].trendline.order, order)
				workbook.close()

	def test_xlsm_is_saved_as_xlsm(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = Path(folder) / "messwerte.xlsm"
			workbook = Workbook()
			worksheet = workbook.active
			worksheet.title = "Messwerte"
			worksheet.append(["Zeit [s]", "Strecke [m]"])
			for row in [(0, 0), (1, 2), (2, 4)]:
				worksheet.append(row)
			workbook.save(path)
			self.assertEqual(main.main([str(path), "--fit", "none"]), 0)
			output = path.with_name("messwerte_physik.xlsm")
			self.assertTrue(output.is_file())
			result = load_workbook(output, keep_vba=True)
			self.assertEqual(len(result["Messwerte"]._charts), 1)
			result.close()
			if result.vba_archive is not None:
				result.vba_archive.close()


if __name__ == "__main__":
	unittest.main()