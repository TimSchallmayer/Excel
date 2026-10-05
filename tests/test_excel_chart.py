from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import main
from models import ChartSpecification


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

	def _run_and_check_chart(
		self,
		path: Path,
		x_column: str,
		y_column: str,
		fit: str = "none",
		y_error_column: str | None = None,
		origin: bool = False,
		output: Path | None = None,
	) -> Path:
		spec = ChartSpecification(
			x_column=x_column,
			y_column=y_column,
			x_label=x_column,
			y_label=y_column,
			trendline=fit,
			y_error_column=y_error_column,
			origin=origin,
			confidence=0.95,
			reason="AI test decision.",
		)
		args = [str(path)]
		result_path = output or path.with_name(f"{path.stem}_PLVS_ULTRA_Graphs{path.suffix.lower()}")
		if result_path.exists():
			base = result_path
			index = 2
			while result_path.exists():
				result_path = base.with_name(f"{base.stem}_{index}{base.suffix}")
				index += 1
		if output is not None:
			args.extend(["--output", str(output)])
		with (
			patch("main.load_config", return_value={"endpoint": "https://example.invalid", "api_key": "", "model": "test"}),
			patch("main.analyze_with_ai", return_value=spec),
		):
			result = main.main(args)
		self.assertEqual(result, 0)
		self.assertTrue(result_path.is_file())
		workbook = load_workbook(result_path)
		worksheet = workbook["Messwerte"]
		self.assertEqual(worksheet["A1"].value, x_column)
		self.assertEqual(len(worksheet._charts), 1)
		chart = worksheet._charts[0]
		self.assertEqual(chart.__class__.__name__, "ScatterChart")
		self.assertEqual(chart.x_axis.title.tx.rich.p[0].r[0].t, x_column)
		self.assertEqual(chart.y_axis.title.tx.rich.p[0].r[0].t, y_column)
		self.assertIsNone(chart.legend)
		workbook.close()
		return result_path

	def test_general_numeric_table_creates_xy_chart(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Week", "Active users"], [(0, 1), (1, 2), (2, 4)])
			self._run_and_check_chart(path, "Week", "Active users")

	def test_non_linear_trendlines_are_written_as_native_excel_fits(self) -> None:
		cases = (
			("exponential", [(0, 1), (1, 2), (2, 4), (3, 8)], "exp"),
			("logarithmic", [(1, -1), (2, 0), (3, 1)], "log"),
			("power", [(1, 1), (2, 4), (3, 9)], "power"),
		)
		for fit, rows, expected_type in cases:
			with self.subTest(fit=fit), tempfile.TemporaryDirectory() as folder:
				path = self._make_workbook(folder, ["Input", "Output"], rows)
				result = self._run_and_check_chart(path, "Input", "Output", fit=fit)
				workbook = load_workbook(result)
				self.assertEqual(workbook["Messwerte"]._charts[0].series[0].trendline.trendlineType, expected_type)
				workbook.close()

	def test_error_column_is_added_as_native_error_bars(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(
				folder,
				["Week", "Active users", "User uncertainty"],
				[(0, 1, 0.1), (1, 2, 0.2), (2, 4, 0.1)],
			)
			result = self._run_and_check_chart(path, "Week", "Active users", y_error_column="User uncertainty")
			workbook = load_workbook(result)
			chart = workbook["Messwerte"]._charts[0]
			self.assertEqual(chart.series[0].errBars.errDir, "y")
			self.assertEqual(chart.series[0].errBars.errValType, "cust")
			workbook.close()

	def test_ai_selected_zero_origin_is_applied(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Input", "Output"], [(1, 1), (2, 2), (3, 4)])
			result = self._run_and_check_chart(path, "Input", "Output", origin=True)
			workbook = load_workbook(result)
			chart = workbook["Messwerte"]._charts[0]
			self.assertEqual(chart.x_axis.scaling.min, 0)
			workbook.close()

	def test_existing_output_gets_a_unique_name_and_source_is_preserved(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Week", "Active users"], [(0, 1), (1, 2), (2, 4)])
			original = path.read_bytes()
			first = self._run_and_check_chart(path, "Week", "Active users")
			second = self._run_and_check_chart(path, "Week", "Active users")
			self.assertNotEqual(first, second)
			self.assertEqual(path.read_bytes(), original)

	def test_existing_explicit_output_gets_a_unique_name(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Week", "Active users"], [(0, 1), (1, 2), (2, 4)])
			requested_output = Path(folder) / "Ergebnis.xlsx"
			requested_output.write_bytes(b"keep existing")
			result = self._run_and_check_chart(path, "Week", "Active users", output=requested_output)
			self.assertNotEqual(result, requested_output)
			self.assertEqual(requested_output.read_bytes(), b"keep existing")

	def test_empty_measurements_are_rejected(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Input", "Output"], [])
			self.assertEqual(main.main([str(path)]), 2)
			self.assertFalse(path.with_name(f"{path.stem}_PLVS_ULTRA_Graphs{path.suffix.lower()}").exists())

	def test_single_numeric_column_is_rejected(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._make_workbook(folder, ["Input", "Note"], [(0, "a"), (1, "b"), (2, "c")])
			self.assertEqual(main.main([str(path)]), 2)
			self.assertFalse(path.with_name(f"{path.stem}_PLVS_ULTRA_Graphs{path.suffix.lower()}").exists())

	def test_xlsm_is_saved_as_xlsm(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = Path(folder) / "messwerte.xlsm"
			workbook = Workbook()
			worksheet = workbook.active
			worksheet.title = "Data"
			worksheet.append(["Input", "Output"])
			for row in [(0, 1), (1, 2), (2, 4)]:
				worksheet.append(row)
			workbook.save(path)
			with (
				patch("main.load_config", return_value={"endpoint": "https://example.invalid", "api_key": "", "model": "test"}),
				patch("main.analyze_with_ai", return_value=ChartSpecification(
					x_column="Input",
					y_column="Output",
					x_label="Input",
					y_label="Output",
					confidence=0.95,
				)),
			):
				self.assertEqual(main.main([str(path)]), 0)
			output = path.with_name(f"{path.stem}_PLVS_ULTRA_Graphs{path.suffix.lower()}")
			self.assertTrue(output.is_file())
			result = load_workbook(output, keep_vba=True)
			self.assertEqual(len(result["Data"]._charts), 1)
			vba_archive = result.vba_archive
			result.close()
			if vba_archive is not None:
				vba_archive.close()


if __name__ == "__main__":
	unittest.main()
