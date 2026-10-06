from __future__ import annotations

import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import main
from ai import AIServiceError
from models import ChartSpecification


class CLITests(unittest.TestCase):
	def _workbook(self, folder: str, headers: list[str]) -> Path:
		path = Path(folder) / "messwerte.xlsx"
		workbook = Workbook()
		sheet = workbook.active
		sheet.append(headers)
		for row in ((0, 1, 2), (1, 2, 4), (2, 3, 6)):
			sheet.append(row[:len(headers)])
		workbook.save(path)
		return path

	def _spec(self) -> ChartSpecification:
		return ChartSpecification(
			x_column="Week",
			y_column="Active users",
			x_label="Week",
			y_label="Active users",
			trendline="exponential",
			confidence=0.92,
			reason="The values grow exponentially.",
		)

	def test_chart_creation_always_uses_ai(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Week", "Active users"])
			spec = self._spec()
			with (
				patch("main.load_config", return_value={"endpoint": "http://localhost", "api_key": "", "model": "test"}),
				patch("main.analyze_with_ai", return_value=spec) as analyze,
				patch("main.save_workbook_with_chart") as save_chart,
			):
				self.assertEqual(main.main([str(path)]), 0)
			analyze.assert_called_once()
			save_chart.assert_called_once()
			self.assertEqual(save_chart.call_args.args[3], spec)
			self.assertEqual(save_chart.call_args.args[4], "exponential")

	def test_ai_error_prevents_chart_creation(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Week", "Active users"])
			with (
				patch("main.load_config", return_value={"endpoint": "http://localhost", "api_key": "", "model": "test"}),
				patch("main.analyze_with_ai", side_effect=AIServiceError("KI nicht verfügbar")),
				patch("main.save_workbook_with_chart") as save_chart,
			):
				self.assertEqual(main.main([str(path)]), 2)
			save_chart.assert_not_called()

	def test_cli_reanalyzes_after_ai_user_input(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Week", "Active users"])
			provisional = ChartSpecification(
				needs_user_input=True,
				user_input_options=[
					{
						"type": "x_column",
						"question": "Welche Spalte soll X sein?",
						"options": ["Week", "Active users"],
					}
				],
			)
			final = self._spec()
			config = {"endpoint": "http://localhost", "api_key": "", "model": "test"}
			with (
				patch("main.load_config", return_value=config),
				patch("main.analyze_with_ai", return_value=provisional) as analyze,
				patch("main.reanalyze_with_user_input", return_value=final) as reanalyze,
				patch("builtins.input", return_value="Week"),
				patch("main.save_workbook_with_chart") as save_chart,
				redirect_stdout(StringIO()),
			):
				self.assertEqual(main.main([str(path)]), 0)

			analyze.assert_called_once()
			self.assertEqual(
				reanalyze.call_args.args[1:],
				(
					analyze.call_args.args[1],
					config,
					provisional,
					{"x_column": "Week"},
				),
			)
			save_chart.assert_called_once()
			self.assertIs(save_chart.call_args.args[3], final)

	def test_cli_does_not_accept_local_axis_or_fit_overrides(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Week", "Active users"])
			with self.assertRaises(SystemExit):
				main.main([str(path), "--x", "Week"])

	def test_test_ai_mode_does_not_require_excel_file(self) -> None:
		with (
			patch("main.load_config", return_value={"endpoint": "http://localhost", "api_key": "", "model": "test"}),
			patch("main.test_ai_connection") as connection_test,
		):
			self.assertEqual(main.main(["--test-ai"]), 0)
		connection_test.assert_called_once()


if __name__ == "__main__":
	unittest.main()
