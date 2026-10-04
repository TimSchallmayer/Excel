from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import main
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

	def test_unambiguous_local_choice_skips_ai(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Zeit [s]", "Spannung [V]"])
			with (
				patch("main.analyze_with_ai", side_effect=AssertionError("AI should not be called")),
				patch("main.create_plot") as create_plot,
			):
				self.assertEqual(main.main([str(path), "--ai"]), 0)
			create_plot.assert_called_once()
			self.assertEqual(create_plot.call_args.args[1].x_column, "Zeit [s]")

	def test_ambiguous_choice_uses_ai_when_enabled(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = self._workbook(folder, ["Zeit [s]", "Spannung [V]", "Stromstärke [A]"])
			spec = ChartSpecification(
				x_column="Zeit [s]",
				y_column="Spannung [V]",
				x_label="Zeit [s]",
				y_label="Spannung [V]",
				confidence=0.9,
				reason="Testdiagnose.",
			)
			with (
				patch("main.load_config", return_value={"endpoint": "http://localhost", "api_key": "", "model": "test"}),
				patch("main.analyze_with_ai", return_value=spec) as analyze,
				patch("main.create_plot"),
			):
				self.assertEqual(main.main([str(path), "--ai"]), 0)
			analyze.assert_called_once()

	def test_test_ai_mode_does_not_require_excel_file(self) -> None:
		with (
			patch("main.load_config", return_value={"endpoint": "http://localhost", "api_key": "", "model": "test"}),
			patch("main.test_ai_connection") as connection_test,
		):
			self.assertEqual(main.main(["--test-ai"]), 0)
		connection_test.assert_called_once()


if __name__ == "__main__":
	unittest.main()