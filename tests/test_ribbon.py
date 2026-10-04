from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class DashboardAndBridgeTests(unittest.TestCase):
	def test_dashboard_uses_workbook_sheet_ui(self) -> None:
		python_source = (ROOT / "excel_addin.py").read_text(encoding="utf-8")
		self.assertIn('PRODUCT_NAME = "PLVS ULTRA Graphs"', python_source)
		for label in ("ANALYSIEREN", "DIAGRAMM ERSTELLEN", "KI-EINSTELLUNGEN", "ANALYSE", "HILFE"):
			self.assertIn(f'"{label}"', python_source)
		self.assertIn("def show_help()", python_source)

	def test_bridge_calls_match_dashboard_actions(self) -> None:
		bridge = (ROOT / "excel_vba" / "PLVSBridge.bas").read_text(encoding="utf-8")
		for action in (
			"PLVS_ULTRA_Analyze",
			"PLVS_ULTRA_CreateChart",
			"PLVS_ULTRA_AISettings",
			"PLVS_ULTRA_ShowAnalysis",
			"PLVS_ULTRA_Help",
		):
			self.assertIn(f"Public Sub {action}()", bridge)
			self.assertIn("RunPython", bridge)

	def test_active_workbook_is_obtained_via_xlwings_caller(self) -> None:
		python_source = (ROOT / "excel_addin.py").read_text(encoding="utf-8")
		self.assertIn("xw.Book.caller()", python_source)


if __name__ == "__main__":
	unittest.main()
