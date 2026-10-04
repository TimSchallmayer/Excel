from __future__ import annotations

import unittest
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]


class DashboardAndBridgeTests(unittest.TestCase):
	def test_ribbon_tab_is_after_xlwings_with_six_real_callbacks(self) -> None:
		ribbon_path = ROOT / "ribbon" / "customUI.xml"
		bridge_path = ROOT / "excel_vba" / "PLVSBridge.bas"
		root = ElementTree.parse(ribbon_path).getroot()
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		tab = root.find("r:ribbon/r:tabs/r:tab", namespace)
		self.assertIsNotNone(tab)
		self.assertEqual(tab.attrib["label"], "PLVS ULTRA Graphs")
		self.assertEqual(tab.attrib["insertAfterQ"], "xw:xlwingsTab")
		groups = tab.findall("r:group", namespace)
		self.assertEqual([group.attrib["label"] for group in groups], ["Diagramm", "KI", "Analyse"])
		buttons = tab.findall(".//r:button", namespace)
		self.assertEqual(
			[button.attrib["label"] for button in buttons],
			["Analysieren", "Diagramm erstellen", "KI testen", "KI-Einstellungen", "Analyse anzeigen", "Hilfe"],
		)
		bridge = bridge_path.read_text(encoding="utf-8")
		python_functions = {
			"PLVS_RibbonAnalyze": "analyze_table",
			"PLVS_RibbonCreateChart": "create_chart",
			"PLVS_RibbonTestAI": "test_ai",
			"PLVS_RibbonAISettings": "open_ai_settings",
			"PLVS_RibbonShowAnalysis": "show_analysis",
			"PLVS_RibbonHelp": "show_help",
		}
		for callback, function in python_functions.items():
			self.assertIn(f"onAction=\"{callback}\"", ribbon_path.read_text(encoding="utf-8"))
			self.assertIn(f"Public Sub {callback}(control As IRibbonControl)", bridge)
			self.assertIn(f"excel_addin.{function}()", bridge)

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
