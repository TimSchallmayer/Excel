from __future__ import annotations

import unittest
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]


class RibbonDefinitionTests(unittest.TestCase):
	def test_ribbon_is_inserted_after_view_and_has_real_callbacks(self) -> None:
		xml_path = ROOT / "ribbon" / "customUI.xml"
		vba_path = ROOT / "ribbon" / "RibbonPhysikAssistent.bas"
		xml_root = ElementTree.parse(xml_path).getroot()
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		tab = xml_root.find(".//r:tab", namespace)
		self.assertIsNotNone(tab)
		self.assertEqual(tab.attrib["label"], "Physik-Assistent")
		self.assertEqual(tab.attrib["insertAfterMso"], "TabView")
		callback_names = {
			"CreateChart", "AnalyzeTable", "TestAI", "OpenAISettings", "ShowAnalysis", "OpenSettings",
		}
		buttons = xml_root.findall(".//r:button", namespace)
		self.assertEqual(len(buttons), 6)
		module = vba_path.read_text(encoding="utf-8")
		python_bridge = (ROOT / "excel_addin.py").read_text(encoding="utf-8")
		for button in buttons:
			macro_name = button.attrib["onAction"].rsplit(".", 1)[1]
			self.assertIn(macro_name, callback_names)
			self.assertIn(f"Public Sub {macro_name}(", module)
			self.assertIn("excel_addin.", module)
			self.assertIn(f"def {self._python_action(macro_name)}(", python_bridge)

	def _python_action(self, macro_name: str) -> str:
		return {
			"CreateChart": "create_chart",
			"AnalyzeTable": "analyze_table",
			"TestAI": "test_ai",
			"OpenAISettings": "open_ai_settings",
			"ShowAnalysis": "show_analysis",
			"OpenSettings": "open_settings",
		}[macro_name]


if __name__ == "__main__":
	unittest.main()
