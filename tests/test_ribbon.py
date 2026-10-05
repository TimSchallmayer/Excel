from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree

import xlwings as xw

from excel_bridge import _add_ribbon_to_package

ROOT = Path(__file__).resolve().parents[1]


class DashboardAndBridgeTests(unittest.TestCase):
	def test_ribbon_tab_has_expected_callbacks_and_analysis(self) -> None:
		ribbon_path = ROOT / "ribbon" / "customUI.xml"
		bridge_path = ROOT / "excel_vba" / "PLVSBridge.bas"
		root = ElementTree.parse(ribbon_path).getroot()
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		tab = root.find("r:ribbon/r:tabs/r:tab", namespace)
		self.assertIsNotNone(tab)
		self.assertEqual(tab.attrib["label"], "PLVS ULTRA Graphs")
		self.assertNotIn("insertAfterQ", tab.attrib)
		self.assertEqual(root.attrib["onLoad"], "PLVS_RibbonOnLoad")
		groups = tab.findall("r:group", namespace)
		self.assertEqual([group.attrib["label"] for group in groups], ["Diagramm", "KI", "Sonstiges"])
		buttons = tab.findall(".//r:button", namespace)
		self.assertEqual(
			[button.attrib["label"] for button in buttons],
			["Diagramm erstellen", "KI-Einstellungen", "KI testen", "Hilfe"],
		)
		test_ai_button = root.find(".//r:button[@id='PLVSTestAI']", namespace)
		self.assertIsNotNone(test_ai_button)
		self.assertEqual(test_ai_button.attrib["imageMso"], "RefreshAll")
		contextual_tab_set = root.find(
			"r:ribbon/r:contextualTabs/r:tabSet[@idMso='TabSetChartTools']",
			namespace,
		)
		self.assertIsNotNone(contextual_tab_set)
		contextual_tab = contextual_tab_set.find("r:tab[@idMso='TabChartDesign']", namespace)
		self.assertIsNotNone(contextual_tab)
		self.assertEqual(contextual_tab.find("r:group", namespace).attrib["label"], "Analyse")
		analysis_labels = contextual_tab.findall(".//r:labelControl", namespace)
		self.assertEqual(
			[label.attrib["id"] for label in analysis_labels],
			[
				"PLVSAnalysisAxes",
				"PLVSAnalysisRelationship",
				"PLVSAnalysisChartFit",
				"PLVSAnalysisMetrics",
				"PLVSAnalysisStatistics",
				"PLVSAnalysisChange",
				"PLVSAnalysisAI",
			],
		)
		self.assertTrue(all(label.attrib["getLabel"] == "PLVS_RibbonGetAnalysis" for label in analysis_labels))
		self.assertFalse(any("Datenquelle" in button.attrib.get("label", "") for button in buttons))
		bridge = bridge_path.read_text(encoding="utf-8")
		python_functions = {
			"PLVS_RibbonCreateChart": "create_chart",
			"PLVS_RibbonAISettings": "open_ai_settings",
			"PLVS_RibbonTestAI": "test_ai",
			"PLVS_RibbonHelp": "show_help",
		}
		for callback, function in python_functions.items():
			self.assertIn(f"onAction=\"{callback}\"", ribbon_path.read_text(encoding="utf-8"))
			self.assertIn(f"Public Sub {callback}(control As IRibbonControl)", bridge)
			self.assertIn(f"excel_addin.{function}()", bridge)
		self.assertIn("Public Sub PLVS_RibbonOnLoad(ribbon As IRibbonUI)", bridge)
		self.assertIn("Set PLVS_ApplicationEvents.ExcelApp = Application", bridge)
		self.assertNotIn("PLVS_IsAnalysisContext", bridge)
		self.assertIn("Public Sub PLVS_RefreshAnalysis()", bridge)
		self.assertRegex(
			bridge,
			r"(?ms)^Public Sub PLVS_RefreshAnalysis\(\)\r?\n.*?^End Sub\s*$",
		)
		app_events = (ROOT / "excel_vba" / "PLVSAppEvents.cls").read_text(encoding="utf-8")
		self.assertIn("ExcelApp_SheetSelectionChange", app_events)
		self.assertIn("ExcelApp_WorkbookOpen", app_events)
		self.assertIn("PLVS_RefreshAnalysis", app_events)
		self.assertIn("PLVS_WatchCharts", app_events)
		chart_events = (ROOT / "excel_vba" / "PLVSChartEvents.cls").read_text(encoding="utf-8")
		self.assertIn("PLVS_Chart_Activate", chart_events)
		self.assertIn("PLVS_Chart_Deactivate", chart_events)
		self.assertIn('analysisName = analysisName & "_" & CStr(activeChart.Parent.Name)', bridge)
		for label_id in (
			"PLVSAnalysisAxes",
			"PLVSAnalysisRelationship",
			"PLVSAnalysisChartFit",
			"PLVSAnalysisMetrics",
			"PLVSAnalysisStatistics",
			"PLVSAnalysisChange",
			"PLVSAnalysisAI",
		):
			self.assertIn(f'PLVS_Ribbon.InvalidateControl "{label_id}"', bridge)
		self.assertIn('Case "PLVSAnalysisMetrics"', bridge)
		self.assertIn('Case "PLVSAnalysisStatistics"', bridge)
		self.assertIn('Case "PLVSAnalysisChange"', bridge)
		self.assertIn('Case "PLVSAnalysisAI"', bridge)
		self.assertIn("data_sheet.api.ChartObjects(chart_name).Activate()", (ROOT / "excel_addin.py").read_text(encoding="utf-8"))

	def test_installed_xlwings_addin_hides_only_its_ribbon_tab(self) -> None:
		addin_path = Path(xw.__file__).resolve().parent / "addin" / "xlwings.xlam"
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		with ZipFile(addin_path) as addin:
			ribbon = ElementTree.fromstring(addin.read("customUI/customUI.xml"))
		tab = ribbon.find(".//r:tab[@id='xlwingsTab']", namespace)
		self.assertIsNotNone(tab)
		self.assertEqual(tab.attrib.get("visible"), "false")
		xlwings_vba = (Path(xw.__file__).resolve().parent / "xlwings_custom_addin.bas").read_text(encoding="utf-8")
		self.assertIn("Public Function RunPython", xlwings_vba)

	def test_ribbon_is_embedded_with_package_relationship_and_xml_default_type(self) -> None:
		package_namespace = "http://schemas.openxmlformats.org/package/2006"
		content_types_namespace = f"{package_namespace}/content-types"
		relationships_namespace = f"{package_namespace}/relationships"
		with TemporaryDirectory() as folder:
			addin_path = Path(folder) / "test.xlam"
			with ZipFile(addin_path, "w", ZIP_DEFLATED) as addin:
				addin.writestr(
					"[Content_Types].xml",
					f"""<?xml version="1.0" encoding="UTF-8"?>
					<Types xmlns="{content_types_namespace}">
						<Default Extension="xml" ContentType="application/xml"/>
						<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
					</Types>""",
				)
				addin.writestr(
					"_rels/.rels",
					f"""<?xml version="1.0" encoding="UTF-8"?>
					<Relationships xmlns="{relationships_namespace}"/>""",
				)

			_add_ribbon_to_package(addin_path)

			with ZipFile(addin_path) as addin:
				ribbon = ElementTree.fromstring(addin.read("customUI/customUI.xml"))
				relationships = ElementTree.fromstring(addin.read("_rels/.rels"))
				content_types = ElementTree.fromstring(addin.read("[Content_Types].xml"))
			namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
			test_button = ribbon.find(".//r:button[@id='PLVSTestAI']", namespace)
			self.assertIsNotNone(test_button)
			self.assertEqual(test_button.attrib["imageMso"], "RefreshAll")
			analysis_group = ribbon.find(".//r:group[@id='PLVSAnalysisGroup']", namespace)
			self.assertIsNotNone(analysis_group)
			self.assertEqual(
				[label.attrib["id"] for label in analysis_group.findall(".//r:labelControl", namespace)],
				[
					"PLVSAnalysisAxes",
					"PLVSAnalysisRelationship",
					"PLVSAnalysisChartFit",
					"PLVSAnalysisMetrics",
					"PLVSAnalysisStatistics",
					"PLVSAnalysisChange",
					"PLVSAnalysisAI",
				],
			)
			contextual_tab_set = ribbon.find(
				".//r:contextualTabs/r:tabSet[@idMso='TabSetChartTools']",
				namespace,
			)
			self.assertIsNotNone(contextual_tab_set)
			self.assertIsNotNone(contextual_tab_set.find("r:tab[@idMso='TabChartDesign']", namespace))
			self.assertTrue((ROOT / "excel_vba" / "PLVSAppEvents.cls").is_file())
			self.assertTrue((ROOT / "excel_vba" / "PLVSChartEvents.cls").is_file())

		relationship = relationships.find(
			f"{{{relationships_namespace}}}Relationship[@Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility']"
		)
		self.assertIsNotNone(relationship)
		self.assertEqual(relationship.attrib["Target"], "customUI/customUI.xml")
		xml_default = content_types.find(
			f"{{{content_types_namespace}}}Default[@Extension='xml']"
		)
		self.assertIsNotNone(xml_default)
		self.assertEqual(xml_default.attrib["ContentType"], "application/xml")
		self.assertIsNone(
			content_types.find(
				f"{{{content_types_namespace}}}Override[@PartName='/customUI/customUI.xml']"
			)
		)
		self.assertEqual(ribbon.tag, "{http://schemas.microsoft.com/office/2006/01/customui}customUI")

	def test_normal_addin_does_not_create_or_use_a_dashboard_sheet(self) -> None:
		python_source = (ROOT / "excel_addin.py").read_text(encoding="utf-8")
		self.assertIn('PRODUCT_NAME = "PLVS ULTRA Graphs"', python_source)
		self.assertNotIn("def _settings_sheet", python_source)
		self.assertNotIn("def ensure_dashboard", python_source)
		self.assertIn("def _set_text_name", python_source)
		self.assertIn("def _write_analysis", python_source)
		self.assertIn("def show_help()", python_source)
		bridge_source = (ROOT / "excel_bridge.py").read_text(encoding="utf-8")
		self.assertNotIn("ensure_dashboard", bridge_source)

	def test_bridge_calls_match_ribbon_actions(self) -> None:
		bridge = (ROOT / "excel_vba" / "PLVSBridge.bas").read_text(encoding="utf-8")
		for action in (
			"PLVS_ULTRA_CreateChart",
			"PLVS_ULTRA_AISettings",
			"PLVS_ULTRA_TestAI",
			"PLVS_ULTRA_Help",
		):
			self.assertIn(f"Public Sub {action}()", bridge)
			self.assertIn("RunPython", bridge)

	def test_active_workbook_is_obtained_via_xlwings_caller(self) -> None:
		python_source = (ROOT / "excel_addin.py").read_text(encoding="utf-8")
		self.assertIn("xw.Book.caller()", python_source)


if __name__ == "__main__":
	unittest.main()
