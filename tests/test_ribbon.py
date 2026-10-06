from __future__ import annotations

import ast
import os
import re
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree

import xlwings as xw
import pandas as pd

from excel_addin import _add_native_chart, _read_chart_state, _save_chart_state, _source_for_chart
from excel_bridge import _add_ribbon_to_package
from models import ChartSpecification

ROOT = Path(__file__).resolve().parents[1]


def _production_addin_path() -> Path:
	return Path(
		os.environ.get(
			"PLVS_TEST_ADDIN_PATH",
			str(ROOT / "dist" / "PLVS ULTRA Graphs.xlam"),
		)
	).resolve()


class DashboardAndBridgeTests(unittest.TestCase):
	def test_ai_choices_remain_available_as_dynamic_ribbon_options(self) -> None:
		ribbon_path = ROOT / "ribbon" / "customUI.xml"
		bridge_path = ROOT / "excel_vba" / "PLVSBridge.bas"
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		ribbon = ElementTree.parse(ribbon_path).getroot()
		choice_control = ribbon.find(".//r:dropDown[@id='PLVSPendingChoice']", namespace)
		self.assertIsNotNone(choice_control)
		self.assertEqual(choice_control.attrib["getVisible"], "PLVS_RibbonGetPendingVisible")
		self.assertEqual(choice_control.attrib["getItemLabel"], "PLVS_RibbonGetPendingItemLabel")
		self.assertEqual(choice_control.attrib["onAction"], "PLVS_RibbonChoosePendingChoice")
		bridge = bridge_path.read_text(encoding="utf-8")
		for callback in (
			"PLVS_RibbonGetPendingVisible",
			"PLVS_RibbonGetPendingItemCount",
			"PLVS_RibbonGetPendingItemLabel",
			"PLVS_RibbonChoosePendingChoice",
		):
			self.assertIn(f"Public Sub {callback}(", bridge)
		for module in ("main.py", "excel_addin.py"):
			source = (ROOT / module).read_text(encoding="utf-8")
			self.assertNotIn("Nummer eingeben", source)
			self.assertNotIn("Optionsnummer", source)

	@unittest.skipUnless(
		os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1",
		"Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed.",
	)
	def test_live_excel_opens_data_source_window_and_uses_clicked_source(self) -> None:
		addin_path = _production_addin_path()
		self.assertTrue(addin_path.is_file(), f"Add-in not found: {addin_path}")
		app = xw.App(visible=True, add_book=True)
		book = app.books.active
		addin = None
		was_installed = False
		try:
			addin = app.api.AddIns.Add(str(addin_path.resolve()), False)
			was_installed = bool(addin.Installed)
			addin.Installed = True
			sheet = book.sheets[0]
			sheet.name = "Daten"
			sheet.range("A1").value = [
				["Zeit", "Wert"],
				[0, 4],
				[1, 7],
				[2, 9],
			]
			sheet.range("D1").value = [
				["Monat", "Umsatz"],
				["Jan", 12],
				["Feb", 15],
				["Mär", 20],
			]
			sheet.range("Z1").select()
			powershell = f"""
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$desktop = [System.Windows.Automation.AutomationElement]::RootElement
$deadline = [DateTime]::UtcNow.AddSeconds(8)
$dialog = $null
while ($null -eq $dialog -and [DateTime]::UtcNow -lt $deadline) {{
    $elements = $desktop.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition
    )
    foreach ($element in $elements) {{
        if ($element.Current.Name -eq 'PLVS ULTRA Graphs - Datenquelle auswählen') {{
            $dialog = $element
            break
        }}
    }}
    if ($null -eq $dialog) {{ Start-Sleep -Milliseconds 100 }}
}}
$dialog = $desktop.FindFirst(
    [System.Windows.Automation.TreeScope]::Children,
    [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        'PLVS ULTRA Graphs - Datenquelle auswählen'
    )
)
if ($null -eq $dialog) {{ throw 'Data-source window did not open in Excel.' }}
$target = $dialog.FindFirst(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        'Daten: A1:B4'
    )
)
if ($null -eq $target) {{ throw 'Expected workbook data source is absent from the dialog.' }}
$target.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
Write-Output 'Visible data-source window opened and first source was clicked.'
"""
			automation = subprocess.Popen(
				["powershell.exe", "-STA", "-NoProfile", "-Command", powershell],
				stdout=subprocess.PIPE,
				stderr=subprocess.PIPE,
				text=True,
			)
			with (
				patch("excel_addin._active_book", return_value=book),
				patch("excel_addin._safe_action", side_effect=lambda action: action(book)),
				patch("excel_addin._get_ai_config", return_value={"endpoint": "test"}),
				patch("excel_addin.analyze_with_ai", return_value=ChartSpecification(
					x_column="Zeit",
					y_column="Wert",
					x_label="Zeit",
					y_label="Wert",
					chart_type="scatter",
				)) as analyze,
				patch("excel_addin._message"),
			):
				from excel_addin import create_chart
				create_chart()
			stdout, stderr = automation.communicate(timeout=30)
			self.assertEqual(automation.returncode, 0, stdout + stderr)
			self.assertIn("Visible data-source window opened", stdout)
			analyze.assert_called_once()
			self.assertEqual(analyze.call_args.args[0]["Wert"].tolist(), [4, 7, 9])
			self.assertEqual(sheet.api.ChartObjects().Count, 1)
			self.assertEqual(
				tuple(sheet.api.ChartObjects(1).Chart.SeriesCollection(1).Values),
				(4.0, 7.0, 9.0),
			)

			sheet.range("D1:E4").select()
			with patch("excel_addin._show_data_source_dialog", side_effect=AssertionError("dialog should not open")):
				selected = _source_for_chart(book)
			self.assertEqual(selected[1].address.replace("$", ""), "D1:E4")
		finally:
			if addin is not None:
				addin.Installed = was_installed
			book.close()
			app.quit()

	@unittest.skipUnless(
		os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1",
		"Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed.",
	)
	def test_fresh_excel_keeps_text_format_controls_available_during_edits(self) -> None:
		addin_path = _production_addin_path()
		self.assertTrue(addin_path.is_file(), f"Add-in not found: {addin_path}")
		app = xw.App(visible=True, add_book=True)
		book = app.books.active
		addin = None
		was_installed = False
		try:
			addin = app.api.AddIns.Add(str(addin_path), False)
			was_installed = bool(addin.Installed)
			addin.Installed = True
			sheet = book.sheets[0]
			rows = [["Input", "Output"], [1, 2], [2, 4], [3, 6]]
			sheet.range("A1").value = rows
			data_range = sheet.range("A1:B4")
			frame = pd.DataFrame(rows[1:], columns=rows[0])
			chart_names = []
			for _ in range(2):
				spec = ChartSpecification(
					x_column="Input",
					y_column="Output",
					x_label="Input",
					y_label="Output",
					chart_type="scatter",
				)
				chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
				chart_names.append(chart_name)
				_save_chart_state(book, chart_name, sheet.name, "A1:B4", spec)
			app.api.Run(f"'{addin_path}'!PLVS_WatchCharts")
			chart_object = sheet.api.ChartObjects(chart_names[0])
			chart_object.Activate()
			chart = chart_object.Chart
			other_chart = sheet.api.ChartObjects(chart_names[1]).Chart
			initial_bold = bool(chart.ChartTitle.Font.Bold)
			initial_italic = bool(chart.ChartTitle.Font.Italic)
			other_font = (
				str(other_chart.ChartTitle.Font.Name),
				float(other_chart.ChartTitle.Font.Size),
				int(other_chart.ChartTitle.Font.Color),
				bool(other_chart.ChartTitle.Font.Bold),
				bool(other_chart.ChartTitle.Font.Italic),
			)
			powershell = f"""
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]{app.api.Hwnd})
function FindByName($parent, $name) {{
    $condition = [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $name
    )
    return $parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}}
$tab = FindByName $root 'PLVS ULTRA Graphs'
if ($null -eq $tab) {{ throw 'PLVS ULTRA Graphs tab is missing.' }}
$textTab = FindByName $root 'Text formatieren'
if ($null -ne $textTab) {{ throw 'An unexpected separate text-formatting tab is present.' }}
$tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
function AssertTextControlsOnMainTab($tab, $root) {{
    if (-not $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Current.IsSelected) {{
        throw 'PLVS ULTRA Graphs tab stopped being selected during formatting.'
    }}
    $textGroup = FindByName $root 'Text'
    if ($null -eq $textGroup) {{ throw 'Text group is absent from the main PLVS tab.' }}
}}
AssertTextControlsOnMainTab $tab $root
Write-Output 'Text formatting controls are visible in the main PLVS Ribbon tab.'
"""
			macro = f"'{addin_path}'!PLVS_RunChartPythonCall"
			format_actions = (
				("excel_addin.toggle_text_italic", None),
				("excel_addin.set_text_font_name", 1),
				("excel_addin.set_text_font_size", 5),
				("excel_addin.toggle_text_bold", None),
				("excel_addin.set_text_color", 255),
			)
			for function_name, argument in format_actions:
				if argument is None:
					app.api.Run(macro, function_name)
				else:
					app.api.Run(macro, function_name, argument)
				result = subprocess.run(
					[
						"powershell.exe",
						"-STA",
						"-NoProfile",
						"-Command",
						powershell,
					],
					capture_output=True,
					check=False,
					text=True,
					timeout=30,
				)
				self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
				self.assertIn("Text formatting controls are visible in the main PLVS Ribbon tab", result.stdout)
			actual_format = (
				str(chart.ChartTitle.Font.Name),
				float(chart.ChartTitle.Font.Size),
				int(chart.ChartTitle.Font.Color),
				bool(chart.ChartTitle.Font.Bold),
				bool(chart.ChartTitle.Font.Italic),
			)
			result = subprocess.run(
				[
					"powershell.exe",
					"-STA",
					"-NoProfile",
					"-Command",
					powershell,
				],
				capture_output=True,
				check=False,
				text=True,
				timeout=30,
			)
			self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
			self.assertEqual(
				actual_format,
				("Arial", 14.0, 255, not initial_bold, not initial_italic),
				result.stdout,
			)
			self.assertEqual(
				(
					str(other_chart.ChartTitle.Font.Name),
					float(other_chart.ChartTitle.Font.Size),
					int(other_chart.ChartTitle.Font.Color),
					bool(other_chart.ChartTitle.Font.Bold),
					bool(other_chart.ChartTitle.Font.Italic),
				),
				other_font,
			)
			self.assertEqual(
				_read_chart_state(book, chart_names[0])["settings"]["text_formats"]["title"],
				{
					"name": "Arial",
					"size": 14.0,
					"bold": not initial_bold,
					"italic": not initial_italic,
					"color": 255,
				},
			)
		finally:
			if addin is not None:
				addin.Installed = was_installed
			book.close()
			app.quit()

	@unittest.skipUnless(
		os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1",
		"Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed.",
	)
	def test_fresh_excel_process_loads_production_ribbon_tab(self) -> None:
		addin_path = _production_addin_path()
		self.assertTrue(addin_path.is_file(), f"Add-in not found: {addin_path}")
		target_path = str(addin_path.resolve())
		app = xw.App(visible=True, add_book=True)
		addin = None
		was_installed = False
		try:
			matching_addins = [
				app.api.AddIns.Item(index)
				for index in range(1, app.api.AddIns.Count + 1)
				if os.path.normcase(str(app.api.AddIns.Item(index).FullName))
				== os.path.normcase(target_path)
			]
			self.assertLessEqual(
				len(matching_addins),
				1,
				"Fresh Excel registered the same PLVS add-in path more than once.",
			)
			addin = matching_addins[0] if matching_addins else app.api.AddIns.Add(target_path, False)
			was_installed = bool(addin.Installed)
			if was_installed and not bool(addin.IsOpen):
				addin.Installed = False
			addin.Installed = True
			self.assertTrue(addin.Installed)
			self.assertTrue(addin.IsOpen)
			self.assertEqual(
				os.path.normcase(str(addin.FullName)),
				os.path.normcase(target_path),
			)
			powershell = f"""
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]{app.api.Hwnd})
$condition = [System.Windows.Automation.PropertyCondition]::new(
    [System.Windows.Automation.AutomationElement]::NameProperty,
    'PLVS ULTRA Graphs'
)
$tab = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
if ($null -eq $tab) {{ throw 'PLVS ULTRA Graphs tab is absent from the live Excel Ribbon.' }}
$pointColorCondition = [System.Windows.Automation.PropertyCondition]::new(
    [System.Windows.Automation.AutomationElement]::NameProperty,
    'Punktfarbe'
)
$pointColor = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $pointColorCondition)
if ($null -ne $pointColor) {{ throw 'The removed Punktfarbe control is still present in the live Ribbon.' }}
$tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
Start-Sleep -Milliseconds 500
$textGroupCondition = [System.Windows.Automation.PropertyCondition]::new(
    [System.Windows.Automation.AutomationElement]::NameProperty,
    'Text'
)
$textGroup = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $textGroupCondition)
if ($null -eq $textGroup) {{ throw 'Text formatting group is absent from the main PLVS tab.' }}
Write-Output "Visible Ribbon tab: $($tab.Current.Name)"
"""
			result = subprocess.run(
				[
					"powershell.exe",
					"-STA",
					"-NoProfile",
					"-Command",
					powershell,
				],
				capture_output=True,
				check=False,
				text=True,
				timeout=30,
			)
			self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
			self.assertIn("Visible Ribbon tab: PLVS ULTRA Graphs", result.stdout)
			self.assertNotIn("Text formatieren", result.stdout)
		finally:
			if addin is not None:
				addin.Installed = was_installed
			app.quit()

	@unittest.skipUnless(
		os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1",
		"Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed.",
	)
	def test_vba_runpython_arguments_are_valid_python_string_literals(self) -> None:
		addin_path = _production_addin_path()
		self.assertTrue(addin_path.is_file(), f"Add-in not found: {addin_path}")
		app = xw.App(visible=False, add_book=True)
		addin = None
		was_installed = False
		try:
			addin = app.api.AddIns.Add(str(addin_path.resolve()), False)
			was_installed = bool(addin.Installed)
			addin.Installed = True
			values = (
				1,
				r"C:\Users\station\Documents\Code\Physik_helfer\Mappe1.xlsx",
				"Tabelle1",
				"PLVS_ULTRA_Chart_3",
				"Messwerte A",
				'Messwerte "Test"',
				r"Daten\2026",
				"O'Brien",
				"Zeile eins\nZeile zwei\täöß 😀",
				True,
				2.5,
			)
			call = app.api.Run(
				f"'{addin_path.resolve()}'!PLVS_BuildPythonCall",
				"excel_addin.change_trendline",
				values,
			)
			parsed = ast.parse(str(call))
			call_expression = parsed.body[1].value
			self.assertEqual(call_expression.func.attr, "change_trendline")
			self.assertEqual(
				[
					eval(compile(ast.Expression(argument), "<RunPython argument>", "eval"), {"chr": chr})
					for argument in call_expression.args
				],
				list(values),
				str(call),
			)
			self.assertRegex(str(call), r"change_trendline\(1, ")
			self.assertNotRegex(str(call), r'change_trendline\("1"')
		finally:
			if addin is not None:
				addin.Installed = was_installed
			app.quit()

	@unittest.skipUnless(
		os.name == "nt" and os.environ.get("PHYSIK_EXCEL_COM_TESTS") == "1",
		"Set PHYSIK_EXCEL_COM_TESTS=1 on Windows with Excel installed.",
	)
	def test_fresh_excel_runs_chart_edits_through_central_runpython_helper(self) -> None:
		addin_path = _production_addin_path()
		self.assertTrue(addin_path.is_file(), f"Add-in not found: {addin_path}")
		app = xw.App(visible=False, add_book=True)
		book = app.books.active
		addin = None
		was_installed = False
		temporary = TemporaryDirectory()
		try:
			addin = app.api.AddIns.Add(str(addin_path.resolve()), False)
			was_installed = bool(addin.Installed)
			addin.Installed = True
			sheet = book.sheets[0]
			sheet.name = "Tabelle1"
			sheet.range("A1").value = [
				["Input", "Output"],
				[1, 2],
				[2, 4],
				[3, 6],
				[4, 8],
				[5, 10],
			]
			workbook_path = Path(temporary.name) / "Mappe1.xlsx"
			book.api.SaveAs(str(workbook_path), FileFormat=51)
			data_range = sheet.range("A1:B6")
			frame = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
			spec = ChartSpecification(
				x_column="Input",
				y_column="Output",
				x_label="Input",
				y_label="Output",
				chart_type="scatter",
				connect_points=True,
			)
			chart_name = _add_native_chart(book, sheet, data_range, frame, spec)
			chart_object = sheet.api.ChartObjects(chart_name)
			chart_object.Name = "PLVS_ULTRA_Chart_3"
			chart_name = str(chart_object.Name)
			_save_chart_state(book, chart_name, sheet.name, "A1:B6", spec)
			chart = chart_object.Chart
			second_spec = ChartSpecification(
				x_column="Input",
				y_column="Output",
				x_label="Input",
				y_label="Output",
				chart_type="line_scatter",
			)
			second_chart_name = _add_native_chart(book, sheet, data_range, frame, second_spec)
			_save_chart_state(book, second_chart_name, sheet.name, "A1:B6", second_spec)
			second_chart = sheet.api.ChartObjects(second_chart_name).Chart
			self.assertEqual(int(second_chart.ChartType), 65)
			chart_object.Activate()
			macro = f"'{addin_path.resolve()}'!PLVS_RunChartPythonCall"
			call = app.api.Run(
				f"'{addin_path.resolve()}'!PLVS_BuildChartPythonCall",
				"excel_addin.change_chart_type",
				3,
			)
			parsed = ast.parse(str(call))
			self.assertEqual(
				[ast.literal_eval(argument) for argument in parsed.body[1].value.args],
				[3, str(workbook_path), sheet.name, chart_name],
				str(call),
			)

			app.api.Run(macro, "excel_addin.change_chart_type", 3)
			self.assertEqual(int(chart.ChartType), 65, str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.change_trendline", 1)
			self.assertEqual(int(chart.SeriesCollection(1).Trendlines(1).Type), -4132, str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.set_legend", True)
			self.assertTrue(bool(chart.HasLegend), str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.set_legend", False)
			self.assertFalse(bool(chart.HasLegend), str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.set_series_color", 255)
			self.assertEqual(int(chart.SeriesCollection(1).Format.Line.ForeColor.RGB), 255, str(app.api.StatusBar))
			self.assertEqual(int(chart.SeriesCollection(1).MarkerForegroundColor), 255, str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.set_connections", True)
			self.assertTrue(bool(chart.SeriesCollection(1).Format.Line.Visible), str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.change_trendline", 0)
			app.api.Run(macro, "excel_addin.change_chart_type", 5)
			self.assertEqual(int(chart.ChartType), 51, str(app.api.StatusBar))
			self.assertFalse(_read_chart_state(book, chart_name)["spec"]["connect_points"])
			app.api.Run(macro, "excel_addin.change_chart_type", 1)
			self.assertEqual(int(chart.ChartType), 74, str(app.api.StatusBar))
			self.assertTrue(bool(chart.SeriesCollection(1).Format.Line.Visible), str(app.api.StatusBar))
			self.assertTrue(_read_chart_state(book, chart_name)["spec"]["connect_points"])
			app.api.Run(macro, "excel_addin.set_connections", False)
			self.assertFalse(bool(chart.SeriesCollection(1).Format.Line.Visible), str(app.api.StatusBar))
			app.api.Run(macro, "excel_addin.change_chart_type", 5)
			app.api.Run(macro, "excel_addin.change_chart_type", 0)
			self.assertEqual(int(chart.ChartType), -4169, str(app.api.StatusBar))
			self.assertFalse(bool(chart.SeriesCollection(1).Format.Line.Visible), str(app.api.StatusBar))
			self.assertFalse(_read_chart_state(book, chart_name)["spec"]["connect_points"])
			app.api.Run(macro, "excel_addin.set_text_font_name", 11)
			self.assertEqual(str(chart.ChartTitle.Font.Name), "Segoe UI", str(app.api.StatusBar))
			refresh_macro = f"'{addin_path.resolve()}'!PLVS_RefreshAnalysis"
			selection_macro = f"'{addin_path.resolve()}'!PLVS_RibbonChartTypeIndex"
			for selection, expected_type, expected_markers, expected_connection in (
				(0, -4169, True, False),
				(1, 74, True, True),
				(3, 65, True, True),
				(2, 4, False, True),
				(0, -4169, True, False),
				(1, 74, True, True),
				(0, -4169, True, False),
			):
				app.api.Run(macro, "excel_addin.change_chart_type", selection)
				app.api.Run(refresh_macro)
				series = chart.SeriesCollection(1)
				self.assertEqual(int(chart.ChartType), expected_type, str(app.api.StatusBar))
				self.assertEqual(int(series.MarkerStyle) != -4142, expected_markers)
				self.assertEqual(bool(series.Format.Line.Visible), expected_connection)
				self.assertEqual(
					int(app.api.Run(selection_macro, expected_type)),
					selection,
				)
				self.assertEqual(
					int(second_chart.ChartType),
					65,
					"Changing the selected chart modified the other chart.",
				)
		finally:
			if addin is not None:
				addin.Installed = was_installed
			book.close()
			app.quit()
			temporary.cleanup()

	def test_dist_addin_contains_current_chart_edit_ribbon(self) -> None:
		addin_path = ROOT / "dist" / "PLVS ULTRA Graphs.xlam"
		if not addin_path.is_file():
			self.skipTest("Build the PLVS add-in to validate its packaged Ribbon XML.")
		namespace = {"r": "http://schemas.microsoft.com/office/2006/01/customui"}
		package_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
		with ZipFile(addin_path) as package:
			ribbon = ElementTree.fromstring(package.read("customUI/customUI.xml"))
			relationships = ElementTree.fromstring(package.read("_rels/.rels"))
		self.assertEqual(
			ribbon.tag,
			"{http://schemas.microsoft.com/office/2006/01/customui}customUI",
		)
		ui_relationships = relationships.findall(
			f"{{{package_namespace}}}Relationship[@Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility']"
		)
		self.assertEqual(len(ui_relationships), 1)
		self.assertEqual(ui_relationships[0].attrib["Target"], "customUI/customUI.xml")
		tabs = ribbon.findall("r:ribbon/r:tabs/r:tab", namespace)
		self.assertEqual([tab.attrib["id"] for tab in tabs], ["PLVSULTRAGraphsTab"])
		all_ids = [element.attrib["id"] for element in ribbon.findall(".//*[@id]", namespace)]
		self.assertEqual(len(all_ids), len(set(all_ids)))
		for group_id in (
			"PLVSAxesGroup",
			"PLVSChartOptionsGroup",
			"PLVSChartAppearanceGroup",
			"PLVSTrendlineGroup",
			"PLVSTextGroup",
		):
			self.assertIsNotNone(ribbon.find(f".//r:group[@id='{group_id}']", namespace))
		self.assertIsNone(ribbon.find(".//r:tab[@id='PLVSTextTab']", namespace))
		self.assertIsNotNone(ribbon.find(
			".//r:tab[@id='PLVSULTRAGraphsTab']/r:group[@id='PLVSTextGroup']",
			namespace,
		))
		swap = ribbon.find(".//r:button[@id='PLVSSwapAxes']", namespace)
		self.assertIsNotNone(swap)
		self.assertEqual(swap.attrib["imageMso"], "ChartSwitchRowColumn")
		self.assertEqual(swap.attrib["onAction"], "PLVS_RibbonSwapAxes")
		self.assertEqual(len(ribbon.findall(".//r:colorPicker", namespace)), 0)
		self.assertIsNone(ribbon.find(".//*[@id='PLVSPointColor']", namespace))
		self.assertFalse(any("Punktfarbe" in control.attrib.get("label", "") for control in ribbon.findall(".//*[@id]", namespace)))
		color_controls = [
			ribbon.find(f".//r:dropDown[@id='{control_id}']", namespace)
			for control_id in (
				"PLVSSeriesColor",
				"PLVSLineColor",
				"PLVSFontColor",
			)
		]
		self.assertTrue(all(control is not None for control in color_controls))
		for control in color_controls:
			self.assertEqual(
				[item.attrib["label"] for item in control.findall("r:item", namespace)],
				[
					"Blau",
					"Orange",
					"Grau",
					"Gold",
					"Hellblau",
					"Grün",
					"Dunkelblau",
					"Rot",
					"Violett",
					"Schwarz",
				],
			)
			self.assertEqual(control.attrib["getSelectedItemIndex"], "PLVS_RibbonGetSelectedItemIndex")
		trendline_dropdown = ribbon.find(".//r:dropDown[@id='PLVSChangeTrendline']", namespace)
		self.assertIsNotNone(trendline_dropdown)
		self.assertEqual(
			trendline_dropdown.attrib["getItemCount"],
			"PLVS_RibbonGetTrendlineCount",
		)
		self.assertEqual(
			trendline_dropdown.attrib["onAction"],
			"PLVS_RibbonChangeTrendline",
		)
		self.assertEqual(
			[control.attrib["id"] for control in (
				ribbon.find(".//r:dropDown[@id='PLVSPolynomialDegree']", namespace),
				ribbon.find(".//r:dropDown[@id='PLVSMovingAveragePeriod']", namespace),
			)],
			["PLVSPolynomialDegree", "PLVSMovingAveragePeriod"],
		)
		self.assertFalse(any("#" in control.attrib.get("label", "") for control in ribbon.findall(".//*[@id]", namespace)))
		validated_image_ids = {
			"AdvancedFileProperties",
			"Bold",
			"ChartAxisTitles",
			"ChartChangeType",
			"ChartDataLabel",
			"ChartGridlines",
			"ChartInsert",
			"ChartLegend",
			"ChartLines",
			"ChartSwitchRowColumn",
			"ChartTitle",
			"ChartTrendline",
			"EditText",
			"Font",
			"FontColorPicker",
			"FontSize",
			"Help",
			"Italic",
			"LineWeightGallery",
			"RefreshAll",
			"ShapeFillColorPicker",
			"ShapeOutlineColorPicker",
		}
		self.assertTrue(
			{element.attrib["imageMso"] for element in ribbon.findall(".//*[@imageMso]", namespace)}
			<= validated_image_ids
		)
		bridge = (ROOT / "excel_vba" / "PLVSBridge.bas").read_text(encoding="utf-8")
		callbacks = {
			value
			for element in ribbon.iter()
			for attribute, value in element.attrib.items()
			if attribute == "onLoad" or attribute == "onAction" or attribute.startswith("get")
		}
		public_callbacks = set(re.findall(r"(?im)^Public\s+Sub\s+(\w+)", bridge))
		self.assertFalse(callbacks - public_callbacks, sorted(callbacks - public_callbacks))

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
		self.assertEqual(
			[group.attrib["label"] for group in groups],
			["Diagramm", "Achsen", "Diagramm", "Darstellung", "Trendlinie", "Text", "KI", "Sonstiges"],
		)
		self.assertIsNone(root.find("r:ribbon/r:tabs/r:tab[@id='PLVSTextTab']", namespace))
		text_group = groups[5]
		self.assertEqual(text_group.attrib["id"], "PLVSTextGroup")
		controls = tab.findall(".//*[@id]", namespace)
		controls_by_id = {control.attrib["id"]: control for control in controls}
		self.assertEqual(controls_by_id["PLVSSwapAxes"].attrib["label"], "X/Y tauschen")
		self.assertEqual(controls_by_id["PLVSSwapAxes"].attrib["imageMso"], "ChartSwitchRowColumn")
		self.assertTrue(controls_by_id["PLVSChangeChartType"].tag.endswith("dropDown"))
		self.assertTrue(controls_by_id["PLVSLineWidth"].tag.endswith("dropDown"))
		self.assertTrue(controls_by_id["PLVSPointSize"].tag.endswith("dropDown"))
		self.assertTrue(controls_by_id["PLVSToggleTitle"].tag.endswith("toggleButton"))
		self.assertEqual(
			controls_by_id["PLVSSeriesColor"].attrib["onAction"],
			"PLVS_RibbonSeriesColor",
		)
		self.assertEqual(
			controls_by_id["PLVSFontColor"].attrib["onAction"],
			"PLVS_RibbonFontColor",
		)
		self.assertEqual(
			[group.attrib["id"] for group in groups],
			[
				"PLVSDiagramGroup",
				"PLVSAxesGroup",
				"PLVSChartOptionsGroup",
				"PLVSChartAppearanceGroup",
				"PLVSTrendlineGroup",
				"PLVSTextGroup",
				"PLVSAIGroup",
				"PLVSUtilityGroup",
			],
		)
		self.assertEqual(
			[
				button.attrib["id"]
				for button in groups[2]
			],
			[
				"PLVSChangeChartType",
				"PLVSToggleTitle",
				"PLVSEditTitle",
				"PLVSToggleLegend",
				"PLVSToggleGridlines",
			],
		)
		self.assertEqual(
			[
				control.attrib["id"]
				for control in groups[3]
			],
			[
				"PLVSSeriesColor",
				"PLVSLineColor",
				"PLVSLineWidth",
				"PLVSPointSize",
				"PLVSTogglePoints",
				"PLVSToggleConnections",
			],
		)
		self.assertEqual(
			[control.attrib["id"] for control in groups[4]],
			["PLVSChangeTrendline", "PLVSPolynomialDegree", "PLVSMovingAveragePeriod"],
		)
		self.assertEqual(
			[control.attrib["id"] for control in groups[1]],
			["PLVSSwapAxes", "PLVSEditXAxisTitle", "PLVSEditYAxisTitle"],
		)
		self.assertEqual(
			[control.attrib["id"] for control in text_group],
			["PLVSTextTarget", "PLVSFontName", "PLVSFontSize", "PLVSFontBold", "PLVSFontItalic", "PLVSFontColor"],
		)
		self.assertEqual(len(groups[3].findall("r:dropDown", namespace)), 4)
		self.assertFalse(root.findall(".//r:colorPicker", namespace))
		self.assertFalse(any("#" in control.attrib.get("label", "") for control in controls))
		self.assertTrue(all(control.tag.endswith(("dropDown", "toggleButton")) for control in (
			controls_by_id["PLVSChangeChartType"],
			controls_by_id["PLVSLineWidth"],
			controls_by_id["PLVSPointSize"],
			controls_by_id["PLVSChangeTrendline"],
			controls_by_id["PLVSTogglePoints"],
			controls_by_id["PLVSToggleConnections"],
			controls_by_id["PLVSSeriesColor"],
			controls_by_id["PLVSLineColor"],
			controls_by_id["PLVSFontColor"],
		)))
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
		self.assertFalse(any("Datenquelle" in control.attrib.get("label", "") for control in controls))
		bridge = bridge_path.read_text(encoding="utf-8")
		python_functions = {
			"PLVS_RibbonCreateChart": "create_chart",
			"PLVS_RibbonSwapAxes": "swap_axes",
			"PLVS_RibbonChangeChartType": "change_chart_type",
			"PLVS_RibbonSeriesColor": "set_series_color",
			"PLVS_RibbonLineColor": "set_line_color",
			"PLVS_RibbonLineWidth": "set_line_width",
			"PLVS_RibbonPointSize": "set_point_size",
			"PLVS_RibbonTogglePoints": "toggle_points",
			"PLVS_RibbonToggleConnections": "set_connections",
			"PLVS_RibbonChangeTrendline": "change_trendline",
			"PLVS_RibbonChangePolynomialDegree": "change_polynomial_degree",
			"PLVS_RibbonChangeMovingAveragePeriod": "change_moving_average_period",
			"PLVS_RibbonEditTitle": "edit_chart_title",
			"PLVS_RibbonEditXAxisTitle": "edit_x_axis_title",
			"PLVS_RibbonEditYAxisTitle": "edit_y_axis_title",
			"PLVS_RibbonToggleLegend": "set_legend",
			"PLVS_RibbonToggleGridlines": "toggle_gridlines",
			"PLVS_RibbonToggleTitle": "toggle_title",
			"PLVS_RibbonTextTarget": "set_text_target",
			"PLVS_RibbonFontName": "set_text_font_name",
			"PLVS_RibbonFontSize": "set_text_font_size",
			"PLVS_RibbonFontColor": "set_text_color",
			"PLVS_RibbonToggleBold": "toggle_text_bold",
			"PLVS_RibbonToggleItalic": "toggle_text_italic",
			"PLVS_RibbonAISettings": "open_ai_settings",
			"PLVS_RibbonTestAI": "test_ai",
			"PLVS_RibbonHelp": "show_help",
		}
		for callback, function in python_functions.items():
			self.assertIn(f"onAction=\"{callback}\"", ribbon_path.read_text(encoding="utf-8"))
			self.assertIn(f"Public Sub {callback}(", bridge)
			self.assertIn(f'"excel_addin.{function}"', bridge)
		self.assertIn("Public Sub PLVS_RibbonOnLoad(ribbon As IRibbonUI)", bridge)
		self.assertIn("Private Function PLVS_RibbonColorFromIndex(ByVal index As Integer) As Long", bridge)
		self.assertIn(
			"Public Sub PLVS_RibbonSeriesColor(control As IRibbonControl, selectedId As String, index As Integer)",
			bridge,
		)
		self.assertIn(
			"Public Sub PLVS_RibbonFontColor(control As IRibbonControl, selectedId As String, index As Integer)",
			bridge,
		)
		for callback in (
			"PLVS_RibbonSeriesColor",
			"PLVS_RibbonLineColor",
			"PLVS_RibbonFontColor",
		):
			self.assertIn(
				f"Public Sub {callback}(control As IRibbonControl, selectedId As String, index As Integer)",
				bridge,
			)
		self.assertIn("Public Sub PLVS_RibbonGetEditEnabled(control As IRibbonControl, ByRef returnedVal)", bridge)
		self.assertIn("Public Sub PLVS_RibbonGetEditPressed(control As IRibbonControl, ByRef returnedVal)", bridge)
		self.assertIn("Public Sub PLVS_RibbonGetEditColor(control As IRibbonControl, ByRef returnedVal)", bridge)
		self.assertIn("Public Sub PLVS_RibbonToggleLegend(control As IRibbonControl, pressed As Boolean)", bridge)
		self.assertIn("Public Sub PLVS_RibbonToggleConnections(control As IRibbonControl, pressed As Boolean)", bridge)
		self.assertNotIn("PLVS_RibbonPointColor", bridge)
		self.assertIn("Private Function PLVS_SelectedTextFont(ByVal activeChart As Chart) As Object", bridge)
		self.assertIn("Public Sub PLVS_RibbonGetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)", bridge)
		self.assertIn("Public Function PLVS_BuildPythonCall(ByVal functionName As String, Optional ByVal arguments As Variant) As String", bridge)
		self.assertIn("Private Function PLVS_PythonArgument(ByVal value As Variant) As String", bridge)
		self.assertIn('Case 34: escaped = escaped & Chr$(39) & " + chr(34) + " & Chr$(39)', bridge)
		self.assertIn("Public Sub PLVS_RunChartPythonCall(ByVal functionName As String, Optional ByVal firstArgument As Variant)", bridge)
		for callback in (
			"PLVS_RibbonGetTrendlineCount",
			"PLVS_RibbonGetTrendlineLabel",
			"PLVS_RibbonGetTrendlineId",
			"PLVS_RibbonGetTrendlineSelectedIndex",
			"PLVS_RibbonGetTrendlineOptionVisible",
			"PLVS_RibbonGetPolynomialDegreeCount",
			"PLVS_RibbonGetPolynomialDegreeLabel",
			"PLVS_RibbonGetPolynomialDegreeId",
			"PLVS_RibbonGetMovingAveragePeriodCount",
			"PLVS_RibbonGetMovingAveragePeriodLabel",
			"PLVS_RibbonGetMovingAveragePeriodId",
		):
			self.assertIn(f"Public Sub {callback}(", bridge)
		self.assertIn("Set PLVS_ApplicationEvents.ExcelApp = Application", bridge)
		self.assertNotIn("PLVS_IsAnalysisContext", bridge)
		self.assertIn("Public Sub PLVS_RefreshAnalysis()", bridge)
		font_names = [
			item.attrib["label"]
			for item in controls_by_id["PLVSFontName"].findall("r:item", namespace)
		]
		self.assertEqual(
			font_names,
			[
				"Aptos", "Arial", "Calibri", "Times New Roman", "Cambria", "Verdana",
				"Tahoma", "Georgia", "Trebuchet MS", "Courier New", "Consolas", "Segoe UI",
			],
		)
		self.assertEqual(
			[item.attrib["id"] for item in controls_by_id["PLVSChangeChartType"].findall("r:item", namespace)],
			[
				"PLVSChartScatter",
				"PLVSChartScatterLines",
				"PLVSChartLine",
				"PLVSChartLineScatter",
				"PLVSChartBar",
				"PLVSChartColumn",
			],
		)
		self.assertNotIn("PLVS_RibbonStoredChartType", bridge)
		for chart_case, index in (
			("-4169", 0),
			("74, 75", 1),
			("4", 2),
			("65", 3),
			("57", 4),
			("51", 5),
		):
			self.assertRegex(
				bridge,
				rf"Case {chart_case}\s*\r?\n\s*PLVS_RibbonChartTypeIndex = {index}",
			)
		self.assertIn('fontNames = Array("Aptos", "Arial", "Calibri", "Times New Roman"', bridge)
		for callback in set(python_functions) - {
			"PLVS_RibbonCreateChart",
			"PLVS_RibbonAISettings",
			"PLVS_RibbonTestAI",
			"PLVS_RibbonHelp",
		}:
			match = re.search(
				rf"(?ms)^Public Sub {callback}\(.*?^End Sub\s*$",
				bridge,
			)
			self.assertIsNotNone(match, callback)
			self.assertIn("PLVS_RunChartPythonCall ", match.group(0), callback)
		self.assertEqual(
			len(re.findall(r"(?m)^\s*RunPython\b", bridge)),
			1,
			"RunPython must only be called by the central serialization helper.",
		)
		self.assertIn("Private Sub PLVS_ExecutePythonCode(ByVal code As String)", bridge)
		self.assertNotIn("PLVS_RibbonTargetArguments", bridge)
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
		for edit_control_id in (
			"PLVSSwapAxes",
			"PLVSChangeChartType",
			"PLVSSeriesColor",
			"PLVSLineColor",
			"PLVSLineWidth",
			"PLVSPointSize",
			"PLVSTogglePoints",
			"PLVSToggleConnections",
			"PLVSChangeTrendline",
			"PLVSPolynomialDegree",
			"PLVSMovingAveragePeriod",
			"PLVSEditTitle",
			"PLVSEditXAxisTitle",
			"PLVSEditYAxisTitle",
			"PLVSToggleLegend",
			"PLVSToggleGridlines",
			"PLVSToggleTitle",
			"PLVSFontBold",
			"PLVSFontItalic",
			"PLVSFontColor",
			"PLVSFontName",
			"PLVSFontSize",
			"PLVSTextTarget",
		):
			self.assertIn(f'PLVS_Ribbon.InvalidateControl "{edit_control_id}"', bridge)
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
			axes_group = ribbon.find(".//r:group[@id='PLVSAxesGroup']", namespace)
			self.assertIsNotNone(axes_group)
			swap_button = axes_group.find("r:button[@id='PLVSSwapAxes']", namespace)
			self.assertIsNotNone(swap_button)
			self.assertEqual(swap_button.attrib["onAction"], "PLVS_RibbonSwapAxes")
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
