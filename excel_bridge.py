"""Create PLVS ULTRA Graphs workbooks and install the workbook-local VBA bridge."""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import xlwings as xw

from excel_addin import PRODUCT_NAME

PROJECT_ROOT = Path(__file__).resolve().parent
BRIDGE_MODULE = PROJECT_ROOT / "excel_vba" / "PLVSBridge.bas"
APP_EVENTS_MODULE = PROJECT_ROOT / "excel_vba" / "PLVSAppEvents.cls"
CHART_EVENTS_MODULE = PROJECT_ROOT / "excel_vba" / "PLVSChartEvents.cls"
RIBBON_DEFINITION = PROJECT_ROOT / "ribbon" / "customUI.xml"
TEMPLATE_PATH = PROJECT_ROOT / "examples" / "PLVS_ULTRA_Graphs.xlsm"
XLWINGS_VBA = Path(xw.__file__).resolve().parent / "xlwings.bas"
XLWINGS_CUSTOM_ADDIN_VBA = Path(xw.__file__).resolve().parent / "xlwings_custom_addin.bas"
XLSM_FILE_FORMAT = 52
XLAM_FILE_FORMAT = 55
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
UI_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X15AC_NAMESPACE = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"

SAMPLE_ROWS = [
	["Zeit [s]", "Strecke [m]", "Strecke Fehler [m]"],
	[0, 0.0, 0.1],
	[1, 2.1, 0.1],
	[2, 4.3, 0.2],
	[3, 6.2, 0.2],
	[4, 8.1, 0.2],
]


def _unique_path(path: Path) -> Path:
	candidate = path
	index = 2
	while candidate.exists():
		candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
		index += 1
	return candidate


def _set_xlwings_config(book: xw.Book) -> None:
	config_name = "xlwings.conf"
	if config_name in book.sheet_names:
		config_sheet = book.sheets[config_name]
	elif "_xlwings.conf" in book.sheet_names:
		config_sheet = book.sheets["_xlwings.conf"]
		config_sheet.name = config_name
	else:
		config_sheet = book.sheets.add(config_name, after=book.sheets[-1])
	config_sheet.range("A1:B2").value = [
		["INTERPRETER_WIN", str(Path(sys.executable).resolve())],
		["PYTHONPATH", str(PROJECT_ROOT)],
	]
	config_sheet.api.Visible = 2


def _install_vba_bridge(book: xw.Book) -> None:
	if not XLWINGS_VBA.is_file():
		raise FileNotFoundError(f"xlwings VBA-Modul fehlt: {XLWINGS_VBA}")
	if not BRIDGE_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Brücke fehlt: {BRIDGE_MODULE}")
	if not APP_EVENTS_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Ereignismodul fehlt: {APP_EVENTS_MODULE}")
	if not CHART_EVENTS_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Diagrammereignismodul fehlt: {CHART_EVENTS_MODULE}")
	project = book.api.VBProject
	if project is None:
		raise RuntimeError(
			"Excel verweigert den VBA-Projektzugriff. Aktiviere einmalig "
			"Datei > Optionen > Trust Center > Makroeinstellungen > "
			"'Zugriff auf das VBA-Projektobjektmodell vertrauen'."
		)
	components = project.VBComponents
	component_names = [components.Item(index).Name for index in range(1, components.Count + 1)]
	if "xlwings" not in component_names:
		components.Import(str(XLWINGS_VBA))
	components = project.VBComponents
	for index in range(components.Count, 0, -1):
		component = components.Item(index)
		if component.Name in {"PLVSBridge", "PLVSAppEvents", "PLVSChartEvents"}:
			components.Remove(component)
	project.VBComponents.Import(str(BRIDGE_MODULE))
	project.VBComponents.Import(str(APP_EVENTS_MODULE))
	project.VBComponents.Import(str(CHART_EVENTS_MODULE))


def _install_addin_vba(book: xw.Book) -> None:
	if not XLWINGS_CUSTOM_ADDIN_VBA.is_file():
		raise FileNotFoundError(f"xlwings Add-in VBA-Modul fehlt: {XLWINGS_CUSTOM_ADDIN_VBA}")
	if not BRIDGE_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Brücke fehlt: {BRIDGE_MODULE}")
	if not APP_EVENTS_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Ereignismodul fehlt: {APP_EVENTS_MODULE}")
	if not CHART_EVENTS_MODULE.is_file():
		raise FileNotFoundError(f"PLVS VBA-Diagrammereignismodul fehlt: {CHART_EVENTS_MODULE}")
	components = book.api.VBProject.VBComponents
	for module_name, module_path in (
		("xlwings", XLWINGS_CUSTOM_ADDIN_VBA),
		("PLVSBridge", BRIDGE_MODULE),
		("PLVSAppEvents", APP_EVENTS_MODULE),
		("PLVSChartEvents", CHART_EVENTS_MODULE),
	):
		for index in range(components.Count, 0, -1):
			component = components.Item(index)
			if component.Name == module_name:
				components.Remove(component)
		components.Import(str(module_path))


def _set_addin_xlwings_config(book: xw.Book, release_mode: bool = False) -> None:
	config_name = "myaddin.conf"
	if config_name in book.sheet_names:
		config_sheet = book.sheets[config_name]
	else:
		config_sheet = book.sheets.add(config_name, before=book.sheets[0])
	interpreter = (
		"%LOCALAPPDATA%\\PLVS ULTRA Graphs\\runtime\\python.exe"
		if release_mode
		else str(Path(sys.executable).resolve())
	)
	pythonpath = "%LOCALAPPDATA%\\PLVS ULTRA Graphs" if release_mode else str(PROJECT_ROOT)
	config_sheet.range("A1:B2").value = [
		["INTERPRETER_WIN", interpreter],
		["PYTHONPATH", pythonpath],
	]
	config_sheet.api.Visible = 2


def _remove_workbook_absolute_path(workbook_xml: bytes) -> bytes:
	alternate_content = re.compile(
		r"<(?P<prefix>[\w.-]+):AlternateContent\b[^>]*>.*?"
		r"</(?P=prefix):AlternateContent\s*>",
		re.DOTALL,
	)

	def remove_absolute_path_block(match: re.Match[str]) -> str:
		block = match.group(0)
		if re.search(r"<[\w.-]+:absPath\b", block):
			return ""
		return block

	text = workbook_xml.decode("utf-8")
	return alternate_content.sub(remove_absolute_path_block, text).encode("utf-8")


def _add_ribbon_to_package(addin_path: Path, release_mode: bool = False) -> None:
	with zipfile.ZipFile(addin_path, "r") as source:
		files = {info.filename: (info, source.read(info.filename)) for info in source.infolist()}
	if "customUI/customUI.xml" in files:
		raise ValueError("Das Add-in enthält bereits eine customUI/customUI.xml-Datei.")

	ribbon_xml = RIBBON_DEFINITION.read_bytes()
	ET.fromstring(ribbon_xml)
	relationships_path = "_rels/.rels"
	content_types_path = "[Content_Types].xml"
	relationships_root = ET.fromstring(files[relationships_path][1])
	existing_ids = {item.get("Id") for item in relationships_root}
	rel_id = "rIdPLVSUI"
	index = 2
	while rel_id in existing_ids:
		rel_id = f"rIdPLVSUI{index}"
		index += 1
	content_types_path = "[Content_Types].xml"
	content_types_root = ET.fromstring(files[content_types_path][1])
	has_xml_default = any(
		item.tag == f"{{{CONTENT_TYPES_NS}}}Default"
		and item.get("Extension", "").casefold() == "xml"
		and item.get("ContentType") == "application/xml"
		for item in content_types_root
	)
	if not has_xml_default:
		raise ValueError("Das Add-in-Paket enthält keinen application/xml-Standardtyp.")

	# Excel ignores the Ribbon part if it has a customUI-specific content-type override.
	ET.SubElement(
		relationships_root,
		f"{{{PACKAGE_REL_NS}}}Relationship",
		{"Id": rel_id, "Type": UI_RELATIONSHIP, "Target": "customUI/customUI.xml"},
	)

	ET.register_namespace("r", PACKAGE_REL_NS)
	files["customUI/customUI.xml"] = (
		zipfile.ZipInfo("customUI/customUI.xml"),
		ribbon_xml,
	)
	files[relationships_path] = (
		files[relationships_path][0],
		ET.tostring(relationships_root, encoding="utf-8", xml_declaration=True),
	)
	if release_mode:
		workbook_path = "xl/workbook.xml"
		files[workbook_path] = (
			files[workbook_path][0],
			_remove_workbook_absolute_path(files[workbook_path][1]),
		)
	with tempfile.NamedTemporaryFile(dir=addin_path.parent, suffix=".xlam", delete=False) as temporary:
		temporary_path = Path(temporary.name)
	try:
		with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
			for name, (info, content) in files.items():
				target.writestr(info, content)
		os.replace(temporary_path, addin_path)
	finally:
		temporary_path.unlink(missing_ok=True)


def create_addin(output_path: Path | None = None, release_mode: bool = False) -> Path:
	"""Erzeugt das PLVS Ribbon-Add-in auf Basis der aktiven xlwings Add-in Vorlage."""
	if output_path is None:
		output_path = PROJECT_ROOT / "dist" / "PLVS ULTRA Graphs.xlam"
	output_path = output_path.resolve()
	output_path.parent.mkdir(parents=True, exist_ok=True)
	app = xw.App(visible=False, add_book=True)
	book = app.books.active
	try:
		app.api.DisplayAlerts = False
		for sheet in list(book.sheets)[1:]:
			sheet.delete()
		book.sheets[0].name = PRODUCT_NAME
		_install_addin_vba(book)
		_set_addin_xlwings_config(book, release_mode=release_mode)
		book.api.SaveAs(str(output_path), FileFormat=XLAM_FILE_FORMAT)
	finally:
		book.close()
		app.quit()
	_add_ribbon_to_package(output_path, release_mode=release_mode)
	return output_path


def _initialize_workbook(book: xw.Book, sample: bool = False) -> None:
	if sample:
		data_sheet = book.sheets[0]
		data_sheet.name = "Messdaten"
		data_sheet.range("A1").value = SAMPLE_ROWS
		data_sheet.range("A1:C1").api.Font.Bold = True
		data_sheet.range("A1:C1").api.Interior.Color = 0xD6E8EF
		for sheet in list(book.sheets)[1:]:
			sheet.delete()
	_install_vba_bridge(book)
	_set_xlwings_config(book)
	book.save()


def create_template(output_path: Path = TEMPLATE_PATH) -> Path:
	"""Erstellt eine makrofähige Beispielarbeitsmappe mit Messdaten."""
	output_path = _unique_path(output_path.resolve())
	output_path.parent.mkdir(parents=True, exist_ok=True)
	app = xw.App(visible=False, add_book=True)
	book = app.books.active
	try:
		book.api.SaveAs(str(output_path), FileFormat=XLSM_FILE_FORMAT)
		_initialize_workbook(book, sample=True)
	finally:
		book.close()
		app.quit()
	return output_path


def install_in_copy(source_path: Path, output_path: Path | None = None) -> Path:
	"""Kopiert eine vorhandene .xlsx/.xlsm-Datei und installiert die VBA-Ribbon-Brücke."""
	source_path = source_path.resolve()
	if not source_path.is_file():
		raise FileNotFoundError(f"Excel-Datei nicht gefunden: {source_path}")
	if source_path.suffix.lower() not in {".xlsx", ".xlsm"}:
		raise ValueError("Es werden .xlsx- und .xlsm-Dateien unterstützt.")
	if output_path is None:
		output_path = source_path.with_name(f"{source_path.stem}_PLVS_ULTRA_Graphs.xlsm")
	output_path = _unique_path(output_path.resolve())
	output_path.parent.mkdir(parents=True, exist_ok=True)
	app = xw.App(visible=False, add_book=False)
	book = None
	try:
		book = app.books.open(str(source_path), update_links=False, read_only=False)
		book.api.SaveAs(str(output_path), FileFormat=XLSM_FILE_FORMAT)
		_initialize_workbook(book, sample=False)
	finally:
		if book is not None:
			book.close()
		app.quit()
	return output_path


def main() -> int:
	parser = argparse.ArgumentParser(description=f"Installationshilfe für {PRODUCT_NAME}.")
	subparsers = parser.add_subparsers(dest="command", required=True)
	template_parser = subparsers.add_parser("create-template", help="Beispielarbeitsmappe erzeugen")
	template_parser.add_argument("--output", type=Path, default=TEMPLATE_PATH)
	install_parser = subparsers.add_parser("install", help="Ribbon-Brücke in eine sichere Workbook-Kopie installieren")
	install_parser.add_argument("workbook", type=Path)
	install_parser.add_argument("--output", type=Path)
	addin_parser = subparsers.add_parser("create-addin", help="Eigenständiges PLVS Ribbon-Add-in erstellen")
	addin_parser.add_argument("--output", type=Path)
	addin_parser.add_argument("--release", action="store_true", help="portable Runtime-Pfade für Installer setzen")
	args = parser.parse_args()
	try:
		if args.command == "create-template":
			result = create_template(args.output)
		elif args.command == "create-addin":
			result = create_addin(args.output, release_mode=args.release)
		else:
			result = install_in_copy(args.workbook, args.output)
		print(f"{PRODUCT_NAME} erstellt: {result}")
		return 0
	except Exception as exc:
		print(f"{PRODUCT_NAME}: {exc}", file=sys.stderr)
		return 2


if __name__ == "__main__":
	raise SystemExit(main())
