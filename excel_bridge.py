"""Create PLVS ULTRA Graphs workbooks and install the workbook-local VBA bridge."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import xlwings as xw

from excel_addin import PRODUCT_NAME, ensure_dashboard

PROJECT_ROOT = Path(__file__).resolve().parent
BRIDGE_MODULE = PROJECT_ROOT / "excel_vba" / "PLVSBridge.bas"
TEMPLATE_PATH = PROJECT_ROOT / "examples" / "PLVS_ULTRA_Graphs.xlsm"
XLWINGS_VBA = Path(xw.__file__).resolve().parent / "xlwings.bas"
XLSM_FILE_FORMAT = 52

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
		if component.Name == "PLVSBridge":
			components.Remove(component)
	project.VBComponents.Import(str(BRIDGE_MODULE))
	workbook_component = next(
		project.VBComponents.Item(index)
		for index in range(1, project.VBComponents.Count + 1)
		if project.VBComponents.Item(index).Type == 100
		)
	workbook_module = workbook_component.CodeModule
	source = workbook_module.Lines(1, workbook_module.CountOfLines) if workbook_module.CountOfLines else ""
	open_event = "Private Sub Workbook_Open()"
	startup_call = 'RunPython "import excel_addin; excel_addin.ensure_dashboard()"'
	if startup_call not in source:
		if open_event.casefold() in source.casefold():
			lines = source.splitlines()
			start_index = next(index for index, line in enumerate(lines) if line.strip().casefold() == open_event.casefold())
			end_index = next(
				(index for index in range(start_index + 1, len(lines)) if lines[index].strip().casefold() == "end sub"),
				len(lines),
			)
			workbook_module.InsertLines(end_index + 1, f"    {startup_call}")
		else:
			workbook_module.AddFromString(f"\n{open_event}\n    {startup_call}\nEnd Sub\n")


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
	ensure_dashboard(book)
	book.sheets[PRODUCT_NAME].activate()
	book.save()


def create_template(output_path: Path = TEMPLATE_PATH) -> Path:
	"""Erstellt eine makrofähige Beispielarbeitsmappe mit Dashboard und Messdaten."""
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
	"""Kopiert eine vorhandene .xlsx/.xlsm-Datei in eine makrofähige Assistentenkopie."""
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
	install_parser = subparsers.add_parser("install", help="Dashboard in eine sichere Workbook-Kopie installieren")
	install_parser.add_argument("workbook", type=Path)
	install_parser.add_argument("--output", type=Path)
	args = parser.parse_args()
	try:
		if args.command == "create-template":
			result = create_template(args.output)
		else:
			result = install_in_copy(args.workbook, args.output)
		print(f"{PRODUCT_NAME}-Arbeitsmappe erstellt: {result}")
		return 0
	except Exception as exc:
		print(f"{PRODUCT_NAME}: {exc}", file=sys.stderr)
		return 2


if __name__ == "__main__":
	raise SystemExit(main())
