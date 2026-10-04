from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import win32com.client

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RIBBON_XML = Path(__file__).with_name("customUI.xml")
VBA_MODULE = Path(__file__).with_name("RibbonPhysikAssistent.bas")
OUTPUT_PATH = PROJECT_ROOT / "dist" / "PhysikAssistent.xlam"


def _generate_template(destination: Path) -> Path:
	cli = Path(sys.executable).with_name("xlwings.exe")
	if not cli.exists():
		raise RuntimeError("xlwings CLI fehlt in der aktiven Python-Umgebung. Installiere requirements.txt erneut.")
	project_name = "physik_assistent_addin"
	subprocess.run(
		[str(cli), "quickstart", project_name, "--addin", "--ribbon"],
		cwd=destination,
		check=True,
	)
	return destination / project_name / f"{project_name}.xlam"


def _validate_ribbon_xml() -> bytes:
	content = RIBBON_XML.read_bytes()
	ElementTree.fromstring(content)
	return content


def _replace_ribbon_package(source: Path, destination: Path, ribbon_xml: bytes) -> None:
	with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(destination, "w") as target_zip:
		for item in source_zip.infolist():
			data = ribbon_xml if item.filename == "customUI/customUI.xml" else source_zip.read(item.filename)
			target_zip.writestr(item, data)


def _install_vba_and_config(addin_path: Path) -> None:
	app = None
	workbook = None
	try:
		app = win32com.client.DispatchEx("Excel.Application")
		app.Visible = False
		app.DisplayAlerts = False
		workbook = app.Workbooks.Open(str(addin_path.resolve()), UpdateLinks=0, ReadOnly=False)
		project = workbook.VBProject
		if project is None:
			raise RuntimeError("VBA-Projektzugriff ist gesperrt oder das xlwings-VBA-Projekt fehlt.")
		components = project.VBComponents
		if components.Count == 0:
			raise RuntimeError("VBA-Projektzugriff ist gesperrt oder das xlwings-VBA-Projekt fehlt.")
		for index in range(components.Count, 0, -1):
			component = components.Item(index)
			if component.Name == "RibbonPhysikAssistent":
				components.Remove(component)
		components.Import(str(VBA_MODULE))

		config_sheet = None
		for sheet in workbook.Worksheets:
			if sheet.Name in {"_myaddin.conf", "myaddin.conf"}:
				config_sheet = sheet
				break
		if config_sheet is None:
			config_sheet = workbook.Worksheets.Add()
			config_sheet.Name = "myaddin.conf"
		elif config_sheet.Name != "myaddin.conf":
			config_sheet.Name = "myaddin.conf"
		values = {
			"interpreter_win": str(Path(sys.executable).resolve()),
			"pythonpath": str(PROJECT_ROOT),
		}
		for row in range(1, config_sheet.UsedRange.Rows.Count + 1):
			key = str(config_sheet.Cells(row, 1).Value or "").strip().casefold()
			if key in values:
				config_sheet.Cells(row, 2).Value = values[key]
		workbook.IsAddin = True
		workbook.Save()
	except Exception as exc:
		message = str(exc).lower()
		if (
			"programmatic access" in message
			or "programmatische zugriff" in message
			or "vbproject" in message
			or "trust" in message
			or "vba-projektzugriff" in message
		):
			raise RuntimeError(
				"Excel blockiert den VBA-Projektzugriff. Aktiviere einmalig Datei > Optionen > "
				"Trust Center > Einstellungen für das Trust Center > Makroeinstellungen > "
				"'Zugriff auf das VBA-Projektobjektmodell vertrauen' und starte den Build erneut."
			) from exc
		raise RuntimeError(f"Das Excel-Add-in konnte nicht fertiggestellt werden: {exc}") from exc
	finally:
		if workbook is not None:
			workbook.Close(SaveChanges=True)
		if app is not None:
			app.Quit()

def _register_excel_addin(addin_path: Path) -> None:
	app = None
	try:
		app = win32com.client.DispatchEx("Excel.Application")
		app.Visible = False
		target = str(addin_path.resolve()).casefold()
		matching_name = None
		addin = None
		for registered in app.AddIns:
			try:
				if str(registered.FullName).casefold() == target:
					addin = registered
					break
				if registered.Name.casefold() == addin_path.name.casefold():
					matching_name = registered
			except Exception:
				continue
		if addin is None:
			addin = matching_name or app.AddIns.Add(str(addin_path.resolve()), False)
		addin.Installed = True
	finally:
		if app is not None:
			app.Quit()


def build_addin(output: Path = OUTPUT_PATH, install: bool = False) -> Path:
	output = output.resolve()
	output.parent.mkdir(parents=True, exist_ok=True)
	ribbon_xml = _validate_ribbon_xml()
	with tempfile.TemporaryDirectory(prefix="physik_assistent_build_") as folder:
		staging = Path(folder)
		template = _generate_template(staging)
		_replace_ribbon_package(template, output, ribbon_xml)
	try:
		_install_vba_and_config(output)
	except Exception:
		output.unlink(missing_ok=True)
		raise
	if install:
		cli = Path(sys.executable).with_name("xlwings.exe")
		subprocess.run([str(cli), "addin", "install", "--file", str(output)], check=True)
		install_path = Path.home() / "AppData" / "Roaming" / "Microsoft" / "Excel" / "XLSTART" / output.name
		_register_excel_addin(install_path if install_path.exists() else output)
	return output


def main() -> int:
	parser = argparse.ArgumentParser(description="Erstellt das Physik-Assistent Excel-Ribbon-Add-in.")
	parser.add_argument("--output", type=Path, default=OUTPUT_PATH)
	parser.add_argument("--install", action="store_true", help="Das fertige Add-in auch in Excel installieren")
	args = parser.parse_args()
	try:
		path = build_addin(args.output, args.install)
		print(f"Excel-Add-in erstellt: {path}")
		if args.install:
			print("Add-in installiert. Excel vollständig schließen und neu starten.")
		return 0
	except (OSError, RuntimeError, subprocess.CalledProcessError, zipfile.BadZipFile) as exc:
		print(f"Add-in-Build fehlgeschlagen: {exc}", file=sys.stderr)
		return 2


if __name__ == "__main__":
	raise SystemExit(main())
