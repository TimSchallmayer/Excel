from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET
from types import SimpleNamespace
from unittest.mock import Mock

from excel_bridge import _remove_workbook_absolute_path, _set_addin_xlwings_config


class ReleasePackagingTests(unittest.TestCase):
	def test_release_addin_strips_excel_absolute_workbook_path(self) -> None:
		workbook_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
 mc:Ignorable="x15">
 <mc:AlternateContent><mc:Choice Requires="x15">
  <x15ac:absPath xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"
   url="C:\\Users\\station\\Documents\\Code\\Physik_helfer\\dist\\" />
 </mc:Choice></mc:AlternateContent><sheets />
</workbook>"""

		result = _remove_workbook_absolute_path(workbook_xml)

		self.assertNotIn(b"Physik_helfer", result)
		self.assertIn(b"sheets", result)
		self.assertIn(b'mc:Ignorable="x15"', result)
		ET.fromstring(result)

	def test_release_addin_configuration_uses_install_relative_placeholders(self) -> None:
		config_range = SimpleNamespace(value=None)
		config_sheet = SimpleNamespace(
			range=Mock(return_value=config_range),
			api=SimpleNamespace(Visible=0),
		)
		book = SimpleNamespace(
			sheet_names=["myaddin.conf"],
			sheets={"myaddin.conf": config_sheet},
		)

		_set_addin_xlwings_config(book, release_mode=True)

		self.assertEqual(
			config_range.value,
			[
				["INTERPRETER_WIN", "@PLVS_INSTALL_DIR@\\runtime\\python.exe"],
				["PYTHONPATH", "@PLVS_INSTALL_DIR@"],
			],
		)
		self.assertEqual(config_sheet.api.Visible, 2)


if __name__ == "__main__":
	unittest.main()
