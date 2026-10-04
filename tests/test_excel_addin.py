from __future__ import annotations

import unittest
from types import SimpleNamespace

from excel_addin import _frame_from_sheet


class ActiveSheetReadTests(unittest.TestCase):
	def test_used_range_becomes_frame_and_duplicate_headers_are_unique(self) -> None:
		sheet = SimpleNamespace(used_range=SimpleNamespace(value=(
			("Zeit [s]", "Spannung [V]", "Zeit [s]"),
			(0, 0, 0),
			(1, 2, 1),
			(None, None, None),
		)))
		frame = _frame_from_sheet(sheet)
		self.assertEqual(list(frame.columns), ["Zeit [s]", "Spannung [V]", "Zeit [s] (2)"])
		self.assertEqual(len(frame), 2)
		self.assertEqual(frame.iloc[1, 1], 2)

if __name__ == "__main__":
	unittest.main()
