from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd

from excel_addin import _analysis_label, _choose_data_source, _connected_regions, _frame_from_values
from models import ChartSpecification


class DataSourceTests(unittest.TestCase):
	def test_data_values_become_frame_and_duplicate_headers_are_unique(self) -> None:
		frame = _frame_from_values((
			("Zeit [s]", "Spannung [V]", "Zeit [s]"),
			(0, 0, 0),
			(1, 2, 1),
			(None, None, None),
		))
		self.assertEqual(list(frame.columns), ["Zeit [s]", "Spannung [V]", "Zeit [s] (2)"])
		self.assertEqual(len(frame), 2)
		self.assertEqual(frame.iloc[1, 1], 2)

	def test_contiguous_regions_detect_separate_candidate_tables(self) -> None:
		regions = _connected_regions(
			[
				["Zeit [s]", "Strecke [m]", None, "Kraft [N]", "Dehnung [mm]"],
				[0, 0, None, 0, 0],
				[1, 2, None, 10, 1.5],
			]
		)
		self.assertEqual(regions, [(0, 0, 2, 1), (0, 3, 2, 4)])

	def test_multiple_sources_require_an_explicit_selection(self) -> None:
		first = (SimpleNamespace(name="Versuch1"), SimpleNamespace(address="$A$1:$B$4"), "Tabelle 1")
		second = (SimpleNamespace(name="Versuch2"), SimpleNamespace(address="$A$1:$B$4"), "Tabelle 2")
		book = SimpleNamespace(app=SimpleNamespace(api=SimpleNamespace(InputBox=Mock(return_value="2"))))
		with (
			patch("excel_addin._discover_data_sources", return_value=[first, second]),
			patch("excel_addin._store_source") as store_source,
		):
			selected = _choose_data_source(book)
		self.assertEqual(selected, second[:2])
		store_source.assert_called_once_with(book, *second)

	def test_empty_multiple_source_selection_is_not_silently_first(self) -> None:
		sources = [
			(SimpleNamespace(name="Versuch1"), SimpleNamespace(address="A1:B4"), "Tabelle 1"),
			(SimpleNamespace(name="Versuch2"), SimpleNamespace(address="A1:B4"), "Tabelle 2"),
		]
		book = SimpleNamespace(app=SimpleNamespace(api=SimpleNamespace(InputBox=Mock(return_value=""))))
		with patch("excel_addin._discover_data_sources", return_value=sources):
			with self.assertRaisesRegex(ValueError, "abgebrochen"):
				_choose_data_source(book)

	def test_analysis_summary_includes_relevant_measurement_statistics(self) -> None:
		frame = pd.DataFrame({"Zeit [s]": [0, 1, 2], "Strecke [m]": [0, 2, 4]})
		spec = ChartSpecification(
			x_column="Zeit [s]",
			y_column="Strecke [m]",
			x_label="Zeit [s]",
			y_label="Strecke [m]",
			x_unit="s",
			y_unit="m",
			trendline="linear",
			confidence=0.95,
		)
		label = _analysis_label(frame, spec)
		for expected in (
			"Unabhängig: Zeit [s]; Abhängig: Strecke [m]",
			"Diagramm: XY; Fit: Linear",
			"Messpunkte: 3; Mittelwert: 2 m",
			"Confidence: 0.95",
		):
			self.assertIn(expected, label)

if __name__ == "__main__":
	unittest.main()
