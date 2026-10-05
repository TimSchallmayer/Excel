from __future__ import annotations

import unittest
import re
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd

from excel_addin import _analysis_label, _choose_data_source, _connected_regions, _display_address, _frame_from_values, _set_text_name, _source_for_chart, _source_label
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
		first = (
			SimpleNamespace(name="Versuch1"),
			SimpleNamespace(address="$A$1:$B$4"),
			"Versuch1: A1:B4",
		)
		second = (
			SimpleNamespace(name="Versuch2"),
			SimpleNamespace(address="$A$1:$B$4"),
			"Versuch2: A1:B4",
		)
		book = SimpleNamespace(app=SimpleNamespace(api=SimpleNamespace(InputBox=Mock(return_value="2"))))
		with (
			patch("excel_addin._discover_data_sources", return_value=[first, second]),
			patch("excel_addin._store_source") as store_source,
		):
			selected = _choose_data_source(book)
		self.assertEqual(selected, second[:2])
		prompt = book.app.api.InputBox.call_args.kwargs["Prompt"]
		self.assertIn("Versuch1: A1:B4", prompt)
		self.assertIn("Versuch2: A1:B4", prompt)
		self.assertNotIn("!", prompt)
		self.assertNotIn("$", prompt)
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

	def test_active_table_is_used_without_opening_source_dialog(self) -> None:
		book = SimpleNamespace()
		active = (SimpleNamespace(name="Aktiv"), SimpleNamespace(address="A1:B4"))
		with (
			patch("excel_addin._current_data_range", return_value=active),
			patch("excel_addin._store_source") as store_source,
			patch("excel_addin._choose_data_source") as choose_source,
		):
			self.assertEqual(_source_for_chart(book), active)
		store_source.assert_called_once_with(book, active[0], active[1], "Aktiv: A1:B4")
		choose_source.assert_not_called()

	def test_source_dialog_is_reached_only_when_no_active_table_exists(self) -> None:
		book = SimpleNamespace()
		source = (SimpleNamespace(name="Tabelle"), SimpleNamespace(address="A1:B4"))
		with (
			patch("excel_addin._current_data_range", return_value=None),
			patch("excel_addin._choose_data_source", return_value=source) as choose_source,
		):
			self.assertEqual(_source_for_chart(book), source)
		choose_source.assert_called_once_with(book)

	def test_display_address_uses_readable_relative_a1_notation(self) -> None:
		self.assertEqual(_display_address(SimpleNamespace(address="$A$1:$B$15")), "A1:B15")

	def test_source_label_uses_sheet_and_range_without_exclamation_mark(self) -> None:
		label = _source_label(
			SimpleNamespace(name="Tabelle1"),
			SimpleNamespace(address="$A$1:$B$8"),
		)
		self.assertEqual(label, "Tabelle1: A1:B8")

	def test_long_analysis_text_uses_excel_safe_formula_chunks(self) -> None:
		value = ('Analyse "KI" und Zusammenhang; ' * 15).strip()
		with patch("excel_addin._set_name") as set_name:
			_set_text_name(SimpleNamespace(), "_PLVS_ULTRA_Analysis", value)
		formula = set_name.call_args.args[2]
		escaped_chunks = re.findall(r'"((?:""|[^"])*)"', formula)
		self.assertGreater(len(escaped_chunks), 1)
		self.assertTrue(all(len(chunk) <= 240 for chunk in escaped_chunks))
		self.assertEqual("".join(chunk.replace('""', '"') for chunk in escaped_chunks), value)

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
			"Zusammenhang: Zeit [s] → Strecke [m]",
			"Diagramm: Punktdiagramm (XY); Trendlinie: linear",
			"Messpunkte: 3 von 3; Mittelwert y: 2 m",
			"Statistik: Median y: 2 m; Min–Max y: 0–4 m; Stdabw. y: 2 m; Korrelation: 1",
			"KI-Einschätzung: 95% —",
		):
			self.assertIn(expected, label)
		self.assertNotIn("Veränderung:", label)

	def test_incomplete_rows_are_excluded_from_measurement_statistics(self) -> None:
		frame = pd.DataFrame({"Zeit": [1, 2, 3], "Messwert": [10, None, 30]})
		spec = ChartSpecification(
			x_column="Zeit",
			y_column="Messwert",
			x_label="Zeit",
			y_label="Messwert",
		)
		label = _analysis_label(frame, spec)
		self.assertIn("Messpunkte: 2 von 3; Mittelwert y: 20", label)
		self.assertIn("Median y: 20; Min–Max y: 10–30", label)

	def test_analysis_values_cannot_break_ribbon_field_separators(self) -> None:
		frame = pd.DataFrame({"Zeit | x": [1, 2], "Messwert | y": [2, 4]})
		spec = ChartSpecification(
			x_column="Zeit | x",
			y_column="Messwert | y",
			x_label="Zeit | x",
			y_label="Messwert | y",
			confidence=0.95,
			reason="Begründung | mit Trenner.",
		)
		label = _analysis_label(frame, spec)
		self.assertEqual(label.count(" | "), 5)
		self.assertIn("Unabhängig: Zeit / x; Abhängig: Messwert / y", label)
		self.assertIn("Begründung / mit Trenner.", label)

	def test_price_analysis_includes_absolute_and_percentage_change(self) -> None:
		frame = pd.DataFrame({"Datum": [3, 1, 2], "Aktienkurs [EUR]": [15, 10, 12]})
		spec = ChartSpecification(
			x_column="Datum",
			y_column="Aktienkurs [EUR]",
			x_label="Datum",
			y_label="Aktienkurs [EUR]",
			y_unit="EUR",
			confidence=0.95,
		)
		label = _analysis_label(frame, spec)
		self.assertIn("Mittelwert y: 12.33 EUR", label)
		self.assertIn("Veränderung: Kurs 10 → 15 EUR; absolut +5 EUR; Prozent: +50%", label)

	def test_price_percentage_change_is_defined_when_start_value_is_nonzero(self) -> None:
		frame = pd.DataFrame({"Woche": [1, 2], "Preis [EUR]": [0, 5]})
		spec = ChartSpecification(
			x_column="Woche",
			y_column="Preis [EUR]",
			x_label="Woche",
			y_label="Preis [EUR]",
			y_unit="EUR",
		)
		self.assertIn(
			"Prozent: nicht definiert (Startwert 0)",
			_analysis_label(frame, spec),
		)

if __name__ == "__main__":
	unittest.main()
