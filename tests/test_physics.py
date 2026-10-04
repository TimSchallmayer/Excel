from __future__ import annotations

import unittest

import pandas as pd

from physics import identify_column, infer_chart_candidates


class PhysicsDatabaseTests(unittest.TestCase):
	def test_extended_quantity_names_and_units(self) -> None:
		cases = {
			"Impuls [kg·m/s]": "Impuls",
			"Drehmoment [N·m]": "Drehmoment",
			"Dichte [kg/m³]": "Dichte",
			"Kapazität [F]": "Kapazität",
			"Magnetische Flussdichte [T]": "Magnetische Flussdichte",
			"Entropie [J/K]": "Entropie",
			"Wellenlänge [m]": "Wellenlänge",
			"Halbwertszeit [s]": "Halbwertszeit",
			"Brennweite [m]": "Brennweite",
			"Spezifische Wärmekapazität [J/(kg·K)]": "Spezifische Wärmekapazität",
			"Magnetischer Fluss [Wb]": "Magnetischer Fluss",
			"Winkelbeschleunigung [rad/s²]": "Winkelbeschleunigung",
		}
		for header, quantity in cases.items():
			with self.subTest(header=header):
				self.assertEqual(identify_column(header)[0], quantity)

	def test_time_measurement_creates_high_confidence_relation_hint(self) -> None:
		frame = pd.DataFrame({"Zeit [s]": [0, 1, 2], "Geschwindigkeit [m/s]": [0, 3, 6]})
		candidates = infer_chart_candidates(frame, list(frame.columns))
		self.assertEqual(candidates[0].x_column, "Zeit [s]")
		self.assertEqual(candidates[0].y_column, "Geschwindigkeit [m/s]")
		self.assertEqual(candidates[0].relationship, "v(t)")

	def test_voltage_current_keeps_both_possible_directions(self) -> None:
		frame = pd.DataFrame({"Spannung [V]": [1, 2, 3], "Stromstärke [A]": [0.1, 0.2, 0.3]})
		candidates = infer_chart_candidates(frame, list(frame.columns))
		orientations = {(item.x_column, item.y_column) for item in candidates}
		self.assertIn(("Spannung [V]", "Stromstärke [A]"), orientations)
		self.assertIn(("Stromstärke [A]", "Spannung [V]"), orientations)

	def test_unrecognized_columns_keep_low_confidence_both_directions(self) -> None:
		frame = pd.DataFrame({"Probe links": [1, 2, 3], "Probe rechts": [4, 5, 6]})
		candidates = infer_chart_candidates(frame, list(frame.columns))
		self.assertEqual(len(candidates), 2)
		self.assertLess(max(item.confidence for item in candidates), 0.5)

	def test_frequency_energy_relation_is_available(self) -> None:
		frame = pd.DataFrame({"Frequenz [Hz]": [1, 2, 3], "Photonenergie [J]": [1, 2, 3]})
		candidates = infer_chart_candidates(frame, list(frame.columns))
		self.assertEqual(candidates[0].relationship, "E(f)")


if __name__ == "__main__":
	unittest.main()