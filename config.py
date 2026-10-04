"""Liest die lokale, optionale API-Konfiguration aus config.json."""

from __future__ import annotations

import json
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import TypedDict


class AIConfig(TypedDict):
	endpoint: str
	api_key: str
	model: str


DEFAULT_CONFIG: AIConfig = {"endpoint": "", "api_key": "", "model": ""}


def load_config(path: Path | None = None) -> AIConfig:
	if path is None:
		path = Path(__file__).with_name("config.json")
	if not path.exists():
		raise FileNotFoundError(
			"Keine config.json gefunden. Kopiere config.json.example nach config.json und trage Endpoint und Modell ein."
		)
	try:
		with path.open(encoding="utf-8") as config_file:
			values = json.load(config_file)
	except json.JSONDecodeError as exc:
		raise ValueError("config.json enthält ungültiges JSON.") from exc
	if not isinstance(values, dict):
		raise ValueError("Die Konfiguration muss ein JSON-Objekt sein.")
	config: AIConfig = {}
	for key, default in DEFAULT_CONFIG.items():
		value = values.get(key, default)
		if not isinstance(value, str):
			raise ValueError(f"Der Konfigurationswert {key} muss Text sein.")
		config[key] = value.strip()
	return config


def save_config(config: AIConfig, path: Path | None = None) -> Path:
	"""Speichert Zugangsdaten lokal; der Aufrufer darf Werte nicht protokollieren."""
	if path is None:
		path = Path(__file__).with_name("config.json")
	path.parent.mkdir(parents=True, exist_ok=True)
	validated: AIConfig = {}
	for key in DEFAULT_CONFIG:
		value = config.get(key, "")
		if not isinstance(value, str):
			raise ValueError(f"Der Konfigurationswert {key} muss Text sein.")
		validated[key] = value.strip()
	with NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, delete=False, suffix=".tmp") as temporary:
		json.dump(validated, temporary, ensure_ascii=False, indent=2)
		temporary.write("\n")
		temporary_path = Path(temporary.name)
	temporary_path.replace(path)
	return path