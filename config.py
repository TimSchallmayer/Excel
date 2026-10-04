"""Liest die lokale, optionale API-Konfiguration aus config.json."""

from __future__ import annotations

import json
from pathlib import Path
from typing import TypedDict


class AIConfig(TypedDict):
	endpoint: str
	api_key: str
	model: str


DEFAULT_CONFIG: AIConfig = {"endpoint": "", "api_key": "", "model": ""}


def load_config(path: Path = Path("config.json")) -> AIConfig:
	if not path.exists():
		return DEFAULT_CONFIG.copy()
	with path.open(encoding="utf-8") as config_file:
		values = json.load(config_file)
	if not isinstance(values, dict):
		raise ValueError("Die Konfiguration muss ein JSON-Objekt sein.")
	return {key: str(values.get(key, DEFAULT_CONFIG[key])) for key in DEFAULT_CONFIG}