from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from config import save_config


class ConfigTests(unittest.TestCase):
	def test_failed_replace_removes_temporary_config_file(self) -> None:
		with tempfile.TemporaryDirectory() as folder:
			path = Path(folder) / "config.json"
			config = {
				"endpoint": "https://example.invalid/v1/chat/completions",
				"api_key": "test-secret",
				"model": "test-model",
			}
			with patch.object(Path, "replace", side_effect=OSError("replace failed")):
				with self.assertRaisesRegex(OSError, "replace failed"):
					save_config(config, path)

			self.assertFalse(path.exists())
			self.assertEqual(list(Path(folder).glob("*.tmp")), [])


if __name__ == "__main__":
	unittest.main()
