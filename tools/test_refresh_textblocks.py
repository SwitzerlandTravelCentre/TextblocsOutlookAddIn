import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, Mock

from refresh_textblocks import download, refresh, validate


class RefreshTests(unittest.TestCase):
    def setUp(self):
        self.data = {"schemaVersion": 1, "blockCount": 1, "blocks": [{
            "id": "sample-de", "sourceItemId": "Sheet1!2", "category": "Test",
            "topic": "Test", "usage": "", "languageLabel": "German",
            "language": "DE", "weight": 0, "text": "Test"}]}

    def test_empty_conversion_cannot_replace_the_data(self):
        self.data.update(blockCount=0, blocks=[])
        with self.assertRaises(ValueError):
            validate(self.data)

    def test_duplicate_ids_are_rejected_even_when_count_matches(self):
        self.data["blocks"].append(copy.deepcopy(self.data["blocks"][0]))
        self.data["blockCount"] = 2
        with self.assertRaises(ValueError):
            validate(self.data)

    def test_nan_weight_is_rejected(self):
        self.data["blocks"][0]["weight"] = float("nan")
        with self.assertRaises(ValueError):
            validate(self.data)

    def test_formatter_cannot_silently_drop_text(self):
        self.data["blocks"][0]["formattedText"] = [{"text": "Tes", "bold": True}]
        with self.assertRaises(ValueError):
            validate(self.data)

    def test_active_content_url_is_rejected(self):
        self.data["blocks"][0]["formattedText"] = [{"text": "Test", "href": "javascript:alert(1)"}]
        with self.assertRaises(ValueError):
            validate(self.data)

    def test_invalid_xlsx_preserves_previous_output(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "broken.xlsx"
            source.write_text("not a workbook")
            output = Path(folder) / "textblocks.json"
            original = json.dumps(self.data).encode()
            output.write_bytes(original)
            with self.assertRaises(Exception):
                refresh(source, output)
            self.assertEqual(output.read_bytes(), original)

    def test_source_permission_failure_never_uses_cached_workbook(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "source.xlsx"
            with patch.dict("os.environ", {"TEXTBLOCKS_GRAPH_TOKEN": "test", "TEXTBLOCKS_DRIVE_ID": "drive", "TEXTBLOCKS_ITEM_ID": "item"}):
                with patch("refresh_textblocks.requests.get", return_value=Mock(status_code=403)):
                    with self.assertRaises(RuntimeError):
                        download(output)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
