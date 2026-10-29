import tempfile
import unittest
from pathlib import Path

from bulk_listing.validator import validate_csv


class ValidatorTests(unittest.TestCase):
    def run_csv(self, contents):
        with tempfile.TemporaryDirectory() as dirname:
            path = Path(dirname) / "data.csv"
            path.write_text(contents, encoding="utf-8")
            return validate_csv(path)

    def test_valid(self):
        result = self.run_csv("sku,title,description,price,stock,image_urls\nA,Good Dress,Simple description,9.99,3,https://example.com/a.jpg\n")
        self.assertEqual(result["error_count"], 0)
        self.assertFalse(result["upload_performed"])

    def test_duplicate_and_blocked_word(self):
        result = self.run_csv("sku,title,description,price,stock,image_urls\nA,Trendy Dress,Good,8,1,https://example.com/a.jpg\nA,Good Dress,seguro,0,-1,invalid\n")
        self.assertGreaterEqual(result["error_count"], 5)

    def test_missing_columns(self):
        result = self.run_csv("sku,title\nA,Dress\n")
        self.assertEqual(result["rows_checked"], 0)
        self.assertGreater(result["error_count"], 0)


if __name__ == "__main__":
    unittest.main()
