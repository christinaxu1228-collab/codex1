"""Local validations only; not official SHEIN platform rules."""
import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlparse

REQUIRED_COLUMNS = ("sku", "title", "description", "price", "stock", "image_urls")
BLOCKED_WORDS = ("trendy", "seguro")


def validate_csv(path: Path) -> dict:
    errors = []
    checked = 0
    seen = set()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        columns = reader.fieldnames or []
        missing = [name for name in REQUIRED_COLUMNS if name not in columns]
        if missing:
            return {"rows_checked": 0, "error_count": len(missing), "errors": [
                {"row": 1, "field": field, "message": "Missing CSV column"} for field in missing
            ], "upload_performed": False}
        for row_num, row in enumerate(reader, start=2):
            checked += 1
            def issue(field, message):
                errors.append({"row": row_num, "field": field, "message": message})
            data = {key: (row.get(key) or "").strip() for key in REQUIRED_COLUMNS}
            for field, value in data.items():
                if not value:
                    issue(field, "Required value missing")
            sku = data["sku"]
            if sku:
                if sku in seen:
                    issue("sku", "Duplicate SKU")
                seen.add(sku)
            for field in ("title", "description"):
                value = data[field].lower()
                for word in BLOCKED_WORDS:
                    import re
                    if re.search(r"\b" + re.escape(word) + r"\b", value):
                        issue(field, f"Blocked word: {word}")
            if data["price"]:
                try:
                    if not Decimal(data["price"]).is_finite() or Decimal(data["price"]) <= 0:
                        issue("price", "Price must be positive and finite")
                except InvalidOperation:
                    issue("price", "Invalid decimal price")
            if data["stock"]:
                try:
                    if int(data["stock"]) < 0:
                        issue("stock", "Stock must be nonnegative")
                except ValueError:
                    issue("stock", "Stock must be an integer")
            if data["image_urls"]:
                for url in data["image_urls"].split("|"):
                    parsed = urlparse(url.strip())
                    if parsed.scheme not in ("http", "https") or not parsed.netloc:
                        issue("image_urls", "Image URL must use http(s) and have a hostname")
    return {"rows_checked": checked, "error_count": len(errors), "errors": errors, "upload_performed": False}
