from __future__ import annotations

import io, json, re, tempfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

app = FastAPI(title="SHEIN 半托管女装批量上品", version="1.0.0")
BASE = Path(__file__).parent

CATEGORIES = {
    "女士两件套": {"id": "95-1780", "min": 95, "max": 1780, "aliases": ["两件套", "two piece", "set"]},
    "女士短连衣裙": {"id": "67-1727", "min": 67, "max": 1727, "aliases": ["短连衣裙", "dress", "连衣裙"]},
    "女士连体裤": {"id": "154-1860", "min": 154, "max": 1860, "aliases": ["连体裤", "jumpsuit", "romper"]},
}
REQUIRED = ["category", "title", "description_en", "sku", "color", "size", "price", "stock", "image_url"]
SENSITIVE = ["fake", "replica", "counterfeit", "仿牌", "假货", "毒品", "weapon"]

def category_for_sheet(title: str) -> str | None:
    value = title.lower()
    for name, info in CATEGORIES.items():
        if name.lower() in value or info["id"] in value or any(a in value for a in info["aliases"]):
            return name
    return None

def detect_sheets(wb) -> list[dict[str, str]]:
    return [{"sheet": ws.title, "category": category_for_sheet(ws.title) or "未识别"} for ws in wb.worksheets]

def parse_products(raw: bytes, filename: str) -> list[dict[str, Any]]:
    if filename.lower().endswith(".json"):
        data = json.loads(raw.decode("utf-8-sig")); return data if isinstance(data, list) else data.get("products", [])
    import csv
    text = raw.decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(text)))

def normalize(row: dict[str, Any]) -> dict[str, Any]:
    aliases = {"标题":"title", "商品标题":"title", "英文描述":"description_en", "描述":"description_en", "货号":"sku", "颜色":"color", "尺码":"size", "供货价":"price", "价格":"price", "库存":"stock", "图片链接":"image_url", "类目":"category"}
    return {aliases.get(str(k).strip(), str(k).strip()): ("" if v is None else v) for k, v in row.items()}

def validate(products: list[dict[str, Any]]) -> list[dict[str, Any]]:
    errors = []; seen = set()
    for i, source in enumerate(products, 1):
        row = normalize(source); sku = str(row.get("sku", "")).strip()
        for field in REQUIRED:
            if not str(row.get(field, "")).strip(): errors.append({"row": i, "field": field, "message": "必填项为空"})
        if sku and sku in seen: errors.append({"row": i, "field": "sku", "message": "重复 SKU"})
        if sku: seen.add(sku)
        for field, value in row.items():
            if any(word.lower() in str(value).lower() for word in SENSITIVE): errors.append({"row": i, "field": field, "message": "包含敏感词"})
        try:
            price = float(row.get("price", 0)); category = str(row.get("category", ""))
            if category in CATEGORIES:
                info = CATEGORIES[category]
                if not (info["min"] <= price <= info["max"]): errors.append({"row": i, "field": "price", "message": f"供货价应为 {info['min']}-{info['max']}"})
        except (TypeError, ValueError): pass
    return errors

def export_workbook(template: bytes, products: list[dict[str, Any]]) -> bytes:
    wb = load_workbook(io.BytesIO(template))
    by_category = {c: [] for c in CATEGORIES}
    for p in products:
        row = normalize(p); by_category.setdefault(row.get("category", ""), []).append(row)
    for ws in wb.worksheets:
        category = category_for_sheet(ws.title)
        if not category: continue
        headers = {str(cell.value).strip(): cell.column for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20)) for cell in row if cell.value}
        # 模板前六行保持不变；第七行及以后按官方字段名或英文别名写入
        start = max(7, next((r for r in range(1, min(ws.max_row, 20)+1) if any(str(c.value).strip().lower() in {"sku", "商品编码", "product sku"} for c in ws[r])), 7))
        if not headers:
            headers = {name: idx for idx, name in enumerate(REQUIRED, 1)}
        for offset, product in enumerate(by_category.get(category, [])):
            target = start + 1 + offset
            for field, value in product.items():
                col = headers.get(field) or headers.get(field.replace("_", " "))
                if col is None:
                    # 仅在模板没有字段时追加列，保留既有格式
                    col = max(headers.values(), default=0) + 1; headers[field] = col; ws.cell(start, col, field)
                ws.cell(target, col, value)
    output = io.BytesIO(); wb.save(output); output.seek(0); return output.getvalue()

@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE / "templates" / "index.html").read_text(encoding="utf-8")

@app.post("/api/inspect")
async def inspect_template(template: UploadFile = File(...)):
    wb = load_workbook(io.BytesIO(await template.read()), read_only=True)
    return {"sheets": detect_sheets(wb), "categories": CATEGORIES}

@app.post("/api/validate")
async def validate_products(products: UploadFile = File(...)):
    rows = parse_products(await products.read(), products.filename or "products.csv")
    return {"count": len(rows), "errors": validate(rows), "preview": [normalize(r) for r in rows[:20]]}

@app.post("/api/export")
async def export(template: UploadFile = File(...), products: UploadFile = File(...)):
    template_bytes = await template.read(); rows = parse_products(await products.read(), products.filename or "products.csv")
    errors = validate(rows)
    if errors: return {"ok": False, "errors": errors, "preview": [normalize(r) for r in rows[:20]]}
    result = export_workbook(template_bytes, rows)
    return StreamingResponse(io.BytesIO(result), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=shein_products.xlsx"})
