from __future__ import annotations

import io, json, re, tempfile
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

app = FastAPI(title="SHEIN 半托管女装批量上品", version="1.0.0")
BASE = Path(__file__).parent

CATEGORIES = {
    "女士两件套": {"id": "95-1780", "aliases": ["两件套", "two piece", "set"]},
    "女士短连衣裙": {"id": "67-1727", "aliases": ["短连衣裙", "dress", "连衣裙"]},
    "女士连体裤": {"id": "154-1860", "aliases": ["连体裤", "jumpsuit", "romper"]},
}
REQUIRED = ["category", "title", "description_en", "sku", "color", "size", "price", "stock", "image_url"]
SENSITIVE = ["fake", "replica", "counterfeit", "仿牌", "假货", "毒品", "weapon"]
DEFAULT_DECLARED_PRICE = Decimal("200.00")
PRICE_STATUSES = ["待提交", "已提交，待核价", "核价通过", "核价未通过", "需要调整报价"]
FIELD_CANDIDATES = {
    "sku": ["supplier_sku", "sku", "货号", "商品货号", "供应商货号"],
    "title": ["default_product_name[en]", "product_name[en]", "商品名称", "商品标题", "title"],
    "description_en": ["default_product_desc[en]", "product_desc[en]", "英文描述", "description_en", "描述"],
    "price": ["cost", "申报供货价", "供货价", "supplier_cost", "price"],
    "image_url": ["main_image", "image", "商品图片", "图片链接", "image_url", "product_image"],
    "color": ["color", "颜色", "主规格", "颜色名称"],
    "size": ["size", "尺码", "尺码名称"],
    "stock": ["warehouse_stock", "stock", "库存", "仓库库存"],
    "category": ["category", "类目", "商品类目"],
}
class TemplateError(ValueError):
    pass

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

def money(value: Any, default: Decimal = DEFAULT_DECLARED_PRICE) -> str:
    """Return a two-decimal declared supplier price; blank values use the default."""
    if value is None or not str(value).strip(): value = default
    try:
        amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return str(value)
    return f"{amount:.2f}"

def prepare_products(products: list[dict[str, Any]], default_price: Any = DEFAULT_DECLARED_PRICE) -> list[dict[str, Any]]:
    try: fallback = Decimal(str(default_price)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError): fallback = DEFAULT_DECLARED_PRICE
    prepared = []
    for source in products:
        row = normalize(source)
        row["price"] = money(row.get("price"), fallback)
        row.setdefault("declared_price", row["price"])
        row.setdefault("shein_feedback_price", "")
        row.setdefault("pricing_status", "待提交")
        row.setdefault("pricing_note", "")
        prepared.append(row)
    return prepared

def validate(products: list[dict[str, Any]]) -> list[dict[str, Any]]:
    errors = []; seen = set()
    for i, source in enumerate(products, 1):
        row = normalize(source); sku = str(row.get("sku", "")).strip()
        if str(row.get("category", "")).strip() not in CATEGORIES:
            errors.append({"row": i, "field": "category", "message": "商品类目必须是支持的三个女装类目"})
        for field in REQUIRED:
            if not str(row.get(field, "")).strip(): errors.append({"row": i, "field": field, "message": "必填项为空"})
        if sku and sku in seen: errors.append({"row": i, "field": "sku", "message": "重复 SKU"})
        if sku: seen.add(sku)
        for field, value in row.items():
            if any(word.lower() in str(value).lower() for word in SENSITIVE): errors.append({"row": i, "field": field, "message": "包含敏感词"})
        # 类目编号（例如 95-1780）仅用于识别工作表，不是供货价范围。
        # 官方模板未提供统一的跨类目价格上下限，因此这里只校验价格是否为数字。
        if str(row.get("price", "")).strip():
            try: float(row["price"])
            except (TypeError, ValueError): errors.append({"row": i, "field": "price", "message": "供货价必须是数字"})
    return errors

def template_columns(ws) -> dict[str, int]:
    """Map internal CSV fields using official template rows 2-3, without writing headers."""
    values: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        labels = [str(ws.cell(row, col).value or "").strip() for row in (2, 3, 4, 5, 6)]
        normalized = {label.lower() for label in labels if label}
        for field, candidates in FIELD_CANDIDATES.items():
            if any(candidate.lower() in normalized for candidate in candidates):
                values[field] = col
    return values

def export_workbook(template: bytes, products: list[dict[str, Any]]) -> bytes:
    wb = load_workbook(io.BytesIO(template))
    found = {category_for_sheet(ws.title) for ws in wb.worksheets}
    missing = [name for name in CATEGORIES if name not in found]
    if missing:
        raise TemplateError("官方模板缺少工作表：" + "、".join(missing))
    by_category = {c: [] for c in CATEGORIES}
    for p in products:
        row = normalize(p); by_category.setdefault(row.get("category", ""), []).append(row)
    for ws in wb.worksheets:
        category = category_for_sheet(ws.title)
        if not category: continue
        columns = template_columns(ws)
        if "sku" not in columns or "price" not in columns:
            raise TemplateError(f"工作表“{ws.title}”缺少官方 SKU 或 cost/供货价字段（请检查第2-6行）")
        for offset, product in enumerate(by_category.get(category, [])):
            target = 7 + offset
            for field, value in product.items():
                if field in {"declared_price", "shein_feedback_price", "pricing_status", "pricing_note"}:
                    if field != "declared_price":
                        continue  # 核价管理字段属于本系统记录，不写入官方商品模板
                    field = "price"
                col = columns.get(field)
                if col is None:
                    continue  # 未在官方第2-6行声明的输入不写入模板，也不制造第7行表头
                if field == "price": value = money(value)
                ws.cell(target, col, float(value) if field == "price" and str(value).replace('.', '', 1).isdigit() else value)
                if field == "price": ws.cell(target, col).number_format = "0.00"
            # Fill a template currency column only when the official sheet exposes one.
            for col in range(1, ws.max_column + 1):
                label = str(ws.cell(2, col).value or ws.cell(3, col).value or "").lower()
                if "currency" in label or "币种" in label:
                    if not ws.cell(target, col).value: ws.cell(target, col, "USD")
    output = io.BytesIO(); wb.save(output); output.seek(0); return output.getvalue()

@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE / "templates" / "index.html").read_text(encoding="utf-8")

@app.post("/api/inspect")
async def inspect_template(template: UploadFile = File(...)):
    wb = load_workbook(io.BytesIO(await template.read()), read_only=True)
    return {"sheets": detect_sheets(wb), "categories": CATEGORIES}

@app.post("/api/validate")
async def validate_products(products: UploadFile = File(...), default_price: str = Form("200.00")):
    rows = prepare_products(parse_products(await products.read(), products.filename or "products.csv"), default_price)
    return {"count": len(rows), "errors": validate(rows), "preview": rows[:20], "pricing_statuses": PRICE_STATUSES}

@app.post("/api/export")
async def export(template: UploadFile = File(...), products: UploadFile = File(...), default_price: str = Form("200.00")):
    template_bytes = await template.read(); rows = prepare_products(parse_products(await products.read(), products.filename or "products.csv"), default_price)
    errors = validate(rows)
    if errors: return {"ok": False, "errors": errors, "preview": rows[:20]}
    try: result = export_workbook(template_bytes, rows)
    except TemplateError as exc: return {"ok": False, "errors": [{"row": 0, "field": "template", "message": str(exc)}], "preview": rows[:20]}
    return StreamingResponse(io.BytesIO(result), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=shein_products.xlsx"})
