from __future__ import annotations

import io, json, re, tempfile, sqlite3
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

app = FastAPI(title="SHEIN 半托管女装批量上品", version="1.0.0")
BASE = Path(__file__).parent
DB_PATH = BASE.parent / "shein_drafts.sqlite3"

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
    "supplier_code": ["supplier_code", "商品货号", "货号", "SPU", "款号"],
    "sku": ["supplier_sku", "sku", "货号", "商品货号", "供应商货号"],
    "title": ["default_product_name[en]", "product_name[en]", "商品名称", "商品标题", "title"],
    "description_en": ["default_product_desc[en]", "product_desc[en]", "英文描述", "description_en", "描述"],
    "price": ["cost", "申报供货价", "供货价", "supplier_cost", "price"],
    "image_url": ["main_image", "main_img", "image", "商品图片", "图片链接", "image_url", "product_image"],
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

def template_metadata(wb) -> dict[str, Any]:
    valid = wb["有效属性值"] if "有效属性值" in wb.sheetnames else None
    options: dict[str, dict[str, list[str]]] = {}
    if valid:
        for r in range(3, valid.max_row + 1):
            category, _, label = (str(valid.cell(r, c).value or "").strip() for c in (1, 2, 3))
            if category not in CATEGORIES or not label: continue
            field = {"颜色": "main_attribute_value", "尺寸": "secondary_attribute_1_value", "面料弹性": "attr_39", "材质": "attr_160"}.get(label)
            if field:
                values = [str(valid.cell(r, c).value).strip() for c in range(6, valid.max_column + 1) if valid.cell(r, c).value not in (None, "")]
                options.setdefault(category, {}).setdefault(field, []).extend(values)
    result = {}
    for ws in wb.worksheets:
        category = category_for_sheet(ws.title)
        if category: result[category] = {"sheet": ws.title, "required": template_required_fields(ws), "options": options.get(category, {})}
    return result

def parse_products(raw: bytes, filename: str) -> list[dict[str, Any]]:
    if filename.lower().endswith(".json"):
        data = json.loads(raw.decode("utf-8-sig")); return data if isinstance(data, list) else data.get("products", [])
    import csv
    text = raw.decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(text)))

def normalize(row: dict[str, Any]) -> dict[str, Any]:
    aliases = {"标题":"title", "商品标题":"title", "英文描述":"description_en", "描述":"description_en", "货号":"supplier_code", "商品货号":"supplier_code", "SPU":"supplier_code", "卖家SKU":"sku", "SKU":"sku", "颜色":"color", "尺码":"size", "供货价":"price", "价格":"price", "库存":"stock", "仓库库存":"stock", "图片链接":"image_url", "商品图片":"image_url", "类目":"category", "主规格":"main_attribute", "主规格内容":"main_attribute_value", "材质":"attr_160", "面料弹性":"attr_39"}
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
        # Keep legacy drafts usable while preserving any explicitly edited value.
        row.setdefault("main_attribute", "颜色")
        row.setdefault("main_attribute_value", row.get("color", ""))
        row.setdefault("declared_price", row["price"])
        row.setdefault("shein_feedback_price", "")
        row.setdefault("pricing_status", "待提交")
        row.setdefault("pricing_note", "")
        prepared.append(row)
    return prepared

def validate(products: list[dict[str, Any]], required_fields: list[str] | None = None) -> list[dict[str, Any]]:
    errors = []; seen = set()
    required_fields = required_fields if required_fields is not None else REQUIRED
    for i, source in enumerate(products, 1):
        row = normalize(source); sku = str(row.get("sku", "")).strip()
        if str(row.get("category", "")).strip() not in CATEGORIES:
            errors.append({"row": i, "field": "category", "message": "商品类目必须是支持的三个女装类目"})
        for field in required_fields:
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
        if row.get("image_url") and not re.match(r"^https?://\S+$", str(row["image_url"])):
            errors.append({"row": i, "field": "image_url", "message": "图片链接必须是可访问的 http/https 地址"})
    return errors

def template_columns(ws) -> dict[str, int]:
    """Map internal CSV fields using official template rows 2-3, without writing headers."""
    values: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        labels = [str(ws.cell(row, col).value or "").strip() for row in (2, 3, 4, 5, 6)]
        normalized = {label.lower() for label in labels if label}
        identifier = str(ws.cell(3, col).value or "").strip()
        if identifier:
            if identifier.startswith("stock["): values.setdefault("stock", col)
            elif identifier.startswith("main_img"): values.setdefault("image_url", col)
            elif identifier.startswith("detail_img") or identifier in {"square_img", "piece_img", "sku_image"}: values.setdefault("image_url", col)
            elif identifier.startswith("attr_") or identifier in {"main_attribute", "main_attribute_value", "supplier_code", "supplier_sku"}:
                values.setdefault(identifier, col)
        for field, candidates in FIELD_CANDIDATES.items():
            if any(candidate.lower() in normalized for candidate in candidates):
                values[field] = col
    return values

def template_required_fields(ws) -> list[str]:
    columns = template_columns(ws); required = []
    for col in range(1, ws.max_column + 1):
        if str(ws.cell(4, col).value or "").strip() != "必填": continue
        identifier = str(ws.cell(3, col).value or "").strip()
        field = identifier
        for candidate, names in FIELD_CANDIDATES.items():
            if identifier in names: field = candidate
        if field not in required: required.append(field)
    return required

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
        # 同一货号的多个 SKU 相邻写入，稳定排序且不生成新货号。
        category_products = sorted(by_category.get(category, []), key=lambda p: (str(p.get("supplier_code", "")), str(p.get("sku", ""))))
        for offset, product in enumerate(category_products):
            target = 7 + offset
            for required in template_required_fields(ws):
                if required in {"supplier_sku", "sku"}: required = "sku"
                if not str(product.get(required, "")).strip():
                    raise TemplateError(f"工作表“{ws.title}”第 {target} 行缺少必填字段：{required}（请按第2-6行填写要求补充）")
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

def verify_export(template: bytes, result: bytes, products: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Read the generated workbook and compare every exported SKU and cost."""
    wb = load_workbook(io.BytesIO(result), data_only=False)
    expected = {(str(p.get("category")), str(p.get("sku"))): money(p.get("price")) for p in products}
    actual = set()
    errors = []
    for ws in wb.worksheets:
        category = category_for_sheet(ws.title)
        if not category: continue
        columns = template_columns(ws)
        for row in range(7, ws.max_row + 1):
            sku = ws.cell(row, columns.get("sku", 0)).value if columns.get("sku") else None
            if not sku: continue
            cost = ws.cell(row, columns.get("price", 0)).value if columns.get("price") else None
            key = (category, str(sku))
            actual.add(key)
            if key not in expected: errors.append({"sku": str(sku), "message": "导出文件出现未确认的 SKU"})
            elif cost is None or money(cost) != expected[key]: errors.append({"sku": str(sku), "message": f"cost 为 {cost}，预期 {expected[key]}"})
    for key in expected:
        if key not in actual: errors.append({"sku": key[1], "message": "导出文件缺少该 SKU"})
    return errors

def init_db():
    with sqlite3.connect(DB_PATH) as db:
        db.execute("CREATE TABLE IF NOT EXISTS drafts (sku TEXT PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)")
init_db()

@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE / "templates" / "index.html").read_text(encoding="utf-8")

@app.post("/api/inspect")
async def inspect_template(template: UploadFile = File(...)):
    wb = load_workbook(io.BytesIO(await template.read()), read_only=False)
    return {"sheets": detect_sheets(wb), "categories": CATEGORIES, "metadata": template_metadata(wb)}

@app.post("/api/validate")
async def validate_products(products: UploadFile = File(...), default_price: str = Form("200.00")):
    rows = prepare_products(parse_products(await products.read(), products.filename or "products.csv"), default_price)
    return {"count": len(rows), "errors": validate(rows), "preview": rows[:20], "pricing_statuses": PRICE_STATUSES}

@app.post("/api/export")
async def export(template: UploadFile = File(...), products: UploadFile = File(...), default_price: str = Form("200.00")):
    template_bytes = await template.read(); rows = prepare_products(parse_products(await products.read(), products.filename or "products.csv"), default_price)
    errors = validate(rows, ["category", "title", "description_en", "price"])
    if errors: return {"ok": False, "errors": errors, "preview": rows[:20]}
    try: result = export_workbook(template_bytes, rows)
    except TemplateError as exc: return {"ok": False, "errors": [{"row": 0, "field": "template", "message": str(exc)}], "preview": rows[:20]}
    verification_errors = verify_export(template_bytes, result, rows)
    if verification_errors:
        return {"ok": False, "errors": [{"row": 0, "field": "export", **error} for error in verification_errors], "preview": rows[:20]}
    return StreamingResponse(io.BytesIO(result), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=shein_products.xlsx"})

@app.post("/api/drafts")
async def save_drafts(payload: dict[str, Any]):
    rows = payload.get("products", [])
    with sqlite3.connect(DB_PATH) as db:
        for row in rows:
            sku = str(row.get("sku", "")).strip()
            if sku: db.execute("INSERT OR REPLACE INTO drafts(sku,payload,updated_at) VALUES(?,?,CURRENT_TIMESTAMP)", (sku, json.dumps(row, ensure_ascii=False)))
    return {"saved": len([r for r in rows if str(r.get("sku", "")).strip()])}

@app.get("/api/drafts")
def load_drafts():
    with sqlite3.connect(DB_PATH) as db:
        rows = db.execute("SELECT payload FROM drafts ORDER BY updated_at DESC").fetchall()
    return {"products": [json.loads(row[0]) for row in rows]}
