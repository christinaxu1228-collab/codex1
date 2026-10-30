from __future__ import annotations

import io, json, re, tempfile, sqlite3, uuid, socket, ipaddress
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import httpx

app = FastAPI(title="SHEIN 半托管女装批量上品", version="1.0.0")
BASE = Path(__file__).parent
DB_PATH = BASE.parent / "shein_drafts.sqlite3"
IMAGE_DIR = BASE.parent / "product_images"
IMAGE_DIR.mkdir(parents=True, exist_ok=True)

CATEGORIES = {
    "女士两件套": {"id": "95-1780", "aliases": ["两件套", "two piece", "set"]},
    "女士短连衣裙": {"id": "67-1727", "aliases": ["短连衣裙", "dress", "连衣裙"]},
    "女士连体裤": {"id": "154-1860", "aliases": ["连体裤", "jumpsuit", "romper"]},
}
REQUIRED = ["category", "title", "description_en", "sku", "color", "size", "price", "stock", "image_url"]
SENSITIVE = ["fake", "replica", "counterfeit", "仿牌", "假货", "毒品", "weapon"]
DEFAULT_DECLARED_PRICE = Decimal("200.00")
PRICE_STATUSES = ["待提交", "已提交，待核价", "核价通过", "核价未通过", "需要调整报价"]
FABRIC_ELASTICITY_VALUES = ["无弹", "低弹", "中弹", "高弹"]
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
        if category:
            result[category] = {"sheet": ws.title, "required": template_required_fields(ws), "options": options.get(category, {}),
                                "image_fields": template_image_fields(ws), "image_rules": template_image_rules(ws)}
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
        if str(row.get("attr_39", "")).strip() and row["attr_39"] not in FABRIC_ELASTICITY_VALUES:
            errors.append({"row": i, "field": "attr_39", "message": "面料弹性必须从官方选项：无弹、低弹、中弹、高弹中选择"})
    return errors

def template_columns(ws) -> dict[str, int]:
    """Map internal CSV fields using official template rows 2-3, without writing headers."""
    values: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        labels = [str(ws.cell(row, col).value or "").strip() for row in (2, 3, 4, 5, 6)]
        normalized = {label.lower() for label in labels if label}
        identifier = str(ws.cell(3, col).value or "").strip()
        if identifier:
            if identifier in {"main_img", "square_img", "piece_img", "sku_image"} or identifier.startswith("detail_img_"):
                values.setdefault(identifier, col)
            if identifier.startswith("stock["): values.setdefault("stock", col)
            elif identifier.startswith("main_img"): values.setdefault("image_url", col)
            elif identifier.startswith("detail_img") or identifier in {"square_img", "piece_img", "sku_image"}: values.setdefault("image_url", col)
            elif identifier.startswith("attr_") or identifier in {"main_attribute", "main_attribute_value", "supplier_code", "supplier_sku"}:
                values.setdefault(identifier, col)
        for field, candidates in FIELD_CANDIDATES.items():
            if any(candidate.lower() in normalized for candidate in candidates):
                values[field] = col
    return values

def template_image_fields(ws) -> list[str]:
    return [str(ws.cell(3, col).value).strip() for col in range(1, ws.max_column + 1)
            if str(ws.cell(3, col).value or "").strip() in {"main_img", "square_img", "piece_img", "sku_image"}
            or str(ws.cell(3, col).value or "").strip().startswith("detail_img_")]

def template_display_fields(ws) -> list[str]:
    """Generic 3:4 display slots; special square/color/SKU slots stay optional."""
    return [field for field in template_image_fields(ws) if field == "main_img" or field.startswith("detail_img_")]

def template_image_rules(ws) -> dict[str, str]:
    rules = {}
    for col in range(1, ws.max_column + 1):
        field = str(ws.cell(3, col).value or "").strip()
        if field in template_image_fields(ws):
            rules[field] = str(ws.cell(5, col).value or "") + " " + str(ws.cell(6, col).value or "")
    return rules

def image_url(value: Any) -> str:
    if isinstance(value, dict): return str(value.get("url", "")).strip()
    return str(value or "").strip()

def product_image_values(product: dict[str, Any], field: str) -> list[Any]:
    images = product.get("images") or {}
    value = images.get(field) if isinstance(images, dict) else None
    if value in (None, "") and field == "main_img": value = product.get("image_url", "")
    if value in (None, ""): value = product.get(field, "")
    if isinstance(value, list): return value
    return [value] if value not in (None, "") else []

def product_images_in_order(product: dict[str, Any]) -> list[Any]:
    """Flatten the user's ordered display-image collection without imposing categories."""
    images = product.get("images") or {}
    values: list[Any] = []
    if isinstance(images, dict):
        for value in images.values():
            values.extend(value if isinstance(value, list) else [value])
    elif isinstance(images, list):
        values.extend(images)
    if not values and product.get("image_url"):
        value = product["image_url"]
        values.extend(value if isinstance(value, list) else [value])
    return [value for value in values if value not in (None, "")]

def validate_image_asset(asset: dict[str, Any]) -> list[str]:
    errors = []
    mime = str(asset.get("mime", "")).lower()
    if mime not in {"image/jpeg", "image/jpg", "image/png", "image/webp"}:
        errors.append("仅支持 JPG、JPEG、PNG、WebP")
    try:
        width, height = int(asset["width"]), int(asset["height"])
        if width < 1 or height < 1: errors.append("图片尺寸无效")
    except (KeyError, TypeError, ValueError): errors.append("无法读取图片尺寸")
    if int(asset.get("size", 0) or 0) > 3 * 1024 * 1024: errors.append("图片大小不能超过 3MB")
    return errors

def image_dimensions(raw: bytes, suffix: str) -> tuple[int, int, str]:
    """Read dimensions without a native imaging dependency."""
    if raw.startswith(b"\x89PNG\r\n\x1a\n") and len(raw) >= 24:
        import struct
        return (*struct.unpack(">II", raw[16:24]), "image/png")
    if raw[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(raw):
            if raw[i] != 0xFF: i += 1; continue
            marker = raw[i + 1]; i += 2
            if marker in (0xD8, 0xD9): continue
            if i + 2 > len(raw): break
            length = int.from_bytes(raw[i:i + 2], "big")
            if marker in range(0xC0, 0xC4) or marker in range(0xC5, 0xC8) or marker in range(0xC9, 0xCC) or marker in range(0xCD, 0xD0):
                if i + 7 <= len(raw): return (int.from_bytes(raw[i + 5:i + 7], "big"), int.from_bytes(raw[i + 3:i + 5], "big"), "image/jpeg")
            i += length
    if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP" and raw[12:16] == b"VP8X" and len(raw) >= 30:
        return (1 + int.from_bytes(raw[24:27], "little"), 1 + int.from_bytes(raw[27:30], "little"), "image/webp")
    raise ValueError("unsupported or corrupt image")

def validate_image_for_field(asset: Any, field: str) -> list[str]:
    if isinstance(asset, str):
        if not re.match(r"^https?://\S+$", asset): return ["必须是可访问的 http/https 图片链接；本地文件不能直接导出"]
        return []
    errors = validate_image_asset(asset if isinstance(asset, dict) else {})
    if isinstance(asset, dict) and asset.get("official_url_status") not in (None, "通过"):
        errors.append("外部图片已下载但尚未确认是 SHEIN 正式认可图片地址")
    try: w, h = int(asset["width"]), int(asset["height"])
    except (KeyError, TypeError, ValueError): return errors
    if field in {"main_img", "square_img", "piece_img", "sku_image"} or field.startswith("detail_img_"):
        ratio = w / h if h else 0
        if abs(ratio - 0.75) > 0.01: errors.append("商品展示图需要 3:4 比例（允许少量误差）")
    if not image_url(asset): errors.append("尚未获得公开图片 URL，不能写入正式 XLSX")
    return errors

def public_https_url(value: str) -> tuple[bool, str]:
    try:
        from urllib.parse import urlparse
        parsed = urlparse(value.strip())
        if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
            return False, "只允许不含账号信息的 HTTPS 图片 URL"
        host = parsed.hostname
        addresses = {item[4][0] for item in socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)}
        for address in addresses:
            ip = ipaddress.ip_address(address)
            if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast or ip.is_unspecified:
                return False, "禁止访问本机、内网、保留地址或云服务元数据地址"
        return True, ""
    except Exception:
        return False, "无法解析公网地址"

async def download_external_image(url: str) -> tuple[dict[str, Any] | None, str | None]:
    allowed, reason = public_https_url(url)
    if not allowed: return None, reason
    try:
        async with httpx.AsyncClient(follow_redirects=False, timeout=15.0, headers={"User-Agent": "SHEIN-product-image-import/1.0"}) as client:
            response = await client.get(url)
        if response.status_code < 200 or response.status_code >= 300:
            return None, f"图片 URL 返回 HTTP {response.status_code}"
        raw = response.content
        if len(raw) > 3 * 1024 * 1024: return None, "图片大小不能超过 3MB"
        width, height, mime = image_dimensions(raw, Path(url).suffix.lower())
        if mime not in {"image/jpeg", "image/png", "image/webp"}:
            return None, "响应内容不是支持的 JPG、PNG 或 WebP 图片"
        asset_id = uuid.uuid4().hex + {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp"}[mime]
        (IMAGE_DIR / asset_id).write_bytes(raw)
        asset = {"id": asset_id, "filename": Path(url).name or asset_id, "path": str((IMAGE_DIR / asset_id).relative_to(IMAGE_DIR.parent)),
                 "url": url, "source_url": url, "official_url_status": "待确认", "mime": mime,
                 "size": len(raw), "width": width, "height": height, "image_type": "display"}
        return asset, None
    except httpx.HTTPError as exc:
        return None, f"下载失败：{exc.__class__.__name__}"
    except Exception as exc:
        return None, f"无法解析实际图片内容：{exc}"

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
        valid_materials = set(template_metadata(wb).get(category, {}).get("options", {}).get("attr_160", []))
        for offset, product in enumerate(category_products):
            target = 7 + offset
            for required in template_required_fields(ws):
                if required in {"supplier_sku", "sku"}: required = "sku"
                if not str(product.get(required, "")).strip():
                    raise TemplateError(f"工作表“{ws.title}”第 {target} 行缺少必填字段：{required}（请按第2-6行填写要求补充）")
            if product.get("attr_160") and valid_materials and product["attr_160"] not in valid_materials:
                raise TemplateError(f"工作表“{ws.title}”第 {target} 行材质不在官方有效属性值中：{product['attr_160']}；请以模板“有效属性值”表为准")
            image_values = product_images_in_order(product)
            image_fields = template_display_fields(ws)
            if image_fields:
                if len(image_fields) > 1 and len(image_values) < 2:
                    raise TemplateError(f"工作表“{ws.title}”第 {target} 行商品展示图片少于 2 张；无需提供固定的正面、侧面、背面或材质图")
                if len(image_values) > len(image_fields):
                    raise TemplateError(f"工作表“{ws.title}”第 {target} 行有 {len(image_values)} 张图片，但当前官方模板只有 {len(image_fields)} 个图片字段；请更换模板或减少图片")
                for index, image in enumerate(image_values):
                    image_field = image_fields[index]
                    image_errors = validate_image_for_field(image, image_field)
                    if image_errors:
                        raise TemplateError(f"工作表“{ws.title}”第 {target} 行第 {index + 1} 张商品展示图：" + "；".join(image_errors))
            for field, value in product.items():
                if field == "declared_price":
                    continue  # 旧 CSV/草稿快照不能覆盖网页最终确认的 price
                if field in {"shein_feedback_price", "pricing_status", "pricing_note"}:
                    continue  # 核价管理字段属于本系统记录，不写入官方商品模板
                col = columns.get(field)
                if col is None:
                    continue  # 未在官方第2-6行声明的输入不写入模板，也不制造第7行表头
                if field == "price": value = money(value)
                ws.cell(target, col, float(value) if field == "price" and str(value).replace('.', '', 1).isdigit() else value)
                if field == "price": ws.cell(target, col).number_format = "0.00"
            # 图片必须使用模板声明的字段，并且只能写公开 URL；本地上传文件仅作为草稿资产保存。
            if template_display_fields(ws):
                for index, image in enumerate(product_images_in_order(product)):
                    ws.cell(target, columns[template_display_fields(ws)[index]], image_url(image))
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

@app.post("/api/images/upload")
async def upload_images(files: list[UploadFile] = File(...), sku: str = Form(""), supplier_code: str = Form(""), image_type: str = Form("main_img")):
    """Persist uploaded assets for drafts. No public URL is fabricated."""
    image_type = image_type.strip() or "display"
    assets, errors = [], []
    for upload in files:
        raw = await upload.read()
        suffix = Path(upload.filename or "").suffix.lower()
        if suffix not in {".jpg", ".jpeg", ".png", ".webp"}:
            errors.append({"file": upload.filename, "message": "仅支持 JPG、JPEG、PNG、WebP"}); continue
        try:
            width, height, detected_mime = image_dimensions(raw, suffix)
            mime = (detected_mime or upload.content_type or "").lower()
        except Exception:
            errors.append({"file": upload.filename, "message": "无法读取图片文件或文件已损坏"}); continue
        asset_id = uuid.uuid4().hex + suffix
        path = IMAGE_DIR / asset_id
        path.write_bytes(raw)
        asset = {"id": asset_id, "filename": upload.filename or asset_id, "path": str(path.relative_to(IMAGE_DIR.parent)),
                 "url": "", "mime": mime, "size": len(raw), "width": width, "height": height,
                 "sku": sku.strip() or Path(upload.filename or "").stem, "supplier_code": supplier_code.strip(), "image_type": image_type}
        asset_errors = validate_image_for_field(asset, "main_img")
        if asset_errors: asset["warnings"] = asset_errors
        assets.append(asset)
    return {"ok": not errors, "assets": assets, "errors": errors,
            "notice": "图片已持久化保存，但尚未获得公开 URL；正式 XLSX 导出前必须补充 SHEIN 接受的图片链接。"}

@app.post("/api/images/import-urls")
async def import_image_urls(payload: dict[str, Any]):
    """Download user-authorized public image URLs; assets remain pending official review."""
    items = payload.get("items", payload.get("urls", []))
    if isinstance(items, str): items = [line.strip() for line in items.splitlines() if line.strip()]
    results, errors = [], []
    for index, item in enumerate(items if isinstance(items, list) else [], 1):
        if isinstance(item, str): item = {"url": item}
        url = str(item.get("url", "")).strip(); sku = str(item.get("sku", "")).strip()
        if not url: errors.append({"row": index, "message": "缺少图片 URL"}); continue
        asset, reason = await download_external_image(url)
        if reason:
            errors.append({"row": index, "sku": sku, "url": url, "message": reason}); continue
        asset.update({"sku": sku, "supplier_code": str(item.get("supplier_code", "")).strip(), "order": item.get("order", index)})
        quality = validate_image_for_field(asset, "main_img")
        asset["quality_errors"] = quality
        results.append(asset)
    return {"ok": not errors, "assets": results, "errors": errors,
            "notice": "下载成功不等于 SHEIN 正式图片地址校验通过；导出前必须完成官方地址确认。"}

@app.post("/api/images/import-csv")
async def import_image_csv(file: UploadFile = File(...)):
    import csv
    try:
        text = (await file.read()).decode("utf-8-sig")
        rows = list(csv.DictReader(io.StringIO(text)))
    except Exception as exc:
        return {"ok": False, "assets": [], "errors": [{"message": f"CSV 读取失败：{exc}"}]}
    items = [{"sku": row.get("SKU") or row.get("sku") or row.get("卖家 SKU") or "",
              "supplier_code": row.get("商品货号") or row.get("supplier_code") or row.get("货号") or "",
              "url": row.get("图片 URL") or row.get("image_url") or row.get("url") or row.get("图片链接") or "",
              "order": row.get("图片顺序") or row.get("order") or i} for i, row in enumerate(rows, 1)]
    return await import_image_urls({"items": items})

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
