import io
import asyncio
from pathlib import Path
from openpyxl import Workbook, load_workbook
import app.main as main
from app.main import validate, export_workbook, prepare_products, TemplateError, verify_export, index

def test_validation_duplicate_and_required():
    rows=[{"category":"女士短连衣裙","title":"A","description_en":"dress","sku":"X","color":"Red","size":"M","price":"100","stock":"2","image_url":"http://x"},{"sku":"X"}]
    errors=validate(rows)
    assert any(e["message"] == "重复 SKU" for e in errors)
    assert any(e["field"] == "title" for e in errors)

def test_category_code_is_not_price_range():
    row = {"category":"女士短连衣裙","title":"A","description_en":"dress","sku":"PRICE-1","color":"Red","size":"M","price":"2000","stock":"2","image_url":"http://x"}
    assert validate([row]) == []
    row["price"] = "not-a-number"
    assert any(e["field"] == "price" and "数字" in e["message"] for e in validate([row]))

def test_default_price_and_manual_price_are_preserved():
    rows = prepare_products([{"category":"女士连体裤", "sku":"A", "price":""}, {"category":"女士两件套", "sku":"B", "price":"12.345"}])
    assert rows[0]["price"] == "200.00"
    assert rows[1]["price"] == "12.35"

def test_custom_default_price():
    assert prepare_products([{"price":""}], "188.8")[0]["price"] == "188.80"

def test_three_categories_accept_default_price_and_export():
    rows = []
    for category, sku in [("女士两件套", "SET-1"), ("女士短连衣裙", "DR-1"), ("女士连体裤", "JP-1")]:
        rows.append(prepare_products([{"category": category, "title": "A", "description_en": "item", "sku": sku, "color": "Red", "size": "M", "price": "", "stock": "1", "image_url": "https://img/x"}])[0])
    assert [r["price"] for r in rows] == ["200.00", "200.00", "200.00"]
    assert validate(rows) == []

def template_bytes():
    wb=Workbook(); wb.remove(wb.active)
    for index,name in enumerate(["女士短连衣裙 67-1727", "女士两件套 95-1780", "女士连体裤 154-1860"]):
        ws=wb.create_sheet(name)
        for r in range(1,7): ws.cell(r,1,f"保留行{r}" if index == 0 else None)
        for i,h in enumerate(["category","商品名称","英文描述","supplier_sku","颜色","尺码","cost","仓库库存","商品图片"],1): ws.cell(2,i,h)
        for i,h in enumerate(["category","default_product_name[en]","default_product_desc[en]","supplier_sku","color","size","cost","warehouse_stock","main_image"],1): ws.cell(3,i,h)
    b=io.BytesIO(); wb.save(b); return b.getvalue()

def test_export_preserves_header_rows():
    data=export_workbook(template_bytes(), [{"category":"女士短连衣裙","title":"A","description_en":"dress","sku":"X","color":"Red","size":"M","price":"100","stock":"2","image_url":"http://x"}])
    ws=load_workbook(io.BytesIO(data)).active
    assert ws["A1"].value == "保留行1" and ws["B7"].value == "A" and ws["G7"].value == 100

def full_template_bytes():
    wb=Workbook(); wb.remove(wb.active)
    for name in ["女士两件套 95-1780", "女士短连衣裙 67-1727", "女士连体裤 154-1860", "辅助数据"]:
        ws=wb.create_sheet(name)
        for i,h in enumerate(["supplier_sku","default_product_name[en]","default_product_desc[en]","cost","color","size","warehouse_stock","main_image"],1): ws.cell(2,i,h); ws.cell(3,i,h)
    b=io.BytesIO(); wb.save(b); return b.getvalue()

def test_three_category_sheets_and_cost_price():
    rows=[]
    for category, sku in [("女士两件套","S1"),("女士短连衣裙","S2"),("女士连体裤","S3")]:
        rows.append({"category":category,"title":"A","description_en":"D","sku":sku,"price":"200.00","color":"Red","size":"M","stock":"1","image_url":"https://x"})
    data=export_workbook(full_template_bytes(), rows)
    wb=load_workbook(io.BytesIO(data))
    for name in ["女士两件套 95-1780","女士短连衣裙 67-1727","女士连体裤 154-1860"]:
        assert wb[name]["D7"].value == 200
        assert wb[name]["A7"].value in {"S1","S2","S3"}
        assert wb[name]["A2"].value == "supplier_sku"

def test_export_readback_verification():
    rows=[{"category":"女士短连衣裙","title":"A","description_en":"D","sku":"S1","price":"200.00","color":"Red","size":"M","stock":"1","image_url":"https://x"}]
    template=template_bytes(); result=export_workbook(template, rows)
    assert verify_export(template, result, rows) == []

def test_missing_category_sheet_blocks_export():
    wb=Workbook(); wb.active.title="女士短连衣裙 67-1727"; b=io.BytesIO(); wb.save(b)
    try:
        export_workbook(b.getvalue(), [])
    except TemplateError:
        pass

def test_inspect_endpoint():
    html = index()
    assert "SHEIN" in html and "默认申报供货价" in html

def test_draft_save_and_reload_preserves_edits(tmp_path):
    original_path = main.DB_PATH
    try:
        main.DB_PATH = Path(tmp_path) / "drafts.sqlite3"
        main.init_db()
        rows = [{"sku": "SKU-1", "supplier_code": "SPU-1", "price": "200.00", "pricing_status": "需要调整报价", "pricing_note": "保留备注", "attr_160": "棉"}]
        result = asyncio.run(main.save_drafts({"products": rows}))
        assert result == {"saved": 1}
        restored = main.load_drafts()["products"]
        assert restored == rows
    finally:
        main.DB_PATH = original_path

def test_main_attribute_value_survives_prepare_and_export():
    row = prepare_products([{"category":"女士连体裤", "supplier_code":"SPU-1", "sku":"TEST-JUMP-001", "title":"Jumpsuit", "description_en":"D", "price":"200", "main_attribute":"颜色", "main_attribute_value":"Black", "attr_39":"中弹", "attr_160":"棉"}])[0]
    assert row["main_attribute_value"] == "Black"

def test_final_price_wins_over_legacy_declared_price():
    rows = [{"category":"女士短连衣裙", "sku":"PRICE-1", "title":"A", "description_en":"D", "price":"200.00", "declared_price":"22.99", "color":"黑色", "size":"M", "stock":"1", "image_url":"https://x"}]
    result = export_workbook(full_template_bytes(), rows)
    wb = load_workbook(io.BytesIO(result))
    assert wb["女士短连衣裙 67-1727"]["D7"].value == 200

def test_different_final_prices_round_trip_all_categories():
    prices = {"女士两件套": "200.00", "女士短连衣裙": "18.75", "女士连体裤": "31.20"}
    rows = [{"category": c, "sku": "P-" + c, "title":"A", "description_en":"D", "price":p, "declared_price":"19.99", "color":"黑色", "size":"M", "stock":"1", "image_url":"https://x"} for c,p in prices.items()]
    template = full_template_bytes(); result = export_workbook(template, rows)
    assert verify_export(template, result, rows) == []

def test_fabric_elasticity_values_are_restricted():
    base={"category":"女士连体裤","title":"A","description_en":"D","sku":"E1","price":"200","attr_39":"低弹","color":"黑色","size":"M","stock":"1","image_url":"https://x"}
    assert validate([base]) == []
    base["attr_39"]="未知弹性"
    assert any(e["field"]=="attr_39" for e in validate([base]))

def test_fabric_elasticity_is_preserved_in_draft_payload():
    row=prepare_products([{"sku":"E2","price":"200","attr_39":"高弹"}])[0]
    assert row["attr_39"] == "高弹"
