import io
from openpyxl import Workbook, load_workbook
from fastapi.testclient import TestClient
from app.main import app, validate, export_workbook, prepare_products

client = TestClient(app)

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
    wb=Workbook(); ws=wb.active; ws.title="女士短连衣裙 67-1727"
    for r in range(1,7): ws.cell(r,1,f"保留行{r}")
    for i,h in enumerate(["category","title","description_en","sku","color","size","price","stock","image_url"],1): ws.cell(7,i,h)
    b=io.BytesIO(); wb.save(b); return b.getvalue()

def test_export_preserves_header_rows():
    data=export_workbook(template_bytes(), [{"category":"女士短连衣裙","title":"A","description_en":"dress","sku":"X","color":"Red","size":"M","price":"100","stock":"2","image_url":"http://x"}])
    ws=load_workbook(io.BytesIO(data)).active
    assert ws["A1"].value == "保留行1" and ws["B8"].value == "A"

def test_inspect_endpoint():
    r=client.post('/api/inspect', files={'template':('x.xlsx',template_bytes(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')})
    assert r.status_code == 200 and r.json()['sheets'][0]['category'] == '女士短连衣裙'
