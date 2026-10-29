import io
from openpyxl import Workbook, load_workbook
from fastapi.testclient import TestClient
from app.main import app, validate, export_workbook

client = TestClient(app)

def test_validation_duplicate_and_required():
    rows=[{"category":"女士短连衣裙","title":"A","description_en":"dress","sku":"X","color":"Red","size":"M","price":"100","stock":"2","image_url":"http://x"},{"sku":"X"}]
    errors=validate(rows)
    assert any(e["message"] == "重复 SKU" for e in errors)
    assert any(e["field"] == "title" for e in errors)

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
