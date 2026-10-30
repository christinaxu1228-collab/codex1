# SHEIN 半托管女装 Excel 批量上品系统

第一版提供中文 FastAPI Web 后台：上传官方 XLSX 模板、识别三个女装类目工作表、导入 CSV/JSON 商品资料、检查必填项/重复 SKU/敏感词并导出 XLSX。导出时保留模板工作表及前六行内容；不会连接 SHEIN API，也不会自动发布商品。

## 本地启动（Mac / Linux）

```bash
cd codex1
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

浏览器打开 <http://127.0.0.1:8000>。云端 Web preview 使用 `0.0.0.0` 监听，并可通过 `PORT` 环境变量指定端口。商品 CSV 需要包含：`category,title,description_en,sku,color,size,price,stock,image_url`；也支持同结构 JSON 数组。类目值使用 `女士两件套`、`女士短连衣裙` 或 `女士连体裤`。

## 测试

```bash
pytest -q
```

## 一键启动

```bash
./start.sh
```
