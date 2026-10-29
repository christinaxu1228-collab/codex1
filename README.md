# SHEIN 电商批量上品助手

面向 SHEIN 女装运营的批量商品预检项目。第一版支持 **CSV 批量导入、必填字段校验、SKU 去重、价格和库存校验、违禁词检查，以及 JSON 报告输出**。

**当前不具备真实上品功能**：尚未对接 SHEIN 授权 API。不要把本地数据验证当成平台核价或审核通过保证。

## 运行

要求 Python 3.10+，无第三方依赖。

```bash
python -m bulk_listing validate examples/products.csv --output report.json
python -m unittest discover -s tests -v
```

CSV 必填列：`sku,title,description,price,stock,image_urls`；可选列：`category,color,size`。多图 URL 用竖线 `|` 分隔。字段只是项目内部格式，不等于平台官方模板。

## 后续开发

- Excel (.xlsx) 导入、SHEIN 模板映射和变体 SKU
- 图像素材上传与图片尺寸检查
- Web 管理后台、批次进度、审核和失败重试
- 经店铺授权后对接平台官方 API，实现提交和状态回查

查看 [AGENTS.md](AGENTS.md) 获取 Codex 任务指引。

## 安全

项目默认仅做 dry-run，不自动提交商品。不要提交账号密码、Cookie、API Token、真实客户资料。正式上架应提供预览、明确确认和可追溯日志。
