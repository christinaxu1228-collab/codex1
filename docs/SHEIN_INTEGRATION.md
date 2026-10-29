# SHEIN-only 批量上品架构与官方接入方案

> 状态：设计稿 + 开发合同（非已开通生产上传）。以店铺实际应用模式、最新官方文档、授权为准。

## 目标流程
1. 导入 Excel/CSV -> 解析 SPU/SKC/SKU/变体/图片 -> 暂存 SQLite/Postgres
2. 选择店铺/业务模式/站点/币种 -> 获取官方类目、属性、字段规范
3. 校验英文标题描述、禁词、图片、销售属性、供应商 SKU 和价格库存
4. 预览并显式确认 -> 创建不可变批次快照、记录操作人和幂等键
5. 图片 upload-pic 或 transform-pic -> 映射平台图片标识
6. 小批量走 publishOrEdit，规模化经适配的 Feed PRODUCT_LISTING V2
7. 记录请求 traceId、平台 SPU/SKC/SKU 和 Feed ID -> 查询审核状态 -> 修复并选择性重试
8. 审核通过不等于已上架，另需按店铺模式核查货架状态

## 官方 OpenAPI 路线

官方中心 https://open.sheincorp.com/zh
发品指南（自运营示例） https://open.sheincorp.com/documents/system/165b51da-c8dd-43b6-a885-4aed2c7fa582
单商品发品 POST /open-api/goods/product/publishOrEdit
发品字段规范 POST /open-api/goods/query-publish-fill-in-standard
类目 /open-api/goods/query-category-tree
类目属性 /open-api/goods/query-attribute-template
图片上传 /open-api/goods/upload-pic
图片转换 /open-api/goods/transform-pic
发品权限 /open-api/goods/product/check-publish-permission
审核状态 /open-api/goods/query-document-state

官方 Feed 方案 https://open.sheincorp.com/documents/system/04fdafe8-9bb7-4fac-bf47-b1f45c859371
步骤：
1. /open-api/sem/feed/createFeedDocument（创建 JSON 文件）
2. /open-api/sem/feed/uploadDocumentContent（上传 JSON，<=100MB）
3. /open-api/sem/feed/getFeedDocument（可选检查）
4. /open-api/sem/feed/createFeed，feedType=PRODUCT_LISTING，version=V2
5. /open-api/sem/feed/getFeed（查询任务及结果文件，逐条解析成败）
任务状态：IN_QUEUE, IN_PROGRESS, FATAL, DONE, CANCELLED。
Feed DONE 是处理完毕，不应自动推断所有商品都审核通过或已上架。

## 授权与签名
官方合作流程：开放平台申请账号 -> 创建应用 -> 审核 -> 店铺授权 -> 接入解决方案。
调用请求包含 x-lt-openKeyId、x-lt-timestamp、x-lt-signature；须严格遵循官方签名算法和密钥轮换要求。密钥只保存在服务端密钥管理服务/环境变量，不在仓库或浏览器保存明文。

## 系统组成建议
前端：Next.js + TypeScript（中文控制台，商品表格、批次预览、修复）
后端：FastAPI + Pydantic（文件导入、映射、校验、店铺绑定、审批）
数据库：PostgreSQL；开发期可 SQLite
异步任务：Redis + Celery/RQ；批次和商品粒度幂等、退避重试
存储：私有对象存储（图片和导入文件，短时签名链接）
适配层：SheinOpenApiAdapter / SheinFeedAdapter / MockAdapter（默认 Mock）

## 数据模型
Shop(id, mode, site, auth_status, secret_ref)
Product(id, merchant_spu, category_id, title_en, desc_en, status)
Variant(id, product_id, merchant_skc, merchant_sku, attrs_json, price, currency, stock)
Asset(id, product_id, storage_ref, platform_asset_ref, status)
Batch(id, shop_id, import_file, snapshot_hash, operator, approved_at, status)
BatchItem(id, batch_id, product_id, idempotency_key, platform_spu, feed_id, status, trace_id, last_error)
AuditEvent(id, actor, action, resource, created_at, detail_json)

## 发布状态
IMPORTED -> INVALID / VALIDATED -> READY -> APPROVED -> QUEUED -> SUBMITTED -> UNDER_REVIEW -> APPROVED_ON_PLATFORM / REJECTED
单独跟踪 ON_SHELF / OFF_SHELF。网络超时后先查询平台，避免重复创建。

## 开发验收
- 100/1000 行文件导入可定位行错误
- 不上传即预检；发布需逐批次人工确认、店铺可发品权限检查
- 幂等、限流、重试、死信队列、审计日志和错误报告
- 测试覆盖 Mock 与真实 payload 合同测试
- 严禁绕过验证码、签名或授权；不要在 CI 使用真实密钥
