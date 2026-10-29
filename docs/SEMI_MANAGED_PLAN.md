# SHEIN 半托管适配方案（预确认）

状态：用户经营方式高度符合半托管，但必须通过商家后台或 SHEIN 官方确认店铺类型；所有生产上传默认关闭。

## 官方定义与核验
半托管：SHEIN 定价、商家不必备货到 SHEIN 仓、商家自行履约。来源：https://open.sheincorp.com/zh/documents/system/7f423ab6-84d4-497c-95c7-cc7bbed5b190

核验路径：
- 店铺后台：我的档案/店铺信息/合作模式（具体菜单以后台实际界面为准）
- 官方开放平台应用类型与可授权商家范围： https://open.sheincorp.com/zh/documents/system/4f5b2c46-839d-4dd2-80e5-65722a060759
- 授权后 POST /open-api/openapi-business-backend/query-store-info 可查店铺基本信息，但 supplierBusinessMode 仅特殊商家返回，不可用字段为空作为半托管确证。

## 必接接口（按官方文档重新确认）
- GET /open-api/goods/product/check-publish-permission —— 发品资格
- POST /open-api/goods-goods-publish-quotas/detail 请勿使用此拼写；**正确路径** /open-api/goods-publish-quotas/detail —— 发品额度；确认 HTTP 方法及权限后接入
- POST /open-api/goods/query-publish-fill-in-standard —— 动态字段规范、语言、供货币种、图片支持
- /open-api/goods/query-category-tree —— 可售类目、product_type_id
- /open-api/goods/query-attribute-template —— 类目属性
- /open-api/goods/query-site-list —— 可售站点
- /open-api/goods/upload-pic、/open-api/goods/transform-pic —— 图片
- POST /open-api/goods/product/publishOrEdit —— 新发 SPU、增添 SKC/SKU、更新已有商品
- /open-api/goods/query-document-state —— 审核状态
- /open-api/goods/update-cost —— 审核通过后的供货价调整

官方：https://open.sheincorp.com/zh/documents/apidoc/detail/3001352-1000001

## 半托管区别
- 新商品/新 SKU 必须提供 cost_info (cost_price,currency)，不应把 price 直接当成消费者零售价。
- 币种必须按发布字段规范返回值选择。
- SKU 包装长宽高、重量等按官方校验规则提供。
- 半托管商品新发需 shelf_way，可能涉及定时上架。
- 支持多销售站点，发布时 site_list，站点由 query-site-list 获取。
- 2026-08 起，自运营及半托管发品额度限制到 SKC 层级；预检权限和剩余额度。
- 商品文案不支持 HTML 和 emoji，禁词检查为附加的用户规则。
- 平台处理完成不等于审核通过、上架或核价通过。

## Codex 开发顺序
1. 店铺模式 UI 标记 SEMI_MANAGED_UNVERIFIED；接入官方授权后才能 VERIFIED。
2. Excel/CSV 导入的数据模型增加 cost_price、cost_currency、包裹尺寸重量、main_site、sub_site_list、shelf_way、供货 SKU。
3. 按类目/店铺动态字段规范校验；不硬编码“所有女装通用必填”。
4. Secret 管理、签名、API 客户端（mock first）；开发时不写入真实密钥。
5. 用户确认商品和店铺后，可执行单品 publishOrEdit；异步审核回查、错误报告、幂等保护。
6. 上述稳定后接入 Feed 批量 API（另行按官方接口文档验证版本和申请权限）。

## 明确未完成
目前仅规划与客户端字段契约，无官方授权、签名实现、生产上传权限、后台页面或完整发布 payload。
