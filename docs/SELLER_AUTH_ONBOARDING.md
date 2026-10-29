# 半托管卖家账号 -> 开发者应用接入清单

店铺模式：半托管（用户已确认）。
当前拥有：SHEIN 商家卖家账号。
尚未确认：是否已注册开发者账号、是否有开发者应用、是否是主账号、应用审核和权限状态。
生产发布：**禁止，直到店铺完成授权、必要接口权限开通、上传前显式确认**。

## 官方接入步骤
1. 访问 https://open.sheincorp.com/zh 注册/登录开发者账号（卖家账号不等同于开发者应用）。
2. 创建应用，应用类型明确选择 **半托管**；根据开放平台流程提交审核。
3. 确认并申请商品发布、商品类目与属性、图片处理、发品权限/额度、发布结果查询的接口权限。
4. 店铺**主账号**通过官方店铺授权链接授权应用（子账号不能授权）。
5. 后端接收一次性 tempToken（官方注明有效期 10 分钟），调用 POST /open-api/auth/get-by-token 交换 openKeyId 和加密的 secretKey，按官方算法安全解密、存储。严禁在浏览器或 GitHub 记录。
6. 半托管正式 API 域名 https://openapi.sheincorp.com；正式应用不得直接使用测试环境域名。授权后从只读接口开始验证再做模拟上传和小批量真实提交。
7. 发布接口 POST /open-api/goods/product/publishOrEdit。上线前与最新官方半托管商品发布方案逐一核对字段和类目要求。

## 参考（官方）
- 开发流程 https://open.sheincorp.com/zh
- 应用类型与授权 https://open.sheincorp.com/zh/documents/system/4f5b2c46-839d-4dd2-80e5-65722a060759
- 店铺授权 https://open.sheincorp.com/zh/documents/system/2169474d-1d4a-41a9-b9fd-427f63f54a63
- 密钥交换 https://open.sheincorp.com/zh/documents/apidoc/detail/3001235
- API 接入及正式域名 https://open.sheincorp.com/zh/documents/system/49b797df-016e-40eb-882b-ab9bbaa8c630
- 商品发布接口 https://open.sheincorp.com/documents/apidoc/detail/3001707-1000001

## Codex 最近的优先级
- 前端“店铺连接”页：显示 未配置/待审核/待授权/已授权/接口无权限/可发布。
- 安全授权回调后端、state 校验、临时 token 过期处理、密钥加密存储与轮换。
- 首先用 MockAdapter 完成商品数据模型、Excel 导入、预检、预览确认及批次审计。
- 仅明确授权后调用真实 API。不要把 SHEIN 核价理解成上传必过，也不要自作主张实现无授权的卖家后台浏览器模拟上传。
