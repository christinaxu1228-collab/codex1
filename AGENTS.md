# Codex 开发指南

## 目标
开发面向 SHEIN 女装店铺的批量上品管理系统。优先实现可靠、可测试、可审计的流程，不虚构 SHEIN API 或审批规则。

## 已有功能
Python 标准库 CLI：读取 CSV -> 校验 -> 输出报告。执行入口：`python -m bulk_listing validate examples/products.csv`。

## 下一阶段任务（按顺序）
1. 基于官方可核验的最新模板规范设计 product/variant/image 数据模型，支持 Excel 工作簿导入。
2. 增加文件上传、批次预览、字段映射、校验结果下载的简单 Web UI。
3. 增加 SQLite/PostgreSQL 持久化，批次状态、幂等键、任务队列、失败重试。
4. 定义 PlatformAdapter 接口，添加 MockAdapter 和受控的上传集成测试。
5. 获得真实官方 SHEIN API 文档、应用权限、店铺授权后，才实现 Production SHEIN Adapter。

## 产品约束
- 标题和描述仅使用英文商品内容；不得含用户指定禁词 trendy、seguro（大小写不敏感）。
- 不根据本地校验结果承诺平台审核或核价通过。
- 所有真实上架动作必须经过显式确认；默认 dry-run。
- 不抓取或绕过验证码、风控、登录限制；不在仓库中写入密钥。
- 更新代码时补充测试和 README，保持命令可运行。
