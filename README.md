# KFC 素材核查 API 服务

这是部署在 Sites Worker 上的真实审核服务，而非浏览器端演示数据。

## 请求流程

1. 用户选择抖音或美团，并上传主导表、对照表、资料文件。
2. 浏览器将三份文件以 `multipart/form-data` 发送到同源 `POST /api/reconcile`。
3. Worker 将文件以 `user_data` 用途上传到 OpenAI Files API，再使用 Responses API 的 `input_file` 传入模型。
4. 模型只能按 JSON Schema 返回券 ID、字段状态、原因和来源；Worker 会补齐缺失字段、拒绝重复券 ID，并重算整券状态。
5. Worker 生成包含冻结表头、筛选和状态字体颜色的 `.xlsx`，以 Base64 连同审核摘要返回浏览器下载。
6. 审核结束后，Worker 会请求删除本次上传到 OpenAI 的临时文件。

## 运行时变量

在 Sites 的运行时环境变量中设置：

- `OPENAI_API_KEY`（密钥，必填）
- `OPENAI_MODEL`（可选；默认 `gpt-4.1`）

不要将密钥写进仓库、前端或 `.openai/hosting.json`。

## 约束

- 每个文件最大 16MB，三份文件合计最大 30MB。
- 支持 `.xlsx`、`.xls`、`.csv` 和需要文字核对时使用的 `.docx`。
- 主导表中没有的券 ID 不会被输出；模型无法可靠关联或判别时，服务端将其标为“待复核”。
