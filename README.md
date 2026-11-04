# KFC 素材核查助手

用于核对 KFC 总表、抖音 Offer 表与平台后台导表中的券素材信息。系统提供美团、抖音两个平台入口，按券 ID 关联数据并生成可下载的 Excel 核查结果。

> 本项目不接入大模型、不调用 OpenAI API，也不需要配置 API Key。所有核对均由服务端的确定性规则完成。

## 功能

- 美团与抖音统一入口，分别保留各自的既有核对规则。
- 上传两份 Excel 表后，按券 ID 比对商品名称、售卖时间、有效期、价格、规则说明、菜品备注等字段。
- 输出通过、待复核、不通过三种结论；价格、时间、有效期等字段归为高风险。
- 页面支持按券 ID、商品 ID、字段名称和双方取值搜索，并可按风险等级筛选。
- 对文本字段提供合理的格式容错，例如平台标签、营销前缀、部分菜品描述差异、重复规则说明等。
- 对不通过的菜品备注和规则说明标记差异文本，便于人工进一步核对。
- 自动生成带冻结表头、筛选和状态颜色的 Excel 文件。

## 核对资料

| 平台 | 第一份资料 | 第二份资料 |
| --- | --- | --- |
| 美团 | KFC 总表 | 美团后台导表 |
| 抖音 | 抖音本地 Offer 表 | 抖音后台商品明细 |

原始上传表不会被修改。生成的结果文件保存到 `unified_tool/data/outputs/`；Docker 部署时会持久化到 `runtime/outputs/`。

## 项目结构

```text
unified_tool/   # 统一 Web 页面和 API 入口
meituan_tool/   # 美团核对规则与 Excel 导出
douyin_tool/    # 抖音核对规则与 Excel 导出
Dockerfile      # 生产镜像
compose.yaml    # Docker Compose 配置
```

## 本地启动

需要 Python 3.12 或兼容版本。

```bash
cd unified_tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload --port 8788
```

浏览器访问：<http://127.0.0.1:8788>

## Docker 部署

服务器项目根目录执行：

```bash
docker compose up -d --build
docker compose ps
```

默认对外端口为 `8790`，访问地址为：

```text
http://服务器地址:8790
```

详细步骤见 [DOCKER-DEPLOY.md](DOCKER-DEPLOY.md)。

## 测试

```bash
cd meituan_tool
python3 -m pytest tests -q

cd ../unified_tool
python3 -m pytest tests -q
```
