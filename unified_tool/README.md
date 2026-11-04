# KFC 素材核查助手

抖音和美团共用一个页面入口。选择平台后，页面会显示该平台原有的两份资料要求与可核查字段；实际生成报告时仍分别调用原来的 `douyin_tool/core.py` 和 `meituan_tool/core.py`，不改变两套核对规则。

## 启动

```bash
cd unified_tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload --port 8788
```

打开 `http://127.0.0.1:8788`。

## 平台资料

| 平台 | 第一份资料 | 第二份资料 |
| --- | --- | --- |
| 抖音 | 抖音本地 Offer 表 | 抖音后台商品明细 |
| 美团 | KFC 总表 | 美团后台导表 |

抖音核对会自动扫描 Offer 中的有效工作表。相同券 ID 出现在多页时，系统会用商品名称、活动价、L3 等字段做唯一确认；证据不足时保留为待复核，不会任意选择历史记录。

美团页面保留“美团价”和“原价”核查；后台可使用售价、商促价、划线价等可明确对应的列。由于价格比例需要两项均有可靠来源，页面不提供“原价/美团价”核查项。

结果文件保存在 `unified_tool/data/outputs/`，原始上传表不会被修改。

