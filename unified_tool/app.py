"""抖音、美团共用入口；各平台仍使用原有的核查规则。"""
from __future__ import annotations

import io
import importlib.util
import json
import sys
from datetime import datetime
from pathlib import Path
from types import ModuleType

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from openpyxl import load_workbook

ROOT = Path(__file__).parent
PROJECT_ROOT = ROOT.parent
OUTPUTS = ROOT / "data" / "outputs"
OUTPUTS.mkdir(parents=True, exist_ok=True)


def load_module(name: str, path: Path) -> ModuleType:
    """加载既有平台核心逻辑，避免两份同名 core.py 相互覆盖。"""
    spec = importlib.util.spec_from_file_location(name, path)
    if not spec or not spec.loader:
        raise RuntimeError(f"无法加载平台核查逻辑：{path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


DOUYIN = load_module("kfc_douyin_core", PROJECT_ROOT / "douyin_tool" / "core.py")
MEITUAN = load_module("kfc_meituan_core", PROJECT_ROOT / "meituan_tool" / "core.py")

PLATFORMS = {
    "douyin": {
        "label": "抖音",
        "source_label": "抖音本地 Offer 表",
        "source_hint": "包含“券ID”、产品名、活动价、L3 和上架时间",
        "export_label": "抖音后台商品明细",
        "export_hint": "包含商品ID和“商家平台券ID”",
        "fields": DOUYIN.FIELDS,
        # 按抖音后台导表确认的核对范围，券ID作为唯一匹配键，不重复勾选。
        "default_fields": DOUYIN.FIELDS,
    },
    "meituan": {
        "label": "美团",
        "source_label": "KFC 总表",
        "source_hint": "包含“优惠券ID”或“券ID”",
        "export_label": "美团后台导表",
        "export_hint": "包含“商家商品id”",
        "fields": MEITUAN.SIMPLE_FIELDS,
    },
}

app = FastAPI(title="KFC 素材核查助手")


def read_xlsx(upload: UploadFile, label: str) -> bytes:
    if Path(upload.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, f"{label}仅支持 XLSX 或 XLSM 文件。")
    content = upload.file.read()
    if not content:
        raise HTTPException(400, f"{label}不能为空。")
    return content


def display_status(value: object) -> str:
    """统一页面的失败状态；抖音旧规则沿用“错误”字样。"""
    return "不通过" if value == "错误" else str(value or "待复核")


def page_results(report: bytes, platform: str, selected_fields: list[str]) -> dict[str, object]:
    """从已生成的结果 Excel 提取轻量页面摘要，避免重新实现平台核对规则。"""
    workbook = load_workbook(io.BytesIO(report), read_only=True, data_only=True)
    sheet = workbook["核对结果表" if platform == "douyin" else "核查结果"]
    headers = [str(cell or "") for cell in next(sheet.iter_rows(values_only=True))]
    positions = {header: index for index, header in enumerate(headers)}
    note_header = "核对说明" if platform == "meituan" else "异常原因"
    rows: list[dict[str, object]] = []

    if platform == "douyin":
        for values in sheet.iter_rows(min_row=2, values_only=True):
            coupon_id = str(values[positions["商家平台券ID"]] or "")
            if not coupon_id:
                continue
            checks = [{
                "field": "商家平台券ID / 券ID",
                "local_value": str(values[positions["业务本地表格-券ID"]] or ""),
                "backend_value": coupon_id,
                "status": display_status(values[positions["商家平台券ID-结果"]]),
            }, *[
                {
                    "field": field,
                    "local_value": str(values[positions[f"业务本地表格-{DOUYIN.offer_output_field_name(field)}"]] or ""),
                    "backend_value": str(values[positions[f"抖音后台表格-{field}"]] or ""),
                    "status": display_status(values[positions[f"{field}-结果"]]),
                }
                for field in selected_fields
            ]]
            rows.append({
                "coupon_id": coupon_id,
                "platform_id": str(values[positions["商品ID"]] or ""),
                "overall": display_status(values[positions["整券结果"]]),
                "checks": checks,
                "reason": str(values[positions[note_header]] or ""),
            })
    else:
        for values in sheet.iter_rows(min_row=2, values_only=True):
            coupon_id = str(values[positions["券ID"]] or "")
            if not coupon_id:
                continue
            checks = [
                {
                    "field": field,
                    "local_value": str(values[positions[f"KFC总表-{field}"]] or ""),
                    "backend_value": str(values[positions[f"美团后台-{field}"]] or ""),
                    "status": display_status(values[positions[f"{field}-结果"]]),
                }
                for field in selected_fields
            ]
            rows.append({
                "coupon_id": coupon_id,
                "coupon_label": coupon_id,
                "platform_id": str(values[positions["美团ID"]] or ""),
                "overall": display_status(values[positions["整券结果"]]),
                "checks": checks,
                "reason": str(values[positions[note_header]] or ""),
            })

    totals = {status: sum(row["overall"] == status for row in rows) for status in ("通过", "待复核", "不通过")}
    return {
        "totals": totals,
        "rows": rows,
        "source_label": PLATFORMS[platform]["source_label"],
        "export_label": PLATFORMS[platform]["export_label"],
    }


@app.get("/", response_class=HTMLResponse)
def home() -> str:
    return (ROOT / "web" / "index.html").read_text(encoding="utf-8")


@app.get("/api/platforms")
def platforms():
    return {key: {name: value for name, value in config.items() if name != "fields"} for key, config in PLATFORMS.items()}


@app.get("/api/platforms/{platform}/fields")
def fields(platform: str):
    config = PLATFORMS.get(platform)
    if not config:
        raise HTTPException(404, "仅支持抖音或美团。")
    return {"fields": config["fields"], "default_fields": config.get("default_fields", config["fields"])}


@app.post("/api/douyin/offer-sheets")
def douyin_offer_sheets(source: UploadFile = File(...), platform_export: UploadFile | None = File(None)):
    content = read_xlsx(source, "抖音本地 Offer 表")
    try:
        export_content = read_xlsx(platform_export, "抖音后台商品明细") if platform_export and platform_export.filename else None
        sheets = DOUYIN.offer_sheet_options(content, export_content)
    except Exception as error:
        raise HTTPException(400, f"Offer 工作表读取失败：{type(error).__name__}：{error}") from error
    if not sheets:
        raise HTTPException(400, "未找到同时包含券ID、产品名称和活动价或 L3 的 Offer 工作表。")
    return {"sheets": sheets}


@app.post("/api/reconcile")
def reconcile(
    platform: str = Form(...),
    source: UploadFile = File(...),
    platform_export: UploadFile = File(...),
    selected_fields: str = Form(...),
    offer_sheets: str = Form(""),
):
    config = PLATFORMS.get(platform)
    if not config:
        raise HTTPException(400, "仅支持抖音或美团。")
    chosen = [field for field in selected_fields.split(",") if field]
    if not chosen:
        raise HTTPException(400, "请至少选择一个核对字段。")
    try:
        selected_offer_sheets = json.loads(offer_sheets) if offer_sheets else []
    except json.JSONDecodeError as error:
        raise HTTPException(400, "Offer 工作表选择格式无效，请重新选择。") from error
    if not isinstance(selected_offer_sheets, list) or not all(isinstance(name, str) and name.strip() for name in selected_offer_sheets):
        raise HTTPException(400, "Offer 工作表选择格式无效，请重新选择。")
    selected_offer_sheets = list(dict.fromkeys(name.strip() for name in selected_offer_sheets))
    try:
        source_content = read_xlsx(source, config["source_label"])
        export_content = read_xlsx(platform_export, config["export_label"])
        if platform == "douyin":
            report, totals = DOUYIN.make_report(
                source_content, export_content, chosen,
                offer_filename=Path(source.filename or config["source_label"]).name,
                backend_filename=Path(platform_export.filename or config["export_label"]).name,
                offer_sheets=selected_offer_sheets or None,
            )
        else:
            report, totals = MEITUAN.create_total_table_report(
                source_content, export_content, chosen,
                master_filename=Path(source.filename or config["source_label"]).name,
                export_filename=Path(platform_export.filename or config["export_label"]).name,
            )
    except HTTPException:
        raise
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        raise HTTPException(400, f"文件读取或核对失败：{type(error).__name__}：{error}") from error

    filename = f"{config['label']}素材核查结果-{datetime.now():%Y%m%d-%H%M%S}.xlsx"
    (OUTPUTS / filename).write_bytes(report)
    return {
        "filename": filename,
        "download": f"/api/outputs/{filename}",
        "totals": totals,
        "page_results": page_results(report, platform, chosen),
    }


@app.get("/api/outputs/{filename}")
def download(filename: str):
    path = OUTPUTS / Path(filename).name
    if not path.is_file():
        raise HTTPException(404, "结果文件不存在。")
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


