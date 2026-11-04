from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from core import SIMPLE_FIELDS, create_total_table_report

ROOT = Path(__file__).parent
OUTPUTS = ROOT / "data" / "outputs"; OUTPUTS.mkdir(parents=True, exist_ok=True)
app = FastAPI(title="KFC 美团素材核查工具")


def read_xlsx(upload: UploadFile) -> bytes:
    if Path(upload.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, "仅支持 XLSX 或 XLSM 文件。")
    content = upload.file.read()
    if not content: raise HTTPException(400, "上传文件不能为空。")
    return content


@app.get("/", response_class=HTMLResponse)
def home() -> str:
    return (ROOT / "web" / "index.html").read_text(encoding="utf-8")


@app.get("/api/fields")
def fields():
    return {"fields": SIMPLE_FIELDS}


@app.post("/api/reconcile")
def reconcile(
    master: UploadFile = File(...),
    meituan_export: UploadFile = File(...),
    selected_fields: str = Form(""),
):
    try:
        selected = [field for field in selected_fields.split(",") if field]
        report, totals = create_total_table_report(
            read_xlsx(master), read_xlsx(meituan_export), selected,
            master_filename=Path(master.filename or "KFC总表").name,
            export_filename=Path(meituan_export.filename or "美团后台导表").name,
        )
    except HTTPException:
        raise
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        # 始终返回 JSON，避免前端将 HTML 500 页面误报为 JSON 语法错误。
        raise HTTPException(400, f"文件读取或核对失败：{type(error).__name__}：{error}") from error
    filename = "美团审核结果表.xlsx"
    (OUTPUTS / filename).write_bytes(report)
    return {"filename": filename, "download": f"/api/outputs/{filename}", "totals": totals}


@app.get("/api/outputs/{filename}")
def download(filename: str):
    path = OUTPUTS / Path(filename).name
    if not path.is_file(): raise HTTPException(404, "结果文件不存在。")
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
