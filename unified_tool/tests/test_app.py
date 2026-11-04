from __future__ import annotations

import io

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app import app


def workbook(headers: list[str], rows: list[list[str]]) -> bytes:
    book = Workbook()
    sheet = book.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def test_platforms_expose_the_two_existing_flows() -> None:
    client = TestClient(app)
    response = client.get("/api/platforms")
    assert response.status_code == 200
    assert set(response.json()) == {"douyin", "meituan"}


def test_douyin_reconcile_uses_the_original_two_file_flow() -> None:
    offer = workbook(
        ["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"],
        [["SZKFCDY-1", "【抖音】套餐 A", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"]],
    )
    export = workbook(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商品总价（商品划线价）"],
        [["套餐 A", "P-1", "SZKFCDY-1", "29.9", "59"]],
    )
    client = TestClient(app)
    sheets = client.post("/api/douyin/offer-sheets", files={"source": ("Offer.xlsx", offer)})
    assert sheets.status_code == 200, sheets.text
    assert sheets.json()["sheets"] == [{"name": "Sheet", "records": 1}]
    scored_sheets = client.post("/api/douyin/offer-sheets", files={"source": ("Offer.xlsx", offer), "platform_export": ("抖音导表.xlsx", export)})
    assert scored_sheets.json()["sheets"] == [{"name": "Sheet", "records": 1, "matches": 1}]
    response = client.post("/api/reconcile", data={"platform": "douyin", "selected_fields": "商品名称,售价", "offer_sheets": '["Sheet"]'}, files={"source": ("Offer.xlsx", offer), "platform_export": ("抖音导表.xlsx", export)})
    assert response.status_code == 200, response.text
    body = response.json()
    assert client.get(body["download"]).status_code == 200
    assert body["page_results"]["totals"] == {"通过": 1, "待复核": 0, "不通过": 0}
    assert body["page_results"]["rows"][0]["checks"] == [
        {"field": "商家平台券ID / 券ID", "local_value": "SZKFCDY-1", "backend_value": "SZKFCDY-1", "status": "通过"},
        {"field": "商品名称", "local_value": "【抖音】套餐 A", "backend_value": "套餐 A", "status": "通过"},
        {"field": "售价", "local_value": "29.9", "backend_value": "29.9", "status": "通过"},
    ]


def test_douyin_automatically_scans_all_offer_sheets() -> None:
    offer = workbook(["券ID", "产品名称", "L3", "活动价"], [["SZKFCDY-1", "套餐 A", "59", "29.9"]])
    export = workbook(["商品名称", "商品ID", "商家平台券ID", "售价"], [["套餐 A", "P-1", "SZKFCDY-1", "29.9"]])
    response = TestClient(app).post("/api/reconcile", data={"platform": "douyin", "selected_fields": "商品名称"}, files={"source": ("Offer.xlsx", offer), "platform_export": ("抖音导表.xlsx", export)})
    assert response.status_code == 200, response.text
    assert response.json()["page_results"]["totals"] == {"通过": 1, "待复核": 0, "不通过": 0}


def test_meituan_reconcile_uses_the_original_two_file_flow() -> None:
    master = workbook(
        ["产品优惠券名称", "产品包含", "渠道", "优惠券ID", "有效期"],
        [["套餐 A", "鸡翅 1 份", "抖音", "SZKFCDY-1", "2026-10-01 至 2026-10-07"], ["", "", "美团", "SZKFCMT-1", ""]],
    )
    export = workbook(
        ["商家商品id", "项目ID", "项目名称", "菜品备注", "券有效期"],
        [["SZKFCMT-1", "P-1", "套餐 A", "鸡翅 1 份", "2026-10-01 至 2026-10-07"]],
    )
    client = TestClient(app)
    response = client.post("/api/reconcile", data={"platform": "meituan", "selected_fields": "项目名称,菜品备注"}, files={"source": ("总表.xlsx", master), "platform_export": ("美团导表.xlsx", export)})
    assert response.status_code == 200, response.text
    body = response.json()
    assert client.get(body["download"]).status_code == 200
    assert body["page_results"]["totals"] == {"通过": 1, "待复核": 0, "不通过": 0}


def test_page_results_normalizes_douyin_error_as_not_passed() -> None:
    offer = workbook(
        ["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"],
        [["SZKFCDY-1", "【抖音】套餐 A", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"]],
    )
    export = workbook(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商品总价（商品划线价）"],
        [["套餐 A", "P-1", "SZKFCDY-1", "19.9", "59"]],
    )
    response = TestClient(app).post("/api/reconcile", data={"platform": "douyin", "selected_fields": "售价", "offer_sheets": '["Sheet"]'}, files={"source": ("Offer.xlsx", offer), "platform_export": ("抖音导表.xlsx", export)})
    assert response.status_code == 200, response.text
    assert response.json()["page_results"]["totals"] == {"通过": 0, "待复核": 0, "不通过": 1}


