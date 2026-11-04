from __future__ import annotations

import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app import app
from core import FIELDS, core_name, make_report, offer_sheet_options


def book(headers: list[str], rows: list[list[str]], title: str = "Sheet1") -> bytes:
    workbook = Workbook(); sheet = workbook.active; sheet.title = title; sheet.append(headers)
    for row in rows: sheet.append(row)
    stream = io.BytesIO(); workbook.save(stream); return stream.getvalue()


def offer() -> bytes:
    return book(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"], [["SZKFCDY-1", "【抖音】套餐 A", "59", "29.9", "T+7\n限09:30-23:00使用", "2026-10-01至2026-10-07"]], "Y26-10月")


def backend() -> bytes:
    return book(["商品名称", "商品ID", "商家平台券ID", "售价", "商品总价（商品划线价）", "售卖开始时间", "售卖结束时间", "商品类型", "每日可用时段"], [["套餐 A", "P-1", "SZKFCDY-1", "29.9元", "59元", "2026-10-01", "2026-10-07", "团购", "0930-2300"]], "商品列表-1")


def test_two_file_report_matches_all_requirement_fields() -> None:
    report, totals = make_report(offer(), backend(), FIELDS)
    workbook = load_workbook(io.BytesIO(report)); result = workbook["核对结果表"]
    assert workbook.sheetnames == ["核对结果表", "异常汇总"]
    assert result["C1"].value == "业务本地表格-券ID"
    assert result["D1"].value == "商家平台券ID-结果"
    assert [result["E1"].value, result["F1"].value] == ["抖音后台表格-商品名称", "业务本地表格-产品名称"]
    assert result["A2"].value == "SZKFCDY-1"
    assert result["B2"].value == "P-1"
    assert result["D2"].value == "通过"
    assert [result.cell(2, i).value for i in range(7, 25, 3)] == ["通过"] * 6
    assert result["W2"].value == "通过"
    assert totals == {"通过": 1, "待复核": 0, "不通过": 0}


def test_duplicate_offer_requires_unique_plain_name_record() -> None:
    duplicate = book(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"], [["SZKFCDY-1", "套餐 A【人气】", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"], ["SZKFCDY-1", "套餐 A【美食局】", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"]])
    report, totals = make_report(duplicate, backend(), ["售价"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["G2"].value == "待复核"
    assert "不唯一" in result["I2"].value
    assert totals["待复核"] == 1


def test_offer_name_fallback_requires_a_unique_core_meal_name() -> None:
    fallback = book(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"], [["SZKFCDY-2", "【抖音】套餐 A", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"]])
    report, totals = make_report(fallback, backend(), ["商品名称", "售价"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert [result["G2"].value, result["J2"].value] == ["通过", "通过"]
    assert result["D2"].value == "不通过"
    assert result["K2"].value == "不通过"
    assert totals["不通过"] == 1


def test_selected_offer_sheet_excludes_historical_duplicate() -> None:
    workbook = Workbook(); current = workbook.active; current.title = "Y26-10月"
    current.append(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"])
    current.append(["SZKFCDY-1", "【抖音】套餐 A", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"])
    historical = workbook.create_sheet("Y25-10月")
    historical.append(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"])
    historical.append(["SZKFCDY-1", "套餐 A【人气】", "59", "29.9", "09:30-23:00", "2025-10-01至2025-10-07"])
    stream = io.BytesIO(); workbook.save(stream)

    assert offer_sheet_options(stream.getvalue()) == [
        {"name": "Y26-10月", "records": 1}, {"name": "Y25-10月", "records": 1}
    ]
    report, totals = make_report(stream.getvalue(), backend(), ["售价"], offer_sheets=["Y26-10月"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["H2"].value == "通过"
    assert totals["通过"] == 1


def test_duplicate_coupon_uses_unique_business_field_evidence() -> None:
    duplicate = book(
        ["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"],
        [["SZKFCDY-1", "套餐 B【人气】", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"], ["SZKFCDY-1", "套餐 A【美食局】", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"]],
    )
    report, totals = make_report(duplicate, backend(), ["商品名称"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["H2"].value == "通过"
    assert totals["通过"] == 1


def test_duplicate_coupon_with_equivalent_offer_rows_is_merged() -> None:
    workbook = Workbook(); current = workbook.active; current.title = "Y26-10月"
    historical = workbook.create_sheet("Y26-11月")
    for sheet in (current, historical):
        sheet.append(["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"])
        sheet.append(["SZKFCDY-1", "【抖音】套餐 A", "59", "29.9", "09:30-23:00", "2026-10-01至2026-10-07"])
    stream = io.BytesIO(); workbook.save(stream)

    report, totals = make_report(stream.getvalue(), backend(), ["售价", "商品总价（商品划线价）"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["K2"].value == "通过"
    assert totals["通过"] == 1


def test_product_name_ignores_platform_display_labels_but_keeps_meal_content() -> None:
    assert core_name("【到手6件】元气随心配") == core_name("元气随心配\n工作日39.9元，周末43.9元")
    assert core_name("东方四黄粥油条2件套【食货局·肯德基早餐】") == core_name("东方四黄粥油条二件套")
    assert core_name("原味鸡青年1块") == core_name("1块原味鸡青年")
    assert core_name("套餐 A【区域上架】") == core_name("套餐 A【门店限定】")
    assert core_name("套餐 A") != core_name("套餐 B")


def test_activity_price_accepts_backend_promo_price() -> None:
    priced_backend = book(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商促价", "商品总价（商品划线价）"],
        [["套餐 A", "P-1", "SZKFCDY-1", "59", "29.9", "59"]],
    )
    report, totals = make_report(offer(), priced_backend, ["售价"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["G2"].value == "通过"
    assert totals["通过"] == 1


def test_tbd_listing_keeps_known_start_date_comparable() -> None:
    local = book(
        ["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"],
        [["SZKFCDY-1", "套餐 A", "59", "29.9", "09:30-23:00", "10/12-TBD"]],
    )
    timed_backend = book(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商品总价（商品划线价）", "售卖开始时间", "售卖结束时间"],
        [["套餐 A", "P-1", "SZKFCDY-1", "29.9", "59", "2026-10-12", "2026-10-12"]],
    )
    report, totals = make_report(local, timed_backend, ["售卖开始时间", "售卖结束时间"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["G2"].value == "通过"
    assert result["J2"].value == "通过"
    assert totals == {"通过": 1, "待复核": 0, "不通过": 0}


def test_any_price_match_between_multiple_prices_passes() -> None:
    local = book(
        ["券ID", "产品名称", "L3", "活动价", "使用效期", "上架时间"],
        [["SZKFCDY-1", "套餐 A", "59", "工作日39.9元，周末43.9元", "09:30-23:00", "2026-10-01至2026-10-07"]],
    )
    priced_backend = book(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商促价"],
        [["套餐 A", "P-1", "SZKFCDY-1", "59", "43.9"]],
    )
    report, totals = make_report(local, priced_backend, ["售价"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result["G2"].value == "通过"
    assert totals == {"通过": 1, "待复核": 0, "不通过": 0}


def test_one_coupon_with_multiple_product_ids_is_merged_into_one_result_row() -> None:
    multi_item_backend = book(
        ["商品名称", "商品ID", "商家平台券ID", "售价", "商品总价（商品划线价）"],
        [["套餐 A", "P-1", "SZKFCDY-1", "29.9", "59"], ["套餐 A", "P-2", "SZKFCDY-1", "29.9", "59"]],
    )
    report, totals = make_report(offer(), multi_item_backend, ["商品名称", "售价"])
    result = load_workbook(io.BytesIO(report))["核对结果表"]
    assert result.max_row == 2
    assert result["A2"].value == "SZKFCDY-1"
    assert result["B2"].value == "P-1、P-2"
    assert [result["G2"].value, result["J2"].value] == ["通过", "通过"]
    assert totals == {"通过": 1, "待复核": 0, "不通过": 0}


def test_upload_endpoint_accepts_two_files() -> None:
    client = TestClient(app)
    response = client.post("/api/reconcile", data={"selected_fields": "商品名称,商品类型"}, files={"offer": ("Offer.xlsx", offer()), "douyin_export": ("后台.xlsx", backend())})
    assert response.status_code == 200, response.text
    assert client.get(response.json()["download"]).status_code == 200
