from __future__ import annotations

import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import TextBlock

from app import app
from core import create_total_table_report, parse_xlsx


def workbook_bytes(headers: list[str], rows: list[list[str]]) -> bytes:
    workbook = Workbook(); sheet = workbook.active; sheet.append(headers)
    for row in rows: sheet.append(row)
    stream = io.BytesIO(); workbook.save(stream); return stream.getvalue()


def master() -> bytes:
    return workbook_bytes(
        ["产品优惠券名称", "产品包含", "渠道", "优惠券ID", "有效期"],
        [["套餐 A", "鸡翅 1 份", "抖音", "SZKFCDY-1", "2026-10-01 至 2026-10-07"], ["", "", "美团", "SZKFCMT-1", ""]],
    )


def backend(price: str = "30") -> bytes:
    return workbook_bytes(
        ["商家商品id", "项目ID", "项目名称", "菜品备注", "券有效期", "美团价", "其他规则补充说明"],
        [["SZKFCMT-1", "P-1", "套餐 A", "鸡翅 1 份", "2026-10-01 至 2026-10-07", price, "规则 A"]],
    )


def test_total_table_uses_inherited_master_values_and_selected_fields() -> None:
    report, totals = create_total_table_report(master(), backend(), ["项目名称", "使用有效期", "菜品备注"])
    sheet = load_workbook(io.BytesIO(report))["核查结果"]
    assert sheet["A2"].value == "SZKFCMT-1"
    assert sheet["B2"].value == "P-1"
    assert [sheet.cell(2, column).value for column in (5, 8, 11)] == ["通过", "通过", "通过"]
    assert totals["通过"] == 1
    assert sheet.cell(2, sheet.max_column).value in (None, "")


def test_check_note_only_keeps_a_short_reason_for_non_passing_fields() -> None:
    export = workbook_bytes(
        ["商家商品id", "项目ID", "项目名称"],
        [["SZKFCMT-1", "P-1", "套餐 B"]],
    )
    report, totals = create_total_table_report(master(), export, ["项目名称"])
    sheet = load_workbook(io.BytesIO(report))["核查结果"]
    assert totals["不通过"] == 1
    assert sheet.cell(2, sheet.max_column).value == "项目名称：不通过。字段值不一致。"


def test_validity_can_be_extracted_from_backend_rule_description() -> None:
    master_data = workbook_bytes(
        ["优惠券ID", "有效期"],
        [["SZKFCMT-1", "10.24-11.1(除周四)"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "使用有效期：2026年10月24日-2026年11月1日内周一、周二、周三、周五、周六、周天使用有效。"]],
    )
    report, _ = create_total_table_report(master_data, backend_data, ["使用有效期"])
    sheet = load_workbook(io.BytesIO(report))["核查结果"]
    assert sheet["E2"].value == "通过"
    assert sheet["D2"].value == "10-24|11-01"


def test_missing_master_or_backend_is_review() -> None:
    extra_backend = workbook_bytes(["商家商品id", "项目ID", "项目名称", "菜品备注"], [["SZKFCMT-2", "P-2", "套餐 B", "鸡腿 1 份"]])
    report, totals = create_total_table_report(master(), extra_backend, ["项目名称"])
    sheet = load_workbook(io.BytesIO(report))["核查结果"]
    assert sheet["A2"].value == "SZKFCMT-2"
    assert sheet["C2"].value == "未匹配券ID"
    assert sheet["E2"].value == "待复核"
    assert totals["KFC总表未匹配"] == 1


def test_upload_endpoint_accepts_two_files_and_selected_fields() -> None:
    client = TestClient(app)
    response = client.post("/api/reconcile", data={"selected_fields": "项目名称,菜品备注"}, files={
        "master": ("KFC总表.xlsx", master()), "meituan_export": ("美团后台.xlsx", backend()),
    })
    assert response.status_code == 200, response.text
    assert client.get(response.json()["download"]).status_code == 200


def test_parser_accepts_data_rows_wider_than_the_header() -> None:
    source = workbook_bytes(["券ID", "渠道"], [["SZKFCMT-1", "美团", "后台多出的尾列"]])
    parsed = parse_xlsx(source)
    assert parsed.rows == [{"券ID": "SZKFCMT-1", "渠道": "美团"}]


def test_price_fields_remain_selectable() -> None:
    report, totals = create_total_table_report(master(), backend(), ["美团价"])
    sheet = load_workbook(io.BytesIO(report))["核查结果"]
    assert sheet["C2"].value is None
    assert sheet["E2"].value == "待复核"
    assert totals["待复核"] == 1


def test_price_formula_is_calculated_before_comparison() -> None:
    master_data = workbook_bytes(
        ["优惠券ID", "原价"],
        [["SZKFCMT-1", "113.00"], ["SZKFCMT-1", "=20*2+13.5*4+9.5+9.5"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目ID", "原价"],
        [["SZKFCMT-1", "P-1", "113.00"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["原价"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["C2"].value == "113.00"
    assert sheet["D2"].value == "113.00"
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_meituan_price_formula_is_calculated_before_comparison() -> None:
    master_data = workbook_bytes(
        ["优惠券ID", "活动价格"],
        [["SZKFCMT-1", "=20*2+13.5"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目ID", "美团价"],
        [["SZKFCMT-1", "P-1", "53.50"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["美团价"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["C2"].value == "53.50"
    assert sheet["D2"].value == "53.50"
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_coupon_id_match_compares_all_time_values_without_project_id() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "项目ID", "售卖开始时间", "售卖结束时间"],
        [
            ["SZKFCMT-1", "P-1", "2026-10-01 09:00:00", "2026-10-31 23:59:59"],
            ["SZKFCMT-1", "P-2", "2026-10-02 09:00:00", "2026-11-01 23:59:59"],
        ],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目ID", "售卖开始时间", "售卖结束时间"],
        [
            ["SZKFCMT-1", "P-9", "2026-10-02 13:30:00", "2026-11-01 23:59:59"],
            ["SZKFCMT-1", "P-8", "2026-10-01 12:00:00", "2026-10-31 23:59:59"],
        ],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["售卖开始时间", "售卖结束时间"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert sheet["H2"].value == "通过"
    assert sheet["C2"].value == "2026-10-01、2026-10-02"
    assert sheet["D2"].value == "2026-10-01、2026-10-02"
    assert totals["通过"] == 1


def test_report_only_lists_coupon_ids_from_meituan_export() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "售卖开始时间"],
        [["SZKFCMT-history", "2026-01-01 00:00:00"], ["SZKFCMT-live", "2026-10-01 00:00:00"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "售卖开始时间"],
        [["SZKFCMT-live", "2026-10-01 09:00:00"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["售卖开始时间"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet.max_row == 2
    assert sheet["A2"].value == "SZKFCMT-live"
    assert totals["通过"] == 1


def test_product_name_passes_when_coupon_dish_notes_match() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "项目名称", "菜品备注"],
        [["SZKFCMT-1", "套餐 A", "鸡翅 1 份"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目名称", "菜品备注"],
        [["SZKFCMT-1", "【美团】套餐 A", "鸡翅 1 份"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["项目名称"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert sheet["F2"].value == "通过"
    assert totals["通过"] == 1


def test_product_name_passes_when_only_test_display_tag_differs() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "项目名称"],
        [["SZKFCMT-1", "1支原味冰淇淋花筒"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目名称"],
        [["SZKFCMT-1", "【TEST】1支原味冰淇淋花筒"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["项目名称"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["C2"].value == "1支原味冰淇淋花筒"
    assert sheet["D2"].value == "【TEST】1支原味冰淇淋花筒"
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_product_name_passes_when_marketing_tag_and_duplicate_name_are_present() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "项目名称"],
        [
            ["SZKFCMT-1", "【最后一次6.9元】经典堡卷（3选1）阶梯10次卡"],
            ["SZKFCMT-1", "经典堡卷（3选1）阶梯10次卡"],
        ],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "项目名称"],
        [["SZKFCMT-1", "经典堡卷（3选1）阶梯10次卡"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["项目名称"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_empty_dish_notes_on_both_sides_do_not_require_review() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", ""]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", ""]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert sheet["F2"].value == "通过"
    assert totals["通过"] == 1


def test_dish_note_flavor_wording_is_treated_as_the_same_product() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "桂花风味酸梅汤（大）"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "桂花酸梅汤(大)"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_dish_note_allows_small_non_material_wording_difference() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "桂花酸梅汤（大杯）"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "桂花风味酸梅汤（大）"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_dish_note_matches_equivalent_combo_formatting() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "鸡肉帕尼尼/轻盈版鸡肉帕尼尼*4+豆浆（冷热）*2"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "产品包含：任选4份【1份芝士鸡肉帕尼尼/1份芝士鸡肉帕尼尼（轻盈版）】+任选2杯【1杯醇豆浆（冰）/1杯醇豆浆（热）】"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_duplicate_original_ice_cream_wording_is_treated_as_the_same_product() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "产品包含：1支原味冰淇淋原味花筒"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "1支原味冰淇淋花筒"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_nearly_identical_rule_wording_passes() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "肯德基在线客服，在线客服工作时间（9:00-22:00）。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "请咨询肯德基在线客服，在线客服工作时间（9:00-22:00）。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_rule_allows_equivalent_service_wording() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本券仅限堂食使用，客服工作时间为9:00-22:00。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "仅限到店消费，客服工作时间9:00-22:00。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_dish_and_rule_texts_are_checked_together() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注", "其他规则补充说明"],
        [["SZKFCMT-1", "鸡肉帕尼尼/轻盈版鸡肉帕尼尼*4+豆浆（冷热）*2", "仅限堂食使用。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注", "其他规则补充说明"],
        [["SZKFCMT-1", "产品包含：任选4份【1份芝士鸡肉帕尼尼/1份芝士鸡肉帕尼尼（轻盈版）】+任选2杯【1杯醇豆浆（冰）/1杯醇豆浆（热）】", "产品包含：任选4份【1份芝士鸡肉帕尼尼/1份芝士鸡肉帕尼尼（轻盈版）】+任选2杯【1杯醇豆浆（冰）/1杯醇豆浆（热）】；仅限到店消费。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["菜品备注", "其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert sheet["H2"].value == "通过"
    assert totals["通过"] == 1


def test_rule_ignores_product_title_preamble_and_exchange_wording() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "bkkp：2个黑糖珍珠圣代。本产品共1张兑换券，可兑换【1】份【美团】1块吮指原味鸡；产品包含：1份吮指原味鸡（1块装）。使用有效期：自购买之日起27天内有效；仅限堂食使用。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本产品共1张兑换券，可兑换1块吮指原味鸡，【美团】1块吮指原味鸡；产品包含：1份吮指原味鸡（1块装）。使用有效期：自购买之日起27天内有效；仅限到店消费。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_rule_passes_when_total_table_repeats_the_full_rule_block() -> None:
    rule = "本产品共1张兑换券。产品包含：1杯桂花酸梅汤（大）。使用有效期：自购买之日起7天内有效；仅限堂食使用。"
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", f"{rule}、{rule}"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本产品共1张兑换券。产品包含：1杯桂花风味酸梅汤（大）。使用有效期：自购买之日起7天内有效；仅限到店消费。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_rule_passes_when_same_coupon_has_duplicate_rule_rows() -> None:
    rule = "本产品共1张兑换券。产品包含：1杯桂花酸梅汤（大）。使用有效期：自购买之日起7天内有效；仅限堂食使用。"
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", rule], ["SZKFCMT-1", rule.replace("。", "；")]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本产品共1张兑换券。产品包含：1杯桂花风味酸梅汤（大）。使用有效期：自购买之日起7天内有效；仅限到店消费。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_rule_text_treats_osmanthus_flavor_drink_as_the_same_product() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "产品包含：1杯桂花风味酸梅汤（大）。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "产品包含：1杯桂花酸梅汤（大）。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "通过"
    assert totals["通过"] == 1


def test_material_rule_difference_is_not_accepted_as_similar() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本券自购买之日起7天内使用有效。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "本券自购买之日起30天内使用有效。"]],
    )
    report, totals = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True)["核查结果"]
    assert sheet["E2"].value == "不通过"
    assert totals["不通过"] == 1


def test_rule_difference_keeps_source_text_but_uses_a_short_note() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "仅限到店使用；客服工作时间9:00-22:00。"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "其他规则补充说明"],
        [["SZKFCMT-1", "仅限外送使用；客服工作时间9:00-22:00。"]],
    )
    report, _ = create_total_table_report(master_data, backend_data, ["其他规则补充说明"])
    sheet = load_workbook(io.BytesIO(report), data_only=True, rich_text=True)["核查结果"]
    assert str(sheet["C2"].value) == "仅限到店使用；客服工作时间9:00-22:00。"
    assert str(sheet["D2"].value) == "仅限外送使用；客服工作时间9:00-22:00。"
    assert sheet["G2"].value == "其他规则补充说明：不通过。文本内容不一致。"
    assert any(isinstance(part, TextBlock) and part.text == "到店" for part in sheet["C2"].value)


def test_dish_note_difference_keeps_source_text_but_uses_a_short_note() -> None:
    master_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "产品包含：1支原味冰淇淋花筒"]],
    )
    backend_data = workbook_bytes(
        ["商家商品id", "菜品备注"],
        [["SZKFCMT-1", "产品包含：2支原味冰淇淋花筒"]],
    )
    report, _ = create_total_table_report(master_data, backend_data, ["菜品备注"])
    sheet = load_workbook(io.BytesIO(report), data_only=True, rich_text=True)["核查结果"]
    assert str(sheet["C2"].value) == "产品包含：1支原味冰淇淋花筒"
    assert str(sheet["D2"].value) == "产品包含：2支原味冰淇淋花筒"
    assert sheet["G2"].value == "菜品备注：不通过。文本内容不一致。"
    assert any(isinstance(part, TextBlock) and part.text == "2" for part in sheet["D2"].value)
