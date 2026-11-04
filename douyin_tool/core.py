from __future__ import annotations

import io
import re
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# 页面和结果表保留抖音后台导表中的原始字段名称。商家平台券ID为唯一匹配键，
# 已作为首列展示，因此不再作为重复的可选核对字段。
FIELDS = ["商品名称", "售价", "商品总价（商品划线价）", "售卖开始时间", "售卖结束时间", "商品类型"]
STATUS_COLORS = {"通过": "008000", "待复核": "FFC000", "不通过": "FF0000"}
TBD_END_DATE = datetime(2026, 12, 31)
BACKEND = {"id": ["商家平台券ID", "商家平台券id"], "item": ["商品ID", "商品id"], "商品名称": ["商品名称"], "券ID": ["商家平台券ID", "商家平台券id"], "活动价": ["售价"], "L3": ["商品总价（商品划线价）", "商品总价(商品划线价)", "划线价"], "售卖开始时间": ["售卖开始时间"], "售卖结束时间": ["售卖结束时间"], "商品类型": ["商品类型"], "每日可用时段": ["每日可用时段"]}
OFFER = {"id": ["券ID", "券id"], "商品名称": ["产品名称", "产品名", "产品名字", "券名"], "券ID": ["券ID", "券id"], "活动价": ["活动价", "活动价格"], "L3": ["L3", "L3划线价", "前端L3\n划线价", "前端L3 / 划线价"], "售卖开始时间": ["上架时间", "上架时间\n（是否预售）", "上架时间 / （是否预售）"], "售卖结束时间": ["上架时间", "上架时间\n（是否预售）", "上架时间 / （是否预售）"], "商品类型": ["产品名称", "产品名", "产品名字", "券名"], "每日可用时段": ["使用效期"]}
FIELD_MAPPING = {
    "商品名称": ("商品名称", "产品名称"),
    "售价": ("活动价", "活动价"),
    "商品总价（商品划线价）": ("L3", "L3"),
    "售卖开始时间": ("售卖开始时间", "上架时间"),
    "售卖结束时间": ("售卖结束时间", "上架时间"),
    "商品类型": ("商品类型", "产品名称"),
}


def text(raw: Any) -> str:
    if raw is None: return ""
    if isinstance(raw, datetime): return raw.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(raw, float) and raw.is_integer(): return str(int(raw))
    return str(raw).strip()


def norm(raw: Any) -> str: return re.sub(r"\s+", "", text(raw)).replace("．", ".")


def find_header(row: dict[str, str], candidates: list[str]) -> str | None:
    choices = {norm(key).lower(): key for key in row}
    return next((choices[norm(item).lower()] for item in candidates if norm(item).lower() in choices), None)


def value(row: dict[str, str], candidates: list[str]) -> str: return text(row.get(find_header(row, candidates) or "", ""))


def coupon_ids(raw: str) -> list[str]: return list(dict.fromkeys(item.upper() for item in re.findall(r"SZKFCDY[A-Za-z0-9_-]+", text(raw), flags=re.I)))


def money(raw: str) -> str:
    found = re.search(r"-?\d+(?:\.\d+)?", text(raw).replace(",", ""))
    return f"{float(found.group()):.2f}" if found else ""


def money_values(raw: str) -> set[str]:
    """提取一个单元格内的全部价格，用于多售价的任一值匹配。"""
    return {f"{float(item):.2f}" for item in re.findall(r"-?\d+(?:\.\d+)?", text(raw).replace(",", ""))}


def time_range(raw: str) -> str:
    found = re.search(r"(?<!\d)(\d{1,2}):?(\d{2})\s*[-~至]\s*(次日\s*)?(\d{1,2}):?(\d{2})(?!\d)", text(raw))
    if not found: return ""
    next_day = "次日" if found.group(3) else ""
    return f"{int(found.group(1)):02d}:{found.group(2)}-{next_day}{int(found.group(4)):02d}:{found.group(5)}"


def dates(raw: str, year: int | None = None) -> list[datetime]:
    result = [datetime(int(y), int(m), int(d)) for y, m, d in re.findall(r"(20\d{2})[年/-](\d{1,2})[月/-](\d{1,2})", text(raw))]
    if result or not year: return result
    return [datetime(year, int(m), int(d)) for m, d in re.findall(r"(?<!\d)(\d{1,2})[月/.-](\d{1,2})(?:日)?", text(raw))]


def year_from(raw: str) -> int | None:
    found = re.search(r"(20\d{2})", text(raw))
    return int(found.group(1)) if found else None


def core_name(raw: str) -> str:
    """Return the stable product title, excluding platform-facing display labels.

    抖音后台常在同一套餐名称前后加渠道、配送、直播或件数营销标签；
    这些标签不改变券的实际套餐内容，不能作为不通过的理由。商品内容、
    数量和未知标签则保留，避免把真正不同的套餐误合并。
    """
    display_terms = (
        "抖音", "美食局", "食货局", "人气", "免配送", "免配", "不免配送",
        "全配", "半配", "备用", "直播", "kfc", "肯德基早餐", "到手", "单件不到",
        "买一送一", "工作日早餐", "可选薄脆金沙翅", "区域上架", "区域限定",
        "门店限定", "指定门店", "门店上架",
    )
    lines = text(raw).replace("KFC |", "").splitlines()
    name = lines[0] if lines else ""

    def strip_display_tag(match: re.Match[str]) -> str:
        tag = norm(match.group(1)).lower()
        return "" if any(term in tag for term in display_terms) else match.group(0)

    name = re.sub(r"【([^】]+)】", strip_display_tag, name)
    # 投放范围是后台/本地表常见的展示差异，不改变套餐商品内容。
    name = re.sub(r"(?:区域上架|区域限定|门店限定|指定门店|门店上架)", "", name)
    name = re.sub(r"(?:[·・、,，]\s*)?可选薄脆金沙翅", "", name)
    name = re.sub(r"^(?:抖音|直播间)[：:|\-\s]*", "", name, flags=re.I)
    name = name.replace("二", "2")
    key = re.sub(r"[\s·・、,，。.!！?？:：()（）\[\]【】]", "", name).lower()
    # “原味鸡青年1块”和“1块原味鸡青年”是后台与 Offer 的常见等价写法。
    quantity_last = re.fullmatch(r"(.+?)(\d+[个只块支])", key)
    return f"{quantity_last.group(2)}{quantity_last.group(1)}" if quantity_last else key


@dataclass
class Record:
    sheet: str
    row: int
    data: dict[str, str]

    @property
    def location(self) -> str: return f"{self.sheet}第 {self.row} 行"


def parse_sheet(worksheet) -> list[Record]:
    matrix = list(worksheet.iter_rows(values_only=True))
    if not matrix: return []
    markers = {norm(x).lower() for x in [*BACKEND["id"], *OFFER["id"], "商品名称", "产品名称", "产品名字", "活动价", "L3"]}
    header_row = max(range(min(30, len(matrix))), key=lambda i: len({norm(cell).lower() for cell in matrix[i] if text(cell)} & markers))
    heads = [text(cell) for cell in matrix[header_row]]
    if not any(heads): return []
    records = []
    for number, cells in enumerate(matrix[header_row + 1:], start=header_row + 2):
        data = {heads[i]: text(cell) for i, cell in enumerate(cells) if i < len(heads) and heads[i]}
        if any(data.values()): records.append(Record(worksheet.title, number, data))
    return records


def parse_xlsx(content: bytes, sheet_names: list[str] | None = None) -> list[Record]:
    book = load_workbook(io.BytesIO(content), read_only=False, data_only=True)
    if sheet_names is None:
        worksheets = book.worksheets
    else:
        missing = [name for name in sheet_names if name not in book.sheetnames]
        if missing:
            raise ValueError(f"Offer 工作表不存在：{'、'.join(missing)}")
        worksheets = [book[name] for name in sheet_names]
    return [record for sheet in worksheets for record in parse_sheet(sheet)]


def offer_sheet_options(content: bytes, backend_content: bytes | None = None) -> list[dict[str, object]]:
    """返回可作为抖音 Offer 来源的工作表，供页面让业务选择本次批次。"""
    book = load_workbook(io.BytesIO(content), read_only=False, data_only=True)
    backend_ids = None
    if backend_content:
        backend_records = parse_xlsx(backend_content)
        backend_ids = {
            coupon_id
            for record in backend_records
            for coupon_id in coupon_ids(value(record.data, BACKEND["id"]))
        }
    options = []
    for sheet in book.worksheets:
        records = parse_sheet(sheet)
        eligible = [
            record for record in records
            if find_header(record.data, OFFER["id"])
            and find_header(record.data, OFFER["商品名称"])
            and (find_header(record.data, OFFER["活动价"]) or find_header(record.data, OFFER["L3"]))
        ]
        if eligible:
            option: dict[str, object] = {"name": sheet.title, "records": len(eligible)}
            if backend_ids is not None:
                offer_ids = {
                    coupon_id
                    for record in eligible
                    for coupon_id in coupon_ids(value(record.data, OFFER["id"]))
                }
                option["matches"] = len(offer_ids & backend_ids)
            options.append(option)
    return options


def offers_by_id(records: list[Record]) -> tuple[dict[str, list[Record]], list[Record]]:
    eligible = [x for x in records if find_header(x.data, OFFER["商品名称"]) and (find_header(x.data, OFFER["活动价"]) or find_header(x.data, OFFER["L3"]))]
    grouped: dict[str, list[Record]] = defaultdict(list)
    for record in eligible:
        ids = coupon_ids(value(record.data, OFFER["id"])) or coupon_ids(" ".join(record.data.values()))
        for coupon_id in ids: grouped[coupon_id].append(record)
    return grouped, eligible


def candidate_score(backend: Record, offer: Record) -> int:
    """同券 ID 出现于多页时，只在业务字段能唯一佐证时自动选中。"""
    score = 0
    backend_name, offer_name = value(backend.data, BACKEND["商品名称"]), value(offer.data, OFFER["商品名称"])
    if backend_name and offer_name and core_name(backend_name) == core_name(offer_name):
        score += 4
    offer_price = money(value(offer.data, OFFER["活动价"]))
    backend_prices = {
        money(value(backend.data, BACKEND["活动价"])),
        money(value(backend.data, ["商促价"])),
    }
    if offer_price and offer_price in backend_prices:
        score += 2
    backend_l3, offer_l3 = money(value(backend.data, BACKEND["L3"])), money(value(offer.data, OFFER["L3"]))
    if backend_l3 and offer_l3 and backend_l3 == offer_l3:
        score += 2
    return score


def offer_signature(offer: Record) -> tuple[str, str, str, str, str]:
    """Fields that determine whether duplicate Offer rows are equivalent."""
    return (
        norm(value(offer.data, OFFER["商品名称"])),
        money(value(offer.data, OFFER["活动价"])),
        money(value(offer.data, OFFER["L3"])),
        norm(value(offer.data, OFFER["售卖开始时间"])),
        norm(value(offer.data, OFFER["每日可用时段"])),
    )


def choose_offer(coupon_id: str, backend: Record, by_id: dict[str, list[Record]], all_offers: list[Record]) -> tuple[Record | None, str, str]:
    candidates = by_id.get(coupon_id, [])
    if len(candidates) == 1: return candidates[0], "券ID匹配", ""
    if len(candidates) > 1:
        # 同一券常同时保留在相邻月份的 Offer 工作表。若全部核对字段一致，
        # 它们只是重复来源，不应降低本次核对的可信度。
        if len({offer_signature(candidate) for candidate in candidates}) == 1:
            return candidates[0], "券ID重复，相同Offer合并确认", ""
        plain = [x for x in candidates if not re.search(r"【[^】]+】\s*$", value(x.data, OFFER["商品名称"]))]
        if len(plain) == 1: return plain[0], "券ID重复，选择无【】后缀券名", ""
        scored = sorted(((candidate_score(backend, candidate), candidate) for candidate in candidates), key=lambda item: item[0], reverse=True)
        if scored and scored[0][0] >= 2 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
            return scored[0][1], "券ID重复，按商品字段唯一确认", ""
        return None, "券ID重复待复核", "Offer 中无【】后缀的券名记录不唯一"
    name = core_name(value(backend.data, BACKEND["商品名称"]))
    fallback = [x for x in all_offers if name and core_name(value(x.data, OFFER["商品名称"])) == name]
    if len(fallback) == 1: return fallback[0], "商品名称唯一回退匹配", "券ID未在 Offer 表找到，已按商品名称唯一匹配"
    return None, "未匹配", "券ID未在 Offer 表找到，且商品名称未找到唯一匹配"


def compare(backend: Record, offer: Record | None, field: str, match_note: str) -> tuple[str, str]:
    if not offer: return "待复核", match_note
    logical_field, _ = FIELD_MAPPING[field]
    actual = value(backend.data, BACKEND[logical_field])
    expected = value(offer.data, OFFER[logical_field])
    if not actual or not expected: return "待复核", "后台或 Offer 字段为空"
    if logical_field == "商品名称":
        return ("通过", "") if core_name(actual) == core_name(expected) else ("不通过", "核心套餐名不一致")
    if logical_field == "活动价":
        offer_prices = money_values(expected)
        backend_prices = money_values(actual) | money_values(value(backend.data, ["商促价"]))
        if not offer_prices or not backend_prices: return "待复核", "后台或业务本地表价格无法解析"
        return ("通过", "") if offer_prices & backend_prices else ("不通过", "售价和商促价均与业务本地表活动价不一致")
    if logical_field == "L3":
        if not money(actual) or not money(expected): return "待复核", "后台或 Offer 价格无法解析"
        return ("通过", "") if money(actual) == money(expected) else ("不通过", "金额不一致")
    if logical_field == "商品类型":
        expected_type = "次卡" if "次卡" in expected else "团购"
        return ("通过", "") if norm(actual) == expected_type else ("不通过", "与 Offer 产品名称推导类型不一致")
    if logical_field == "每日可用时段":
        if not time_range(expected): return "待复核", "Offer 使用效期未写明每日可用时段"
        return ("通过", "") if time_range(actual) == time_range(expected) else ("不通过", "与 Offer 使用效期中的每日可用时段不一致")
    # “10/12-TBD”已明确上架起点。结束时间为 TBD 时，按业务约定使用
    # 2026-12-31 作为默认结束边界。
    offer_dates = [TBD_END_DATE] if "tbd" in expected.lower() and logical_field == "售卖结束时间" else dates(expected, year_from(actual))
    backend_dates = dates(actual)
    if not offer_dates or not backend_dates: return "待复核", "Offer 上架时间边界无法解析"
    actual_date = min(backend_dates) if logical_field == "售卖开始时间" else max(backend_dates)
    expected_date = min(offer_dates) if logical_field == "售卖开始时间" else max(offer_dates)
    valid = actual_date >= expected_date if logical_field == "售卖开始时间" else actual_date <= expected_date
    return ("通过", "") if valid else ("不通过", "超出 Offer 已知上架时间范围")


def joined_values(rows: list[Record], headers: list[str]) -> str:
    """保留同券下所有不同后台值，供业务确认多商品ID的实际内容。"""
    return "、".join(dict.fromkeys(value(row.data, headers) for row in rows if value(row.data, headers)))


def choose_coupon_offer(coupon_id: str, rows: list[Record], by_id: dict[str, list[Record]], all_offers: list[Record]) -> tuple[Record | None, str, str]:
    """同一券ID可对应多条商品ID，只有候选一致时才确认同一个 Offer。"""
    attempts = [choose_offer(coupon_id, row, by_id, all_offers) for row in rows]
    selected = [(offer, method) for offer, method, _ in attempts if offer]
    if not selected:
        return None, attempts[0][1], attempts[0][2]
    locations = {(offer.sheet, offer.row) for offer, _ in selected}
    if len(locations) == 1:
        offer, method = selected[0]
        return offer, method, ""
    if len({offer_signature(offer) for offer, _ in selected}) == 1:
        return selected[0][0], "同券多商品ID，相同Offer合并确认", ""
    return None, "券ID重复待复核", "同一券ID对应的 Offer 候选不一致"


def aggregate_field_result(rows: list[Record], offer: Record | None, field: str, match_note: str) -> tuple[str, str]:
    if field == "售价" and offer:
        local_prices = money_values(value(offer.data, OFFER["活动价"]))
        backend_prices = set().union(*(money_values(value(row.data, BACKEND["活动价"])) | money_values(value(row.data, ["商促价"])) for row in rows))
        if local_prices and backend_prices and local_prices & backend_prices:
            return "通过", ""
    outcomes = [compare(row, offer, field, match_note) for row in rows]
    statuses = [status for status, _ in outcomes]
    reasons = list(dict.fromkeys(reason for _, reason in outcomes if reason))
    status = "不通过" if "不通过" in statuses else "待复核" if "待复核" in statuses else "通过"
    return status, "；".join(reasons)


def coupon_id_result(coupon_id: str, offer: Record | None, match_note: str) -> tuple[str, str]:
    """商家平台券ID与业务本地表券ID的单独核对结果。"""
    if not offer:
        return "待复核", match_note
    local_coupon_ids = coupon_ids(value(offer.data, OFFER["券ID"]))
    if coupon_id and coupon_id in local_coupon_ids:
        return "通过", ""
    return "不通过", "商家平台券ID与业务本地表券ID不一致"


def offer_output_field_name(field: str) -> str:
    """返回业务本地表的真实字段名。"""
    return FIELD_MAPPING[field][1]


def make_report(offer_content: bytes, backend_content: bytes, selected_fields: list[str], offer_filename: str = "抖音本地Offer表", backend_filename: str = "抖音后台导表", offer_sheets: list[str] | None = None) -> tuple[bytes, dict[str, int]]:
    if not selected_fields: raise ValueError("请至少选择一个核对字段。")
    if unknown := [x for x in selected_fields if x not in FIELDS]: raise ValueError(f"不支持的核对字段：{'、'.join(unknown)}")
    offers = parse_xlsx(offer_content, offer_sheets)
    if not offers:
        selected = "、".join(offer_sheets or [])
        raise ValueError(f"Offer 工作表“{selected}”没有可读取的记录。" if selected else "Offer 表没有可读取的记录。")
    by_id, all_offers = offers_by_id(offers)
    backend = [x for x in parse_xlsx(backend_content) if find_header(x.data, BACKEND["id"]) and value(x.data, BACKEND["item"])]
    if not backend: raise ValueError("抖音后台导表未识别到商品ID和商家平台券ID列。")
    book = Workbook(); sheet = book.active; sheet.title = "核对结果表"
    # 与美团结果表保持相同的对照方式：每个字段都并列显示两侧原始值和结果。
    heads = ["商家平台券ID", "商品ID", "业务本地表格-券ID", "商家平台券ID-结果"]
    for field in selected_fields:
        heads += [f"抖音后台表格-{field}", f"业务本地表格-{offer_output_field_name(field)}", f"{field}-结果"]
    heads += ["整券结果", "异常原因"]
    sheet.append(heads); sheet.freeze_panes = "A2"; totals = {"通过": 0, "待复核": 0, "不通过": 0}; exceptions = []
    backend_groups: dict[str, list[Record]] = defaultdict(list)
    for row in backend:
        ids = coupon_ids(value(row.data, BACKEND["id"]))
        # 无法解析为标准券ID的记录也保留在结果中，便于人工补充；标准券ID则严格一券一行。
        group_id = ids[0] if len(ids) == 1 else value(row.data, BACKEND["id"]) or f"未识别券ID-{row.sheet}-{row.row}"
        backend_groups[group_id].append(row)
    for group_id, rows in backend_groups.items():
        coupon_id = group_id if len(coupon_ids(group_id)) == 1 else ""
        offer, _method, match_note = choose_coupon_offer(coupon_id, rows, by_id, all_offers) if coupon_id else (None, "未匹配", "后台券ID为空或无法解析")
        coupon_status, coupon_reason = coupon_id_result(coupon_id, offer, match_note)
        statuses, reasons = [coupon_status], []
        if coupon_reason: reasons.append(f"商家平台券ID：{coupon_reason}")
        product_ids = joined_values(rows, BACKEND["item"])
        row_output = [coupon_id or group_id, product_ids, value(offer.data, OFFER["券ID"]) if offer else "", coupon_status]
        for field in selected_fields:
            status, reason = aggregate_field_result(rows, offer, field, match_note)
            statuses.append(status)
            if reason: reasons.append(f"{field}：{reason}")
            logical_field, _ = FIELD_MAPPING[field]
            row_output += [
                joined_values(rows, BACKEND[logical_field]),
                value(offer.data, OFFER[logical_field]) if offer else "",
                status,
            ]
        overall = "不通过" if "不通过" in statuses else "待复核" if "待复核" in statuses else "通过"; totals[overall] += 1
        sheet.append([*row_output, overall, "；".join(reasons)])
        if reasons: exceptions.append((product_ids, coupon_id or group_id, overall, "；".join(reasons)))
    for cell in sheet[1]: cell.font = Font(bold=True, color="65738C")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if cell.value in STATUS_COLORS: cell.font = Font(color=STATUS_COLORS[cell.value])
        alignment = copy(row[-1].alignment); alignment.wrap_text = True; alignment.vertical = "top"; row[-1].alignment = alignment
    for i, label in enumerate(heads, 1):
        sheet.column_dimensions[get_column_letter(i)].width = 62 if label == "异常原因" else 28 if "-" in label else 22
    sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    summary = book.create_sheet("异常汇总"); summary.append(["商品ID", "商家平台券ID", "异常类型", "具体原因"])
    for item in exceptions: summary.append(item)
    summary.freeze_panes = "A2"; summary.auto_filter.ref = f"A1:D{summary.max_row}"
    for cell in summary[1]: cell.font = Font(bold=True, color="65738C")
    summary.column_dimensions["A"].width = 22; summary.column_dimensions["B"].width = 24; summary.column_dimensions["C"].width = 14; summary.column_dimensions["D"].width = 80
    stream = io.BytesIO(); book.save(stream)
    return stream.getvalue(), totals
