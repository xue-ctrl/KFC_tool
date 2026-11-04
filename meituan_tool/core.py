from __future__ import annotations

import ast
import io
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from zipfile import BadZipFile, ZipFile

RESULT_FIELDS = ["售卖开始时间", "售卖结束时间", "项目名称", "其他规则补充说明", "美团价", "原价", "菜品备注", "原价/美团价"]
OUTPUT_HEADERS = ["券ID", "美团ID", *RESULT_FIELDS, "整券结果", "异常原因"]
# 保留单项价格核查。后台导表可使用“售价 / 商促价 / 划线价”等同义列；
# 价格比例需要两项都存在可靠来源，当前不提供为独立核查项。
SIMPLE_FIELDS = ["售卖开始时间", "售卖结束时间", "项目名称", "使用有效期", "其他规则补充说明", "美团价", "原价", "菜品备注"]
SIMPLE_ALIASES = {
    "售卖开始时间": (["售卖开始时间", "上架时间", "*上架时间", "开始时间"], ["售卖开始时间", "上架时间", "开始时间"]),
    "售卖结束时间": (["售卖结束时间", "下架时间", "*上架时间", "结束时间"], ["售卖结束时间", "下架时间", "结束时间"]),
    "项目名称": (["产品优惠券名称", "券名", "*券名", "项目名称", "商品名称", "产品标题（不一定等于券名）"], ["项目名称", "商品名称", "券名"]),
    "使用有效期": (["有效期", "使用效期", "*使用效期", "券有效期", "有效期(和券说明内容一致）", "券说明文字"], ["券有效期", "使用有效期", "有效期"]),
    "其他规则补充说明": (["完整券说明", "券说明", "其他规则补充说明", "备注"], ["其他规则补充说明", "使用规则", "项目备注", "备注"]),
    "美团价": (["活动价格", "美团价", "售价", "L3价格"], ["美团价", "售价", "商促价"]),
    "原价": (["L3价格", "原价", "划线价", "商品总价（商品划线价）", "*售价（元）", "售价（元）"], ["原价", "划线价", "商品总价（商品划线价）"]),
    "菜品备注": (["产品包含", "菜品备注", "产品内容", "产品内容（和券说明产品包含一致）", "券说明文字"], ["菜品备注", "产品包含", "产品内容"]),
}
MASTER_ID_HEADERS = ["优惠券ID", "券ID", "券id", "商家商品id", "商家平台券ID", "平台券ID", "美团券ID", "券码ID", "券码"]
MASTER_CHANNEL_HEADERS = ["渠道", "投放渠道"]
EXPORT_ID_HEADERS = ["商家商品id", "商家商品ID"]
EXPORT_PROJECT_HEADERS = ["项目ID", "项目id"]
EXPORT_FIELD_HEADERS = {
    "售卖开始时间": ["售卖开始时间"], "售卖结束时间": ["售卖结束时间"], "项目名称": ["项目名称"],
    "其他规则补充说明": ["其他规则补充说明"], "美团价": ["美团价"], "原价": ["原价"], "菜品备注": ["菜品备注"],
}
OFFER_HEADERS = {"id": ["券ID", "券id"], "name": ["*券名", "券名"], "content": ["产品内容"], "delivery": ["宅急送内容"], "validity": ["*使用效期", "使用效期"], "listing": ["*上架时间", "上架时间"], "delivery_fee": ["外送费优惠"], "sale_price": ["*售价（元）", "售价（元）"], "strike_price": ["划线价", "*划线价"]}


def text(value: Any) -> str:
    if value is None: return ""
    if isinstance(value, float) and value.is_integer(): return str(int(value))
    return str(value).strip()


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", text(value)).replace("．", ".")


def find_header(row: dict[str, str], candidates: list[str]) -> str | None:
    keys = {norm(key).lower(): key for key in row}
    for candidate in candidates:
        if found := keys.get(norm(candidate).lower()): return found
    return None


def value(row: dict[str, str], candidates: list[str]) -> str:
    header = find_header(row, candidates)
    return text(row.get(header, "")) if header else ""


@dataclass
class Sheet:
    rows: list[dict[str, str]]
    numbers: list[int]


def parse_worksheet(worksheet, offer: bool = False) -> Sheet:
    worksheet.reset_dimensions(); matrix = list(worksheet.iter_rows(values_only=True))
    if not matrix: return Sheet([], [])
    if offer and len(matrix) >= 3:
        header_index = 2
    else:
        markers = {"商家商品id", "项目id", "其他规则补充说明", "菜品备注", *{header.lower() for header in MASTER_ID_HEADERS}}
        header_index = max(range(min(30, len(matrix))), key=lambda index: len({norm(cell).lower() for cell in matrix[index] if text(cell)} & markers))
    headers = [text(cell) for cell in matrix[header_index]]
    rows, numbers = [], []
    for number, cells in enumerate(matrix[header_index + 1:], start=header_index + 2):
        # 某些后台导表的数据行会比表头多出空白/隐藏尾列；只读取有表头的范围。
        row = {headers[index]: text(cell) for index, cell in enumerate(cells) if index < len(headers) and headers[index]}
        if any(row.values()): rows.append(row); numbers.append(number)
    return Sheet(rows, numbers)


def parse_xlsx(content: bytes, offer: bool = False) -> Sheet:
    workbook = load_workbook(io.BytesIO(content), read_only=True, data_only=False)
    worksheet = workbook["offer表"] if offer and "offer表" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    return parse_worksheet(worksheet, offer)


def parse_xlsx_sheets(content: bytes) -> list[tuple[str, Sheet]]:
    workbook = load_workbook(io.BytesIO(content), read_only=True, data_only=False)
    return [(worksheet.title, parse_worksheet(worksheet)) for worksheet in workbook.worksheets]


def group_export(sheet: Sheet) -> tuple[dict[str, list[tuple[int, dict[str, str]]]], int]:
    grouped: dict[str, list[tuple[int, dict[str, str]]]] = defaultdict(list); missing = 0
    for number, row in zip(sheet.numbers, sheet.rows):
        coupon_id = value(row, EXPORT_ID_HEADERS)
        if coupon_id: grouped[coupon_id].append((number, row))
        else: missing += 1
    return grouped, missing


def field_values(rows: list[tuple[int, dict[str, str]]], headers: list[str], transform=lambda raw: norm(raw)) -> tuple[str | None, list[int]]:
    values, numbers = set(), []
    for number, row in rows:
        current = transform(value(row, headers))
        if current: values.add(current); numbers.append(number)
    return (next(iter(values)) if len(values) == 1 else None), numbers


def dates(raw: str) -> list[str]:
    return [f"{year}-{int(month):02d}-{int(day):02d}" for year, month, day in re.findall(r"(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})", text(raw))]


def date_boundary(raw: str, end: bool) -> str:
    found = dates(raw)
    return found[-1] if end and found else found[0] if found else ""


def formula_amount(raw: str) -> str:
    """计算仅由数字和四则运算组成的 Excel 价格公式，拒绝单元格引用和函数。"""
    expression = str(raw).strip()
    if not expression.startswith("="):
        return ""
    try:
        tree = ast.parse(expression[1:], mode="eval")

        def calculate(node: ast.AST) -> Decimal:
            if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
                return Decimal(str(node.value))
            if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
                value_ = calculate(node.operand)
                return value_ if isinstance(node.op, ast.UAdd) else -value_
            if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
                left, right = calculate(node.left), calculate(node.right)
                if isinstance(node.op, ast.Add): return left + right
                if isinstance(node.op, ast.Sub): return left - right
                if isinstance(node.op, ast.Mult): return left * right
                if right == 0: raise ZeroDivisionError
                return left / right
            raise ValueError("unsupported formula")

        return f"{calculate(tree.body):.2f}"
    except (SyntaxError, ValueError, InvalidOperation, ZeroDivisionError):
        return ""


def amount(raw: str) -> str:
    formula_value = formula_amount(raw)
    if formula_value:
        return formula_value
    try: return f"{float(str(raw).replace('元', '').strip()):.2f}"
    except (TypeError, ValueError): return ""


def title_key(raw: str) -> str:
    # 名称前的【美团】、[TEST]、【最后一次6.9元】等均为渠道/营销展示标签，
    # 不改变实际商品名称；仅移除开头连续出现的标签，正文中的方括号内容保留。
    cleaned = re.sub(r"^(?:\s*[【\[][^】\]]{1,40}[】\]]\s*)+", "", norm(raw), flags=re.I)
    cleaned = cleaned.replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", "", cleaned).lower()


def parse_docx(content: bytes) -> dict[str, list[str]]:
    try:
        with ZipFile(io.BytesIO(content)) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
    except (BadZipFile, KeyError, ET.ParseError) as error:
        raise ValueError("券说明 DOCX 无法读取，请重新导出后上传。") from error
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    paragraphs = []
    for paragraph in root.iter(f"{namespace}p"):
        current = "".join(node.text or "" for node in paragraph.iter(f"{namespace}t"))
        if text(current): paragraphs.append(text(current))
    blocks: list[str] = []; current: list[str] = []; current_ids: set[str] = set()
    for paragraph in paragraphs:
        found = set(re.findall(r"SZKFCMT[A-Za-z0-9_-]+", paragraph))
        if found and current_ids and found != current_ids:
            blocks.append("\n".join(current)); current = []; current_ids = set()
        current.append(paragraph); current_ids.update(found)
    if current: blocks.append("\n".join(current))
    indexed: dict[str, list[str]] = defaultdict(list)
    for block in blocks:
        for coupon_id in set(re.findall(r"SZKFCMT[A-Za-z0-9_-]+", block)): indexed[coupon_id].append(block)
    return indexed


def applicable_doc_text(block: str) -> str:
    """Keep common and Meituan-specific clauses; exclude explicitly other-channel-only lines."""
    other_channels = ("抖音", "支付宝", "高德", "淘闪", "快手")
    lines = []
    for line in block.splitlines():
        if any(channel in line for channel in other_channels) and "美团" not in line: continue
        lines.append(line)
    return "\n".join(lines)


def extract_after(label: str, source: str) -> str:
    match = re.search(re.escape(label) + r"\s*([^\n\r]+)", source)
    return text(match.group(1)) if match else ""


def product_content(source: str) -> str:
    match = re.search(r"产品包含[：:]\s*(.*?)(?=\s*(?:使用有效期|仅限|本券|不适用|$))", source, re.S)
    return norm(match.group(1)) if match else ""


def offer_candidates(coupon_id: str, export_rows: list[tuple[int, dict[str, str]]], offers: Sheet) -> tuple[list[tuple[int, dict[str, str]]], str]:
    exact = [(number, row) for number, row in zip(offers.numbers, offers.rows) if value(row, OFFER_HEADERS["id"]) == coupon_id]
    if exact: return exact, "券ID精确关联"
    project_name, _ = field_values(export_rows, EXPORT_FIELD_HEADERS["项目名称"])
    if not project_name: return [], "导表项目名称冲突或为空"
    candidates = [(number, row) for number, row in zip(offers.numbers, offers.rows) if title_key(value(row, OFFER_HEADERS["name"])) == title_key(project_name)]
    # 券名只是候选条件；券 ID 缺失时至少还需一个可验证的关键字段。
    start, _ = field_values(export_rows, EXPORT_FIELD_HEADERS["售卖开始时间"], lambda raw: date_boundary(raw, False))
    original, _ = field_values(export_rows, EXPORT_FIELD_HEADERS["原价"], amount)
    food, _ = field_values(export_rows, EXPORT_FIELD_HEADERS["菜品备注"])
    confirmed = []
    for number, row in candidates:
        signals = 0
        if start and start == date_boundary(value(row, OFFER_HEADERS["listing"]), False): signals += 1
        if original and original == amount(value(row, OFFER_HEADERS["sale_price"])): signals += 1
        if food and norm(food) == norm(value(row, OFFER_HEADERS["content"])): signals += 1
        if signals >= 1: confirmed.append((number, row))
    if len(confirmed) == 1: return confirmed, "券名及关键字段候选关联"
    if len(candidates) == 1: return candidates, "券名唯一但缺少价格、日期或产品内容等关键佐证"
    if not candidates: return [], "Offer 中未找到同券ID或同名候选"
    return candidates, f"券名匹配到 {len(candidates)} 个候选 Offer 行，关键字段无法唯一确认"


def export_conflict(rows: list[tuple[int, dict[str, str]]], headers: list[str], transform=lambda raw: norm(raw)) -> tuple[bool, list[int]]:
    seen = {transform(value(row, headers)) for _, row in rows if transform(value(row, headers))}
    return len(seen) > 1, [number for number, _ in rows]


def check(condition: bool, reliable: bool, reason_if_false: str, reason_if_unreliable: str) -> tuple[str, str]:
    if not reliable: return "待复核", reason_if_unreliable
    return ("通过", "") if condition else ("不通过", reason_if_false)


def make_report(export_content: bytes, offer_content: bytes, reference_content: bytes, export_filename: str = "待核对导表") -> tuple[bytes, dict[str, int]]:
    export, offers = parse_xlsx(export_content), parse_xlsx(offer_content, offer=True)
    if not export.rows:
        raise ValueError("待核对导表没有可读取的数据行。")
    missing_headers = [header for header in ("商家商品id", "其他规则补充说明", "菜品备注") if not find_header(export.rows[0], [header])]
    if missing_headers:
        raise ValueError(f"待核对导表缺少必需表头：{'、'.join(missing_headers)}。请确认上传的是完整的美团导表。")
    groups, unindexed = group_export(export); documents = parse_docx(reference_content)
    workbook = Workbook(); sheet = workbook.active; sheet.title = "美团审核结果"; sheet.freeze_panes = "A2"; sheet.append(OUTPUT_HEADERS)
    colors = {"通过": "008000", "待复核": "FFC000", "不通过": "FF0000"}; totals = {"通过": 0, "待复核": 0, "不通过": 0, "无法索引": unindexed}
    for coupon_id, export_rows in groups.items():
        project_ids = sorted({value(row, EXPORT_PROJECT_HEADERS) for _, row in export_rows if value(row, EXPORT_PROJECT_HEADERS)})
        candidates, evidence = offer_candidates(coupon_id, export_rows, offers)
        offer_row = candidates[0][1] if len(candidates) == 1 else None
        offer_line = candidates[0][0] if offer_row else None
        doc_blocks = documents.get(coupon_id, [])
        document = applicable_doc_text(doc_blocks[0]) if len(doc_blocks) == 1 else ""
        reasons: list[str] = []
        statuses: dict[str, str] = {}

        def source_ready(needs_doc: bool = False) -> tuple[bool, str]:
            if len(candidates) != 1: return False, evidence
            # Offer 券ID为空时，即使标题/价格等形成唯一候选，也不能确认属于同一批次；
            # 该候选仅用于给出人工复核线索，不能据此判定不通过或通过。
            if evidence != "券ID精确关联": return False, f"{evidence}，Offer 未以券ID确认同批次关联"
            if needs_doc and not document: return False, "未找到同券ID的唯一券说明块"
            return True, ""

        start, start_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["售卖开始时间"], lambda raw: date_boundary(raw, False))
        expected_start = date_boundary(value(offer_row or {}, OFFER_HEADERS["listing"]), False)
        ready, unavailable = source_ready(); statuses["售卖开始时间"], reason = check(start == expected_start, bool(start and expected_start and ready), f"与 Offer 表第 {offer_line} 行的上架起点不一致（导表 {start}，Offer {expected_start}）", unavailable or "导表或 Offer 上架时间无法解析")
        if reason: reasons.append(f"[{statuses['售卖开始时间']}] 《{export_filename}》与《Offer表》第 {offer_line or '未知'} 行：售卖开始时间，{reason}")

        end, end_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["售卖结束时间"], lambda raw: date_boundary(raw, True))
        expected_end = date_boundary(value(offer_row or {}, OFFER_HEADERS["listing"]), True)
        ready, unavailable = source_ready(); statuses["售卖结束时间"], reason = check(end == expected_end, bool(end and expected_end and ready), f"与 Offer 表第 {offer_line} 行的上架终点不一致（导表 {end}，Offer {expected_end}）", unavailable or "导表或 Offer 上架时间无法解析")
        if reason: reasons.append(f"[{statuses['售卖结束时间']}] 《{export_filename}》与《Offer表》第 {offer_line or '未知'} 行：售卖结束时间，{reason}")

        project_name, project_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["项目名称"])
        expected_name = value(offer_row or {}, OFFER_HEADERS["name"])
        ready, unavailable = source_ready(needs_doc=True)
        name_ok = title_key(project_name) == title_key(expected_name) and title_key(project_name) in title_key(document)
        statuses["项目名称"], reason = check(name_ok, bool(project_name and expected_name and ready), "项目名称与同券 Offer 标题或券说明标题不一致", unavailable or "项目名称缺失")
        if reason: reasons.append(f"[{statuses['项目名称']}] 《{export_filename}》与《Offer表/券说明》第 {offer_line or '未知'} 行：项目名称，{reason}")

        actual_rules, rule_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["其他规则补充说明"])
        offer_delivery = value(offer_row or {}, OFFER_HEADERS["delivery"]); offer_has_delivery = bool(offer_delivery and norm(offer_delivery) != "/")
        actual_has_delivery = "自理" in actual_rules
        doc_has_delivery = "自理" in document
        actual_validity = extract_after("使用有效期：", actual_rules); doc_validity = extract_after("使用有效期：", document); offer_validity = value(offer_row or {}, OFFER_HEADERS["validity"])
        validity_ok = bool(actual_validity and doc_validity and offer_validity and norm(actual_validity) == norm(doc_validity) and dates(actual_validity) == dates(offer_validity))
        rules_text_ok = bool(document and norm(actual_rules) == norm(document))
        delivery_ok = bool(offer_delivery and actual_has_delivery == offer_has_delivery and actual_has_delivery == doc_has_delivery)
        ready, unavailable = source_ready(needs_doc=True)
        rules_ok = delivery_ok and validity_ok and rules_text_ok
        statuses["其他规则补充说明"], reason = check(rules_ok, ready, "宅急送/有效期/券说明文字存在已证实差异", unavailable or "规则来源缺失或无法可靠分离美团适用条款")
        if reason: reasons.append(f"[{statuses['其他规则补充说明']}] 《{export_filename}》与《Offer表/券说明》第 {offer_line or '未知'} 行：其他规则补充说明，{reason}")

        mt_price, mt_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["美团价"], amount)
        strike = amount(value(offer_row or {}, OFFER_HEADERS["strike_price"]))
        ready, unavailable = source_ready(); statuses["美团价"], reason = check(mt_price == strike, bool(mt_price and strike and ready), f"美团价与 Offer 可确认划线价不一致（导表 {mt_price}，Offer {strike}）", unavailable or "Offer 未单列可确认的划线价，不能在 L4 ALC 与 L3 ALC 中自行选择")
        if reason: reasons.append(f"[{statuses['美团价']}] 《{export_filename}》与《Offer表》第 {offer_line or '未知'} 行：美团价，{reason}")

        original, original_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["原价"], amount)
        sale = amount(value(offer_row or {}, OFFER_HEADERS["sale_price"]))
        ready, unavailable = source_ready(); statuses["原价"], reason = check(original == sale, bool(original and sale and ready), f"原价与 Offer 售价不一致（导表 {original}，Offer {sale}）", unavailable or "导表原价或 Offer 售价缺失")
        if reason: reasons.append(f"[{statuses['原价']}] 《{export_filename}》与《Offer表》第 {offer_line or '未知'} 行：原价，{reason}")

        food, food_lines = field_values(export_rows, EXPORT_FIELD_HEADERS["菜品备注"])
        offer_food = norm(value(offer_row or {}, OFFER_HEADERS["content"])); doc_food = product_content(document)
        ready, unavailable = source_ready(needs_doc=True)
        food_ok = bool(food and offer_food and doc_food and food == offer_food and food == doc_food)
        statuses["菜品备注"], reason = check(food_ok, ready, "菜品备注与 Offer 产品内容或同券券说明的产品包含不一致", unavailable or "菜品备注对照来源缺失")
        if reason: reasons.append(f"[{statuses['菜品备注']}] 《{export_filename}》与《Offer表/券说明》第 {offer_line or '未知'} 行：菜品备注，{reason}")

        statuses["原价/美团价"], reason = check(bool(original and mt_price and float(mt_price) != 0), bool(original and mt_price and float(mt_price) != 0), "原价/美团价计算失败", "原价或美团价为空，不能计算原价/美团价")
        if reason: reasons.append(f"[{statuses['原价/美团价']}] 《{export_filename}》：原价/美团价，{reason}")

        # 同券 ID 的导表记录不能任取其中一行；冲突字段一律交由人工复核。
        conflict_specs = {
            "售卖开始时间": (EXPORT_FIELD_HEADERS["售卖开始时间"], lambda raw: date_boundary(raw, False)),
            "售卖结束时间": (EXPORT_FIELD_HEADERS["售卖结束时间"], lambda raw: date_boundary(raw, True)),
            "项目名称": (EXPORT_FIELD_HEADERS["项目名称"], norm),
            "其他规则补充说明": (EXPORT_FIELD_HEADERS["其他规则补充说明"], norm),
            "美团价": (EXPORT_FIELD_HEADERS["美团价"], amount),
            "原价": (EXPORT_FIELD_HEADERS["原价"], amount),
            "菜品备注": (EXPORT_FIELD_HEADERS["菜品备注"], norm),
        }
        for field, (headers, transform) in conflict_specs.items():
            conflict, lines = export_conflict(export_rows, headers, transform)
            if conflict:
                statuses[field] = "待复核"
                reasons = [entry for entry in reasons if f"：{field}，" not in entry]
                reasons.append(f"[待复核] 《{export_filename}》第 {'、'.join(map(str, lines))} 行：{field}，同一券ID下存在互相冲突的导表值，不能任选一行判定。")
        if statuses["美团价"] == "待复核" or statuses["原价"] == "待复核":
            statuses["原价/美团价"] = "待复核"
            reasons = [entry for entry in reasons if "：原价/美团价，" not in entry]
            reasons.append(f"[待复核] 《{export_filename}》：原价/美团价，价格字段尚未可靠确认，不能据此计算折扣。")
        overall = "不通过" if "不通过" in statuses.values() else "待复核" if "待复核" in statuses.values() else "通过"
        totals[overall] += 1
        sheet.append([coupon_id, "、".join(project_ids), *[statuses[field] for field in RESULT_FIELDS], overall, "\n".join(reasons)])
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}{sheet.max_row}"
    for cell in sheet[1]: cell.font = Font(bold=True, color="65738C")
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=11):
        for cell in row: cell.font = Font(color=colors[cell.value])
    sheet.column_dimensions["A"].width = 24; sheet.column_dimensions["B"].width = 24; sheet.column_dimensions["L"].width = 68
    for cells in sheet.iter_rows(min_col=12, max_col=12):
        alignment = copy(cells[0].alignment); alignment.wrap_text = True; alignment.vertical = "top"; cells[0].alignment = alignment
    stream = io.BytesIO(); workbook.save(stream)
    return stream.getvalue(), totals


def is_meituan_coupon(row: dict[str, str], coupon_id: str) -> bool:
    """总表通常以一行公共资料、随后多行渠道券 ID 的形式保存。"""
    return coupon_id.upper().startswith("SZKFCMT") or "美团" in value(row, MASTER_CHANNEL_HEADERS)


def inherited_master_rows(sheet: Sheet) -> list[tuple[int, dict[str, str]]]:
    """将 KFC 总表中渠道券 ID 行补足其上方公共资料，但不伪造券 ID。"""
    previous: dict[str, str] = {}
    expanded: list[tuple[int, dict[str, str]]] = []
    for number, row in zip(sheet.numbers, sheet.rows):
        coupon_id = value(row, MASTER_ID_HEADERS)
        completed = {header: text(row.get(header)) or previous.get(header, "") for header in row}
        # 只为当前有值的券 ID 建行；空白 ID 不会继承为虚假券 ID。
        if coupon_id:
            completed[find_header(completed, MASTER_ID_HEADERS) or "券ID"] = coupon_id
            if is_meituan_coupon(completed, coupon_id): expanded.append((number, completed))
        for header, current in row.items():
            if text(current): previous[header] = text(current)
    return expanded


def comparable(raw: str, field: str) -> str:
    if field == "售卖开始时间":
        return date_boundary(raw, False) or norm(raw)
    if field == "售卖结束时间":
        return date_boundary(raw, True) or norm(raw)
    if field == "使用有效期":
        parsed = validity_dates(raw)
        return "|".join(parsed) if parsed else normalize_validity_text(raw)
    if field in {"美团价", "原价"}:
        return amount(raw) or norm(raw)
    return norm(raw)


def validity_dates(raw: str) -> list[str]:
    """Return validity boundaries without requiring both sides to use the same year format."""
    tokens: list[str] = []
    for _year, month, day in re.findall(r"(20\d{2})[年/-](\d{1,2})[月/-](\d{1,2})", text(raw)):
        token = f"{int(month):02d}-{int(day):02d}"
        if token not in tokens:
            tokens.append(token)
    for month, day in re.findall(r"(?<!\d)(\d{1,2})[月./-](\d{1,2})(?:日)?", text(raw)):
        token = f"{int(month):02d}-{int(day):02d}"
        if token not in tokens:
            tokens.append(token)
    return tokens[:2]


def normalize_validity_text(raw: str) -> str:
    """Normalize repeated duration labels without masking a genuinely different validity rule."""
    normalized = norm(raw)
    # 例如“30天、自购买次日起30天，购买当天默认可用”中的首个“30天”
    # 只是总表的重复展示，后面的完整规则才是实际有效期说明。
    normalized = re.sub(
        r"^\s*(\d+\s*天)\s*[、,，;；]\s*(?=自购买(?:次日|之日)起\s*\1)",
        "",
        normalized,
    )
    # 总表常仅写“14天”，后台则补充“自购买次日起14天，购买当天默认可用”。
    # 日期范围已在上层单独严格处理；无日期时按相同天数视为同一有效期。
    if t_plus := re.search(r"\bT\s*\+\s*(\d+)", normalized, flags=re.I):
        return f"{t_plus.group(1)}天"
    if duration_only := re.match(r"^\s*(\d+)\s*天", normalized):
        days = duration_only.group(1)
        suffix = normalized[duration_only.end():]
        # “14天0930-2300”或“5天、5天0930-2300”中的时段不属于有效期；
        # 仅当其余内容确实只有同一天数与时段时才忽略，避免掩盖其他规则差异。
        duration_pattern = rf"{re.escape(days)}\s*天"
        suffix = re.sub(rf"(?:{duration_pattern})?\s*\d{{1,2}}:?\d{{2}}\s*[-至]\s*\d{{1,2}}:?\d{{2}}", "", suffix)
        suffix = re.sub(duration_pattern, "", suffix)
        if not re.sub(r"[\s、,，;；]", "", suffix):
            return f"{days}天"
    if duration := re.search(r"自购买(?:次日|之日)起\s*(\d+)\s*天", normalized):
        return f"{duration.group(1)}天"
    return re.sub(r"[\s、,，;；]", "", normalized)


def rule_validity_value_set(rows: list[tuple[int, dict[str, str]]]) -> tuple[list[str], list[int]]:
    """Extract the validity period embedded in Meituan's rule-description text."""
    values: set[str] = set()
    lines: list[int] = []
    aliases = SIMPLE_ALIASES["其他规则补充说明"][1]
    for number, row in rows:
        rules = field_raw_value(row, aliases, "其他规则补充说明")
        match = re.search(r"使用有效期\s*[：:]\s*(.{0,180})", rules, re.S)
        parsed = validity_dates(match.group(1) if match else "")
        if parsed:
            values.add("|".join(parsed))
            lines.append(number)
    return sorted(values), lines


def field_raw_value(row: dict[str, str], aliases: list[str], field: str) -> str:
    """读取总表字段；当值来自券说明文字时，仅提取该字段对应的片段。"""
    matched_header = find_header(row, aliases)
    raw = text(row.get(matched_header)) if matched_header else ""
    if not raw:
        return ""
    if field == "使用有效期" and matched_header == "券说明文字":
        match = re.search(r"使用有效期[：:]\s*(.*?)(?=\s*(?:仅限|本券|不适用|产品包含|$))", raw, re.S)
        if match:
            return text(match.group(1))
    if field == "菜品备注" and matched_header == "券说明文字":
        content = product_content(raw)
        if content:
            return content
    return raw


def grouped_values(rows: list[tuple[int, dict[str, str]]], aliases: list[str], field: str) -> tuple[str | None, list[int], bool]:
    values = {comparable(field_raw_value(row, aliases, field), field) for _, row in rows if field_raw_value(row, aliases, field)}
    lines = [number for number, _ in rows]
    return (next(iter(values)) if len(values) == 1 else None), lines, len(values) > 1


def coupon_value_set(rows: list[tuple[int, dict[str, str]]], aliases: list[str], field: str) -> tuple[list[str], list[int]]:
    """Return every distinct value under one coupon ID.

    A coupon can legitimately be bound to several live Meituan projects.  The
    total-table reconciliation is keyed only by coupon ID, so those values must
    be compared as a set instead of treating more than one value as an
    automatic conflict.
    """
    values = sorted({
        comparable(field_raw_value(row, aliases, field), field)
        for _, row in rows
        if comparable(field_raw_value(row, aliases, field), field)
    })
    return values, [number for number, _ in rows]


def coupon_project_name_sets_match(master_rows: list[tuple[int, dict[str, str]]], backend_rows: list[tuple[int, dict[str, str]]]) -> bool:
    """将美团/测试展示标签排除后，比较同券的完整项目名称集合。"""
    master_aliases, backend_aliases = SIMPLE_ALIASES["项目名称"]
    master_names = {title_key(field_raw_value(row, master_aliases, "项目名称")) for _, row in master_rows if field_raw_value(row, master_aliases, "项目名称")}
    backend_names = {title_key(field_raw_value(row, backend_aliases, "项目名称")) for _, row in backend_rows if field_raw_value(row, backend_aliases, "项目名称")}
    return bool(master_names and backend_names) and master_names == backend_names


def display_coupon_values(values: list[str]) -> str:
    """Keep every coupon-level value visible in the Excel result."""
    return "、".join(values)


def normalize_dish_note(value_: str) -> str:
    """Normalize non-material product naming variants used by the platforms."""
    normalized = norm(value_).replace("风味", "")
    normalized = normalized.replace("原味冰淇淋原味花筒", "原味冰淇淋花筒")
    normalized = normalized.replace("大杯", "大").replace("小杯", "小")
    # 单品括号内的“4块/2根/3串”等包装规格常在一侧省略，不改变菜品本身；
    # 套餐的任选份数、加购数量仍保留参与核对。
    normalized = re.sub(r"[（(]\s*\d+\s*(?:块|根|串|个|杯|份)(?:装)?\s*[）)]", "", normalized)
    normalized = re.sub(r"^产品包含[：:]", "", normalized)
    return re.sub(r"[\s，,、。.;；:：()（）]", "", normalized)


def numeric_facts(value_: str) -> tuple[str, ...]:
    """Numbers are material for quantities, prices, dates and duration rules."""
    return tuple(re.findall(r"\d+(?:\.\d+)?", value_))


def dish_notes_similar(source: str, target: str) -> bool:
    """Allow small naming differences while preserving every quantity."""
    def material_numbers(value_: str) -> tuple[str, ...]:
        return numeric_facts(re.sub(r"[（(]\s*\d+\s*(?:块|根|串|个|杯|份)(?:装)?\s*[）)]", "", value_))

    if material_numbers(source) != material_numbers(target):
        return False
    return SequenceMatcher(None, source, target).ratio() >= 0.88


def dish_group_counts(value_: str) -> tuple[str, ...]:
    """Extract selectable-combo quantities, e.g. 任选4份 / 鸡肉帕尼尼*4."""
    normalized = norm(value_)
    counts = re.findall(r"任选\s*(\d+)", normalized)
    counts.extend(re.findall(r"[×*]\s*(\d+)", normalized))
    return tuple(sorted(counts, key=int))


def dish_core_items(value_: str) -> set[str]:
    """Extract core food names while ignoring platform-only item descriptors."""
    normalized = norm(value_).lower()
    normalized = re.sub(r"产品包含\s*[：:]?", "", normalized)
    normalized = re.sub(r"任选\s*\d+\s*(?:份|杯|个|盒|瓶|串)?", "", normalized)
    normalized = re.sub(r"[×*]\s*\d+", "", normalized)
    normalized = re.sub(r"\d+\s*(?:份|杯|个|盒|瓶|串)", "", normalized)
    items: set[str] = set()
    for part in re.split(r"[，,、/／+＋【】\[\]()（）]", normalized):
        core = part.strip()
        if not core:
            continue
        core = core.replace("风味", "").replace("芝士", "").replace("醇", "").replace("轻盈版", "")
        core = re.sub(r"(?:冰|热|冷热)$", "", core)
        core = re.sub(r"\s+", "", core)
        if len(core) >= 2:
            items.add(core)
    return items


def combo_dish_notes_match(source: str, target: str) -> bool:
    """Match differently formatted combo menus by their core items and selection totals."""
    source_counts, target_counts = dish_group_counts(source), dish_group_counts(target)
    if not source_counts or source_counts != target_counts:
        return False
    source_items, target_items = dish_core_items(source), dish_core_items(target)
    return bool(source_items and source_items == target_items)


def coupon_dish_note_sets_match(
    master_rows: list[tuple[int, dict[str, str]]],
    backend_rows: list[tuple[int, dict[str, str]]],
) -> bool:
    """Compare dish-note sets after removing display-only product wording."""
    master_values, _ = coupon_value_set(master_rows, SIMPLE_ALIASES["菜品备注"][0], "菜品备注")
    backend_values, _ = coupon_value_set(backend_rows, SIMPLE_ALIASES["菜品备注"][1], "菜品备注")
    if not master_values or len(master_values) != len(backend_values):
        return False
    remaining = [normalize_dish_note(item) for item in backend_values]
    for source in (normalize_dish_note(item) for item in master_values):
        candidates = [(index, target) for index, target in enumerate(remaining) if dish_notes_similar(source, target) or combo_dish_notes_match(source, target)]
        if not candidates:
            return False
        best_index, _ = max(candidates, key=lambda item: SequenceMatcher(None, source, item[1]).ratio())
        remaining.pop(best_index)
    return True


def normalize_rule_text(value_: str) -> str:
    """Normalize harmless wording and punctuation differences in rule text."""
    normalized = canonical_rule_source(value_).lower()
    normalized = re.sub(r"^(?:如有疑问[，,]?)?(?:请)?(?:联系|咨询)", "", normalized)
    normalized = normalized.replace("桂花风味酸梅汤", "桂花酸梅汤")
    normalized = normalized.replace("在线客服工作时间", "客服工作时间")
    normalized = normalized.replace("到店消费", "堂食").replace("到店", "堂食")
    normalized = normalized.replace("本券", "").replace("可使用", "").replace("使用", "").replace("消费", "").replace("为", "")
    normalized = normalized.replace("堂食堂食", "堂食")
    return re.sub(r"[\s，,。.;；:：、()（）\[\]【】'\"“”‘’!！?？]", "", normalized)


def canonical_rule_source(value_: str) -> str:
    """Discard an optional product-title preamble and duplicate full coupon-rule blocks."""
    source = norm(value_)
    match = re.search(r"(?:本产品共|本券|可兑换|产品包含|使用有效期)", source)
    source = source[match.start():] if match else source
    # 总表偶尔会在同一单元格中重复拼接整段券说明。第二段通常再次以“本产品共”
    # 开始；保留首段可避免把重复的时间、次数等误判为规则差异。
    repeated = list(re.finditer(r"本产品共\s*\d+\s*张", source))
    if len(repeated) > 1:
        source = source[:repeated[1].start()].rstrip("、,，;； ")
    return source


def critical_rule_numbers(value_: str) -> tuple[str, ...]:
    """Only dates, periods, times and amounts are risk-bearing rule numbers; food counts are checked separately."""
    source = canonical_rule_source(value_)
    patterns = [
        r"20\d{2}[年/-]\d{1,2}[月/-]\d{1,2}(?:日)?",
        r"\d{1,2}:\d{2}",
        r"\d+(?:\.\d+)?\s*(?:天|日|小时|分钟)",
        r"\d+(?:\.\d+)?\s*元",
    ]
    return tuple(token for pattern in patterns for token in re.findall(pattern, source))


def rule_material_facts(value_: str) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Keep risk-bearing facts strict even when the surrounding rule wording differs."""
    normalized = canonical_rule_source(value_).lower()
    scopes = []
    for name, variants in {
        "堂食": ("堂食", "到店", "门店"),
        "外送": ("外送", "宅急送", "配送"),
        "自助点餐": ("自助点餐",),
        "全国": ("全国",),
        "不可": ("不可", "不适用", "不得", "禁止", "不能"),
        "可叠加": ("可叠加",),
        "不可叠加": ("不可叠加",),
    }.items():
        if any(variant in normalized for variant in variants):
            scopes.append(name)
    return critical_rule_numbers(normalized), tuple(scopes)


def coupon_rule_sets_similar(
    master_rows: list[tuple[int, dict[str, str]]],
    backend_rows: list[tuple[int, dict[str, str]]],
    threshold: float = 0.90,
) -> bool:
    """Accept near-identical rule wording without hiding material rule changes."""
    def canonical_values(rows: list[tuple[int, dict[str, str]]], aliases: list[str]) -> list[str]:
        # 同券在总表中可能存在两条仅标点、平台标签或重复拼接不同的说明记录；
        # 按标准化内容去重后再做集合比对，避免一对多造成误判。
        values: dict[str, str] = {}
        for _, row in rows:
            raw = field_raw_value(row, aliases, "其他规则补充说明")
            if raw:
                canonical = canonical_rule_source(raw)
                values.setdefault(normalize_rule_text(canonical), canonical)
        return list(values.values())

    master_values = canonical_values(master_rows, SIMPLE_ALIASES["其他规则补充说明"][0])
    backend_values = canonical_values(backend_rows, SIMPLE_ALIASES["其他规则补充说明"][1])
    if not master_values or len(master_values) != len(backend_values):
        return False
    remaining = [normalize_rule_text(item) for item in backend_values]
    remaining_facts = [rule_material_facts(item) for item in backend_values]
    for raw_source in master_values:
        source, source_facts = normalize_rule_text(raw_source), rule_material_facts(raw_source)
        candidates = [(index, target) for index, target in enumerate(remaining) if source_facts == remaining_facts[index]]
        if not candidates:
            return False
        best_index, best_score = max(
            ((index, SequenceMatcher(None, source, target).ratio()) for index, target in candidates),
            key=lambda item: item[1],
        )
        if best_score < threshold:
            return False
        remaining.pop(best_index)
        remaining_facts.pop(best_index)
    return True


def rule_outline(value_: str) -> str:
    """Remove the product-content clause so it can be checked together with dish notes."""
    outline = canonical_rule_source(value_)
    outline = re.sub(
        r"可兑换.*?(?=\s*(?:产品包含|使用有效期|仅限|本券|不适用|适用范围|$))",
        "",
        outline,
        flags=re.S,
    )
    return re.sub(
        r"产品包含[：:]\s*(.*?)(?=\s*(?:使用有效期|仅限|本券|不适用|适用范围|$))",
        "",
        outline,
        flags=re.S,
    )


def combined_dish_rule_match(
    master_rows: list[tuple[int, dict[str, str]]],
    backend_rows: list[tuple[int, dict[str, str]]],
) -> bool:
    """Allow food details to move between dish notes and rule descriptions without hiding material rule changes."""
    def combined_values(rows: list[tuple[int, dict[str, str]]], dish_aliases: list[str], rule_aliases: list[str]) -> tuple[set[str], tuple[str, ...], str]:
        food_texts, outlines = [], []
        for _, row in rows:
            dish = field_raw_value(row, dish_aliases, "菜品备注")
            rules = field_raw_value(row, rule_aliases, "其他规则补充说明")
            if dish:
                food_texts.append(dish)
            if rules:
                content = product_content(rules)
                if content:
                    food_texts.append(content)
                outlines.append(rule_outline(rules))
        items = set().union(*(dish_core_items(item) for item in food_texts)) if food_texts else set()
        counts = tuple(sorted({count for item in food_texts for count in dish_group_counts(item)}, key=int))
        return items, counts, " ".join(outline for outline in outlines if outline)

    master_dish_aliases, backend_dish_aliases = SIMPLE_ALIASES["菜品备注"]
    master_rule_aliases, backend_rule_aliases = SIMPLE_ALIASES["其他规则补充说明"]
    master_items, master_counts, master_outline = combined_values(master_rows, master_dish_aliases, master_rule_aliases)
    backend_items, backend_counts, backend_outline = combined_values(backend_rows, backend_dish_aliases, backend_rule_aliases)
    if not master_items or master_items != backend_items or master_counts != backend_counts:
        return False
    if not master_outline and not backend_outline:
        return True
    if rule_material_facts(master_outline) != rule_material_facts(backend_outline):
        return False
    return SequenceMatcher(None, normalize_rule_text(master_outline), normalize_rule_text(backend_outline)).ratio() >= 0.90


def field_difference_ranges(master_value: str, backend_value: str) -> tuple[list[tuple[int, int]], list[tuple[int, int]]]:
    """Locate non-matching text ranges in the two source values."""
    matcher = SequenceMatcher(None, master_value, backend_value)
    master_ranges: list[tuple[int, int]] = []
    backend_ranges: list[tuple[int, int]] = []
    for tag, start_a, end_a, start_b, end_b in matcher.get_opcodes():
        if tag != "equal":
            if start_a != end_a:
                master_ranges.append((start_a, end_a))
            if start_b != end_b:
                backend_ranges.append((start_b, end_b))
    return master_ranges, backend_ranges


def field_difference_hint(master_value: str, backend_value: str, limit: int = 120) -> str:
    """Return changed fragments so reviewers do not need to compare full text."""
    master_ranges, backend_ranges = field_difference_ranges(master_value, backend_value)
    master_parts: list[str] = []
    backend_parts: list[str] = []
    for start, end in master_ranges:
        master_parts.append(master_value[start:end])
    for start, end in backend_ranges:
        backend_parts.append(backend_value[start:end])
    def compact(parts: list[str]) -> str:
        result = "、".join(dict.fromkeys(part.strip() for part in parts if part.strip()))
        return (result[:limit] + "…") if len(result) > limit else (result or "（无）")
    return f"差异片段：KFC总表「{compact(master_parts)}」；美团后台「{compact(backend_parts)}」。"


def rich_text_with_difference_highlights(value_: str, ranges: list[tuple[int, int]]) -> str | CellRichText:
    """Keep matching text unchanged and render only changed fragments in red."""
    if not ranges:
        return value_
    rich_text = CellRichText()
    position = 0
    red_font = InlineFont(color="FFFF0000")
    for start, end in ranges:
        if position < start:
            rich_text.append(value_[position:start])
        rich_text.append(TextBlock(red_font, value_[start:end]))
        position = end
    if position < len(value_):
        rich_text.append(value_[position:])
    return rich_text


def price_ratio(original: str | None, meituan_price: str | None) -> str:
    """返回“原价/美团价”的可比较数值；分母为空或为 0 时不可计算。"""
    if not original or not meituan_price:
        return ""
    try:
        denominator = float(meituan_price)
        if denominator == 0:
            return ""
        # 价格统一为两位小数后再计算，避免 Excel 展示格式影响比较。
        return f"{float(original) / denominator:.4f}"
    except (TypeError, ValueError, ZeroDivisionError):
        return ""


def offer_mode_groups(master: Sheet, export_groups: dict[str, list[tuple[int, dict[str, str]]]]) -> dict[str, list[tuple[int, dict[str, str]]]]:
    """Offer 表常在券创建前导出，券ID列存在但为空；以美团渠道、标题和售价建立候选。"""
    offer_name_headers = ["*券名", "券名"]
    if not master.rows or not find_header(master.rows[0], offer_name_headers): return {}
    channel_headers = ["*售卖渠道", "售卖渠道", "渠道"]
    offer_rows = [(number, row) for number, row in zip(master.numbers, master.rows) if "美团" in value(row, channel_headers)]
    linked: dict[str, list[tuple[int, dict[str, str]]]] = {}
    for coupon_id, backend_rows in export_groups.items():
        backend_name, _, conflict = grouped_values(backend_rows, SIMPLE_ALIASES["项目名称"][1], "项目名称")
        backend_price, _, _ = grouped_values(backend_rows, SIMPLE_ALIASES["原价"][1], "原价")
        if not backend_name or conflict: continue
        candidates = [(number, row) for number, row in offer_rows if comparable(value(row, offer_name_headers), "项目名称") == backend_name]
        price_matched = [(number, row) for number, row in candidates if backend_price and comparable(value(row, SIMPLE_ALIASES["原价"][0]), "原价") == backend_price]
        # 标题同名时优先使用售价一致者；没有价格佐证时仅接受唯一标题候选。
        if len(price_matched) == 1: linked[coupon_id] = price_matched
        elif len(candidates) == 1: linked[coupon_id] = candidates
    return linked


def create_total_table_report(master_content: bytes, export_content: bytes, selected_fields: list[str], master_filename: str = "KFC总表", export_filename: str = "美团后台导表") -> tuple[bytes, dict[str, int]]:
    if not selected_fields:
        raise ValueError("请至少选择一个核对字段。")
    unknown = [field for field in selected_fields if field not in SIMPLE_FIELDS]
    if unknown: raise ValueError(f"不支持的核对字段：{'、'.join(unknown)}")
    export = parse_xlsx(export_content)
    if not export.rows: raise ValueError("美团后台导表没有可读取的数据行。")
    if not find_header(export.rows[0], EXPORT_ID_HEADERS): raise ValueError("美团后台导表缺少商家商品id列。")
    export_groups, export_unindexed = group_export(export)
    def master_sheet_score(candidate: Sheet) -> int:
        if not candidate.rows: return -1
        coupon_ids = {value(row, MASTER_ID_HEADERS) for row in candidate.rows if value(row, MASTER_ID_HEADERS)}
        overlap = len(coupon_ids & set(export_groups))
        structural = sum(bool(find_header(candidate.rows[0], aliases[0])) for aliases in SIMPLE_ALIASES.values())
        return overlap * 1000 + structural
    candidates = parse_xlsx_sheets(master_content)
    selected_sheet_name, master = max(candidates, key=lambda item: master_sheet_score(item[1]), default=("", Sheet([], [])))
    if not master.rows: raise ValueError("KFC总表没有可读取的数据行。")
    master_filename = f"{master_filename}（工作表：{selected_sheet_name}）"
    is_offer_mode = bool(find_header(master.rows[0], ["*券名", "券名"]))
    if not find_header(master.rows[0], MASTER_ID_HEADERS) and not is_offer_mode:
        detected = "、".join(list(master.rows[0])[:12]) or "未读取到表头"
        raise ValueError(f"KFC总表未识别到券ID列。支持：优惠券ID、券ID、商家平台券ID、美团券ID、券码。已读取表头：{detected}")
    master_groups: dict[str, list[tuple[int, dict[str, str]]]] = defaultdict(list)
    for number, row in inherited_master_rows(master):
        coupon_id = value(row, MASTER_ID_HEADERS)
        if coupon_id: master_groups[coupon_id].append((number, row))
    offer_fallback = is_offer_mode and not master_groups
    if offer_fallback:
        master_groups = defaultdict(list, offer_mode_groups(master, export_groups))
    elif master_groups:
        # 主工作表优先。仅当某个后台券 ID 在主表完全不存在时，才从其他
        # 含券 ID 的工作表补齐，避免历史工作表覆盖主表的当前数据。
        missing_coupon_ids = set(export_groups) - set(master_groups)
        for sheet_name, candidate in candidates:
            if not missing_coupon_ids or sheet_name == selected_sheet_name or not candidate.rows:
                continue
            if not find_header(candidate.rows[0], MASTER_ID_HEADERS):
                continue
            for number, row in inherited_master_rows(candidate):
                coupon_id = value(row, MASTER_ID_HEADERS)
                if coupon_id in missing_coupon_ids:
                    master_groups[coupon_id].append((number, row))

    # 美团后台导表是本次“已上线商品”的审核范围。总表中不在后台的历史券
    # 不写入结果，避免出现后台时间看似为空的无效记录。
    coupon_ids = sorted(export_groups)
    headers = ["券ID", "美团ID"]
    for field in selected_fields: headers += [f"KFC总表-{field}", f"美团后台-{field}", f"{field}-结果"]
    headers += ["整券结果", "核对说明"]
    workbook = Workbook(); sheet = workbook.active; sheet.title = "核查结果"; sheet.freeze_panes = "A2"; sheet.append(headers)
    colors = {"通过": "008000", "待复核": "FFC000", "不通过": "FF0000"}
    totals = {"通过": 0, "待复核": 0, "不通过": 0, "KFC总表未匹配": 0, "美团导表未匹配": 0, "无法索引": export_unindexed}
    for coupon_id in coupon_ids:
        master_rows, backend_rows = master_groups.get(coupon_id, []), export_groups.get(coupon_id, [])
        combined_texts_match = combined_dish_rule_match(master_rows, backend_rows)
        project_ids = sorted({value(row, EXPORT_PROJECT_HEADERS) for _, row in backend_rows if value(row, EXPORT_PROJECT_HEADERS)})
        row_output: list[str] = [coupon_id, "、".join(project_ids)]; results: list[str] = []; reasons: list[str] = []
        text_highlights: dict[str, tuple[list[tuple[int, int]], list[tuple[int, int]]]] = {}
        for field in selected_fields:
            if field == "原价/美团价":
                master_original, master_original_lines, master_original_conflict = grouped_values(master_rows, SIMPLE_ALIASES["原价"][0], "原价")
                master_price, master_price_lines, master_price_conflict = grouped_values(master_rows, SIMPLE_ALIASES["美团价"][0], "美团价")
                backend_original, backend_original_lines, backend_original_conflict = grouped_values(backend_rows, SIMPLE_ALIASES["原价"][1], "原价")
                backend_price, backend_price_lines, backend_price_conflict = grouped_values(backend_rows, SIMPLE_ALIASES["美团价"][1], "美团价")
                master_value = price_ratio(master_original, master_price)
                backend_value = price_ratio(backend_original, backend_price)
                master_lines = sorted(set(master_original_lines + master_price_lines))
                backend_lines = sorted(set(backend_original_lines + backend_price_lines))
                master_conflict = master_original_conflict or master_price_conflict
                backend_conflict = backend_original_conflict or backend_price_conflict
            else:
                master_aliases, backend_aliases = SIMPLE_ALIASES[field]
                # 核对键始终只有券 ID。相同券 ID 下的多条后台项目记录不再
                # 被误判为“冲突”，而是比较两侧字段值的完整集合。
                master_values, master_lines = coupon_value_set(master_rows, master_aliases, field)
                backend_values, backend_lines = coupon_value_set(backend_rows, backend_aliases, field)
                if field == "使用有效期" and not backend_values:
                    # 美团后台常把有效期写入“其他规则补充说明”而不单列有效期。
                    # 此时提取说明中的日期范围，与总表的有效期按同一规则核对。
                    backend_values, backend_lines = rule_validity_value_set(backend_rows)
                master_value = display_coupon_values(master_values)
                backend_value = display_coupon_values(backend_values)
                master_conflict = backend_conflict = False
            if not master_rows:
                status = "待复核"; totals["KFC总表未匹配"] += 1
                reason = f"{field}：待复核。KFC总表未找到相同券ID的记录。"
            elif not backend_rows:
                status = "待复核"; totals["美团导表未匹配"] += 1
                reason = f"{field}：待复核。美团后台未找到相同券ID的记录。"
            elif master_conflict:
                status = "待复核"; reason = f"{field}：待复核。KFC总表同一券ID下存在多个不同值（第 {'、'.join(map(str, master_lines))} 行），请确认。"
            elif backend_conflict:
                status = "待复核"; reason = f"{field}：待复核。美团后台同一券ID下存在多个不同值（第 {'、'.join(map(str, backend_lines))} 行），请确认。"
            elif field == "菜品备注" and not master_value and not backend_value:
                # 两侧都没有维护菜品备注时，没有内容可比对，也不应把整券
                # 拉入待复核队列。
                status = "通过"; reason = "菜品备注：通过。双方均为空，无需复核。"
            elif not master_value or not backend_value:
                empty_side = "KFC总表" if not master_value else "美团后台"
                status = "待复核"; reason = f"{field}：待复核。{empty_side}字段为空或无法按格式解析。"
            elif offer_fallback:
                status = "待复核"; reason = f"{field}：待复核。Offer券ID为空，当前仅通过名称和价格建立候选关联，需确认是否为同批次券。"
            elif master_value == backend_value:
                status = "通过"; reason = f"{field}：通过。KFC总表与美团后台字段一致。"
            elif field == "项目名称" and coupon_project_name_sets_match(master_rows, backend_rows):
                status = "通过"; reason = "项目名称：通过。仅平台或测试展示标签不同。"
            elif field == "项目名称" and coupon_dish_note_sets_match(master_rows, backend_rows):
                # 名称可以因平台展示文案不同而变化。只要同券 ID 下的菜品备注
                # 完整一致，就以菜品内容作为同一商品的可靠佐证。
                status = "通过"; reason = "项目名称：通过。名称展示不同，但菜品备注一致。"
            elif field == "菜品备注" and coupon_dish_note_sets_match(master_rows, backend_rows):
                status = "通过"; reason = "菜品备注：通过。内容表达不同，但菜品数量和名称一致。"
            elif field == "其他规则补充说明" and coupon_rule_sets_similar(master_rows, backend_rows):
                # 对客服用语、标点等小幅文案差异作容错；相似度不足时仍按
                # 原有不通过规则处理，避免掩盖有效期、使用范围等实质变更。
                status = "通过"; reason = "其他规则补充说明：通过。文字表达不同，但关键规则一致。"
            elif field in {"菜品备注", "其他规则补充说明"} and combined_texts_match:
                status = "通过"; reason = f"{field}：通过。与{'其他规则补充说明' if field == '菜品备注' else '菜品备注'}综合核对一致。"
            else:
                if field in {"其他规则补充说明", "菜品备注"}:
                    text_highlights[field] = field_difference_ranges(master_value, backend_value)
                    summary = "文本内容不一致。"
                elif field in {"美团价", "原价"}:
                    summary = "价格数值不一致。"
                else:
                    summary = "字段值不一致。"
                # 原始值已在相邻对照列中保留。核对说明只提供人工处理所需的
                # 简短结论，避免把长券说明和菜品内容重复写入该列。
                status = "不通过"; reason = f"{field}：不通过。{summary}"
            master_display = master_value or ("未匹配券ID" if not master_rows else "")
            backend_display = backend_value or ("未匹配券ID" if not backend_rows else "")
            row_output += [master_display, backend_display, status]; results.append(status)
            if status != "通过":
                reasons.append(reason)
        overall = "不通过" if "不通过" in results else "待复核" if "待复核" in results else "通过"
        totals[overall] += 1; row_output += [overall, "\n".join(reasons)]; sheet.append(row_output)
        for index, field in enumerate(selected_fields):
            if field not in text_highlights:
                continue
            master_ranges, backend_ranges = text_highlights[field]
            sheet.cell(sheet.max_row, 3 + index * 3).value = rich_text_with_difference_highlights(row_output[2 + index * 3], master_ranges)
            sheet.cell(sheet.max_row, 4 + index * 3).value = rich_text_with_difference_highlights(row_output[3 + index * 3], backend_ranges)
    sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    for cell in sheet[1]: cell.font = Font(bold=True, color="65738C")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value in colors:
                cell.font = Font(color=colors[cell.value])
        reason_cell = row[-1]; alignment = copy(reason_cell.alignment); alignment.wrap_text = True; alignment.vertical = "top"; reason_cell.alignment = alignment
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = 52 if header == "核对说明" else 22
    stream = io.BytesIO(); workbook.save(stream)
    return stream.getvalue(), totals
