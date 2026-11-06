#!/usr/bin/env python3
"""Generate a bonded-terminal operation forecast from a vessel particulars A form."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

try:
    from openpyxl import load_workbook
except ImportError as exc:  # pragma: no cover - environment diagnostic
    raise SystemExit(
        "Missing dependency. Run with the Codex bundled Python runtime containing "
        f"openpyxl: {exc}"
    )

from parse_a_table import parse_particulars


DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "综保区码头船舶作业预报单.xlsx"
REQUIRED = ("ship_name", "imo", "loa", "dwt")
FLAG_CN = {
    "PANAMA": "巴拿马", "CHINA": "中国", "LIBERIA": "利比里亚",
    "MARSHALL ISLANDS": "马绍尔群岛", "SINGAPORE": "新加坡",
    "HONG KONG": "中国香港", "HONGKONG": "中国香港", "HK": "中国香港",
    "MALTA": "马耳他", "GREECE": "希腊", "CYPRUS": "塞浦路斯",
    "INDONESIA": "印度尼西亚", "VIET NAM": "越南", "VIETNAM": "越南",
    "PHILIPPINES": "菲律宾", "MALAYSIA": "马来西亚", "INDIA": "印度",
    "JAPAN": "日本", "KOREA": "韩国", "BRAZIL": "巴西",
    "UNITED STATES": "美国", "USA": "美国", "AUSTRALIA": "澳大利亚",
    "NEW ZEALAND": "新西兰", "PORTUGAL": "葡萄牙", "SPAIN": "西班牙",
    "GERMANY": "德国", "TURKEY": "土耳其", "THAILAND": "泰国",
    "MYANMAR": "缅甸", "NORWAY": "挪威", "DENMARK": "丹麦",
    "BAHRAIN": "巴林", "OMAN": "阿曼", "ECUADOR": "厄瓜多尔", "TONGA": "汤加",
}


def without_unit(value: object, unit: str) -> str:
    text = re.sub(rf"\s*{re.escape(unit)}\s*$", "", str(value or ""), flags=re.IGNORECASE)
    text = text.replace(",", "").replace("，", "").strip()
    if re.fullmatch(r"-?\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def flag_cn(value: object) -> str:
    text = str(value or "").strip().upper()
    if text in FLAG_CN:
        return FLAG_CN[text]
    compact = text.replace(" ", "")
    for key, result in FLAG_CN.items():
        if key.replace(" ", "") == compact:
            return result
    return text


def extract_particulars(source: Path) -> dict[str, str]:
    if not source.exists():
        raise FileNotFoundError(f"Input does not exist: {source}")
    raw = parse_particulars(str(source))
    values = {
        "ship_name": raw.get("ship_name", ""), "flag": raw.get("flag", ""),
        "built_year": raw.get("built", ""), "ship_type": raw.get("shiptype", ""),
        "imo": raw.get("imo", ""), "gross_tonnage": raw.get("gt", ""),
        "net_tonnage": raw.get("nt", ""), "dwt": raw.get("dwt", ""),
        "call_sign": raw.get("callsign", ""), "loa": raw.get("loa", ""),
        "breadth": raw.get("breadth", ""), "depth": raw.get("depth", ""),
    }
    missing = [field for field in REQUIRED if not values.get(field)]
    if missing:
        raise ValueError("Required fields not found in particulars A form: " + ", ".join(missing))
    return values


def media_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(
            1
            for name in archive.namelist()
            if name.startswith("xl/media/") and not name.endswith("/")
        )


def generate(
    particulars: dict[str, str], template: Path, output: Path, declaration_date: str
) -> dict[str, object]:
    if not re.fullmatch(r"\d{8}", declaration_date):
        raise ValueError("Declaration date must use YYYYMMDD")
    if not template.exists():
        raise FileNotFoundError(f"Template does not exist: {template}")

    template_media = media_count(template)
    workbook = load_workbook(template, data_only=False)
    sheet = workbook.active

    loa = without_unit(particulars.get("loa", ""), "M")
    breadth = without_unit(particulars.get("breadth", ""), "M")
    depth = without_unit(particulars.get("depth", ""), "M")
    dwt = without_unit(particulars.get("dwt", ""), "MT")

    mapped = {
        "J6": declaration_date,
        "B7": f"({particulars['ship_name']})- 加载计划",
        "C9": particulars["ship_name"],
        "D9": flag_cn(particulars.get("flag", "")),
        "E9": "中远海运",
        "F9": "中远海运",
        "H9": loa,
        "I9": dwt,
        "C15": particulars.get("imo", ""),
        "F15": particulars.get("call_sign", ""),
        "J15": particulars.get("built_year", ""),
        "C16": loa,
        "F16": breadth,
        "J16": depth,
        "C17": without_unit(particulars.get("gross_tonnage", ""), "MT"),
        "F17": without_unit(particulars.get("net_tonnage", ""), "MT"),
        "J17": dwt,
    }
    for cell, value in mapped.items():
        sheet[cell] = value

    blank_cells = [
        "B9", "G9", "J9", "K9",
        "B11", "C11", "D11", "E11", "F11", "G11",
    ]
    blank_cells.extend(f"{column}18" for column in "CDEFGHIJK")
    blank_cells.extend(f"{column}{row}" for row in (22, 23) for column in "CDEFGHIJK")
    blank_cells.append("C24")
    for cell in blank_cells:
        sheet[cell] = None

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)

    result_media = media_count(output)
    checked = load_workbook(output, data_only=False)
    checked_sheet = checked.active
    for cell, expected in mapped.items():
        if checked_sheet[cell].value != expected:
            raise AssertionError(
                f"Verification failed at {cell}: {checked_sheet[cell].value!r} != {expected!r}"
            )
    nonblank = {cell: checked_sheet[cell].value for cell in blank_cells if checked_sheet[cell].value not in (None, "")}
    if nonblank:
        raise AssertionError(f"Required blank cells contain values: {nonblank}")
    if result_media != template_media:
        raise AssertionError(
            f"Embedded media changed: template={template_media}, output={result_media}"
        )

    formulas = [
        cell.coordinate
        for row in checked_sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    ]
    return {
        "status": "ok",
        "output": str(output.resolve()),
        "declaration_date": declaration_date,
        "ship_name": particulars["ship_name"],
        "extracted": particulars,
        "written_cells": mapped,
        "verified_blank_cells": blank_cells,
        "template_media_count": template_media,
        "output_media_count": result_media,
        "formula_cells": formulas,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="Particulars A .doc/.docx/.xls/.xlsx/.xlsm/.pdf")
    parser.add_argument("--output", required=True, type=Path, help="Output .xlsx path")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument(
        "--date",
        dest="declaration_date",
        default=None,
        help="Declaration date in YYYYMMDD; defaults to today in Asia/Shanghai",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    declaration_date = args.declaration_date or datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y%m%d")
    try:
        particulars = extract_particulars(args.input.resolve())
        summary = generate(
            particulars,
            args.template.resolve(),
            args.output.resolve(),
            declaration_date,
        )
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
