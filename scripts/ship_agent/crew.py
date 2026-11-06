from __future__ import annotations

import csv
import json
import re
from datetime import datetime, time
from pathlib import Path
from random import Random
from typing import Any

from openpyxl import load_workbook


RESOURCE_DIR = Path(__file__).resolve().parents[2] / "references" / "mappings"


def _load_map(name: str) -> dict[str, str]:
    try:
        return json.loads((RESOURCE_DIR / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


NATIONALITY = {
    "CN": "CN-中国", "CHN": "CN-中国", "CHINA": "CN-中国",
    "VN": "VN-越南", "VIETNAM": "VN-越南", "PH": "PH-菲律宾", "PHILIPPINES": "PH-菲律宾",
    "ID": "ID-印度尼西亚", "INDONESIA": "ID-印度尼西亚", "IN": "IN-印度", "INDIA": "IN-印度",
    "BD": "BD-孟加拉国", "BANGLADESH": "BD-孟加拉国", "MM": "MM-缅甸", "MYANMAR": "MM-缅甸",
    "RU": "RU-俄罗斯", "RUSSIA": "RU-俄罗斯", "KR": "KR-韩国", "KOREA": "KR-韩国",
    "JP": "JP-日本", "JAPAN": "JP-日本", "US": "US-美国", "USA": "US-美国",
    "EG": "EG-埃及", "EGYPT": "EG-埃及", "EGYPTIAN": "EG-埃及",
    "UA": "UA-乌克兰", "UKRAINE": "UA-乌克兰", "UKRAINIAN": "UA-乌克兰",
    "TZ": "TZ-坦桑尼亚", "TANZANIA": "TZ-坦桑尼亚", "TANZANIAN": "TZ-坦桑尼亚",
    "CY": "CY-塞浦路斯", "CYPRUS": "CY-塞浦路斯", "TR": "TR-土耳其", "TURKEY": "TR-土耳其",
    "IQ": "IQ-伊拉克", "IRAQ": "IQ-伊拉克", "KZ": "KZ-哈萨克斯坦", "KAZAKHSTAN": "KZ-哈萨克斯坦",
}
NATIONALITY.update(_load_map("nationality_map.json"))
PORT_MAP = _load_map("port_map.json")
POC_COUNTRY = _load_map("poc-country-map.json")

PORT_OVERRIDES = {
    "HAY POINT": "AUHAY-海角(HAY POINT)",
    "ZHOUSHAN": "CNZOS-舟山(Zhoushan)",
    "CAOFEIDIAN": "CNCFD-曹妃甸港",
    "HUANGHUA": "CNHUH-黄骅(Huanghua)",
    "RIZHAO": "CNRZH-日照(Rizhao)",
    "SINGAPORE": "SGSIN-新加坡(Singapore)",
    "NINGBO": "CNNBO-CNNBO-宁波(Ningbo)",
}

DUTY = {
    "MASTER": "51-船长", "CAPT": "51-船长", "CAPTAIN": "51-船长",
    "C/O": "52-大副", "CHIEF OFFICER": "52-大副", "1/O": "52-大副", "FIRST OFFICER": "52-大副",
    "2/O": "53-二副", "SECOND OFFICER": "53-二副", "3/O": "54-三副", "THIRD OFFICER": "54-三副",
    "C/E": "61-轮机长", "CHIEF ENGINEER": "61-轮机长", "1/E": "62-大管轮", "FIRST ENGINEER": "62-大管轮",
    "2/E": "62-大管轮", "SECOND ENGINEER": "62-大管轮", "3/E": "63-二管轮", "THIRD ENGINEER": "63-二管轮",
    "4/E": "64-三管轮", "FOURTH ENGINEER": "64-三管轮", "A/B": "56-高级值班水手", "AB": "56-高级值班水手", "ABLE SEAMAN": "56-高级值班水手",
    "BOSUN": "55-值班水手", "BSN": "55-值班水手", "COOK": "65-值班机工", "STEWARD": "65-值班机工",
    "FITTER": "65-值班机工", "WIPER": "65-值班机工", "MESSMAN": "65-值班机工", "MOTORMAN": "65-值班机工",
    "ETO": "66-高级值班机工", "E/E": "66-高级值班机工", "ELECTRICIAN": "66-高级值班机工",
}


def _norm_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9一-龥]", "", str(value or "").lower())


def _rows(path: Path) -> list[dict[str, Any]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        values = list(sheet.values)
        return _tabular_rows(values)
    if path.suffix.lower() == ".xls":
        import xlrd

        workbook = xlrd.open_workbook(path)
        sheet = workbook.sheet_by_index(0)
        values = []
        for row_index in range(sheet.nrows):
            row = []
            for column_index in range(sheet.ncols):
                cell = sheet.cell(row_index, column_index)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(value, workbook.datemode)
                row.append(value)
            values.append(tuple(row))
        return _tabular_rows(values)
    if path.suffix.lower() == ".pdf":
        import pdfplumber
        rows: list[dict[str, Any]] = []
        with pdfplumber.open(path) as document:
            for page in document.pages:
                for table in page.extract_tables():
                    if not table:
                        continue
                    headers = [str(value or "") for value in table[0]]
                    if not any("PORT" in header.upper() for header in headers):
                        continue
                    rows.extend(dict(zip(headers, row)) for row in table[1:] if row and any(row))
        return rows
    raise ValueError(f"仅支持 CSV/XLS/XLSX Crew List 或 Port of Call，收到: {path.name}")


def _tabular_rows(values: list[tuple[Any, ...]]) -> list[dict[str, Any]]:
    if not values:
        return []
    header_index = 0
    for index, row in enumerate(values):
        joined = " ".join(str(x or "") for x in row)
        if "Family name" in joined or "Rank or rating" in joined or "Port of Call" in joined:
            header_index = index
            break
    headers = [str(x or "") for x in values[header_index]]
    rows: list[dict[str, Any]] = []
    imo_layout = any("Family name" in header for header in headers) and any("Rank or rating" in header for header in headers)
    for row in values[header_index + 1:]:
        if not any(x not in (None, "") for x in row):
            continue
        # IMO crew list data rows start with a numeric sequence in column B.
        if header_index and not isinstance(row[1] if len(row) > 1 else None, (int, float)):
            continue
        if imo_layout:
            rows.append({
                "No.": row[1], "Family name": row[2], "Given names": row[3], "Sex": row[4],
                "Rank": row[5], "Nationality": row[6], "Place of Birth": row[8], "Date of Birth": row[10],
                "Passport": row[11], "Seaman's Book": row[13], "Port of Embarkation": row[15],
                "Date of Embarkation": row[16],
            })
        else:
            rows.append(dict(zip(headers, row)))
    return rows


def find_input(files: list[Path], kind: str) -> Path | None:
    candidates = []
    for path in files:
        if path.suffix.lower() not in {".csv", ".xls", ".xlsx", ".xlsm", ".pdf"}:
            continue
        name = path.name.lower()
        if kind == "crew" and any(token in name for token in ("crew", "船员", "imo")):
            candidates.append(path)
        if kind == "port" and any(token in name for token in ("port", "poc", "call", "last 10", "last ten", "港口", "靠港")):
            candidates.append(path)
    if not candidates:
        return None
    if kind == "crew":
        return sorted(candidates, key=lambda p: (p.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}, p.name.lower()))[0]
    return sorted(candidates, key=lambda p: (not any(token in p.name.lower() for token in ("last 10", "last ten", "poc")), p.suffix.lower() == ".pdf", p.name.lower()))[0]


def _get(row: dict[str, Any], *keys: str) -> Any:
    normalized = {_norm_header(k): v for k, v in row.items()}
    for key in keys:
        key_norm = _norm_header(key)
        for header, value in normalized.items():
            if key_norm in header or header in key_norm:
                return value
    return ""


def normalize_name(value: Any, nationality: str) -> str:
    text = str(value or "").strip()
    if nationality == "CN":
        chinese = "".join(re.findall(r"[一-龥]+", text))
        return chinese or text
    return text.upper()


def normalize_gender(value: Any) -> str:
    text = str(value or "").strip().lower()
    if text in {"f", "female", "女", "2"}:
        return "2-女"
    return "1-男"


def normalize_date(value: Any, compact: bool = False) -> str:
    if value in (None, ""):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y%m%d" if compact else "%Y-%m-%d")
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d.%m.%Y", "%y-%m-%d", "%Y%m%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d" if compact else "%Y-%m-%d")
        except ValueError:
            continue
    return text


def duty_code(value: Any, index: int = 0) -> str:
    text = str(value or "").strip().upper()
    for key, result in DUTY.items():
        if key in text:
            return result
    if any(token in text for token in ("E", "ENGINE", "机", "轮")):
        return "66-高级值班机工" if index < 3 else "65-值班机工"
    return "56-高级值班水手" if index < 3 else "55-值班水手"


def convert_crew(path: Path, embark_date: Any = "", embark_port: str = "") -> list[list[str]]:
    output: list[list[str]] = []
    for index, row in enumerate(_rows(path), start=1):
        nat_raw = str(_get(row, "Nationality", "国籍", "Nation") or "").strip().upper()
        nat_code = nat_raw.split()[0] if nat_raw else ""
        nat = NATIONALITY.get(nat_code, NATIONALITY.get(nat_raw, ""))
        if nat.startswith("CN-"):
            nat_code = "CN"
        family = _get(row, "Family name", "姓")
        given = _get(row, "given names", "Given name", "名")
        name_value = f"{family} {given}".strip() if family or given else _get(row, "Full Name", "Name", "姓名")
        name = normalize_name(name_value, nat_code)
        passport = _get(row, "Passport", "护照")
        seaman = _get(row, "Seaman's Book", "Seaman Book", "海员证")
        document = seaman if nat_code == "CN" and seaman else passport
        birth_country = nat.split("-", 1)[1] if "-" in nat else ""
        output.append([
            str(index), name, normalize_gender(_get(row, "Gender", "Sex", "性别")),
            duty_code(_get(row, "Rank", "Position", "职务"), index - 1), nat,
            normalize_date(_get(row, "Date of Birth", "Date and place of birth", "Birth", "出生日期"), compact=True),
            birth_country,
            "17-海员证" if nat_code == "CN" else "14-普通护照", str(document or ""), "", "", "", normalize_date(embark_date, compact=True), embark_port, "",
        ])
        row_embark_date = normalize_date(_get(row, "Date of Embarkation", "Join Date", "登船日期"), compact=True)
        row_embark_port = str(_get(row, "Port of Embarkation", "Join Port", "登船口岸") or "")
        output[-1][13] = row_embark_date or normalize_date(embark_date, compact=True)
        output[-1][14] = match_port(row_embark_port or embark_port)
    return output


def _port_code(row: dict[str, Any]) -> str:
    value = str(_get(row, "UNLOCODE", "UN LOCODE", "UNCTAD LOCODE", "Port Code", "Code", "港口代码") or "").upper().replace(" ", "")
    return value


def match_port(value: Any, code: str = "") -> str:
    raw = str(value or "").strip().upper().replace(",", " ")
    normalized_code = code.upper().replace(" ", "")
    if normalized_code == "CNNGB":
        normalized_code = "CNNBO"
    if normalized_code in PORT_MAP:
        return PORT_MAP[normalized_code]
    for key, result in PORT_OVERRIDES.items():
        if key in raw:
            return result
    clean = re.sub(r"\b(CHINA|AUSTRALIA|PRC|USA|UK)\b", "", raw).strip()
    for full in PORT_MAP.values():
        if len(clean) >= 4 and clean in full.upper():
            return full
    return ""


def convert_ports(path: Path) -> list[list[str]]:
    output: list[list[str]] = []
    rng = Random(20260809)
    for index, row in enumerate(_rows(path), start=1):
        arrival = normalize_date(_get(row, "Arrival", "Arrival Date", "进港", "抵港")).replace("-", "/")
        departure = normalize_date(_get(row, "Departure", "Departure Date", "离港", "离港时间")).replace("-", "/")
        arrival_time = time(rng.randint(0, 11), rng.randint(0, 59), rng.randint(0, 59))
        departure_time = time(rng.randint(12, 23), rng.randint(0, 59), rng.randint(0, 59))
        port = str(_get(row, "Port", "Port Name", "港口", "停靠港口") or "").strip()
        code = _port_code(row)
        mapped_port = match_port(port, code)
        country_raw = str(_get(row, "Country", "国家", "Country/Area") or "").strip().upper()
        if not country_raw and "," in port:
            country_raw = port.rsplit(",", 1)[-1].strip().upper()
        country = POC_COUNTRY.get(country_raw, NATIONALITY.get(country_raw, NATIONALITY.get(code[:2], "")))
        output.append([
            str(index), f"{arrival} {arrival_time:%H:%M:%S}" if arrival else "", f"{departure} {departure_time:%H:%M:%S}" if departure else "",
            country, "1-1级", "", mapped_port or (f"{code}-{port}" if code else port), "1-1级",
        ])
    return output
