from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pdfplumber

from .certificates import CertificateRecord


CERT_NAMES = {
    "NATIONALITY / REGISTRY": "01_INTERNATIONAL_REGISTRY_CERT",
    "TONNAGE": "02_TONNAGE_CERT",
    "LOAD LINE": "03_LOAD_LINE",
    "IOPP": "04_IOPP_FORM_B",
    "SAFETY CONSTRUCTION": "05_SAFETY_CONSTRUCTION",
    "SAFETY EQUIPMENT": "06_SAFETY_EQUIPMENT_FORM_E",
    "SAFETY RADIO": "07_SAFETY_RADIO_FORM_R",
    "MINIMUM SAFE MANNING": "08_MINIMUM_MANNING",
    "CLC": "09_CLC",
    "DOC": "10_DOC",
    "SMC": "11_SMC",
    "ISSC": "12_ISSC",
    "WRC": "13_WRC",
}

RANK_TO_DUTY = {
    "CAPTAIN": "51-船长", "C / OFFICER": "52-大副", "2 / OFFICER": "53-二副", "3 / OFFICER": "54-三副",
    "C / ENGINEER": "61-轮机长", "2 / ENGINEER": "62-大管轮", "3 / ENGINEER": "63-二管轮", "4 / ENGINEER": "64-三管轮",
}


def _date(value: object) -> str:
    text = str(value or "").strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def _key(value: object) -> str:
    return " ".join(str(value or "").replace("\n", " ").upper().split())


def find_particulars_a(files: list[Path]) -> Path | None:
    candidates = [p for p in files if p.suffix.lower() == ".pdf" and "PARTICULARS" in p.name.upper()]
    return next((p for p in candidates if "(A)" in p.name.upper()), candidates[0] if candidates else None)


def read_particulars(path: Path | None) -> tuple[dict[str, str], dict[str, tuple[str, str, str]], dict[str, tuple[str, str]]]:
    info: dict[str, str] = {}
    certificate_dates: dict[str, tuple[str, str, str]] = {}
    competencies: dict[str, tuple[str, str]] = {}
    if not path:
        return info, certificate_dates, competencies
    with pdfplumber.open(path) as document:
        tables = document.pages[0].extract_tables() if document.pages else []
    if tables:
        for row in tables[0]:
            for column in range(0, len(row) - 1, 2):
                name = _key(row[column])
                value = " ".join(str(row[column + 1] or "").split())
                if name and value:
                    info[name] = value
    if len(tables) > 1:
        for row in tables[1][1:]:
            for offset in (0, 4):
                if len(row) < offset + 4:
                    continue
                category = CERT_NAMES.get(_key(row[offset]))
                if category:
                    certificate_dates[category] = (_date(row[offset + 1]), _date(row[offset + 2]), _date(row[offset + 3]))
    if len(tables) > 2:
        for row in tables[2][1:]:
            for offset in (0, 2):
                if len(row) < offset + 2:
                    continue
                duty = RANK_TO_DUTY.get(_key(row[offset]))
                value = str(row[offset + 1] or "").strip()
                match = re.search(r"(.+?)\s*/\s*(\d{2}-\d{2}-\d{4})\s*$", value)
                if duty and match:
                    competencies[duty] = (match.group(1).strip(), _date(match.group(2)))
    return info, certificate_dates, competencies


def apply_certificate_dates(records: list[CertificateRecord], values: dict[str, tuple[str, str, str]]) -> None:
    for record in records:
        if record.category not in values:
            continue
        issue, expiry, annual = values[record.category]
        record.issue_date = issue or record.issue_date
        record.expiry_date = expiry or record.expiry_date
        record.annual_verification_date = annual or record.annual_verification_date


def apply_competencies(crew_rows: list[list[str]], values: dict[str, tuple[str, str]]) -> None:
    for row in crew_rows:
        certificate = values.get(row[3])
        if certificate:
            row[10], row[11] = certificate


def taskbook_fields(info: dict[str, str], mail_text: str, ship_name: str) -> dict[str, str]:
    def pick(*names: str) -> str:
        return next((info.get(name, "") for name in names if info.get(name)), "")

    eta = ""
    eta_match = re.search(r"ETA\s+(\d{1,2}[./-]\d{1,2}[./-]\d{4})\s+(?:AT\s+)?(\d{3,4})\s*LT", mail_text, re.I)
    if eta_match:
        eta = f"{eta_match.group(1)} {eta_match.group(2)}LT"
    next_port = ""
    next_match = re.search(r"NEXT\s+PORT\s*[:：]\s*([A-Z][A-Z ]{1,30})", mail_text, re.I)
    if next_match:
        next_port = next_match.group(1).strip()
    draft = ""
    draft_match = re.search(r"ARRIVAL\s+DRAFT\s*\(([^)]+)\)\s*&?\s*DEPARTURE\s+DRAFT\s*\(([^)]+)\)", mail_text, re.I)
    if draft_match:
        def compact(value: str) -> str:
            numbers = re.findall(r"\d+(?:\.\d+)?", value)
            return "/".join(numbers[:3])
        draft = f"F/M/A 抵 {compact(draft_match.group(1))}；离 {compact(draft_match.group(2))}m"
    return {
        "船名中文": ship_name,
        "英文": ship_name,
        "国籍": pick("FLAG", "NATIONALITY", "PORT OF REGISTRY"),
        "建造年月": pick("BUILT YEAR &MONTH"),
        "船型": pick("SHIP-TYPE"),
        "呼号": pick("CALL SIGN"),
        "IMO": pick("IMO NO."),
        "MMSI": pick("MMSI", "MMSI NO."),
        "总吨": pick("GROSS TONNAGE"),
        "净吨": pick("NET TONNAGE"),
        "载重吨": pick("DWT /TEU"),
        "总长": pick("LOA"),
        "型宽": pick("BREADTH MOULDED"),
        "型深": pick("DEPTH MOULDED"),
        "夏季满吃水：": pick("SUMMER MEAN DRAFT"),
        "ETA": eta,
        "下一港：": next_port,
        "吃水": draft,
        "货物/修理情况:": "卸煤" if "DISCHARG" in mail_text.upper() and "COAL" in mail_text.upper() else "",
    }
