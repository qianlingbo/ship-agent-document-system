from __future__ import annotations

import re
import unicodedata
from pathlib import Path


GENERIC_WORDS = {"certificate", "cert", "document", "docs", "scan", "copy", "附件", "资料", "文件", "pdf"}


def normalize_text(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).upper().split())


def safe_ship_name(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "_", value).strip(" .")
    return value[:100] or "UNKNOWN_SHIP"


def extract_ship_name(subject: str, body: str = "", filenames: list[str] | None = None) -> str | None:
    text = f"{subject}\n{body}"
    patterns = [
        r"(?:SHIP\s*NAME|VESSEL|船名)\s*[:：]\s*([A-Za-z0-9][A-Za-z0-9 _.-]{2,})",
        r"\b(?:M/V|MV|S/S|SS)\s+([A-Za-z0-9][A-Za-z0-9 _.-]{2,})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return safe_ship_name(match.group(1).strip(" .-_"))
    subject_clean = re.sub(r"\b(?:RE|FW|FWD)\s*:\s*", "", subject, flags=re.I).strip()
    if subject_clean and len(subject_clean) >= 3 and not re.fullmatch(r"[\W_]+", subject_clean):
        candidate = re.sub(r"\[[^]]+\]", "", subject_clean).strip(" -_")
        if candidate and not re.search(r"(crew|certificate|port of call|documents?)", candidate, re.I):
            return safe_ship_name(candidate)
    for filename in filenames or []:
        stem = Path(filename).stem
        tokens = [t for t in re.split(r"[ _.-]+", stem) if t and t.lower() not in GENERIC_WORDS]
        if tokens:
            candidate = " ".join(tokens)
            if len(candidate) >= 3:
                return safe_ship_name(candidate)
    return None


CERTIFICATE_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("01_INTERNATIONAL_REGISTRY_CERT", ("INTERNATIONAL REGISTRY", "REGISTRY CERT", "CERTIFICATE OF REGISTRY", "NATIONALITY CERT")),
    ("02_TONNAGE_CERT", ("TONNAGE",)),
    ("03_LOAD_LINE", ("LOAD LINE",)),
    ("04_IOPP_FORM_B", ("IOPP", "I.O.P.P", "FORM A", "FORM B")),
    ("05_SAFETY_CONSTRUCTION", ("SAFETY CONSTRUCTION",)),
    ("06_SAFETY_EQUIPMENT_FORM_E", ("SAFETY EQUIPMENT", "FORM E")),
    ("07_SAFETY_RADIO_FORM_R", ("SAFETY RADIO", "FORM R")),
    ("08_MINIMUM_MANNING", ("MINIMUM MANNING", "SAFE MANNING")),
    ("09_CLC", ("CLC", "BUNKER OIL",)),
    ("10_DOC", ("D.O.C", "DOC WITH", "DOC OF COMPLIANCE", "DOCUMENT OF COMPLIANCE")),
    ("11_SMC", ("S.M.C", "SMC", "SAFETY MANAGEMENT CERT")),
    ("12_ISSC", ("I.S.S.C", "ISSC", "SHIP SECURITY CERT")),
    ("13_WRC", ("WRC", "REMOVAL OF WRECK", "REMOVEL OF WRECK")),
]


def classify_document(path: str | Path, text: str = "") -> str:
    haystack = normalize_text(f"{path} {text}")
    for category, keywords in CERTIFICATE_RULES:
        if any(normalize_text(keyword) in haystack for keyword in keywords):
            return category
    if any(token in haystack for token in ("CREW LIST", "CREWLIST", "IMO CREW", "船员名单")):
        return "CREW_LIST"
    if any(token in haystack for token in ("PORT OF CALL", "POC", "港口挂靠", "靠港")):
        return "PORT_OF_CALL"
    if any(token in haystack for token in ("COMPETENCY", "CERT OF COMPETENCY", "COC", "适任")):
        return "CREW_COMPETENCY"
    if "DANGEROUS" in haystack or "DG" in haystack or "危险品" in haystack:
        return "DANGEROUS_GOODS"
    return "OTHER"
