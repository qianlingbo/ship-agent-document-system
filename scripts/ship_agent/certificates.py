from __future__ import annotations

import json
import hashlib
import re
import shutil
import subprocess
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from pypdf import PdfReader, PdfWriter

from .classify import CERTIFICATE_RULES, classify_document


MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec"
DATE_PATTERN = re.compile(
    rf"(?<!\d)(?:\d{{1,2}}(?:st|nd|rd|th)\s+day\s+of\s+(?:{MONTHS})\s+\d{{4}}|\d{{1,2}}\s+(?:{MONTHS})\s+\d{{4}}|(?:{MONTHS})\s+\d{{1,2}},?\s+\d{{4}}|"
    r"\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2}[./-]\d{4}|(?:19|20)\d{6}|"
    r"\d{4}年\d{1,2}月\d{1,2}日)(?!\d)",
    re.I,
)


@dataclass
class CertificateRecord:
    source: str
    category: str
    certificate_number: str = ""
    issue_date: str = ""
    expiry_date: str = ""
    annual_verification_date: str = ""
    pages: int = 0
    notes: str = ""


def _normalize_date(raw: str) -> str:
    raw = " ".join(raw.strip().replace(",", "").split())
    raw = re.sub(r"^(\d{1,2})(?:st|nd|rd|th)\s+day\s+of\s+", r"\1 ", raw, flags=re.I)
    chinese = re.fullmatch(r"(\d{4})年(\d{1,2})月(\d{1,2})日", raw)
    if chinese:
        raw = "-".join(chinese.groups())
    formats = ("%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y%m%d")
    for fmt in formats:
        try:
            parsed = datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
        if 1900 <= parsed.year <= 2100:
            return parsed.isoformat()
    return ""


def _date_tokens(text: str, start: int = 0, end: int | None = None) -> list[tuple[int, str]]:
    result = []
    matches = DATE_PATTERN.finditer(text, start) if end is None else DATE_PATTERN.finditer(text, start, end)
    for match in matches:
        value = _normalize_date(match.group(0))
        if value:
            result.append((match.start(), value))
    return result


def extract_pdf_text(path: str | Path) -> str:
    try:
        reader = PdfReader(str(path))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        if text.strip() or not shutil.which("pdftoppm") or not shutil.which("tesseract"):
            return text
        with TemporaryDirectory(prefix="ship-cert-ocr-") as temporary:
            prefix = str(Path(temporary) / "page")
            render = subprocess.run(
                ["pdftoppm", "-png", "-r", "200", "-f", "1", "-l", str(min(len(reader.pages), 6)), str(path), prefix],
                capture_output=True, text=True,
            )
            if render.returncode:
                return ""
            parts = []
            for image in sorted(Path(temporary).glob("page-*.png")):
                ocr = subprocess.run(["tesseract", str(image), "stdout", "-l", "eng+chi_sim"], capture_output=True, text=True)
                if ocr.returncode == 0:
                    parts.append(ocr.stdout)
            return "\n".join(parts)
    except Exception:
        return ""


def extract_certificate_dates(text: str) -> tuple[str, str, str]:
    if not text:
        return "", "", ""
    upper = text.upper()
    annual_start = next((pos for phrase in ("ENDORSEMENT FOR ANNUAL", "ENDORSEMENT FOR INTERMEDIATE", "ANNUAL VERIFICATION", "年度检验", "年检") if (pos := upper.find(phrase)) >= 0), len(text))
    main_dates = _date_tokens(text, 0, annual_start)
    annual_dates = _date_tokens(text, annual_start) if annual_start < len(text) else []

    expiry_anchor = max((upper.find(phrase) for phrase in ("VALID UNTIL", "VALIDITY UNTIL", "DATE OF EXPIRY", "EXPIRATION", "有效期至") if upper.find(phrase) >= 0), default=-1)
    expiry_candidates = [value for pos, value in main_dates if pos > expiry_anchor] if expiry_anchor >= 0 else []
    expiry = max(expiry_candidates, default="")

    issue_anchor = min((pos for phrase in ("DATE OF ISSUE", "ISSUED AT", "ISSUED ON", "签发日期") if (pos := upper.find(phrase)) >= 0), default=-1)
    issue_candidates = [value for pos, value in main_dates if pos > issue_anchor and value != expiry] if issue_anchor >= 0 else [value for _, value in main_dates if value != expiry]
    issue = max(issue_candidates, default="")

    today = date.today().isoformat()
    annual_valid = [value for _, value in annual_dates if value <= today and value != expiry]
    annual = max(annual_valid, default="")
    return issue, expiry, annual


def extract_certificate_number(text: str) -> str:
    patterns = (
        r"CERTIFICATE\s*(?:NO\.?|NUMBER)\s*[:：]?\s*([A-Z0-9][A-Z0-9./_-]{3,49})",
        r"OFFICIAL\s*(?:NO\.?|NUMBER)\s*[:：]?\s*([A-Z0-9][A-Z0-9./_-]{3,49})",
        r"证书编号\s*[:：]?\s*([A-Z0-9][A-Z0-9./_-]{3,49})",
    )
    upper = text.upper()
    for pattern in patterns:
        match = re.search(pattern, upper, re.I)
        if match:
            return match.group(1).strip(" .")[:50]
    return ""


def inspect_certificate(
    path: str | Path, category: str | None = None, text: str | None = None,
) -> CertificateRecord:
    source = Path(path)
    text = extract_pdf_text(source) if text is None else text
    category = category or classify_document(source, text)
    issue, expiry, annual = extract_certificate_dates(text)
    notes = "" if text else "未提取到 PDF 文本层；扫描件可能需要 OCR"
    try:
        pages = len(PdfReader(str(source)).pages)
    except Exception:
        pages = 0
    return CertificateRecord(str(source), category, extract_certificate_number(text), issue, expiry, annual, pages, notes)


def _merge(paths: list[Path], destination: Path) -> None:
    writer = PdfWriter()
    for path in paths:
        try:
            reader = PdfReader(str(path))
            for page in reader.pages:
                writer.add_page(page)
        except Exception:
            continue
    if len(writer.pages) == 0:
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        writer.write(handle)


def build_certificate_outputs(files: list[Path], maritime_dir: str | Path) -> list[CertificateRecord]:
    maritime = Path(maritime_dir)
    cert_dir = maritime / "证书原件"
    cert_dir.mkdir(parents=True, exist_ok=True)
    records: list[CertificateRecord] = []
    categorized: dict[str, list[Path]] = {}
    allowed_categories = {category for category, _ in CERTIFICATE_RULES}
    seen: set[str] = set()
    for path in files:
        if path.suffix.lower() != ".pdf":
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest in seen:
            continue
        seen.add(digest)
        text = extract_pdf_text(path)
        category = classify_document(path, text)
        if category not in allowed_categories:
            continue
        record = inspect_certificate(path, category, text)
        records.append(record)
        categorized.setdefault(category, []).append(path)
        target = cert_dir / path.name
        if path.resolve() != target.resolve():
            target.write_bytes(path.read_bytes())

    ordered = [category for category, _ in CERTIFICATE_RULES]
    vessel_files = [p for category in ordered for p in categorized.get(category, [])]
    _merge(vessel_files, maritime / "船舶证书.pdf")
    competency = sorted(
        (p for p in files if p.suffix.lower() == ".pdf" and classify_document(p, extract_pdf_text(p)) == "CREW_COMPETENCY"),
        key=lambda path: [int(token) if token.isdigit() else token.lower() for token in re.split(r"(\d+)", path.name)],
    )
    _merge(competency, maritime / "适任证书.pdf")
    dangerous_dir = maritime / "危险品证书"
    dangerous_dir.mkdir(parents=True, exist_ok=True)
    dangerous_groups = (
        ("国籍证书 INTERNATIONAL REGISTRY CERT.pdf", ["01_INTERNATIONAL_REGISTRY_CERT"]),
        ("IOPP及附件 I.O.P.P. AND FORM A OR B.pdf", ["04_IOPP_FORM_B"]),
        ("安全证书 SAFETY CONSTRUCTION_EQUIPMENT_RADIO.pdf", [
            "05_SAFETY_CONSTRUCTION", "06_SAFETY_EQUIPMENT_FORM_E", "07_SAFETY_RADIO_FORM_R",
        ]),
        ("CLC.pdf", ["09_CLC"]),
    )
    for filename, categories in dangerous_groups:
        _merge([path for category in categories for path in categorized.get(category, [])], dangerous_dir / filename)
    return records


def save_certificate_manifest(records: list[CertificateRecord], path: str | Path) -> None:
    Path(path).write_text(json.dumps([asdict(record) for record in records], ensure_ascii=False, indent=2), encoding="utf-8")
