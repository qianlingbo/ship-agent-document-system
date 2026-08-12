from __future__ import annotations

import json
import hashlib
import html
import re
import shutil
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from .archives import ArchiveError, extract_archive
from .agent_mail import AgentMailClient
from .certificates import build_certificate_outputs, save_certificate_manifest
from .classify import classify_document, extract_ship_name, safe_ship_name
from .crew import convert_crew, convert_ports, find_input
from .crew_info_export import export_crew_info
from .email_ingest import EmailRecord, load_eml
from .excel_output import build_workbook
from .passports import apply_passport_mrz_names
from .ship_info import (
    apply_certificate_dates,
    apply_competencies,
    find_particulars_a,
    read_particulars,
    taskbook_fields,
)
from .taskbook import build_taskbook


def _unique_path(directory: Path, filename: str) -> Path:
    candidate = directory / Path(filename).name
    stem, suffix = candidate.stem, candidate.suffix
    counter = 1
    while candidate.exists():
        candidate = directory / f"{stem}_{counter}{suffix}"
        counter += 1
    return candidate


def _digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def _copy_once(source: Path, directory: Path) -> Path:
    candidate = directory / source.name
    if candidate.exists() and _digest(candidate) == _digest(source):
        return candidate
    target = _unique_path(directory, source.name)
    shutil.copy2(source, target)
    return target


def _plain_text(value: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", value or "")).split())


def _field(text: str, label: str) -> str:
    match = re.search(rf"{re.escape(label)}\s*[:：]\s*([^\n\r;,，；]+)", text, re.I)
    return match.group(1).strip() if match else ""


def _numbered_folder(files: list[Path], suffixes: set[str], tokens: tuple[str, ...]) -> Path | None:
    candidates: dict[Path, set[int]] = {}
    for path in files:
        if path.suffix.lower() not in suffixes:
            continue
        match = re.match(r"\s*0*(\d{1,3})(?:\D|$)", path.name)
        if match:
            candidates.setdefault(path.parent, set()).add(int(match.group(1)))
    if not candidates:
        return None
    ranked = sorted(
        candidates.items(),
        key=lambda item: (
            -sum(token in str(item[0]).upper() for token in tokens),
            -len(item[1]),
            str(item[0]),
        ),
    )
    return ranked[0][0]


class ShipPipeline:
    def __init__(
        self,
        output_root: str | Path,
        template_path: str | Path,
        crew_info_template: str | Path | None = None,
    ):
        self.output_root = Path(output_root)
        self.template_path = Path(template_path)
        self.crew_info_template = Path(crew_info_template) if crew_info_template else None

    def process_eml(self, path: str | Path, ship_name: str | None = None) -> Path:
        return self.process_email(load_eml(path), ship_name=ship_name)

    def process_directory(self, path: str | Path, ship_name: str | None = None) -> Path:
        source = Path(path)
        files = [p for p in source.rglob("*") if p.is_file()]
        pseudo = EmailRecord(source.name, "local", datetime.now(timezone.utc), "", [])
        return self._process_files(pseudo, files, ship_name=ship_name, source_root=source)

    def process_email(self, record: EmailRecord, ship_name: str | None = None) -> Path:
        with TemporaryDirectory(prefix="ship-agent-") as temporary:
            staging = Path(temporary)
            source_files: list[Path] = []
            for attachment in record.attachments:
                destination = _unique_path(staging, attachment.filename or "attachment.bin")
                destination.write_bytes(attachment.content)
                source_files.append(destination)
            (staging / "邮件正文.txt").write_text(record.body, encoding="utf-8")
            return self._process_files(record, source_files, ship_name=ship_name)

    def process_agent_mail(self, client: AgentMailClient, message_id: str, ship_name: str | None = None) -> Path:
        """Read one Agent Mail message and download its ordinary attachments."""
        record = client.read_message(message_id)
        with TemporaryDirectory(prefix="ship-agent-mail-") as temporary:
            downloaded = client.download_attachments(message_id, [a.filename for a in record.attachments], temporary)
            return self._process_files(record, downloaded, ship_name=ship_name)

    def _process_files(
        self,
        record: EmailRecord,
        source_files: list[Path],
        ship_name: str | None = None,
        source_root: Path | None = None,
    ) -> Path:
        ship_name = safe_ship_name(ship_name or extract_ship_name(record.subject, record.body, [p.name for p in source_files]) or f"UNKNOWN_SHIP_{record.received_at:%Y%m%d_%H%M%S}")
        ship_dir = self.output_root / ship_name
        raw_dir = ship_dir / "原始文件"
        maritime_dir = ship_dir / "海事"
        raw_dir.mkdir(parents=True, exist_ok=True)
        maritime_dir.mkdir(parents=True, exist_ok=True)

        errors: list[str] = []
        for source in source_files:
            if source_root:
                try:
                    relative = source.relative_to(source_root)
                    destination = raw_dir / relative
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if not destination.exists() or _digest(destination) != _digest(source):
                        shutil.copy2(source, destination)
                except ValueError:
                    _copy_once(source, raw_dir)
            else:
                _copy_once(source, raw_dir)
        # Keep the original mail text with the ship file set.
        if record.body:
            mail_dir = raw_dir / "邮件正文"
            mail_dir.mkdir(exist_ok=True)
            mail_key = safe_ship_name(record.message_id or record.received_at.strftime("%Y%m%d_%H%M%S"))
            (mail_dir / f"{mail_key}.html").write_text(record.body, encoding="utf-8")

        history_path = ship_dir / "邮件记录.json"
        try:
            history = json.loads(history_path.read_text(encoding="utf-8")) if history_path.exists() else []
        except (OSError, ValueError):
            history = []
        if not any(item.get("message_id") == record.message_id for item in history):
            history.append({
                "message_id": record.message_id, "subject": record.subject, "sender": record.sender,
                "received_at": record.received_at.isoformat(),
            })
            history_path.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")

        all_files = [p for p in raw_dir.rglob("*") if p.is_file() and "邮件正文" not in p.parts]

        # Expand archives recursively. The extracted copies are retained for audit.
        queue = list(all_files)
        seen_archives: set[Path] = set()
        while queue:
            current = queue.pop(0)
            if current.suffix.lower() not in {".zip", ".rar", ".tar", ".gz", ".tgz", ".bz2", ".xz"} or current in seen_archives:
                continue
            seen_archives.add(current)
            destination = raw_dir / "解压" / current.stem
            try:
                expanded = extract_archive(current, destination)
                all_files.extend(expanded)
                queue.extend(expanded)
            except ArchiveError as exc:
                errors.append(f"{current.name}: {exc}")

        usable = list({p.resolve(): p for p in all_files if p.is_file()}.values())
        particulars, certificate_dates, competencies = read_particulars(find_particulars_a(usable))
        crew_path = find_input(usable, "crew")
        port_path = find_input(usable, "port")
        crew_rows = convert_crew(crew_path) if crew_path else []
        previous_checks: list[dict[str, str]] = []
        try:
            previous_checks = json.loads((ship_dir / "处理清单.json").read_text(encoding="utf-8")).get("passport_name_checks", [])
        except (OSError, ValueError):
            pass
        try:
            manual_overrides = json.loads((ship_dir / "护照姓名人工核对.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            manual_overrides = {}
        passport_name_checks = apply_passport_mrz_names(crew_rows, usable, previous_checks, manual_overrides)
        manual_mrz = [item for item in passport_name_checks if item.get("verification") == "needs_manual_review"]
        if manual_mrz:
            errors.append(
                "护照 MRZ 未自动确认，需人工核对序号："
                + ", ".join(str(item.get("sequence") or "?") for item in manual_mrz)
            )
        apply_competencies(crew_rows, competencies)
        port_rows = convert_ports(port_path) if port_path else []
        certificates = build_certificate_outputs(usable, maritime_dir)
        apply_certificate_dates(certificates, certificate_dates)
        workbook_path = build_workbook(crew_rows, port_rows, certificates, ship_dir / "单一窗口录入.xlsx")
        save_certificate_manifest(certificates, maritime_dir / "证书登记.json")

        crew_info_path = ship_dir / "船员信息.xlsx"
        if crew_rows and self.crew_info_template and self.crew_info_template.exists():
            photo_folder = _numbered_folder(usable, {".jpg", ".jpeg", ".png"}, ("CREW PHOTO", "PHOTO", "照片"))
            passport_folder = _numbered_folder(usable, {".pdf"}, ("TRAVEL PASSPORT", "PASSPORT", "护照"))
            if photo_folder and passport_folder:
                try:
                    export_crew_info(
                        self.crew_info_template, workbook_path, photo_folder, passport_folder, crew_info_path,
                    )
                except (OSError, RuntimeError, ValueError) as exc:
                    errors.append(f"船员信息表生成失败: {exc}")
            else:
                missing = []
                if not photo_folder:
                    missing.append("按序号命名的船员照片文件夹")
                if not passport_folder:
                    missing.append("按序号命名的护照 PDF 文件夹")
                errors.append("未生成船员信息表：缺少" + "、".join(missing))

        mail_bodies = [record.body]
        mail_dir = raw_dir / "邮件正文"
        if mail_dir.exists():
            mail_bodies.extend(path.read_text(encoding="utf-8", errors="ignore") for path in mail_dir.glob("*.html"))
        body = _plain_text("\n".join(dict.fromkeys(mail_bodies)))
        task_fields = taskbook_fields(particulars, body, ship_name)
        task_fields.update({
            "船名中文": _field(body, "船名中文") or _field(body, "船名") or ship_name,
            "英文": _field(body, "英文") or ship_name,
            "航次": _field(body, "航次"),
            "船东事宜等其他注意事项": _field(body, "船东事宜等其他注意事项"),
        })
        taskbook_path = ship_dir / "船舶计划任务书.docx"
        if self.template_path.exists():
            build_taskbook(self.template_path, taskbook_path, task_fields, {"海事", "危险品", "单证"} if certificates else set())
        else:
            errors.append(f"未找到计划任务书模板: {self.template_path}")

        manifest = {
            "ship_name": ship_name,
            "messages": history,
            "files": [{"path": str(path.relative_to(ship_dir)), "category": classify_document(path)} for path in usable],
            "certificates": [asdict(c) for c in certificates],
            "passport_name_checks": passport_name_checks,
            "outputs": {
                "single_window_excel": str(workbook_path.relative_to(ship_dir)),
                "crew_information_excel": str(crew_info_path.relative_to(ship_dir)) if crew_info_path.exists() else "",
                "vessel_pdf": "海事/船舶证书.pdf" if (maritime_dir / "船舶证书.pdf").exists() else "",
                "competency_pdf": "海事/适任证书.pdf" if (maritime_dir / "适任证书.pdf").exists() else "",
                "dangerous_goods_folder": "海事/危险品证书" if any((maritime_dir / "危险品证书").glob("*.pdf")) else "",
                "taskbook": str(taskbook_path.relative_to(ship_dir)) if taskbook_path.exists() else "",
            },
            "errors": errors,
        }
        (ship_dir / "处理清单.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return ship_dir
