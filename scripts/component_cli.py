#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from ship_agent.agent_mail import AgentMailClient, AgentMailError
from ship_agent.archives import ArchiveError, extract_archive
from ship_agent.certificates import CertificateRecord, build_certificate_outputs, inspect_certificate, save_certificate_manifest
from ship_agent.classify import CERTIFICATE_RULES, classify_document, extract_ship_name, safe_ship_name
from ship_agent.crew import convert_crew, convert_ports
from ship_agent.crew_info_export import export_crew_info
from ship_agent.email_ingest import EmailRecord, load_eml
from ship_agent.excel_output import DEFAULT_TEMPLATE, build_workbook
from ship_agent.passports import apply_passport_mrz_names
from ship_agent.ship_info import apply_certificate_dates, find_particulars_a, read_particulars, taskbook_fields
from ship_agent.taskbook import build_taskbook


SKILL_ROOT = Path(__file__).resolve().parent.parent
TASKBOOK_TEMPLATE = SKILL_ROOT / "assets" / "船舶计划任务书.docx"
CREW_INFO_TEMPLATE = SKILL_ROOT / "assets" / "船员信息模板.xlsx"
ARCHIVE_SUFFIXES = {".zip", ".rar", ".tar", ".gz", ".tgz", ".bz2", ".xz"}
CERTIFICATE_CATEGORIES = {category for category, _ in CERTIFICATE_RULES} - {"13_WRC"}


def _dump(value: object, path: Path | None = None) -> None:
    text = json.dumps(value, ensure_ascii=False, indent=2)
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
    print(text)


def _hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _copy_unique(source: Path, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() and _hash(source) == _hash(destination):
        return destination
    candidate = destination
    counter = 1
    while candidate.exists():
        candidate = destination.with_name(f"{destination.stem}_{counter}{destination.suffix}")
        counter += 1
    shutil.copy2(source, candidate)
    return candidate


def _expand_archives(raw_dir: Path) -> tuple[list[Path], list[str]]:
    files = [path for path in raw_dir.rglob("*") if path.is_file()]
    queue = list(files)
    seen: set[str] = set()
    errors: list[str] = []
    while queue:
        archive = queue.pop(0)
        if archive.suffix.lower() not in ARCHIVE_SUFFIXES:
            continue
        digest = _hash(archive)
        if digest in seen:
            continue
        seen.add(digest)
        destination = raw_dir / "解压" / archive.stem
        try:
            expanded = extract_archive(archive, destination)
        except (ArchiveError, OSError) as exc:
            errors.append(f"{archive.name}: {exc}")
            continue
        files.extend(expanded)
        queue.extend(expanded)
    unique = {path.resolve(): path for path in files if path.is_file()}
    return list(unique.values()), errors


def _organize(
    source_files: list[Path], output_root: Path, ship_name: str, source_root: Path | None = None,
    record: EmailRecord | None = None,
) -> dict[str, object]:
    ship_dir = output_root / safe_ship_name(ship_name)
    raw_dir = ship_dir / "原始文件"
    raw_dir.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for source in source_files:
        relative = source.name
        if source_root:
            try:
                relative = str(source.relative_to(source_root))
            except ValueError:
                pass
        copied.append(_copy_unique(source, raw_dir / relative))
    if record:
        body_dir = raw_dir / "邮件正文"
        body_dir.mkdir(exist_ok=True)
        key = safe_ship_name(record.message_id or record.received_at.strftime("%Y%m%d_%H%M%S"))
        (body_dir / f"{key}.html").write_text(record.body or "", encoding="utf-8")
        history_path = ship_dir / "邮件记录.json"
        try:
            history = json.loads(history_path.read_text(encoding="utf-8")) if history_path.exists() else []
        except (OSError, ValueError):
            history = []
        if not any(item.get("message_id") == record.message_id for item in history):
            history.append({
                "message_id": record.message_id,
                "subject": record.subject,
                "sender": record.sender,
                "received_at": record.received_at.isoformat(),
            })
            history_path.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    expanded, errors = _expand_archives(raw_dir)
    manifest = {
        "ship_name": safe_ship_name(ship_name),
        "ship_directory": str(ship_dir),
        "copied_files": [str(path.relative_to(ship_dir)) for path in copied],
        "available_files": [str(path.relative_to(ship_dir)) for path in expanded],
        "errors": errors,
    }
    _dump(manifest, ship_dir / "整理清单.json")
    return manifest


def _all_files(source: Path) -> list[Path]:
    return [path for path in source.rglob("*") if path.is_file()] if source.is_dir() else [source]


def _load_rows(path: Path, key: str, converter) -> list[list[str]]:
    if path.suffix.lower() == ".json":
        payload = json.loads(path.read_text(encoding="utf-8"))
        rows = payload.get(key, payload) if isinstance(payload, dict) else payload
        if not isinstance(rows, list):
            raise ValueError(f"JSON 缺少 {key}")
        return rows
    return converter(path)


def _certificate_records(path: Path | None) -> list[CertificateRecord]:
    if not path:
        return []
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows = payload.get("certificates", payload) if isinstance(payload, dict) else payload
    return [CertificateRecord(**row) for row in rows]


def command_organize(args: argparse.Namespace) -> int:
    source = Path(args.source).resolve()
    if source.suffix.lower() == ".eml":
        record = load_eml(source)
        ship_name = args.ship_name or extract_ship_name(record.subject, record.body, [a.filename for a in record.attachments])
        with TemporaryDirectory(prefix="ship-organize-") as temporary:
            staging = Path(temporary)
            files = []
            for attachment in record.attachments:
                target = staging / Path(attachment.filename or "attachment.bin").name
                target.write_bytes(attachment.content)
                files.append(target)
            _organize(files, Path(args.output), ship_name or source.stem, record=record)
        return 0
    files = _all_files(source)
    ship_name = args.ship_name or (source.name if source.is_dir() else source.stem)
    _organize(files, Path(args.output), ship_name, source if source.is_dir() else source.parent)
    return 0


def command_mail(args: argparse.Namespace) -> int:
    client = AgentMailClient(args.mail_cli)
    if args.mail_action == "list":
        rows = [asdict(item) for item in client.list_messages(not args.include_read, args.limit)]
        _dump({"messages": rows, "count": len(rows)})
        return 0
    record = client.read_message(args.message_id)
    ship_name = args.ship_name or extract_ship_name(record.subject, record.body, [a.filename for a in record.attachments])
    with TemporaryDirectory(prefix="ship-mail-intake-") as temporary:
        files = client.download_attachments(
            args.message_id, [attachment.filename for attachment in record.attachments], temporary,
        )
        _organize(files, Path(args.output), ship_name or record.subject or "UNKNOWN_SHIP", record=record)
    return 0


def command_cert_pdf(args: argparse.Namespace) -> int:
    files = _all_files(Path(args.source))
    output = Path(args.output)
    records = build_certificate_outputs(files, output)
    manifest = output / "证书登记.json"
    save_certificate_manifest(records, manifest)
    _dump({
        "certificates": [asdict(record) for record in records],
        "vessel_pdf": str(output / "船舶证书.pdf") if (output / "船舶证书.pdf").exists() else "",
        "competency_pdf": str(output / "适任证书.pdf") if (output / "适任证书.pdf").exists() else "",
        "dangerous_goods": [str(path) for path in sorted((output / "危险品证书").glob("*.pdf"))],
        "manifest": str(manifest),
    })
    return 0


def command_cert_data(args: argparse.Namespace) -> int:
    files = _all_files(Path(args.source))
    records = []
    seen: set[str] = set()
    for path in files:
        if path.suffix.lower() != ".pdf":
            continue
        digest = _hash(path)
        category = classify_document(path)
        if digest in seen or category not in CERTIFICATE_CATEGORIES:
            continue
        seen.add(digest)
        records.append(inspect_certificate(path, category))
    _, dates, _ = read_particulars(find_particulars_a(files))
    apply_certificate_dates(records, dates)
    payload = {"certificates": [asdict(record) for record in sorted(records, key=lambda item: item.category)]}
    _dump(payload, Path(args.output))
    return 0


def command_crew(args: argparse.Namespace) -> int:
    rows = convert_crew(Path(args.source), args.embark_date, args.embark_port)
    _dump({"crew_rows": rows, "count": len(rows)}, Path(args.output))
    return 0


def command_ports(args: argparse.Namespace) -> int:
    rows = convert_ports(Path(args.source))
    _dump({"port_rows": rows, "count": len(rows)}, Path(args.output))
    return 0


def command_mrz(args: argparse.Namespace) -> int:
    crew_rows = _load_rows(Path(args.crew), "crew_rows", convert_crew)
    files = _all_files(Path(args.passports))
    overrides = {}
    if args.overrides:
        overrides = json.loads(Path(args.overrides).read_text(encoding="utf-8"))
    checks = apply_passport_mrz_names(crew_rows, files, [], overrides)
    manual = [item for item in checks if item.get("verification") == "needs_manual_review"]
    payload = {
        "crew_rows": crew_rows,
        "passport_name_checks": checks,
        "manual_review_sequences": [item.get("sequence") for item in manual],
    }
    _dump(payload, Path(args.output))
    return 2 if manual else 0


def command_single_window(args: argparse.Namespace) -> int:
    crew = _load_rows(Path(args.crew), "crew_rows", convert_crew) if args.crew else []
    ports = _load_rows(Path(args.ports), "port_rows", convert_ports) if args.ports else []
    certificates = _certificate_records(Path(args.certificates) if args.certificates else None)
    output = build_workbook(
        crew, ports, certificates, Path(args.output), Path(args.template) if args.template else DEFAULT_TEMPLATE,
    )
    _dump({
        "output": str(output), "crew_count": len(crew), "port_count": len(ports),
        "certificate_count": len([item for item in certificates if item.category != "13_WRC"]),
    })
    return 0


def command_taskbook(args: argparse.Namespace) -> int:
    particulars = Path(args.particulars) if args.particulars else None
    info, _, _ = read_particulars(particulars)
    mail_text = Path(args.mail_body).read_text(encoding="utf-8", errors="ignore") if args.mail_body else ""
    fields = taskbook_fields(info, mail_text, args.ship_name)
    if args.fields:
        fields.update(json.loads(Path(args.fields).read_text(encoding="utf-8")))
    completed = set(args.completed or [])
    output = build_taskbook(
        Path(args.template) if args.template else TASKBOOK_TEMPLATE, Path(args.output), fields, completed,
    )
    _dump({"output": str(output), "fields": fields, "completed": sorted(completed)})
    return 0


def command_crew_photo(args: argparse.Namespace) -> int:
    result = export_crew_info(
        Path(args.template) if args.template else CREW_INFO_TEMPLATE,
        Path(args.crew), Path(args.photos), Path(args.passports), Path(args.output),
    )
    _dump(result)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ship-agent-component", description="船舶代理单证细分组件")
    sub = parser.add_subparsers(dest="component", required=True)

    organize = sub.add_parser("organize", help="按船名归档并递归解压")
    organize.add_argument("source")
    organize.add_argument("--output", required=True)
    organize.add_argument("--ship-name")
    organize.set_defaults(handler=command_organize)

    mail = sub.add_parser("mail", help="Agent Mail 读取与下载")
    mail_sub = mail.add_subparsers(dest="mail_action", required=True)
    mail_list = mail_sub.add_parser("list")
    mail_list.add_argument("--limit", type=int, default=50)
    mail_list.add_argument("--include-read", action="store_true")
    mail_list.add_argument("--mail-cli")
    mail_list.set_defaults(handler=command_mail)
    mail_download = mail_sub.add_parser("download")
    mail_download.add_argument("message_id")
    mail_download.add_argument("--output", required=True)
    mail_download.add_argument("--ship-name")
    mail_download.add_argument("--mail-cli")
    mail_download.set_defaults(handler=command_mail)

    cert_pdf = sub.add_parser("cert-pdf", help="合并船舶、适任和危险品 PDF")
    cert_pdf.add_argument("source")
    cert_pdf.add_argument("--output", required=True)
    cert_pdf.set_defaults(handler=command_cert_pdf)

    cert_data = sub.add_parser("cert-data", help="提取证书号与日期")
    cert_data.add_argument("source")
    cert_data.add_argument("--output", required=True)
    cert_data.set_defaults(handler=command_cert_data)

    crew = sub.add_parser("crew", help="转换 Crew List")
    crew.add_argument("source")
    crew.add_argument("--output", required=True)
    crew.add_argument("--embark-date", default="")
    crew.add_argument("--embark-port", default="")
    crew.set_defaults(handler=command_crew)

    ports = sub.add_parser("ports", help="转换 Port of Call")
    ports.add_argument("source")
    ports.add_argument("--output", required=True)
    ports.set_defaults(handler=command_ports)

    mrz = sub.add_parser("mrz", help="用护照 MRZ 核验船员姓名")
    mrz.add_argument("--crew", required=True)
    mrz.add_argument("--passports", required=True)
    mrz.add_argument("--output", required=True)
    mrz.add_argument("--overrides")
    mrz.set_defaults(handler=command_mrz)

    single = sub.add_parser("single-window", help="组装单一窗口 Excel")
    single.add_argument("--crew")
    single.add_argument("--ports")
    single.add_argument("--certificates")
    single.add_argument("--template")
    single.add_argument("--output", required=True)
    single.set_defaults(handler=command_single_window)

    taskbook = sub.add_parser("taskbook", help="生成船舶计划任务书")
    taskbook.add_argument("--ship-name", required=True)
    taskbook.add_argument("--particulars")
    taskbook.add_argument("--mail-body")
    taskbook.add_argument("--fields")
    taskbook.add_argument("--completed", nargs="*", choices=("海事", "危险品", "单证", "PASS"))
    taskbook.add_argument("--template")
    taskbook.add_argument("--output", required=True)
    taskbook.set_defaults(handler=command_taskbook)

    photo = sub.add_parser("crew-photo", help="生成船员照片信息表")
    photo.add_argument("--crew", required=True)
    photo.add_argument("--photos", required=True)
    photo.add_argument("--passports", required=True)
    photo.add_argument("--template")
    photo.add_argument("--output", required=True)
    photo.set_defaults(handler=command_crew_photo)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.handler(args)
    except (AgentMailError, ArchiveError, OSError, ValueError, RuntimeError) as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
