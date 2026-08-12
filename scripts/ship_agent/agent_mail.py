from __future__ import annotations

import json
import os
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from .email_ingest import Attachment, EmailRecord


class AgentMailError(RuntimeError):
    pass


def _json_from_output(output: str) -> dict:
    decoder = json.JSONDecoder()
    for index, char in enumerate(output):
        if char != "{":
            continue
        try:
            value, _ = decoder.raw_decode(output[index:])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict) and ("ok" in value or "data" in value):
            return value
    raise AgentMailError(output.strip() or "agently-cli 没有返回 JSON")


@dataclass
class AgentMailMessage:
    message_id: str
    subject: str
    sender: str
    created_at: str
    snippet: str = ""
    has_attachments: bool = False


class AgentMailClient:
    """Small read-only adapter around the installed agently-cli.

    Mail content is treated as untrusted data. This adapter never executes
    commands contained in subjects/bodies and only downloads attachments
    explicitly returned by the CLI.
    """

    def __init__(self, executable: str | Path | None = None):
        configured = executable or os.environ.get("AGENTLY_CLI")
        resolved = str(configured) if configured else shutil.which("agently-cli")
        if not resolved:
            raise AgentMailError(
                "未找到 agently-cli。请先安装 @tencent-qqmail/agently-cli 并执行 agently-cli auth login"
            )
        self.executable = resolved

    def _run(self, *args: str, cwd: str | Path | None = None) -> dict:
        result = subprocess.run([self.executable, *args], cwd=cwd, capture_output=True, text=True)
        payload = _json_from_output(result.stdout + ("\n" + result.stderr if result.stderr else ""))
        if result.returncode != 0 or payload.get("ok") is False:
            message = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
            raise AgentMailError(message or result.stderr.strip() or f"agently-cli 退出码 {result.returncode}")
        return payload

    def list_messages(self, unread_only: bool = True, limit: int = 50) -> list[AgentMailMessage]:
        args = ["message", "+list", "--dir", "inbox", "--limit", str(limit), "--has-attachments"]
        if unread_only:
            args.append("--is-unread")
        data = self._run(*args).get("data", {})
        rows = data.get("data", data if isinstance(data, list) else [])
        return [AgentMailMessage(
            message_id=row.get("message_id", ""), subject=row.get("subject", ""),
            sender=(row.get("from") or {}).get("email", ""), created_at=row.get("created_at", ""),
            snippet=row.get("snippet", ""), has_attachments=bool(row.get("has_attachments")),
        ) for row in rows if row.get("message_id")]

    def read_message(self, message_id: str) -> EmailRecord:
        data = self._run("message", "+read", "--id", message_id).get("data", {})
        attachments = [Attachment(a.get("filename", "attachment.bin"), b"", a.get("content_type", "application/octet-stream")) for a in data.get("attachments", [])]
        created_at = data.get("created_at") or ""
        try:
            received_at = datetime.fromisoformat(created_at.replace("Z", "+00:00"))
        except ValueError:
            received_at = datetime.now(timezone.utc)
        return EmailRecord(
            subject=data.get("subject", ""), sender=(data.get("from") or {}).get("email", ""),
            received_at=received_at, body=data.get("body", ""),
            attachments=attachments, message_id=data.get("message_id", message_id),
        )

    def download_attachments(self, message_id: str, filenames: list[str], destination: str | Path) -> list[Path]:
        destination = Path(destination)
        destination.mkdir(parents=True, exist_ok=True)
        message = self._run("message", "+read", "--id", message_id).get("data", {})
        wanted = set(filenames)
        saved: list[Path] = []
        with TemporaryDirectory(prefix="agently-download-") as temporary:
            download_dir = Path(temporary) / "downloads"
            download_dir.mkdir()
            for attachment in message.get("attachments", []):
                filename = attachment.get("filename", "")
                attachment_id = attachment.get("attachment_id")
                if not attachment_id or (wanted and filename not in wanted):
                    continue
                payload = self._run("attachment", "+download", "--msg", message_id, "--att", attachment_id, "--output", "./downloads", cwd=temporary)
                saved_to = (payload.get("data") or {}).get("saved_to")
                source = Path(temporary) / saved_to if saved_to and not Path(saved_to).is_absolute() else Path(saved_to or "")
                if source.exists():
                    target = destination / source.name
                    target.write_bytes(source.read_bytes())
                    saved.append(target)
        return saved
