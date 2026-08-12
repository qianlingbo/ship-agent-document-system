from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass(frozen=True)
class Settings:
    email: str = ""
    imap_host: str = "imap.qq.com"
    imap_port: int = 993
    imap_folder: str = "INBOX"
    imap_ssl: bool = True
    imap_password: str = ""
    work_root: Path = Path("data")
    poll_seconds: int = 7200
    mark_seen: bool = False
    template_path: Path = Path(__file__).resolve().parents[2] / "assets" / "船舶计划任务书.docx"

    @classmethod
    def from_env(cls, root: str | Path | None = None) -> "Settings":
        base = Path(root or os.environ.get("SHIP_AGENT_WORK_ROOT", "data"))
        return cls(
            email=os.environ.get("SHIP_AGENT_EMAIL", cls.email),
            imap_host=os.environ.get("SHIP_AGENT_IMAP_HOST", cls.imap_host),
            imap_port=int(os.environ.get("SHIP_AGENT_IMAP_PORT", cls.imap_port)),
            imap_folder=os.environ.get("SHIP_AGENT_IMAP_FOLDER", cls.imap_folder),
            imap_ssl=_bool(os.environ.get("SHIP_AGENT_IMAP_SSL"), cls.imap_ssl),
            imap_password=os.environ.get("SHIP_AGENT_IMAP_PASSWORD", ""),
            work_root=base,
            poll_seconds=int(os.environ.get("SHIP_AGENT_POLL_SECONDS", cls.poll_seconds)),
            mark_seen=_bool(os.environ.get("SHIP_AGENT_MARK_SEEN"), cls.mark_seen),
            template_path=Path(os.environ.get("SHIP_AGENT_TEMPLATE", str(cls.template_path))),
        )
