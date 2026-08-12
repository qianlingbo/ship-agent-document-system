from __future__ import annotations

import email
import email.policy
import imaplib
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from email.header import decode_header
from email.message import Message
from pathlib import Path


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    parts = []
    for text, charset in decode_header(value):
        if isinstance(text, bytes):
            parts.append(text.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(text)
    return "".join(parts)


@dataclass
class Attachment:
    filename: str
    content: bytes
    content_type: str


@dataclass
class EmailRecord:
    subject: str
    sender: str
    received_at: datetime
    body: str
    attachments: list[Attachment]
    message_id: str = ""


def _body(message: Message) -> str:
    texts: list[str] = []
    parts = message.walk() if message.is_multipart() else [message]
    for part in parts:
        if part.get_content_disposition() == "attachment":
            continue
        if part.get_content_type() not in {"text/plain", "text/html"}:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        charset = part.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if part.get_content_type() == "text/html":
            text = re.sub(r"<[^>]+>", " ", text)
        texts.append(text)
    return "\n".join(texts).strip()


def parse_message(message: Message, received_at: datetime | None = None) -> EmailRecord:
    attachments: list[Attachment] = []
    for part in message.walk() if message.is_multipart() else [message]:
        filename = decode_mime(part.get_filename())
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if payload is not None:
            attachments.append(Attachment(filename, payload, part.get_content_type()))
    return EmailRecord(
        subject=decode_mime(message.get("Subject")),
        sender=decode_mime(message.get("From")),
        received_at=received_at or datetime.now(timezone.utc),
        body=_body(message),
        attachments=attachments,
        message_id=message.get("Message-ID", ""),
    )


def load_eml(path: str | Path) -> EmailRecord:
    source = Path(path)
    with source.open("rb") as handle:
        message = email.message_from_binary_file(handle, policy=email.policy.default)
    return parse_message(message, datetime.fromtimestamp(source.stat().st_mtime, timezone.utc))


class ImapClient:
    def __init__(self, host: str, port: int, username: str, password: str, folder: str = "INBOX", ssl: bool = True):
        self.host, self.port, self.username, self.password = host, port, username, password
        self.folder, self.ssl = folder, ssl
        self._connection: imaplib.IMAP4 | imaplib.IMAP4_SSL | None = None

    def __enter__(self) -> "ImapClient":
        self._connection = imaplib.IMAP4_SSL(self.host, self.port) if self.ssl else imaplib.IMAP4(self.host, self.port)
        self._connection.login(self.username, self.password)
        self._connection.select(self.folder)
        return self

    def __exit__(self, *_: object) -> None:
        if self._connection is not None:
            try:
                self._connection.close()
            finally:
                self._connection.logout()

    def unread(self, mark_seen: bool = False) -> list[EmailRecord]:
        if self._connection is None:
            raise RuntimeError("IMAP 客户端尚未连接")
        status, data = self._connection.search(None, "UNSEEN")
        if status != "OK":
            raise RuntimeError(f"IMAP 搜索失败: {status}")
        records = []
        for number in data[0].split():
            status, raw = self._connection.fetch(number, "(RFC822 INTERNALDATE)")
            if status != "OK":
                continue
            payload = next((item[1] for item in raw if isinstance(item, tuple)), None)
            if not payload:
                continue
            records.append(parse_message(email.message_from_bytes(payload, policy=email.policy.default)))
            if mark_seen:
                self._connection.store(number, "+FLAGS", "\\Seen")
        return records

