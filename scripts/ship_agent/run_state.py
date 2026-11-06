from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4


FINAL_STATES = {"completed", "needs_manual_review", "failed"}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


class RunStateStore:
    """Persist a small, auditable history of mail processing runs."""

    def __init__(self, path: str | Path, keep: int = 200):
        self.path = Path(path)
        self.keep = keep

    def _load(self) -> dict:
        try:
            value = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            value = {"schema_version": 1, "runs": []}
        if not isinstance(value, dict) or not isinstance(value.get("runs"), list):
            return {"schema_version": 1, "runs": []}
        return value

    def _save(self, value: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(self.path)

    def start(self, message_id: str, subject: str = "") -> str:
        value = self._load()
        run_id = uuid4().hex
        value["runs"].append({
            "run_id": run_id,
            "message_id": message_id,
            "subject": subject,
            "state": "running",
            "started_at": _now(),
            "finished_at": "",
            "ship_folder": "",
            "error": "",
        })
        value["runs"] = value["runs"][-self.keep:]
        self._save(value)
        return run_id

    def finish(
        self,
        run_id: str,
        state: str,
        ship_folder: str | Path | None = None,
        error: str = "",
    ) -> None:
        if state not in FINAL_STATES:
            raise ValueError(f"Unsupported final state: {state}")
        value = self._load()
        for item in reversed(value["runs"]):
            if item.get("run_id") == run_id:
                item.update({
                    "state": state,
                    "finished_at": _now(),
                    "ship_folder": str(ship_folder or ""),
                    "error": error,
                })
                self._save(value)
                return
        raise KeyError(f"Unknown run_id: {run_id}")


def output_state(ship_folder: str | Path) -> tuple[str, str]:
    """Map a completed pipeline output to an explicit workflow state."""
    manifest_path = Path(ship_folder) / "处理清单.json"
    try:
        errors = json.loads(manifest_path.read_text(encoding="utf-8")).get("errors", [])
    except (OSError, ValueError, AttributeError):
        return "needs_manual_review", "处理清单缺失或无法读取"
    if errors:
        return "needs_manual_review", "；".join(str(item) for item in errors)
    return "completed", ""
