#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
COMPONENTS = ROOT / "components"
COMPONENT_NAMES = (
    "ship-bonded-terminal-forecast",
    "ship-certificate-data",
    "ship-certificate-pdf",
    "ship-crew-list",
    "ship-crew-photo-workbook",
    "ship-document-organizer",
    "ship-mail-intake",
    "ship-passport-mrz",
    "ship-port-of-call",
    "ship-single-window-excel",
    "ship-taskbook",
)


def _ignore(_: str, names: list[str]) -> set[str]:
    ignored = {name for name in names if name in {".git", "__pycache__"} or name.endswith((".pyc", ".zip"))}
    return ignored


def _runtime_ignore(directory: str, names: list[str]) -> set[str]:
    ignored = _ignore(directory, names)
    if Path(directory).resolve() == ROOT.resolve():
        ignored.add("components")
    return ignored


def main() -> int:
    parser = argparse.ArgumentParser(description="安装船舶代理总控 Skill 和全部细分 Skills")
    parser.add_argument(
        "--target", type=Path, default=Path.home() / ".codex" / "skills",
        help="Skill 目录；Codex 默认 ~/.codex/skills",
    )
    args = parser.parse_args()
    target = args.target.expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)

    runtime_target = target / "ship-agent-document-system"
    if ROOT.resolve() != runtime_target.resolve():
        shutil.copytree(ROOT, runtime_target, dirs_exist_ok=True, ignore=_runtime_ignore)
    installed = [runtime_target]
    if COMPONENTS.exists():
        component_sources = [COMPONENTS / name for name in COMPONENT_NAMES]
    else:
        component_sources = [ROOT.parent / name for name in COMPONENT_NAMES]
    for component in component_sources:
        if not (component / "SKILL.md").exists():
            raise FileNotFoundError(f"缺少细分 Skill: {component}")
        destination = target / component.name
        if component.resolve() != destination.resolve():
            shutil.copytree(component, destination, dirs_exist_ok=True, ignore=_ignore)
        installed.append(destination)

    print(f"已安装 {len(installed)} 个 Skills 到 {target}")
    for path in installed:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
