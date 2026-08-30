#!/usr/bin/env python3
from __future__ import annotations
import sys
from pathlib import Path

for root in (Path(__file__).resolve().parents[3], Path(__file__).resolve().parents[2] / "ship-agent-document-system"):
    if (root / "scripts" / "component_cli.py").exists():
        sys.path.insert(0, str(root / "scripts"))
        from component_cli import main
        raise SystemExit(main(["cert-data", *sys.argv[1:]]))
raise SystemExit("未找到 ship-agent-document-system 共享运行时")
