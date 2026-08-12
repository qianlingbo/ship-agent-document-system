#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from ship_agent.crew_info_export import export_crew_info


def main() -> int:
    parser = argparse.ArgumentParser(description="单独生成船员信息及照片 Excel")
    parser.add_argument("--template", required=True)
    parser.add_argument("--crew", required=True, help="已生成的单一窗口 Excel")
    parser.add_argument("--photos", required=True)
    parser.add_argument("--passports", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    try:
        result = export_crew_info(
            Path(args.template), Path(args.crew), Path(args.photos), Path(args.passports), Path(args.output),
        )
    except (OSError, RuntimeError, ValueError) as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False))
        return 1
    print(json.dumps({"status": "success", **result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
