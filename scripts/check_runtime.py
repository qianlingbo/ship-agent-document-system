#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import json
import shutil
import sys


def main() -> int:
    packages = {
        name: importlib.util.find_spec(module) is not None
        for name, module in {
            "openpyxl": "openpyxl", "Pillow": "PIL", "numpy": "numpy",
            "pypdf": "pypdf", "pdfplumber": "pdfplumber", "python-docx": "docx", "xlrd": "xlrd",
        }.items()
    }
    required_commands = {name: shutil.which(name) for name in ("pdftoppm", "tesseract", "magick")}
    result = {
        "python": sys.version.split()[0],
        "python_ok": sys.version_info >= (3, 10),
        "python_packages": packages,
        "required_commands": required_commands,
        "agent_mail_cli": shutil.which("agently-cli"),
        "archive_commands": {name: shutil.which(name) for name in ("7zz", "7z", "unar")},
    }
    core_ready = result["python_ok"] and all(packages.values()) and all(required_commands.values())
    result["status"] = "ready" if core_ready else "missing_dependencies"
    result["mail_ready"] = bool(result["agent_mail_cli"])
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if core_ready else 1


if __name__ == "__main__":
    raise SystemExit(main())
