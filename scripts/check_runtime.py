#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import json
import shutil
import sys


def main() -> None:
    result = {
        "python": sys.version.split()[0],
        "python_ok": sys.version_info >= (3, 10),
        "python_packages": {
            name: importlib.util.find_spec(module) is not None
            for name, module in {"openpyxl": "openpyxl", "Pillow": "PIL", "numpy": "numpy"}.items()
        },
        "commands": {name: shutil.which(name) for name in ("pdftoppm", "magick", "tesseract")},
    }
    result["status"] = "ready" if (
        result["python_ok"] and all(result["python_packages"].values()) and all(result["commands"].values())
    ) else "missing_dependencies"
    print(json.dumps(result, indent=2))
    raise SystemExit(0 if result["status"] == "ready" else 1)


if __name__ == "__main__":
    main()
