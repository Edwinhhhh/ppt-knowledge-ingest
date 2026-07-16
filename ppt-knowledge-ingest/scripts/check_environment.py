#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import json
import shutil
import sys


def has_module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def main() -> int:
    result = {
        "python": sys.version.split()[0],
        "modules": {
            "pptx": has_module("pptx"),
            "PIL": has_module("PIL"),
        },
        "binaries": {
            "soffice": shutil.which("soffice"),
            "pdftoppm": shutil.which("pdftoppm"),
            "osascript": shutil.which("osascript"),
        },
    }
    result["ok_for_text_extraction"] = result["modules"]["pptx"] and result["modules"]["PIL"]
    result["ok_for_soffice_visual_rendering"] = bool(result["binaries"]["soffice"] and result["binaries"]["pdftoppm"])
    result["ok_for_mac_powerpoint_batch_rendering"] = bool(result["binaries"]["osascript"])
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok_for_text_extraction"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
