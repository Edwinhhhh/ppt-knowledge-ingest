#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import argparse
import json
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parents[1]
CONVERTER = SKILL_ROOT / "scripts" / "ppt_to_knowledge.py"
DEFAULT_OUTPUT_ROOT = Path.home() / "ppt-knowledge-library"
LIBRARY_ROOT = DEFAULT_OUTPUT_ROOT
STATUS_PATH = LIBRARY_ROOT / "rendered_batch_status.jsonl"


def load_converter():
    spec = importlib.util.spec_from_file_location("ppt_to_knowledge", CONVERTER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot import converter: {CONVERTER}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def append_status(row: dict[str, Any]) -> None:
    row = {"time": datetime.now(timezone.utc).isoformat(), **row}
    STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with STATUS_PATH.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(row, ensure_ascii=False) + "\n")


def unique_deck_dirs() -> list[Path]:
    catalog_path = LIBRARY_ROOT / "catalog.jsonl"
    seen: set[Path] = set()
    deck_dirs: list[Path] = []
    for line in catalog_path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        row = json.loads(line)
        output_dir = row.get("output_dir")
        if not output_dir:
            continue
        deck_dir = Path(output_dir).resolve()
        if deck_dir in seen:
            continue
        seen.add(deck_dir)
        deck_dirs.append(deck_dir)
    return deck_dirs


def escape_applescript_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def export_pdf_with_powerpoint(pptx_path: Path, pdf_path: Path, timeout_seconds: int = 420) -> None:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    script = f"""
tell application "Microsoft PowerPoint"
  activate
  with timeout of 900 seconds
    open POSIX file {escape_applescript_string(str(pptx_path))}
    delay 1
    set thePresentation to active presentation
    save thePresentation in POSIX file {escape_applescript_string(str(pdf_path))} as save as PDF
    close thePresentation saving no
  end timeout
end tell
"""
    clicker = subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "--click-powerpoint-dialogs"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            text=True,
            capture_output=True,
            timeout=max(timeout_seconds, 960),
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or f"osascript exited {result.returncode}")
    finally:
        clicker.terminate()
        try:
            clicker.wait(timeout=3)
        except subprocess.TimeoutExpired:
            clicker.kill()
    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        raise RuntimeError(f"PowerPoint did not create PDF: {pdf_path}")


def reset_powerpoint() -> None:
    try:
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft PowerPoint" to quit saving no'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
        )
    except Exception:
        pass
    time.sleep(2)
    subprocess.run(["pkill", "-x", "Microsoft PowerPoint"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(2)


def click_powerpoint_dialogs() -> int:
    buttons = [
        "Grant Access",
        "Allow",
        "OK",
        "Open",
        "Continue",
        "Replace",
        "Save",
        "Choose...",
        "允许",
        "好",
        "确定",
        "打开",
        "继续",
        "替换",
        "保存",
        "选择...",
    ]
    button_list = ", ".join(json.dumps(button, ensure_ascii=False) for button in buttons)
    script = f"""
set buttonNames to {{{button_list}}}
tell application "System Events"
  repeat with buttonName in buttonNames
    try
      tell process "Microsoft PowerPoint"
        if exists window 1 then
          click button (buttonName as text) of window 1
          return
        end if
      end tell
    end try
  end repeat
end tell
"""
    while True:
        subprocess.run(["osascript", "-e", script], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(0.7)


def main() -> int:
    global LIBRARY_ROOT, STATUS_PATH
    parser = argparse.ArgumentParser(description="Render visual HTML for an existing PPT knowledge library.")
    parser.add_argument("--output-root", type=Path, default=DEFAULT_OUTPUT_ROOT, help="Knowledge library output root containing catalog.jsonl and decks/.")
    args = parser.parse_args()
    LIBRARY_ROOT = args.output_root.expanduser().resolve()
    STATUS_PATH = LIBRARY_ROOT / "rendered_batch_status.jsonl"

    converter = load_converter()
    deck_dirs = unique_deck_dirs()
    completed = 0
    skipped = 0
    failed = 0

    print(json.dumps({"deck_count": len(deck_dirs), "status": str(STATUS_PATH)}, ensure_ascii=False))
    for index, deck_dir in enumerate(deck_dirs, start=1):
        metadata_path = deck_dir / "metadata.json"
        try:
            metadata = read_json(metadata_path)
            source_file = metadata.get("source_file")
            if not source_file:
                raise RuntimeError("metadata.json has no source_file")
            pptx_path = Path(source_file)
            if not pptx_path.exists():
                raise FileNotFoundError(pptx_path)

            rendered_html = deck_dir / "rendered.html"
            if rendered_html.exists():
                skipped += 1
                append_status({"status": "skipped_existing", "index": index, "deck_dir": str(deck_dir), "source_file": str(pptx_path)})
                print(json.dumps({"index": index, "status": "skipped_existing", "deck_dir": str(deck_dir)}, ensure_ascii=False), flush=True)
                continue

            pdf_path = deck_dir / "assets" / "rendered_source.pdf"
            started = time.time()
            if not pdf_path.exists() or pdf_path.stat().st_size == 0:
                export_pdf_with_powerpoint(pptx_path, pdf_path)

            result = converter.process_deck(
                pptx_path=pptx_path,
                output_dir=deck_dir,
                archive_source=False,
                tag_keywords=None,
                source_sha256=metadata.get("sha256"),
                standalone_html=False,
                rendered_html=True,
                rendered_standalone_html=False,
                rendered_source_pdf=pdf_path,
            )
            completed += 1
            append_status(
                {
                    "status": "rendered",
                    "index": index,
                    "deck_dir": str(deck_dir),
                    "source_file": str(pptx_path),
                    "slide_count": result.get("slide_count"),
                    "seconds": round(time.time() - started, 1),
                }
            )
            print(
                json.dumps(
                    {
                        "index": index,
                        "status": "rendered",
                        "slide_count": result.get("slide_count"),
                        "deck_dir": str(deck_dir),
                    },
                    ensure_ascii=False,
                ),
                flush=True,
            )
        except Exception as exc:
            failed += 1
            append_status({"status": "failed", "index": index, "deck_dir": str(deck_dir), "error": str(exc)})
            print(json.dumps({"index": index, "status": "failed", "deck_dir": str(deck_dir), "error": str(exc)}, ensure_ascii=False), flush=True)
            reset_powerpoint()

    manifest_path = LIBRARY_ROOT / "library_manifest.json"
    if manifest_path.exists():
        manifest = read_json(manifest_path)
        manifest["visual_rendered_count"] = len(
            [p for p in deck_dirs if (p / "rendered.html").exists()]
        )
        manifest["visual_rendered_mode"] = "linked_html"
        manifest["visual_standalone_html_count"] = len([p for p in deck_dirs if (p / "rendered.standalone.html").exists()])
        manifest["visual_rendered_updated_at"] = datetime.now(timezone.utc).isoformat()
        manifest["rendered_batch_status"] = str(STATUS_PATH)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({"completed": completed, "skipped": skipped, "failed": failed}, ensure_ascii=False))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--click-powerpoint-dialogs":
        sys.exit(click_powerpoint_dialogs())
    sys.exit(main())
