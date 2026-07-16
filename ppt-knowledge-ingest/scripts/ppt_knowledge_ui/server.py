#!/usr/bin/env python3
from __future__ import annotations

import json
import hashlib
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
APP_DIR = Path(__file__).resolve().parent
PYTHON = Path("/Users/bytedance/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3")
DEFAULT_OUTPUT_ROOT = ROOT / "outputs" / "ppt-knowledge-library"
SCRIPT_CANDIDATES = [
    APP_DIR.parent / "ppt_to_knowledge.py",
    ROOT / "outputs" / "ppt-knowledge-ingest" / "scripts" / "ppt_to_knowledge.py",
    ROOT / "scripts" / "ppt_to_knowledge.py",
]
SCRIPT = next((path for path in SCRIPT_CANDIDATES if path.exists()), SCRIPT_CANDIDATES[0])
LARK_CLI_CANDIDATES = [Path("/Users/bytedance/.trae/binaries/node/versions/v24.14.0/bin/lark-cli")]
if shutil.which("lark-cli"):
    LARK_CLI_CANDIDATES.append(Path(shutil.which("lark-cli") or ""))
LARK_CLI = next((path for path in LARK_CLI_CANDIDATES if path and path.exists()), None)

TASKS: dict[str, dict[str, Any]] = {}


def jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows = []
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def safe_path(raw: str) -> Path:
    return Path(raw).expanduser().resolve()


def is_http_url(raw: str) -> bool:
    parsed = urllib.parse.urlparse(raw)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def is_lark_url(raw: str) -> bool:
    host = urllib.parse.urlparse(raw).netloc.lower()
    return any(domain in host for domain in ["larkoffice.com", "feishu.cn", "larksuite.com"])


def safe_name(text: str) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "-", text).strip()
    return re.sub(r"\s+", " ", text)[:120] or "deck"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def single_deck_output_dir(output_root: Path, source_sha256: str, pptx_path: Path) -> Path:
    slug = safe_name(pptx_path.stem).replace(" ", "-")
    return output_root / "decks" / f"{source_sha256[:16]}-{slug[:72]}"


def filename_from_url(url: str, headers: Any) -> str:
    disposition = headers.get("Content-Disposition", "") if headers else ""
    match = re.search(r"filename\\*=UTF-8''([^;]+)", disposition, re.I)
    if match:
        return urllib.parse.unquote(match.group(1)).strip("\"")
    match = re.search(r'filename="?([^";]+)"?', disposition, re.I)
    if match:
        return urllib.parse.unquote(match.group(1)).strip("\"")
    parsed = urllib.parse.urlparse(url)
    name = Path(urllib.parse.unquote(parsed.path)).name
    if not name or "." not in name:
        name = f"online-{hashlib.sha1(url.encode('utf-8')).hexdigest()[:10]}.pptx"
    if not name.lower().endswith(".pptx"):
        name = f"{Path(name).stem or 'online-deck'}.pptx"
    return safe_name(name)


def validate_pptx_file(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
    except zipfile.BadZipFile as exc:
        raise ValueError("链接下载到的不是有效 PPTX 文件，可能是网页、登录页或权限页。") from exc
    if "[Content_Types].xml" not in names or not any(name.startswith("ppt/slides/") for name in names):
        raise ValueError("链接下载到的文件不像 PowerPoint PPTX，请确认链接是直接下载地址。")


def parse_lark_cli_json(stdout: str) -> dict[str, Any]:
    decoder = json.JSONDecoder()
    for match in re.finditer(r"{", stdout):
        candidate = stdout[match.start() :]
        try:
            result, _ = decoder.raw_decode(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(result, dict):
            return result
    raise RuntimeError(stdout.strip() or "lark-cli did not return JSON")


def run_lark_cli(args: list[str]) -> dict[str, Any]:
    if not LARK_CLI:
        raise RuntimeError("未找到 lark-cli，无法解析飞书/Wiki 在线链接。")
    env = {
        **os.environ,
        "PATH": f"/Users/bytedance/.trae/binaries/node/versions/v24.14.0/bin:{'/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:/usr/sbin:/sbin'}",
        "LARKSUITE_CLI_NO_UPDATE_NOTIFIER": "1",
        "LARKSUITE_CLI_NO_SKILLS_NOTIFIER": "1",
    }
    proc = subprocess.run([str(LARK_CLI), *args], cwd=str(ROOT), env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)
    output = "\n".join(part for part in [proc.stdout, proc.stderr] if part)
    if proc.returncode != 0:
        raise RuntimeError(output.strip() or f"lark-cli exited {proc.returncode}")
    result = parse_lark_cli_json(output)
    if result.get("ok") is False:
        raise RuntimeError(json.dumps(result.get("error", result), ensure_ascii=False))
    return result


def move_lark_download_to_output(downloaded_path: Path, output_root: Path, url: str, title: str) -> Path:
    incoming_dir = output_root / "_online_sources"
    incoming_dir.mkdir(parents=True, exist_ok=True)
    filename = safe_name(title or downloaded_path.name or filename_from_url(url, {}))
    if not filename.lower().endswith(".pptx"):
        filename = f"{Path(filename).stem or 'lark-deck'}.pptx"
    target = incoming_dir / f"{int(time.time())}-{filename}"
    shutil.move(str(downloaded_path), target)
    validate_pptx_file(target)
    return target


def download_lark_pptx(url: str, output_root: Path, task: dict[str, Any]) -> Path:
    task["progress"] = {"total": 1, "current": 0, "percent": 5, "status": "running", "label": "解析飞书链接"}
    inspected = run_lark_cli(["drive", "+inspect", "--as", "user", "--url", url, "--json"])
    data = inspected.get("data", {})
    doc_type = data.get("type")
    token = data.get("token")
    title = data.get("title") or "lark-deck.pptx"
    if not token:
        raise RuntimeError("飞书链接解析失败：没有拿到资源 token。")

    download_dir = ROOT / "tmp_lark_online_sources"
    download_dir.mkdir(parents=True, exist_ok=True)
    safe_title = safe_name(title)
    if doc_type == "file":
        if not safe_title.lower().endswith(".pptx"):
            raise RuntimeError(f"这个飞书 Wiki 指向的是文件，但不是 PPTX：{title}")
        relative_output = Path("tmp_lark_online_sources") / f"{int(time.time())}-{safe_title}"
        task["progress"] = {"total": 1, "current": 0, "percent": 20, "status": "running", "label": "下载飞书 PPTX 文件"}
        downloaded = run_lark_cli(["drive", "+download", "--as", "user", "--file-token", token, "--output", str(relative_output), "--overwrite", "--json"])
        saved_path = Path(downloaded.get("data", {}).get("saved_path", ROOT / relative_output))
        return move_lark_download_to_output(saved_path, output_root, url, title)

    if doc_type == "slides":
        relative_dir = Path("tmp_lark_online_sources")
        task["progress"] = {"total": 1, "current": 0, "percent": 20, "status": "running", "label": "导出飞书 Slides 为 PPTX"}
        exported = run_lark_cli(
            [
                "drive",
                "+export",
                "--as",
                "user",
                "--url",
                url,
                "--file-extension",
                "pptx",
                "--file-name",
                safe_title if safe_title.lower().endswith(".pptx") else f"{safe_title}.pptx",
                "--output-dir",
                str(relative_dir),
                "--overwrite",
                "--json",
            ]
        )
        saved = exported.get("data", {}).get("saved_path") or exported.get("data", {}).get("file_path") or exported.get("data", {}).get("output")
        if not saved:
            candidates = sorted(download_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
            if not candidates:
                raise RuntimeError("飞书 Slides 导出完成但没有找到本地 PPTX 文件。")
            saved_path = candidates[0]
        else:
            saved_path = Path(saved)
        return move_lark_download_to_output(saved_path, output_root, url, title)

    raise RuntimeError(f"暂不支持把这个飞书资源类型转成 PPTX：{doc_type}。请先导出为 PPTX/PDF 后再导入。")


def download_online_pptx(url: str, output_root: Path, task: dict[str, Any]) -> Path:
    if not is_http_url(url):
        raise ValueError("请输入 http 或 https 开头的在线链接。")
    if is_lark_url(url):
        return download_lark_pptx(url, output_root, task)
    incoming_dir = output_root / "_online_sources"
    incoming_dir.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 PPT-Knowledge-Ingest/1.0",
            "Accept": "application/vnd.openxmlformats-officedocument.presentationml.presentation,application/octet-stream,*/*",
        },
    )
    task["progress"] = {"total": 1, "current": 0, "percent": 3, "status": "running", "label": "连接在线链接"}
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            final_url = response.geturl()
            headers = response.headers
            filename = filename_from_url(final_url, headers)
            target = incoming_dir / f"{int(time.time())}-{filename}"
            total = int(headers.get("Content-Length") or 0)
            downloaded = 0
            with target.open("wb") as out:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    out.write(chunk)
                    downloaded += len(chunk)
                    if total:
                        percent = 5 + round(min(downloaded / total, 1) * 35)
                        task["progress"] = {
                            "total": 1,
                            "current": 0,
                            "percent": percent,
                            "status": "running",
                            "label": f"下载在线 PPT · {round(downloaded / 1024 / 1024, 1)} MB",
                        }
            validate_pptx_file(target)
            return target
    except Exception as exc:
        raise RuntimeError(
            "在线链接无法直接解析为 PPTX。若这是 WPS/腾讯文档/飞书文档等在线编辑链接，通常需要先在网页中导出为 PPTX/PDF，或接入对应平台授权后再导出。"
            f" 原始错误：{exc}"
        ) from exc


def catalog_row(pptx_path: Path, source_sha256: str, status: str, **extra: Any) -> dict[str, Any]:
    stat = pptx_path.stat()
    row = {
        "path": str(pptx_path),
        "filename": pptx_path.name,
        "size_bytes": stat.st_size,
        "modified_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(stat.st_mtime)),
        "sha256": source_sha256,
        "status": status,
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    row.update(extra)
    return row


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")


def rebuild_all_chunks(output_root: Path, rows: list[dict[str, Any]]) -> int:
    all_chunks_path = output_root / "all_chunks.jsonl"
    seen_dirs: set[Path] = set()
    count = 0
    all_chunks_path.parent.mkdir(parents=True, exist_ok=True)
    with all_chunks_path.open("w", encoding="utf-8") as out:
        for row in rows:
            deck_dir_raw = row.get("output_dir")
            if row.get("status") not in {"processed", "skipped_existing", "duplicate"} or not deck_dir_raw:
                continue
            deck_dir = Path(deck_dir_raw).resolve()
            if deck_dir in seen_dirs:
                continue
            seen_dirs.add(deck_dir)
            chunks_path = deck_dir / "chunks.jsonl"
            metadata = read_json(deck_dir / "metadata.json")
            if not chunks_path.exists():
                continue
            for chunk in jsonl(chunks_path):
                chunk["deck_dir"] = str(deck_dir)
                if metadata.get("source_file"):
                    chunk["source_file"] = metadata["source_file"]
                    if isinstance(chunk.get("source"), dict):
                        chunk["source"].setdefault("file", metadata["source_file"])
                if chunk.get("slide_no") is not None:
                    chunk["source_slide"] = chunk["slide_no"]
                    if isinstance(chunk.get("source"), dict):
                        chunk["source"].setdefault("slide", chunk["slide_no"])
                if metadata.get("sha256"):
                    chunk["source_sha256"] = metadata["sha256"]
                out.write(json.dumps(chunk, ensure_ascii=False) + "\n")
                count += 1
    return count


def refresh_library_index(output_root: Path) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)
    catalog_path = output_root / "catalog.jsonl"
    rows = jsonl(catalog_path)
    failures = [row for row in rows if row.get("status") == "failed"]
    write_jsonl(output_root / "failed.jsonl", failures)
    chunk_count = rebuild_all_chunks(output_root, rows)
    summary = {
        "library_root": "",
        "output_root": str(output_root),
        "dry_run": False,
        "exclude_patterns": [],
        "pptx_count": len(rows),
        "pending_count": sum(1 for row in rows if row.get("status") == "pending"),
        "processed_count": sum(1 for row in rows if row.get("status") == "processed"),
        "skipped_existing_count": sum(1 for row in rows if row.get("status") == "skipped_existing"),
        "duplicate_count": sum(1 for row in rows if row.get("status") == "duplicate"),
        "failed_count": len(failures),
        "all_chunks_count": chunk_count,
        "catalog": str(catalog_path),
        "failed": str(output_root / "failed.jsonl"),
        "all_chunks": str(output_root / "all_chunks.jsonl"),
    }
    (output_root / "library_manifest.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def upsert_catalog_row(output_root: Path, row: dict[str, Any]) -> None:
    catalog_path = output_root / "catalog.jsonl"
    rows = [item for item in jsonl(catalog_path) if item.get("path") != row.get("path")]
    rows.append(row)
    write_jsonl(catalog_path, rows)
    refresh_library_index(output_root)


def progress_from_event(event: dict[str, Any]) -> dict[str, Any]:
    total = int(event.get("total") or 0)
    current = int(event.get("current") or 0)
    percent = round((current / total) * 100) if total else 0
    file_name = Path(event.get("file", "")).name if event.get("file") else ""
    status = event.get("status") or "running"
    if event.get("event") == "progress_complete":
        label = "完成"
    elif total and file_name:
        label = f"{current}/{total} · {file_name}"
    elif status == "running":
        label = "准备中"
    else:
        label = str(status)
    return {
        "event": event.get("event"),
        "total": total,
        "current": current,
        "percent": max(0, min(100, percent)),
        "status": status,
        "file": event.get("file", ""),
        "label": label,
    }


def choose_path(kind: str) -> dict[str, Any]:
    prompts = {
        "source_folder": "选择包含 PPTX 的文件夹",
        "source_file": "选择一个 PPTX 文件",
        "output_folder": "选择输出知识库目录",
    }
    if kind not in prompts:
        raise ValueError(f"unsupported browse kind: {kind}")
    if kind == "source_file":
        script = f'POSIX path of (choose file with prompt "{prompts[kind]}")'
    else:
        script = f'POSIX path of (choose folder with prompt "{prompts[kind]}")'
    proc = subprocess.run(["osascript", "-e", script], text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        return {"cancelled": True, "error": proc.stderr.strip()}
    selected = safe_path(proc.stdout.strip())
    if kind == "source_file" and selected.suffix.lower() != ".pptx":
        raise ValueError("请选择 .pptx 文件")
    return {"cancelled": False, "path": str(selected)}


def update_single_file_catalog(task: dict[str, Any], success: bool) -> None:
    pptx_path = Path(task["source_path"])
    output_root = Path(task["output_root"])
    source_sha256 = task.get("source_sha256") or sha256_file(pptx_path)
    output_dir = Path(task["output_dir"])
    source_url = task.get("source_url")
    if success:
        metadata = read_json(output_dir / "metadata.json")
        if source_url:
            metadata["source_url"] = source_url
            (output_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
            chunks_path = output_dir / "chunks.jsonl"
            if chunks_path.exists():
                chunks = jsonl(chunks_path)
                for chunk in chunks:
                    chunk["source_url"] = source_url
                    if isinstance(chunk.get("source"), dict):
                        chunk["source"]["url"] = source_url
                write_jsonl(chunks_path, chunks)
        row = catalog_row(
            pptx_path,
            source_sha256,
            "processed",
            output_dir=str(output_dir),
            title=metadata.get("title") or pptx_path.stem,
            slide_count=metadata.get("slide_count"),
            source_url=source_url,
        )
    else:
        row = catalog_row(pptx_path, source_sha256, "failed", output_dir=str(output_dir), error=task.get("stderr", ""), source_url=source_url)
    upsert_catalog_row(output_root, row)


def run_single_scan(task_id: str) -> None:
    task = TASKS[task_id]
    task["status"] = "running"
    task["started_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        pptx_path = Path(task["source_path"])
        output_root = Path(task["output_root"])
        source_sha256 = sha256_file(pptx_path)
        output_dir = single_deck_output_dir(output_root, source_sha256, pptx_path)
        status = "skipped_existing" if (output_dir / "manifest.json").exists() else "pending"
        row = catalog_row(pptx_path, source_sha256, status, output_dir=str(output_dir))
        upsert_catalog_row(output_root, row)
        task["source_sha256"] = source_sha256
        task["output_dir"] = str(output_dir)
        task["progress"] = {"total": 1, "current": 1, "percent": 100, "status": status, "label": f"1/1 · {pptx_path.name}"}
        task["status"] = "done"
    except Exception as exc:
        task["status"] = "failed"
        task["stderr"] = str(exc)
        task["progress"] = {"total": 1, "current": 0, "percent": 0, "status": "failed", "label": str(exc)}
    task["finished_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    task["library"] = load_library(Path(task["output_root"]))


def run_online_task(task_id: str, kind: str, options: dict[str, Any]) -> None:
    task = TASKS[task_id]
    task["status"] = "running"
    task["started_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    output_root = Path(task["output_root"])
    try:
        pptx_path = download_online_pptx(task["source_url"], output_root, task)
        source_sha256 = sha256_file(pptx_path)
        output_dir = single_deck_output_dir(output_root, source_sha256, pptx_path)
        task.update(
            {
                "source_path": str(pptx_path),
                "library_root": str(pptx_path),
                "source_sha256": source_sha256,
                "output_dir": str(output_dir),
            }
        )
        if kind == "scan":
            status = "skipped_existing" if (output_dir / "manifest.json").exists() else "pending"
            row = catalog_row(pptx_path, source_sha256, status, output_dir=str(output_dir), source_url=task["source_url"])
            upsert_catalog_row(output_root, row)
            task["progress"] = {"total": 1, "current": 1, "percent": 100, "status": status, "label": f"1/1 · {pptx_path.name}"}
            task["status"] = "done"
        else:
            cmd = [
                str(PYTHON if PYTHON.exists() else sys.executable),
                str(SCRIPT),
                "--pptx",
                str(pptx_path),
                "--output-dir",
                str(output_dir),
                "--progress-jsonl",
            ]
            if options.get("rendered"):
                cmd.append("--rendered-html")
            if options.get("rendered_standalone"):
                cmd.append("--rendered-standalone-html")
            if options.get("standalone"):
                cmd.append("--standalone-html")
            task["cmd"] = cmd
            run_task(task_id, cmd, output_root)
            return
    except Exception as exc:
        task["status"] = "failed"
        task["stderr"] = str(exc)
        task["progress"] = {"total": 1, "current": 0, "percent": 0, "status": "failed", "label": "在线链接解析失败"}
    task["finished_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    task["library"] = load_library(output_root)


def run_task(task_id: str, cmd: list[str], output_root: Path) -> None:
    task = TASKS[task_id]
    task["status"] = "running"
    task["started_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    output_lines: list[str] = []
    try:
        proc = subprocess.Popen(cmd, cwd=str(ROOT), text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=1)
        assert proc.stdout is not None
        for line in proc.stdout:
            output_lines.append(line)
            stripped = line.strip()
            try:
                event = json.loads(stripped)
            except Exception:
                event = {}
            if isinstance(event, dict) and str(event.get("event", "")).startswith(("progress", "file_")):
                task["progress"] = progress_from_event(event)
            task["stdout"] = "".join(output_lines[-300:])
        proc.wait()
        task["returncode"] = proc.returncode
        task["stdout"] = "".join(output_lines)
        task["stderr"] = "" if proc.returncode == 0 else "".join(output_lines[-80:])
        task["status"] = "done" if proc.returncode == 0 else "failed"
    except Exception as exc:
        task["status"] = "failed"
        task["stderr"] = str(exc)
    if task.get("source_type") in {"file", "online"}:
        try:
            update_single_file_catalog(task, task["status"] == "done")
        except Exception as exc:
            task["stderr"] = (task.get("stderr") or "") + f"\nCatalog update failed: {exc}"
    task["finished_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    if task["status"] == "done":
        task["progress"] = {**task.get("progress", {}), "percent": 100, "status": "done"}
    task["library"] = load_library(output_root)


def start_task(kind: str, payload: dict[str, Any]) -> dict[str, Any]:
    source_type = payload.get("source_type") or "folder"
    raw_source = (payload.get("library_root") or "").strip()
    output_root = safe_path(payload.get("output_root") or str(DEFAULT_OUTPUT_ROOT))
    excludes = [item for item in payload.get("exclude", []) if item]
    rendered = bool(payload.get("rendered_html"))
    rendered_standalone = bool(payload.get("rendered_standalone_html"))
    standalone = bool(payload.get("standalone_html"))

    if source_type == "online":
        if not raw_source:
            raise ValueError("请粘贴在线 PPT 链接")
        task_id = f"{int(time.time() * 1000)}"
        TASKS[task_id] = {
            "id": task_id,
            "kind": kind,
            "status": "queued",
            "cmd": [],
            "output_root": str(output_root),
            "library_root": raw_source,
            "source_path": "",
            "source_url": raw_source,
            "source_type": "online",
            "source_sha256": "",
            "output_dir": "",
            "progress": {"total": 1, "current": 0, "percent": 0, "status": "queued", "label": "排队下载在线链接"},
        }
        options = {"rendered": rendered, "rendered_standalone": rendered_standalone, "standalone": standalone}
        thread = threading.Thread(target=run_online_task, args=(task_id, kind, options), daemon=True)
        thread.start()
        return TASKS[task_id]

    source_path = safe_path(raw_source)
    if source_type == "file":
        if not source_path.is_file() or source_path.suffix.lower() != ".pptx":
            raise ValueError("请选择一个 .pptx 文件")
        source_sha256 = sha256_file(source_path)
        output_dir = single_deck_output_dir(output_root, source_sha256, source_path)
        cmd = [str(PYTHON if PYTHON.exists() else sys.executable), str(SCRIPT), "--pptx", str(source_path), "--output-dir", str(output_dir), "--progress-jsonl"]
    else:
        if not source_path.is_dir():
            raise NotADirectoryError(source_path)
        source_sha256 = ""
        output_dir = None
        cmd = [str(PYTHON if PYTHON.exists() else sys.executable), str(SCRIPT), "--library-root", str(source_path), "--output-root", str(output_root), "--progress-jsonl"]
        if kind == "scan":
            cmd.append("--dry-run")
    for item in excludes:
        if source_type != "file":
            cmd.extend(["--exclude", item])
    if rendered:
        cmd.append("--rendered-html")
    if rendered_standalone:
        cmd.append("--rendered-standalone-html")
    if standalone:
        cmd.append("--standalone-html")

    task_id = f"{int(time.time() * 1000)}"
    TASKS[task_id] = {
        "id": task_id,
        "kind": kind,
        "status": "queued",
        "cmd": cmd,
        "output_root": str(output_root),
        "library_root": str(source_path),
        "source_path": str(source_path),
        "source_type": source_type,
        "source_sha256": source_sha256,
        "output_dir": str(output_dir) if output_dir else "",
        "progress": {"total": 0, "current": 0, "percent": 0, "status": "queued", "label": "排队中"},
    }
    target = run_single_scan if kind == "scan" and source_type == "file" else run_task
    args = (task_id,) if target is run_single_scan else (task_id, cmd, output_root)
    thread = threading.Thread(target=target, args=args, daemon=True)
    thread.start()
    return TASKS[task_id]


def load_library(output_root: Path) -> dict[str, Any]:
    output_root = output_root.resolve()
    manifest = read_json(output_root / "library_manifest.json")
    catalog = jsonl(output_root / "catalog.jsonl")
    failed = jsonl(output_root / "failed.jsonl")
    decks = []
    for row in catalog:
        out = row.get("output_dir")
        if not out:
            continue
        deck_dir = Path(out)
        metadata = read_json(deck_dir / "metadata.json")
        manifest_row = read_json(deck_dir / "manifest.json")
        outputs = metadata.get("outputs", {}) if isinstance(metadata.get("outputs"), dict) else {}
        imported_at = metadata.get("generated_at") or row.get("updated_at") or row.get("modified_at")
        source_url = metadata.get("source_url") or row.get("source_url")
        display_source = source_url or metadata.get("source_file") or row.get("path")
        decks.append(
            {
                "path": row.get("path"),
                "filename": row.get("filename"),
                "status": row.get("status"),
                "sha256": row.get("sha256"),
                "imported_at": imported_at,
                "updated_at": row.get("updated_at"),
                "modified_at": row.get("modified_at"),
                "source_file": display_source,
                "source_url": source_url,
                "output_dir": str(deck_dir),
                "title": metadata.get("title") or row.get("title") or row.get("filename"),
                "slide_count": metadata.get("slide_count") or row.get("slide_count"),
                "has_rendered_html": (deck_dir / "rendered.html").exists(),
                "has_chunks": (deck_dir / "chunks.jsonl").exists(),
                "has_markdown": (deck_dir / "content.md").exists(),
                "visual_mode": "linked" if (deck_dir / "rendered.html").exists() else "none",
                "outputs": outputs,
                "manifest": manifest_row,
            }
        )
    return {
        "output_root": str(output_root),
        "manifest": manifest,
        "catalog": catalog,
        "failed": failed,
        "decks": decks,
    }


def deck_detail(deck_dir: Path) -> dict[str, Any]:
    deck_dir = deck_dir.resolve()
    metadata = read_json(deck_dir / "metadata.json")
    manifest = read_json(deck_dir / "manifest.json")
    content_md = deck_dir / "content.md"
    return {
        "deck_dir": str(deck_dir),
        "metadata": metadata,
        "manifest": manifest,
        "markdown": content_md.read_text(encoding="utf-8") if content_md.exists() else "",
        "links": {
            key: local_file_url(deck_dir / value)
            for key, value in manifest.items()
            if isinstance(value, str) and value.endswith((".html", ".md", ".json", ".jsonl"))
        },
    }


def local_file_url(path: Path) -> str:
    return "/local" + urllib.parse.quote(str(path.resolve()))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def send_json(self, data: Any, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        if parsed.path == "/":
            self.serve_file(APP_DIR / "index.html")
        elif parsed.path == "/app.js":
            self.serve_file(APP_DIR / "app.js")
        elif parsed.path == "/styles.css":
            self.serve_file(APP_DIR / "styles.css")
        elif parsed.path == "/api/task":
            task_id = qs.get("id", [""])[0]
            self.send_json(TASKS.get(task_id, {"error": "task not found"}), 200 if task_id in TASKS else 404)
        elif parsed.path == "/api/library":
            output_root = safe_path(qs.get("output_root", [str(DEFAULT_OUTPUT_ROOT)])[0])
            self.send_json(load_library(output_root))
        elif parsed.path == "/api/deck":
            deck_dir = safe_path(qs.get("deck_dir", [""])[0])
            self.send_json(deck_detail(deck_dir))
        elif parsed.path == "/file":
            target = safe_path(qs.get("path", [""])[0])
            self.serve_file(target)
        elif parsed.path.startswith("/local/"):
            target = safe_path(urllib.parse.unquote(parsed.path.removeprefix("/local")))
            self.serve_file(target)
        else:
            self.send_json({"error": "not found"}, 404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/scan":
            try:
                self.send_json(start_task("scan", self.read_body()))
            except Exception as exc:
                self.send_json({"error": str(exc)}, 400)
        elif parsed.path == "/api/convert":
            try:
                self.send_json(start_task("convert", self.read_body()))
            except Exception as exc:
                self.send_json({"error": str(exc)}, 400)
        elif parsed.path == "/api/browse":
            try:
                self.send_json(choose_path(self.read_body().get("kind", "")))
            except Exception as exc:
                self.send_json({"error": str(exc)}, 400)
        else:
            self.send_json({"error": "not found"}, 404)

    def serve_file(self, path: Path) -> None:
        if not path.exists() or not path.is_file():
            self.send_json({"error": f"file not found: {path}"}, 404)
            return
        mime = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8787
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"PPT Knowledge UI running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
