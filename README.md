# PPT Knowledge Ingest

Turn local PowerPoint decks into AI-ready knowledge assets and a local review workbench.

This project packages a Codex Skill plus a local Web UI for converting `.pptx` files into:

- `content.md` for human-readable slide content
- `content.html` for structured browsing
- `rendered.html` for visual review of the original slide layout
- `chunks.jsonl` for retrieval/RAG ingestion
- `slides.json` for structured slide data
- `metadata.json` and `manifest.json` for provenance and automation
- `assets/` for extracted and rendered images

## What It Supports

- Local folder batch conversion with resume, duplicate detection, failure queues, and a combined `all_chunks.jsonl`.
- Single local `.pptx` conversion.
- Online links that directly download `.pptx`.
- Feishu/Lark Wiki or Drive links that point to `.pptx` files, when `lark-cli` is installed and authenticated.
- Visual HTML output via `rendered.html` when a renderer is available.
- A local desktop-feeling Web UI for scanning, converting, sorting, and previewing outputs.

The Web UI runs locally. It can read local file paths only on the machine where the server is running.

## Install As A Codex Skill

Download or clone this repository, then copy the skill folder into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R ppt-knowledge-ingest ~/.codex/skills/
```

Restart Codex so the skill list refreshes.

After installation, you can ask Codex things like:

```text
用 ppt-knowledge-ingest 扫描 ~/Downloads 里的 PPT，输出到 ~/ppt-knowledge-library
```

## Run Directly From The Repository

```bash
python3 -m pip install -r requirements.txt
python3 ppt-knowledge-ingest/scripts/ppt_knowledge_ui/server.py 8787
```

Open:

```text
http://127.0.0.1:8787/
```

## Run The Local Web UI

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/ppt_knowledge_ui/server.py 8787
```

Open:

```text
http://127.0.0.1:8787/
```

The UI runs locally. It does not upload PPT files anywhere.

The local UI supports both typed paths and native file pickers:

- Choose a folder for recursive library conversion.
- Choose a single `.pptx` file for one-deck conversion.
- Paste an online URL that directly downloads a `.pptx` file.
- Paste a Feishu/Lark Wiki or Drive file URL that points to a `.pptx`, if `lark-cli` is installed and logged in as a user.
- Choose an output folder for the generated knowledge library.
- Watch a live progress bar while files are scanned or converted.

Online editor links such as WPS or Tencent Docs often return an HTML page or require login. In that case, export the document as `.pptx`/PDF first, or add a platform-specific authorized exporter.

For Feishu/Lark links, authenticate first:

```bash
lark-cli auth login --domain drive
```

## Command Line Usage

Single deck:

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/ppt_to_knowledge.py \
  --pptx /path/to/deck.pptx \
  --output-dir ~/ppt-knowledge-library/deck-name
```

Recursive library:

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/ppt_to_knowledge.py \
  --library-root ~/Downloads \
  --output-root ~/ppt-knowledge-library
```

Progress events for custom wrappers:

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/ppt_to_knowledge.py \
  --library-root ~/Downloads \
  --output-root ~/ppt-knowledge-library \
  --progress-jsonl
```

Visual HTML:

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/ppt_to_knowledge.py \
  --pptx /path/to/deck.pptx \
  --output-dir ~/ppt-knowledge-library/deck-name \
  --rendered-html
```

For large archives, prefer `rendered.html + assets/rendered_slides/`. Avoid `rendered.standalone.html` unless you need a single shareable file for one deck.

## Dependencies

Required for extraction:

- Python 3.10+
- `python-pptx`
- `Pillow`

Often available in Codex Desktop's bundled Python runtime. If using system Python:

```bash
python3 -m pip install python-pptx Pillow
```

Check the local environment:

```bash
python3 ~/.codex/skills/ppt-knowledge-ingest/scripts/check_environment.py
```

Optional for visual rendering:

- LibreOffice or `soffice`
- `pdftoppm` from Poppler
- Or Microsoft PowerPoint on macOS for `scripts/batch_render_visual_html.py`

## Output Model

Each deck becomes a folder. The important files are:

- `content.md`: readable Markdown
- `content.html`: structured browser view
- `rendered.html`: original-layout visual browser view
- `chunks.jsonl`: AI retrieval chunks, one row per slide
- `slides.json`: full structured extraction
- `metadata.json`: source path, SHA-256, page count, generation metadata
- `manifest.json`: generated artifact index
- `assets/images/`: embedded images extracted from PPTX
- `assets/rendered_slides/`: rendered slide images used by `rendered.html`

In library mode, the root output also includes:

- `catalog.jsonl`: discovered PPTX inventory and status
- `all_chunks.jsonl`: combined retrieval chunks
- `failed.jsonl`: files that failed conversion
- `library_manifest.json`: summary counts and pointers

## Notes

- This tool is designed for `.pptx`, not legacy `.ppt`.
- Image-only slide text requires OCR or a later vision pass.
- SmartArt and OLE objects are detected but not fully parsed.
- `source_file` and `source_slide` are included in retrieval chunks so AI search results can trace back to the original PPT path and slide number.
- For online Feishu/Lark imports, `source_url` is also recorded so AI search results can trace back to the original cloud document link.
