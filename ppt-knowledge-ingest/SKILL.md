---
name: ppt-knowledge-ingest
description: Convert PowerPoint .pptx decks, especially consulting decks and knowledge-heavy slideware, into AI-ready knowledge assets: human-readable Markdown, static HTML, JSONL retrieval chunks, structured slide JSON, metadata, approximate slide layout structure, chart data, extracted embedded images, and resumable local knowledge libraries. Use when users ask to ingest, archive, batch-convert, deduplicate, catalog, index, search-enable, or prepare PPT/PPTX files for AI/RAG/knowledge-base workflows.
---

# PPT Knowledge Ingest

## Overview

Use this skill to turn `.pptx` decks into a lightweight, AI-readable asset folder. The default output favors later retrieval and browsing over visual-perfect slide reproduction.

Default output:

```text
deck-name/
  content.md
  content.html
  content.standalone.html  # optional with --standalone-html
  rendered.html             # optional with --rendered-html
  rendered.standalone.html  # optional with --rendered-standalone-html; avoid for large archives
  chunks.jsonl
  slides.json
  metadata.json
  manifest.json
  assets/images/
```

Optional archival output:

```text
deck-name/
  source/original.pptx
```

Library-mode output:

```text
knowledge-library/
  catalog.jsonl
  failed.jsonl
  all_chunks.jsonl
  library_manifest.json
  decks/
    sha-prefix-deck-slug/
      content.md
      content.html
      chunks.jsonl
      slides.json
      metadata.json
      manifest.json
      assets/images/
```

## When To Use

Use this skill when the user wants to:

- Convert PPTX decks into Markdown, HTML, JSON, or JSONL.
- Prepare consulting/project decks for AI retrieval, RAG, or a knowledge base.
- Extract slide-level text, tables, chart summaries, notes, tags, and embedded images.
- Create a browser-friendly HTML view of slide content without generating a new PPT.
- Create a single self-contained HTML file per PPT with embedded image data.
- Create visual-fidelity HTML where each original slide is rendered as an image, preserving PPT layout and proportions.
- Batch-process a folder of `.pptx` files into a consistent asset structure.
- Scan a local hard drive/folder of PPTX files, deduplicate by SHA-256, resume prior work, and produce one combined retrieval file.
- Run a local desktop-feeling Web UI for scanning, converting, browsing, and reviewing the PPT knowledge library.
- Import an online `.pptx` URL, including Feishu/Lark Wiki or Drive links that resolve to a `.pptx` file when `lark-cli` is installed and authenticated.

Do not use this skill to create a new presentation or visually redesign slides. For visual slide generation/editing, use the presentation-focused skills instead.

## Quick Start

Prefer the bundled script for deterministic conversion:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name
```

The script requires Python with `python-pptx` available. In Codex desktop workspaces, prefer the bundled Python runtime when available.

To launch the local Web UI:

```bash
python3 scripts/ppt_knowledge_ui/server.py 8787
```

Then open:

```text
http://127.0.0.1:8787/
```

The Web UI is a local-only tool. It reads local PPTX files, runs the bundled conversion script, and serves generated artifacts from the output folder.

The Web UI also supports an online-link mode. Direct `.pptx` download links are fetched over HTTP. Feishu/Lark Wiki or Drive links are resolved with `lark-cli drive +inspect` and downloaded with `lark-cli drive +download`; the original cloud URL is written as `source_url` in metadata, catalog rows, and retrieval chunks.

For batch conversion:

```bash
python3 scripts/ppt_to_knowledge.py \
  --input-dir /absolute/path/to/pptx-folder \
  --output-root /absolute/path/to/output-root
```

For a full local knowledge library with recursive scan, deduplication, resume, catalog, and combined chunks:

```bash
python3 scripts/ppt_to_knowledge.py \
  --library-root /absolute/path/to/local-ppt-archive \
  --output-root /absolute/path/to/knowledge-library
```

Library mode resumes by default and skips duplicate SHA-256 files by default. Use these switches only when needed:

```bash
python3 scripts/ppt_to_knowledge.py \
  --library-root /absolute/path/to/local-ppt-archive \
  --output-root /absolute/path/to/knowledge-library \
  --dry-run \
  --exclude "Archive/*" \
  --exclude ".DS_Store"
```

Other library-mode switches:

```bash
python3 scripts/ppt_to_knowledge.py \
  --library-root /absolute/path/to/local-ppt-archive \
  --output-root /absolute/path/to/knowledge-library \
  --no-resume \
  --include-duplicates
```

To keep a copy of the source PPTX for no-loss archival:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name \
  --archive-source
```

To apply deterministic keyword tags, pass a newline-delimited text file or JSON list:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name \
  --tag-keywords /absolute/path/to/tags.txt
```

To create one self-contained HTML file per PPT, add:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name \
  --standalone-html
```

This writes `content.standalone.html` in addition to `content.html`.

To preserve the original PPT visual layout and proportions, render each slide as an image-backed HTML:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name \
  --rendered-html
```

This writes `rendered.html`, which references `assets/rendered_slides/`.

Only add `--rendered-standalone-html` for single-deck sharing. It embeds rendered slide images into one large HTML file and is not recommended for large archives.

Use rendered HTML when the user wants "same PPT, different format." Use `content.html` when the user wants AI-readable browsing and structure.

If PPTX-to-PDF rendering is unavailable or unreliable, provide an existing PDF exported from PowerPoint/Keynote:

```bash
python3 scripts/ppt_to_knowledge.py \
  --pptx /absolute/path/to/deck.pptx \
  --output-dir /absolute/path/to/output/deck-name \
  --rendered-source-pdf /absolute/path/to/deck.pdf \
  --rendered-html
```

To generate linked visual HTML for an existing library via Microsoft PowerPoint on macOS:

```bash
python3 scripts/batch_render_visual_html.py \
  --output-root /absolute/path/to/knowledge-library
```

## Workflow

1. Identify whether the user wants lightweight mode or archival mode.
   - Lightweight mode is the default and does not copy the original PPTX.
   - Archival mode uses `--archive-source` and writes `source/original.pptx`.
2. Choose output location.
   - For a single deck, use a clear deck-specific folder.
   - For batch input, use `--output-root`; the script creates one folder per deck.
   - For a large local archive, prefer `--library-root` over `--input-dir`.
3. Decide whether the scale requires library mode.
   - Use `--input-dir` for a small, flat folder.
   - Use `--library-root` for many decks, nested folders, deduplication, resume, `catalog.jsonl`, and `all_chunks.jsonl`.
4. For a large or unfamiliar drive, run `--dry-run` first.
   - Use `--exclude` one or more times to skip folders or filename patterns.
   - Inspect `catalog.jsonl` before running conversion.
5. Decide whether deterministic tags are needed.
   - Default tagging is empty; prefer later LLM labeling for broad or unknown domains.
   - Use `--tag-keywords` only when the user provides a controlled vocabulary.
6. Run the script.
7. Inspect `manifest.json`, `metadata.json`, and the start of `content.md` or `content.html`.
8. In library mode, inspect `library_manifest.json`, `catalog.jsonl`, `failed.jsonl`, and `all_chunks.jsonl`.
9. Spot-check `slides.json` or `chunks.jsonl` for layout groups and chart data on a representative slide.
10. For visual review, prefer `rendered.html + assets/rendered_slides/` over `rendered.standalone.html` when processing many PPTs.
11. Report the output folder and note any limitations, especially image-only slide text or missing visual slide rendering.

## Output Semantics

Read `references/output-schema.md` when the user asks what each file means, wants to customize the schema, or plans to publish/integrate the output format.

Core meanings:

- `content.md`: human-readable slide content.
- `content.html`: browser-friendly static view.
- `content.standalone.html`: optional single-file HTML with embedded images, written with `--standalone-html`.
- `rendered.html`: optional visual-fidelity HTML where each slide is a rendered image, written with `--rendered-html`.
- `rendered.standalone.html`: optional self-contained visual-fidelity HTML, written with `--rendered-standalone-html`.
- `chunks.jsonl`: one JSON object per slide for retrieval/RAG ingestion. Each chunk includes `source_file` and `source_slide` so search hits can point back to the original PPTX path and slide number.
- Online imports also include `source_url` so search hits can point back to the original cloud document link.
- `slides.json`: structured intermediate layer used to regenerate other formats; includes rich text blocks, paragraph levels, bounding boxes, layout groups, chart data, and detected special objects when available.
- `metadata.json`: file identity and processing record.
- `assets/images/`: extracted embedded image assets.
- `source/original.pptx`: optional original deck copy for no-loss archival.
- `catalog.jsonl`: library-mode file inventory with path, SHA-256, status, duplicate mapping, and output folder.
- `failed.jsonl`: library-mode conversion failures that did not stop the whole run.
- `all_chunks.jsonl`: library-mode combined retrieval chunks across all unique processed decks, preserving original PPTX path and slide number on each line.
- `library_manifest.json`: library-mode summary counts and pointers.

## Current Limits

- The script extracts editable PPTX text, paragraph levels, approximate layout groups, chart categories/values when accessible, and embedded images; it does not render full slide screenshots.
- Full visual HTML requires `soffice` and `pdftoppm`; when available, rendered HTML preserves slide layout by using full-slide images.
- If `soffice` is unavailable, use `--rendered-source-pdf` with a PDF exported by PowerPoint/Keynote.
- Image-only text requires OCR or a later vision pass.
- Visual-perfect HTML reproduction is outside the default scope. The generated HTML is a knowledge browsing view, not a pixel-perfect PPT clone.
- SmartArt and OLE objects are detected but not fully parsed.
- The script expects `.pptx`, not legacy `.ppt`.
