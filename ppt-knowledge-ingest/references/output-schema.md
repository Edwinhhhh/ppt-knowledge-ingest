# Output Schema

This skill writes one folder per deck.

In library mode, it also writes a root-level catalog and combined retrieval file.

## Files

`content.md`

- Human-readable Markdown.
- Organized by slide.
- Keeps slide title, extracted text, paragraph indentation, approximate layout groups, tables, chart summaries/data, notes, tags, and source slide number.

`content.html`

- Static browser view.
- Uses the same structured slide data as `content.md`.
- Shows extracted embedded images from `assets/images/`.
- Shows approximate horizontal layout groups for side-by-side consulting slide structures.
- Intended for browsing and review, not pixel-perfect PowerPoint reproduction.

`content.standalone.html`

- Optional self-contained browser view written with `--standalone-html`.
- Embeds extracted images as Base64 data URIs.
- Useful when the user wants one PPT to become one portable HTML file.
- Larger than `content.html`, but easier to share or archive because it does not depend on the `assets/images/` folder.

`rendered.html`

- Optional visual-fidelity HTML written with `--rendered-html`.
- Converts the PPTX to PDF with `soffice`, renders each slide to PNG with `pdftoppm`, then places each slide image in HTML.
- Can also use `--rendered-source-pdf` to render from an existing PDF exported by PowerPoint/Keynote.
- Best for "same PPT, different format" use cases where layout, proportions, typography, and visual composition matter more than text editability.
- Depends on `assets/rendered_slides/`.

`rendered.standalone.html`

- Optional self-contained visual-fidelity HTML written with `--rendered-standalone-html`.
- Embeds rendered slide PNGs as Base64 data URIs.
- Usually larger than `rendered.html`, but it is the closest one-file artifact for sharing or archival.

`chunks.jsonl`

- Retrieval/RAG layer.
- JSON Lines format: one JSON object per slide.
- Better for batch ingestion than one large JSON file because each line can be streamed, retried, embedded, or indexed independently.

Typical chunk:

```json
{
  "id": "deck-title#slide-004",
  "deck_id": "deck-title",
  "slide_no": 4,
  "source_file": "/absolute/path/to/original-deck.pptx",
  "source_slide": 4,
  "title": "Slide title",
  "text": "Extracted slide text",
  "layout_groups": [
    {
      "group_no": 2,
      "layout": "columns",
      "block_ids": [3, 4, 5],
      "texts": ["Current state", "Challenge", "Recommendation"]
    }
  ],
  "charts": [
    {
      "chart_type": "COLUMN_CLUSTERED (51)",
      "categories": ["2022", "2023", "2024"],
      "series": [{"name": "Revenue", "values": [100.0, 123.0, 156.0]}]
    }
  ],
  "tags": ["tag"],
  "source": {"type": "pptx", "file": "/absolute/path/to/original-deck.pptx", "slide": 4},
  "assets": {"images": ["assets/images/slide-004-image-01.png"]}
}
```

Key provenance fields:

- `source_file`: absolute path of the original PPTX at conversion time.
- `source_slide`: original PowerPoint slide number for this chunk.
- `source_url`: original online document URL when the deck was imported from a supported cloud link such as Feishu/Lark Wiki or Drive.
- `source`: compact provenance object for integrations that prefer grouped source metadata.

`slides.json`

- Structured intermediate layer.
- Keeps fuller per-slide extraction than `chunks.jsonl`.
- Includes:
  - `rich_text_blocks`: text box bounding boxes, paragraph levels, and simple run formatting.
  - `layout_groups`: approximate row/column grouping based on text box coordinates.
  - `charts`: chart type, categories, series names, and series values when `python-pptx` exposes them.
  - `special_objects`: detected SmartArt/OLE objects that need visual/OXML follow-up.
- Use it to regenerate Markdown, HTML, retrieval chunks, summaries, or custom exports without reparsing the PPTX.

`metadata.json`

- Deck identity and processing record.
- Includes title, source path, file size, SHA-256 hash, slide count, generation timestamp, core Office properties, output paths, and limitations.
- Includes `source_url` when imported from a supported online link.
- Includes PowerPoint sections when present and tagging mode metadata.
- SHA-256 is a digital fingerprint used to detect duplicate or changed source files; it is not encryption and cannot recover the original content.

`manifest.json`

- Small pointer file listing generated artifacts.
- Useful for automation that wants to discover available outputs without knowing every filename in advance.

`catalog.jsonl`

- Library-mode inventory.
- One JSON object per discovered PPTX.
- Tracks original path, filename, size, modified time, SHA-256, status, output directory, duplicate mapping, and errors if any.
- Tracks `source_url` for supported online imports.
- Status values:
  - `pending`: discovered by `--dry-run` and not yet converted.
  - `processed`: converted in the current run.
  - `skipped_existing`: skipped because `manifest.json` already exists and resume mode is enabled.
  - `duplicate`: not converted because another file with the same SHA-256 was already handled.
  - `failed`: conversion failed, but the run continued.

`failed.jsonl`

- Library-mode failures only.
- Empty when all files were processed, skipped, or deduplicated successfully.
- Useful for retry queues.

`all_chunks.jsonl`

- Library-mode combined retrieval file.
- Concatenates chunks from all unique processed/existing deck output folders.
- Keeps `source_file` and `source_slide` on every chunk so downstream search results can open or locate the original PPTX and slide directly.
- Adds `deck_dir` and `source_sha256` to each chunk so downstream systems can also trace retrieval hits back to a deck folder and source fingerprint.

`library_manifest.json`

- Library-mode summary.
- Includes input/output roots, scanned PPTX count, processed count, skipped count, duplicate count, failed count, all-chunks count, and paths to catalog/failure/chunk files.

`assets/images/`

- Embedded image files extracted from the PPTX.
- These are not full slide screenshots.
- Use them for HTML display, visual follow-up analysis, or asset reuse.

## Tags

Default tagging is empty. Pass `--tag-keywords` with a newline-delimited text file or JSON list for deterministic keyword tags. For broad consulting libraries, prefer a later LLM labeling pass so tags can be project- and domain-aware.

`source/original.pptx`

- Optional source copy written only in archival mode.
- Use when no-loss provenance and future reprocessing matter more than storage size.

## Recommended Modes

Lightweight mode:

```text
content.md
content.html
content.standalone.html  # optional with --standalone-html
rendered.html             # optional with --rendered-html
rendered.standalone.html  # optional with --rendered-standalone-html
chunks.jsonl
slides.json
metadata.json
manifest.json
assets/images/
```

Archival mode adds:

```text
source/original.pptx
```

Library mode:

```text
catalog.jsonl
failed.jsonl
all_chunks.jsonl
library_manifest.json
decks/
```

Use library mode for large local drives or nested folders:

```bash
python3 scripts/ppt_to_knowledge.py \
  --library-root /path/to/local-ppt-archive \
  --output-root /path/to/knowledge-library
```

Library mode defaults:

- Resume is on: existing deck folders with `manifest.json` are marked `skipped_existing`.
- Duplicate skipping is on: repeated SHA-256 files are marked `duplicate`.

Optional switches:

- `--dry-run`: scan files and write `catalog.jsonl` / `library_manifest.json` without converting decks or rebuilding `all_chunks.jsonl`.
- `--exclude PATTERN`: exclude paths by glob pattern; can be repeated. The pattern is matched against the relative path, filename, and path parts.
- `--standalone-html`: write `content.standalone.html` for each converted deck.
- `--rendered-html`: write visual-fidelity `rendered.html` and `assets/rendered_slides/`.
- `--rendered-standalone-html`: write one-file visual-fidelity `rendered.standalone.html`.
- `--rendered-source-pdf PATH`: use an existing PDF for visual HTML rendering; supported with single `--pptx` conversions.
- `--no-resume`: force reprocessing even when output already exists.
- `--include-duplicates`: process duplicate source files instead of cataloging them as duplicates.
