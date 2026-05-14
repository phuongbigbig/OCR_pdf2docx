# OCR_pdf2docx

# Local PDF-to-Word Converter with Table Reconstruction

A command-line tool that converts scanned or digitally-created PDFs into editable Word (.docx) files, preserving document structure including tables, headings, lists, and images. Runs entirely offline using neural layout analysis — no cloud APIs, no data leaves your machine.

Built on [Docling](https://github.com/docling-project/docling) (IBM Research) for document understanding and [python-docx](https://python-docx.readthedocs.io/) for Word output.

## Why This Tool Exists

Standard PDF-to-Word converters (Adobe Acrobat, online tools) often fail on:

- **Scanned documents** where text is baked into images and tables are just lines on a page
- **Complex table layouts** with merged cells, multi-line headers, or nested structures
- **Mixed-language documents** (e.g., Vietnamese + English reports) where OCR engines need explicit language hints
- **Confidential files** that cannot be uploaded to cloud conversion services

This tool addresses all four by running neural network models locally for layout detection, OCR, and table structure recognition, then reconstructing genuine Word tables with editable cell text.

## Two OCR Engine Variants

| File | OCR Engine | Best For |
|------|-----------|----------|
| `ocr_table_rebuild_v2.py` | **EasyOCR** (PyTorch) | Cross-platform (macOS, Linux, Windows). Broad language coverage. Works on any machine with Python. |
| `ocr_table_rebuild_v2_ocrmac.py` | **macOS Vision** (OcrMac) | Apple Silicon Macs only. Hardware-accelerated, faster startup, no PyTorch overhead for OCR. Built-in benchmark timers. |

Both scripts share the same document reconstruction logic and produce identical Word output for the same parsed content. The difference is purely in the OCR step.

## Installation

### Prerequisites

- Python 3.10 or higher
- macOS, Linux, or Windows (OcrMac variant requires macOS)

### Install dependencies

For the EasyOCR variant (cross-platform):

```bash
pip install docling docling-core python-docx pymupdf pandas
```

For the OcrMac variant (macOS only):

```bash
pip install "docling[ocrmac]" docling-core python-docx pymupdf pandas
```

First run will download Docling's neural model weights from HuggingFace (~1–2 GB). Subsequent runs use the cached models.

### Verify installation

```bash
python -c "from docling.document_converter import DocumentConverter; print('Ready')"
```

## Usage

### Command-line arguments

```
-i, --input       Path to the source PDF file (required, or prompted interactively)
-o, --output      Custom output path for the .docx file (default: <input>_structured.docx)
-l, --lang        Space-separated OCR language codes (default: en vi)
-p, --pages       Page range to process, e.g., '2-5' or '4' (default: all pages)
    --scanned     Force full-page OCR on every page (for scanned/image-based PDFs)
    --no-open     Do not auto-open the output file after conversion
```

### Basic examples

Convert an entire PDF with default settings (English + Vietnamese, all pages):

```bash
python3 ocr_table_rebuild_v2.py -i report.pdf
```

Convert only pages 3 through 7:

```bash
python3 ocr_table_rebuild_v2.py -i report.pdf -p 3-7
```

Convert a single page:

```bash
python3 ocr_table_rebuild_v2.py -i invoice.pdf -p 1
```

Convert with French and English OCR:

```bash
python3 ocr_table_rebuild_v2.py -i document.pdf -l fr en
```

Convert a scanned PDF (forces OCR on every page):

```bash
python3 ocr_table_rebuild_v2.py -i scanned_contract.pdf --scanned
```

Custom output path, no auto-open:

```bash
python3 ocr_table_rebuild_v2.py -i data.pdf -o ~/Desktop/converted.docx --no-open
```

### Interactive mode

Run without `-i` to enter interactive mode, which prompts for file path, page range, languages, and scan mode step by step:

```bash
python3 ocr_table_rebuild_v2.py
```

### Benchmarking EasyOCR vs OcrMac

Run both scripts on the same PDF and compare the timing summaries:

```bash
python3 ocr_table_rebuild_v2.py -i test.pdf --no-open
python3 ocr_table_rebuild_v2_ocrmac.py -i test.pdf --no-open
```

The OcrMac variant prints a benchmark summary at the end:

```
⏱️  BENCHMARK SUMMARY (OcrMac / macOS Vision)
=======================================================
   Converter init :     2.34s
   PDF parsing    :    12.07s
   DOCX rebuild   :     0.45s
   ─────────────────────────────
   Total wall-clock:   14.86s
   Pages processed : 5
   Speed           : 2.41s per page (parsing only)
   Tables found    : 3
   Text blocks     : 47
=======================================================
```

## Supported Languages

### EasyOCR variant

Uses standard ISO language codes. Common codes:

| Code | Language | Code | Language |
|------|----------|------|----------|
| `en` | English | `vi` | Vietnamese |
| `fr` | French | `ja` | Japanese |
| `de` | German | `ko` | Korean |
| `es` | Spanish | `zh` | Chinese (Simplified) |
| `pt` | Portuguese | `th` | Thai |
| `ru` | Russian | `ar` | Arabic |

Full list: [EasyOCR supported languages](https://www.jaided.ai/easyocr/)

### OcrMac variant

Uses macOS locale codes. Short codes (en, vi, fr) are automatically converted:

| Short code | Resolves to | Short code | Resolves to |
|-----------|-------------|-----------|-------------|
| `en` | `en-US` | `vi` | `vi-VT` |
| `fr` | `fr-FR` | `ja` | `ja-JP` |
| `de` | `de-DE` | `ko` | `ko-KR` |
| `es` | `es-ES` | `zh` | `zh-Hans` |
| `pt` | `pt-BR` | `th` | `th-TH` |
| `ru` | `ru-RU` | `ar` | `ar-SA` |
| `it` | `it-IT` | `tr` | `tr-TR` |
| `nl` | `nl-NL` | `pl` | `pl-PL` |

You can also pass full locale codes directly (e.g., `-l zh-Hant` for Traditional Chinese).

## What Gets Reconstructed

| PDF Element | Word Output | Notes |
|------------|-------------|-------|
| Tables | Native Word tables (`Table Grid` style) | Editable cells, bolded headers when detected |
| Headings | Word heading styles (Heading 1–9) | Preserves hierarchy levels |
| Body text | Normal paragraphs | Reading order preserved |
| Bullet/numbered lists | `List Bullet` style paragraphs | Falls back to `•` prefix if style unavailable |
| Images/figures | Embedded PNG images | Centered, 5-inch width; requires page images in Docling output |
| Captions | Bracketed text placeholder | When image extraction fails |

### What is NOT reconstructed

- **Font styles** (bold/italic/underline within body text) — all text renders in the default Word font
- **Font sizes and colors** — not carried over from PDF
- **Exact spatial positioning** — content follows reading order, not pixel-level placement
- **Headers and footers** — Docling excludes these by default
- **Form fields** — interactive PDF forms are not converted
- **Annotations and comments** — PDF markup is not carried over
- **Multi-column layouts** — columns are linearized into single-column reading order

## Suitable Use Cases

### Strong fit

- **Digitizing scanned reports and contracts** — extracting editable text and tables from image-based PDFs that other tools cannot parse
- **Vietnamese + English mixed documents** — government reports, academic papers, internal memos with bilingual content
- **Data tables in research PDFs** — extracting tabular results from published papers into editable Word tables for further analysis
- **Confidential document conversion** — legal contracts, medical records, financial statements that must not leave the local machine
- **Batch processing with scripting** — wrapping in a shell loop to convert entire directories of PDFs

### Moderate fit

- **Simple digitally-created PDFs** — works but may be overkill; simpler tools (e.g., `pandoc`, LibreOffice) may suffice
- **Presentation slides** — Docling can parse them but the Word output loses slide layout context
- **PDFs with heavy mathematical formulas** — Docling detects formulas but this script does not render them as equation objects in Word

### Poor fit

- **Pixel-perfect reproduction** — if you need the Word file to look identical to the PDF, use Adobe Acrobat or a dedicated DTP tool
- **Interactive PDF forms** — use a PDF form filler instead
- **PDFs with DRM or password protection** — the script cannot open encrypted files

## Pipeline Architecture

```
PDF file
  │
  ├─ PyMuPDF ──────────── page count, validation
  │
  └─ Docling ──────────── neural document understanding
       │
       ├─ Page preprocessing (image rendering)
       ├─ OCR engine (EasyOCR or macOS Vision)
       ├─ Layout model (Heron — detects regions: text, table, figure, heading)
       └─ Table structure model (TableFormer ACCURATE — detects rows, columns, cells)
       │
       ▼
  DoclingDocument (structured representation)
       │
       ├─ iterate_items() in reading order
       │    ├─ SectionHeaderItem → doc.add_heading()
       │    ├─ TextItem          → doc.add_paragraph()
       │    ├─ ListItem          → doc.add_paragraph(style="List Bullet")
       │    ├─ TableItem         → export_to_dataframe() → doc.add_table()
       │    └─ PictureItem       → get_image() → doc.add_picture()
       │
       ▼
  Word (.docx) file
```

## Configuration Details

### Table structure model

Both scripts use `TableFormerMode.ACCURATE` by default — this is slower than `FAST` mode but significantly better at detecting complex table structures (merged cells, multi-row headers, sparse tables). To switch to fast mode for quick-and-dirty conversions, change this line in the script:

```python
pipeline_options.table_structure_options.mode = TableFormerMode.FAST
```

### Full-page OCR

The `--scanned` flag sets `force_full_page_ocr = True`, which tells the OCR engine to process every page as a full image rather than selectively OCR-ing regions that appear to contain bitmap text. Use this when:

- The PDF is entirely scanned (no embedded text layer)
- Auto-detection misses text in certain regions
- You see blank or garbled text in the output

For digitally-created PDFs (where text is already embedded), leave this off — it adds processing time without benefit.

## Troubleshooting

**"Missing core packages" error on startup**
Run the installation command for your variant. If you see import errors for `docling_core.types.doc`, make sure both `docling` and `docling-core` are installed and at compatible versions.

**Empty or garbled table cells**
Try adding `--scanned` to force full-page OCR. If the table structure itself is wrong (wrong number of rows/columns), the issue is in Docling's TableFormer model — complex nested tables or tables without visible borders are harder to detect.

**OcrMac language error ("Invalid language preference")**
Use the short codes (en, vi, fr) — the script auto-converts them to macOS locale format. If you need Traditional Chinese specifically, pass `-l zh-Hant` directly.

**Slow conversion on large PDFs**
Use `-p` to process only the pages you need. Each page goes through layout detection + OCR + table structure, so processing time scales linearly with page count.

**Model download hangs or fails**
Docling downloads models from HuggingFace on first run. If you are behind a firewall or proxy, pre-download with:
```bash
docling-tools models download
```

**Output file cannot be opened automatically**
The auto-open feature uses `open` (macOS), `xdg-open` (Linux), or `os.startfile` (Windows). Use `--no-open` to skip this step and open the file manually.

## License

This script is provided as-is for personal and research use. Docling is MIT-licensed. Individual model licenses (TableFormer, layout models) are governed by their respective packages.
