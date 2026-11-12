# Workbook interior generator

Generates the Darkroom Log workbook interiors as Word (`.docx`) files and exports them to PDF.

## Files

| File | Purpose |
|---|---|
| `build_workbook.js` | builds one variant as `.docx` (layout, tables, pages) |
| `text_en.js` | all English texts and example entries |
| `build_all.py` | builds all variants, adds mirrored margins (ring binding), exports PDFs, checks page counts |
| `preview.py` | builds only selected page types in A5 and A4 and renders them as PNG (`dist/preview/`) |

## Requirements

- Node.js 18+ with `docx`, `sharp` (torn-paper edges, circled IDs) and `qrcode`: run `npm install` in the repository root
- Python 3 with Pillow (`pip install -r requirements.txt`)
- LibreOffice (`soffice`) for the PDF export. Set `SOFFICE_CMD` to use a different command.
- `pdfinfo` (poppler-utils) for the page-count check
- Fonts Courier Prime and Caveat installed (see `fonts/README.md`)
- Online versions: covers in `dist/covers/` (`python3 src/covers/make_covers.py`)

## Usage

From the repository root:

```
python3 src/workbook/build_all.py            # all variants
python3 src/workbook/build_all.py online     # only names containing "online"
```

Output is written to `dist/workbook/`:

| File | Variant |
|---|---|
| `darkroom-log_en_A5_ring-80` / `_A4_ring-80` | ring binding, 80 pages, 3 mm bleed, mirrored margins (inner 20 mm, outer 10 mm) |
| `darkroom-log_en_A5_booklet-40` / `_A4_booklet-40` | booklet, 40 pages, 3 mm bleed |
| `darkroom-log_en_A5_online-40_<light\|dark\|red>` / `_A4_…` | online / self-print, 40 pages + front and back cover page in the chosen cover theme, no bleed, clickable contents |
| `darkroom-log_en_A4_howto-extended` | extended 2-page how-to (A4, online only) |

## How the layout works

- **Widths** are designed for A5 (124 mm content width) and scaled to A4 (186 mm). Type sizes and fixed heights scale by 1.25 on A4.
- **Page height:** tables that end mid-page are extended with rows or note lines to the bottom margin.
- **How-to pages** are fitted automatically: they are scaled up or down in small steps until they fill the page (the console reports e.g. "fitted page 2 at 94 %").
- **Example graphics** (torn-paper edges, circled print IDs, QR code) are rendered as images during the build, in a first pass that collects them.
- **QR code / links:** `T.ext.howtoUrl` and `T.ext.repoUrl` in `text_en.js`.
- **Mirrored margins** are not supported by the `docx` library, so `build_all.py` adds them to the finished file.

## Editing

- **Texts:** edit `text_en.js`.
- **Layout:** the page functions in `build_workbook.js` (`pageSetup`, `sessionBlock`, `pageQR1` …).
- **Small one-off fixes:** edit the generated Word file and export it to PDF from Word (embed the fonts). These changes are lost when you rebuild.
