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
| `…_ring-80_epubli` / `…_booklet-40_epubli` | the same print versions with epubli's production and GPSR contact lines filled in |
| `darkroom-log_en_A5_online-40_<light\|dark\|red>` / `_A4_…` | online / self-print, 40 pages + front and back cover page in the chosen cover theme, no bleed, clickable contents |
| `darkroom-log_en_A4_howto-extended` | extended 2-page how-to (A4, online only) |
| `darkroom-log_en_A5_online-40_self-print-duplex` (PDF only) | the A5 online version (light cover) as a booklet for home printing: A4 landscape, 2 × A5 per side, sorted for duplex printing and stapling |
| `darkroom-log_en_A5_ring-80_epubli_INT` | personal example: epubli ring version with "INT" pre-printed in the free column |

## Options

- `freeCol: 'INT'` pre-prints a label (max. 6 characters) in the free column after FILTER, e.g. INT for light intensity. Default: empty.
- `printer: 'epubli'` fills in the production and GPSR lines of the imprint (print versions).
- `cover: 'light' | 'dark' | 'red'` chooses the cover pages of the online versions.

Add a line to `VARIANTS` in `build_all.py` to build your own combination.

## Imprint of the print versions

The print versions without suffix leave two fields open for the print service: "Production:" and "Contact according to the EU General Product Safety Regulation (GPSR):". The `_epubli` versions have them filled in (texts in `T.imprint` in `text_en.js`; switch with `printer: 'epubli'` in `build_all.py`).

## How the layout works

- **Widths** are designed for A5 (124 mm content width) and scaled to A4 (186 mm). Type sizes and fixed heights scale by 1.25 on A4.
- **Page height:** tables that end mid-page are extended with rows or note lines to the bottom margin.
- **How-to pages** are fitted automatically: they are scaled up or down in small steps until they fill the page (the console reports e.g. "fitted page 2 at 94 %").
- **Example graphics** (torn-paper edges, circled print IDs, QR code) are rendered as images during the build, in a first pass that collects them.
- **QR code / links:** `T.ext.howtoUrl` and `T.ext.repoUrl` in `text_en.js`.
- **No transparency:** all generated images are flattened onto white without alpha channel, because print services reject transparent elements.
- **Self-print booklet:** `self_print_duplex()` in `build_all.py` (needs `pypdf`): page 1 = back | front cover, page 2 = blank, then nested sheets (40 | 1), (2 | 39), (38 | 3) …
- **Mirrored margins** are not supported by the `docx` library, so `build_all.py` adds them to the finished file.

## Editing

- **Texts:** edit `text_en.js`.
- **Layout:** the page functions in `build_workbook.js` (`pageSetup`, `sessionBlock`, `pageQR1` …).
- **Small one-off fixes:** edit the generated Word file and export it to PDF from Word (embed the fonts). These changes are lost when you rebuild.
