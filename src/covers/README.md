# Darkroom Log – Cover Generator

`make_covers.py` draws the front and back covers of the *Darkroom Log* workbook. Every graphic is built from simple shapes in SVG, with no external images. The script then renders each cover to a print-ready JPG at 350 dpi.

## Output

48 JPG files in `dist/covers/` (relative to the repository root), each in a **dark**, a **light** and a **red** (safelight) theme.

**Single front and back files** (24): for ring binding (print, 3 mm bleed) and for the online versions (no bleed).

| | Print (3 mm bleed) | Online / self-print (no bleed) |
|---|---|---|
| **A5** | 154 × 216 mm (2123 × 2977 px) | 148 × 210 mm (2039 × 2894 px) |
| **A4** | 216 × 303 mm (2977 × 4176 px) | 210 × 297 mm (2894 × 4093 px) |

`cover_<A5|A4>_<print|online>_<theme>_front.jpg`, `…_print_<theme>_back_ring80.jpg`, `…_online_<theme>_back.jpg`

**One-piece covers** (24): back · spine · front in one file, as epubli's "Cover als Gesamtdatei". Sizes from epubli's cover calculator, set in `SPREADS`:

| Binding | A5: page · bleed · spine → total | A4: page · bleed · spine → total |
|---|---|---|
| booklet, 40 pages | 148 × 210 · 3 · 0 → 302 × 216 mm | 210 × 297 · 3 · 0 → 426 × 303 mm |
| softcover, 80 pages | 148 × 210 · 3 · 5 → 307 × 216 mm | 210 × 297 · 3 · 5 → 431 × 303 mm |
| hardcover, 80 pages | 157 × 228 · 11 · 9 → 345 × 250 mm | 219 × 315 · 11 · 9 → 469 × 337 mm |
| ring binding, 80 pages | 148 × 210 · 3 · 10 → 312 × 216 mm | 210 × 297 · 3 · 10 → 436 × 303 mm |

`cover_<A5|A4>_<booklet40|softcover80|hardcover80|ring80>_<theme>_spread.jpg`

The design is scaled to the page width and centred vertically when the page is relatively taller (hardcover). The back stays free at the bottom left (no ISBN barcode). If epubli shows other values (different page count or paper), change them in `SPREADS`.

The files follow epubli's requirements: JPG, at least 300 dpi, no transparency, no crop marks, and the background running all the way into the bleed.

## Requirements

1. **Python 3** with **Pillow** (`pip install -r requirements.txt`)
2. **rsvg-convert** (part of librsvg)
   - macOS: `brew install librsvg`
   - Debian/Ubuntu: `sudo apt install librsvg2-bin`
   - Windows: easiest via WSL (Ubuntu), then as above
3. **Font "Courier Prime"**, Regular and Bold, installed on the system (see `fonts/README.md`). If the font is missing, rsvg-convert silently substitutes another font.

## Usage

From the repository root:

```
pip install -r requirements.txt
python3 src/covers/make_covers.py
```

The JPGs are written to `dist/covers/`. Existing files with the same name are overwritten.

To build only some covers, pass words that must appear in the file name, e.g. `python3 src/covers/make_covers.py A5_print_red`. The `dist/` folder is not tracked by git; publish the files via a GitHub Release.

## Adjusting the design

All coordinates are in **millimetres on an A5 page including 3 mm bleed** (154 × 216 mm). The trimmed page runs from x = 3 to 151 and from y = 3 to 213. The A4 versions are the same design scaled by 210/148.

| What | Where in the script |
|---|---|
| Position, angle (17°) and size of the enlarger head | `HEAD_TF` |
| Test strip angle (10°) | `STRIP_ANGLE` |
| Height where the strip's upper edge leaves the right trim edge | `STRIP_EXIT` |
| Strip width (32.5 mm) | `h=` in `bands()` |
| Strip shift from the top (3.5 mm) | `translate(0,3.5)` in `STRIP_TF` |
| Strip tones, dark theme | `DARK` (first definition) |
| Strip tones, light theme (8 s band at #999999, 16 steps apart) | the second `DARK = [...]` line before `FRONT_LIGHT` |
| C/M/Y values shown on the color head | `head()`: `('C', '000'), ('M', '065'), ('Y', '040')` |
| Title and subtitle on the front | `title()` |
| Back text (centered block) | `para = [...]` in `back()`: one list entry per line, `""` for an empty line, `&amp;` instead of `&` |
| Line at the bottom of the back (pages · size · binding) | `backs = {...}` |
| Edition and license line | in `back()` |
| Background and text colors, dark theme | `BG`, `TXT`, and the grey steps `G1`–`G5` |

### How the light theme is made

The light theme is generated from the dark one:

- **Greys:** every grey is inverted and then made 18 steps darker.
- **Background:** pure white.
- **Film strip on the back:** fixed at `#a0a0a0`.
- **Test strip:** not inverted. It uses its own tone scale (`FRONT_LIGHT`).

To change the light theme, edit `invert()` and `light()`.

### How the red (safelight) theme is made

The red theme keeps the dark background and turns every grey into a safelight red of similar lightness: red = grey × 1.3, green = grey × 0.2, blue = grey × 0.14 (function `red()`). The pencil note on the test strip ("gr 3 · f/8", grey `#717171`) is set brighter in this theme.

## Variants and back texts

The line at the bottom of the back names pages, size and binding: for the single files it is set in `backs`, for the one-piece covers in `SPREADS` (`pages`).

## License

- **Script** (`make_covers.py`): AGPL-3.0-or-later
- **Cover design and generated images:** CC BY-NC-SA 4.0, with the additional permission for workshops and courses

See `LICENSE` in the repository root.
