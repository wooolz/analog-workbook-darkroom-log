# Fonts

The workbook and covers use two free fonts. Both are licensed under the **SIL Open Font License 1.1** by their authors. They are not part of this project's license.

| Font | Used for | Download |
|---|---|---|
| Courier Prime (Regular, Bold) | all layout text, covers | https://fonts.google.com/specimen/Courier+Prime |
| Caveat (Regular) | example entries ("handwriting") | https://fonts.google.com/specimen/Caveat |

## Setup

1. Download both font families.
2. Install them on your system (Windows: right-click the file → Install; macOS: double-click → Install Font; Linux: copy to `~/.fonts` and run `fc-cache -f`).
3. Optionally, put the `.ttf` files and their `OFL.txt` in this folder.

By default, `.ttf`/`.otf` files in this folder are ignored by git (see `.gitignore`). The OFL allows redistributing the fonts. If you want to commit them, remove those lines from `.gitignore` and include each font's `OFL.txt` next to it.

## Word → PDF

When exporting PDFs from the Word files, make sure the fonts are embedded (File → Options → Save → "Embed fonts in the file", with both sub-options unticked).
