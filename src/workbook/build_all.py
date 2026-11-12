#!/usr/bin/env python3
# SPDX-License-Identifier: AGPL-3.0-or-later
"""Build all Darkroom Log workbook variants (Word + PDF).

Requires: Node.js with the "docx" package (npm install), Python 3 with Pillow,
LibreOffice (soffice) for the PDF export, fonts Courier Prime + Caveat installed,
and the online covers in dist/covers (python3 src/covers/make_covers.py).

Run from the repository root:  python3 src/workbook/build_all.py [filter]
  optional filter: only build variants whose name contains this text, e.g. "A5_ring"
Environment: SOFFICE_CMD overrides the LibreOffice command (default: "soffice").
"""
import json, os, random, re, shlex, subprocess, sys, zipfile, shutil
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(ROOT, 'dist', 'workbook')
ASSETS = os.path.join(ROOT, 'dist', 'tmp')
COVERS = os.path.join(ROOT, 'dist', 'covers')

VARIANTS = [  # name, config
    ('darkroom-log_en_A5_ring-80', dict(size='A5', binding='ring')),
    ('darkroom-log_en_A4_ring-80', dict(size='A4', binding='ring')),
    ('darkroom-log_en_A5_booklet-40', dict(size='A5', binding='booklet')),
    ('darkroom-log_en_A4_booklet-40', dict(size='A4', binding='booklet')),
    *[(f'darkroom-log_en_{sz}_online-40_{th}', dict(size=sz, binding='online', cover=th))     # online: one file per cover theme
      for sz in ('A5', 'A4') for th in ('light', 'dark', 'red')],
    ('darkroom-log_en_A4_howto-extended', dict(size='A4', binding='online', kind='howto')),
]
EXPECTED_PAGES = {'ring': 80, 'booklet': 40, 'online': 42}


def torn_edges():
    """Torn-paper edge strips for the how-to examples (PNG, 600 dpi, white background)."""
    os.makedirs(ASSETS, exist_ok=True)
    px = 600 / 25.4
    for size, width, S in (('A5', 124, 1.0), ('A4', 186, 1.25)):
        h = 2.4 * S
        for top in (True, False):
            rnd = random.Random(7 if top else 11)
            W, H = int(width * px), int(h * px)
            im = Image.new('RGB', (W, H), 'white'); d = ImageDraw.Draw(im)
            pts, x = [], 0
            while x <= W:
                y = H * (0.45 + rnd.uniform(-0.28, 0.28)); pts.append((x, y)); x += rnd.uniform(0.6, 1.8) * px
            pts.append((W, H * 0.45))
            shade = [(x, y + (2 if top else -2)) for x, y in pts]
            d.line(shade, fill=(236, 236, 236), width=int(0.5 * px))
            d.line(pts, fill=(140, 140, 140), width=max(2, int(0.18 * px)))
            fw = max(2, int(0.21 * px))  # outer frame lines continue to the torn edge
            for fx in (fw // 2, W - fw // 2 - 1):
                ys = [y for x, y in pts if abs(x - fx) < 1.5 * px] or [H * 0.45]
                y0 = sum(ys) / len(ys)
                d.line([(fx, y0), (fx, H if top else 0)], fill=(110, 110, 110), width=fw)
            im.save(os.path.join(ASSETS, f'torn_{size}_{width}_{"top" if top else "bottom"}.png'), dpi=(600, 600))


def mirror_margins(docx):
    """Add <w:mirrorMargins/> (inner/outer margins) – docx-js has no option for it."""
    tmp = docx + '.tmp'
    with zipfile.ZipFile(docx) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == 'word/settings.xml':
                s = data.decode('utf-8')
                s = s.replace('<w:displayBackgroundShape/>', '<w:displayBackgroundShape/><w:mirrorMargins/>', 1) \
                    if '<w:displayBackgroundShape/>' in s else re.sub(r'(<w:settings[^>]*>)', r'\1<w:mirrorMargins/>', s, 1)
                data = s.encode('utf-8')
            zo.writestr(item, data)
    shutil.move(tmp, docx)


def pdf_pages(pdf):
    out = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
    m = re.search(r'Pages:\s+(\d+)', out)
    return int(m.group(1)) if m else -1


def main():
    flt = sys.argv[1] if len(sys.argv) > 1 else ''
    os.makedirs(OUT, exist_ok=True)
    torn_edges()
    soffice = shlex.split(os.environ.get('SOFFICE_CMD', 'soffice'))
    for name, conf in VARIANTS:
        if flt not in name:
            continue
        docx = os.path.join(OUT, name + '.docx')
        conf = dict(conf, kind=conf.get('kind', 'workbook'), out=docx, assets=ASSETS, covers=COVERS)
        subprocess.run(['node', os.path.join(HERE, 'build_workbook.js'), json.dumps(conf)], check=True)
        if conf['binding'] == 'ring':
            mirror_margins(docx)
        subprocess.run(soffice + ['--headless', '--convert-to', 'pdf', '--outdir', OUT, docx], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        n = pdf_pages(os.path.join(OUT, name + '.pdf'))
        exp = 2 if conf['kind'] == 'howto' else EXPECTED_PAGES[conf['binding']]
        print(f'  -> {name}.pdf: {n} pages' + ('' if n == exp else f'   !! expected {exp}'))


if __name__ == '__main__':
    main()
