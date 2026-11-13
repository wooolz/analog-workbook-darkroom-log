#!/usr/bin/env python3
# SPDX-License-Identifier: AGPL-3.0-or-later
"""Quick previews of single page types in A5 and A4 (ring binding geometry), as PNG.

Run from the repository root:
  python3 src/workbook/preview.py title session            -> dist/preview/<kind>_A5.png, _A4.png, sheet.png
Page kinds: title howto1 howto2 setup papers projidx sessidx learn session qr1 qr2 qr3 notes
Option --binding booklet|online (default ring), --dpi N (default 110), --lang en|de (default en)
"""
import json, os, subprocess, sys, shlex, glob
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_all as B

args = sys.argv[1:]
binding, dpi, lang = 'ring', 110, 'en'
if '--binding' in args: i = args.index('--binding'); binding = args[i + 1]; del args[i:i + 2]
if '--lang' in args: i = args.index('--lang'); lang = args[i + 1]; del args[i:i + 2]
if '--dpi' in args: i = args.index('--dpi'); dpi = int(args[i + 1]); del args[i:i + 2]
kinds = args or ['title']
OUT = os.path.join(B.ROOT, 'dist', 'preview')
os.makedirs(OUT, exist_ok=True)
for f in glob.glob(os.path.join(OUT, '*')): os.remove(f)
B.torn_edges()
soffice = shlex.split(os.environ.get('SOFFICE_CMD', 'soffice'))
order = ['title', 'howto1', 'howto2', 'setup', 'papers', 'projidx', 'sessidx', 'learn', 'session', 'qr1', 'qr2', 'qr3', 'notes']
pngs = {}
for size in ('A5', 'A4'):
    docx = os.path.join(OUT, f'preview_{size}.docx')
    conf = dict(lang=lang, size=size, binding=binding, kind='workbook', only=kinds, out=docx, assets=B.ASSETS, covers=B.COVERS)
    subprocess.run(['node', os.path.join(B.HERE, 'build_workbook.js'), json.dumps(conf)], check=True, stdout=subprocess.DEVNULL)
    if binding == 'ring': B.mirror_margins(docx)
    subprocess.run(soffice + ['--headless', '--convert-to', 'pdf', '--outdir', OUT, docx], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = docx[:-5] + '.pdf'
    built = [k for k in order if k in kinds]
    n = B.pdf_pages(pdf) - (2 if binding == 'online' else 0)
    off = 1 if binding == 'online' else 0
    for i, k in enumerate(built[:n]):
        base = os.path.join(OUT, f'{k}_{size}')
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), '-singlefile', '-f', str(i + 1 + off), '-l', str(i + 1 + off), pdf, base], check=True, stderr=subprocess.DEVNULL)
        pngs.setdefault(k, []).append(base + '.png')
# one sheet: rows = kinds, A5 next to A4 (same height)
rows = []
for k in [k for k in order if k in pngs]:
    ims = [Image.open(p) for p in pngs[k]]
    h = 900; ims = [im.resize((int(im.size[0] * h / im.size[1]), h)) for im in ims]
    row = Image.new('RGB', (sum(i.size[0] for i in ims) + 30, h + 20), '#888'); x = 10
    for im in ims: row.paste(im, (x, 10)); x += im.size[0] + 10
    rows.append(row)
W = max(r.size[0] for r in rows); sheet = Image.new('RGB', (W, sum(r.size[1] for r in rows)), '#888'); y = 0
for r in rows: sheet.paste(r, (0, y)); y += r.size[1]
sheet.save(os.path.join(OUT, 'sheet.png')); print(os.path.join(OUT, 'sheet.png'))
