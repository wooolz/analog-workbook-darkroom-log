#! python3

"""Darkroom Log - cover generator.

Draws front and back covers as SVG (all graphics drawn from simple shapes) and renders
them to JPG at 350 dpi with rsvg-convert. Variants: A5/A4 x print (3 mm bleed) / online
(no bleed) x dark/light; backs per binding (ring 80 p., booklet 40 p., online 40 p.).

Requirements: Python 3 with Pillow, rsvg-convert (librsvg), font "Courier Prime" installed.
Run (from anywhere): python3 src/covers/make_covers.py   -> writes JPGs into <repo>/dist/covers
Design coordinates are in mm on an A5 page incl. 3 mm bleed (154 x 216); A4 is scaled.
"""
import math, subprocess, os, re
W, H = 154, 216          # A5 + 3 mm bleed each side
BG, TXT = '#0e0e0e', '#ececec'
G1, G2, G3, G4, G5 = '#1b1b1b', '#262626', '#303030', '#3c3c3c', '#4d4d4d'
FONT = "font-family='Courier Prime'"

def svg(body):
    return f"<svg xmlns='http://www.w3.org/2000/svg' width='{W}mm' height='{H}mm' viewBox='0 0 {W} {H}'>" \
           f"<rect width='{W}' height='{H}' fill='{BG}'/>{body}</svg>"

# ---------- TEXTS (per language) ----------
TEXT = {
    'en': dict(pre='', title='DARKROOM LOG', sub='A PRINTING WORKBOOK',
               para=["A workbook for the darkroom.", "",
                     "Log your printing sessions", "frame by frame, print by", "print: filtration, time,", "aperture, dodging and",
                     "burning. Keep your processes,", "papers and enlargers notes", "in one place.", "",
                     "With session and project", "indexes, pages for your own", "learning notes and quick", "references for exposure,",
                     "contrast and color.", "",
                     "For black &amp; white and RA-4", "color printing."],
               pages='pages', bind=dict(ring80='ring binding', booklet40='booklet', softcover80='softcover', hardcover80='hardcover',
                                        online='for screen &amp; self-print'),
               edition='Edition 1.0 · Nov 2026 · CC BY-NC-SA 4.0'),
    'de': dict(pre='LABORBUCH', title='DUNKELKAMMER', sub='EIN ARBEITSBUCH FÜR ABZÜGE',
               para=["Notizen deiner Sessions,", "Abzug für Abzug:", "Belichtungszeit, Filterung,", "Blende, Abwedeln und",
                     "Nachbelichten sowie", "Prozesse, Papiere und", "Vergrößerer – alles", "an einem Ort.", "",
                     "Dazu Session- und Projekt-", "verzeichnisse, Raum für", "eigene Anmerkungen sowie", "Referenztabellen zu",
                     "Belichtung, Kontrast", "und Farbe.", "",
                     "Für Schwarzweiß- und", "RA-4-Farbabzüge."],
               pages='Seiten', bind=dict(ring80='Ringbindung', booklet40='Heftbindung', softcover80='Taschenbuch', hardcover80='Hardcover',
                                         online='Digital &amp; Selbstdruck'),
               edition='Edition 1.0 · Nov. 2026 · CC BY-NC-SA 4.0'),
}
LANGS = ('en', 'de')

def spaced(t):   # letter-spaced text, right-aligned: no extra space after the last letter
    return f"{t[:-1]}<tspan letter-spacing='0'>{t[-1]}</tspan>"

def title(lang='en', x=139):
    T = TEXT[lang]
    pre = (f"<text x='{x}' y='26.5' {FONT} font-size='4.2' letter-spacing='0.9' text-anchor='end' fill='#bdbdbd'>{spaced(T['pre'])}</text>"
           if T['pre'] else '')
    return (pre + f"<text x='{x}' y='36' {FONT} font-weight='bold' font-size='10.5' letter-spacing='1.6' text-anchor='end' fill='{TXT}'>{spaced(T['title'])}</text>"
            f"<text x='{x}' y='44.5' {FONT} font-size='4.2' letter-spacing='0.9' text-anchor='end' fill='#bdbdbd'>{spaced(T['sub'])}</text>")

def head():
    p = []
    # lamp housing with vents
    p.append(f"<rect x='10' y='-15' width='58' height='15' rx='3' fill='{G3}' stroke='{G4}' stroke-width='0.4'/>")
    for i in range(8): p.append(f"<rect x='{15+i*6.4:.1f}' y='-11' width='3.6' height='7' rx='1' fill='{G2}'/>")
    # main color head body
    p.append(f"<rect x='0' y='0' width='96' height='40' rx='3' fill='{G4}' stroke='{G5}' stroke-width='0.5'/>")
    for i in range(5): p.append(f"<rect x='7' y='{8+i*5}' width='44' height='1.6' rx='0.8' fill='{G3}'/>")   # side vents
    # control panel right side: C/M/Y windows + knobs
    p.append(f"<rect x='58' y='4' width='33' height='32' rx='2' fill='{G3}'/>")
    for k, (lab, val) in enumerate([('C', '000'), ('M', '065'), ('Y', '040')]):
        y = 7.5 + k * 9.5
        p.append(f"<text x='61.2' y='{y+4.6}' {FONT} font-size='3.4' font-weight='bold' fill='#707070'>{lab}</text>")
        p.append(f"<rect x='65' y='{y}' width='13' height='6' rx='0.6' fill='#0b0b0b' stroke='{G5}' stroke-width='0.3'/>")
        p.append(f"<text x='71.5' y='{y+4.5}' {FONT} font-size='3.8' text-anchor='middle' fill='#8c8c8c'>{val}</text>")
        p.append(f"<circle cx='85' cy='{y+3}' r='3.3' fill='{G2}' stroke='{G5}' stroke-width='0.4'/>")
        p.append(f"<line x1='85' y1='{y+3}' x2='{85+2.4*math.sin(k*1.7+0.6):.2f}' y2='{y+3-2.4*math.cos(k*1.7+0.6):.2f}' stroke='#6a6a6a' stroke-width='0.6' stroke-linecap='round'/>")
    # mixing box, negative stage, carrier handle
    p.append(f"<rect x='14' y='40' width='56' height='13' fill='{G3}'/>")
    p.append(f"<rect x='3' y='53' width='78' height='6.5' fill='{G3}' stroke='{G4}' stroke-width='0.4'/>")
    p.append(f"<rect x='-10' y='54.6' width='14' height='3.2' rx='1.6' fill='{G4}'/>")
    # bellows
    for i in range(6):
        y0 = 59.5 + i * 4; w0 = 46 - i * 2; w1 = 46 - (i + 1) * 2
        p.append(f"<polygon points='{42-w0/2},{y0} {42+w0/2},{y0} {42+w1/2},{y0+4} {42-w1/2},{y0+4}' fill='{G3 if i%2 else G2}'/>")
    p.append(f"<rect x='20' y='83.5' width='44' height='4.5' fill='{G4}'/>")
    p.append(f"<rect x='33' y='88' width='18' height='11' fill='{G2}' stroke='{G5}' stroke-width='0.4'/>")
    for yy in (91, 94, 97): p.append(f"<line x1='33' y1='{yy}' x2='51' y2='{yy}' stroke='{G4}' stroke-width='0.4'/>")
    p.append(f"<rect x='31' y='99' width='22' height='4' fill='{G4}'/>")
    return ''.join(p)

CONE = "33,103 51,103 175,330 -95,330"
def bands(times, cols, lab=True, w=24, h=32.5):
    p = []
    for i, (t, c) in enumerate(zip(times, cols)):
        p.append(f"<rect x='{i*w}' y='0' width='{w}' height='{h}' fill='{c}'/>")
        if lab and t: p.append(f"<text x='{i*w+3}' y='{h-3.5}' {FONT} font-size='3.5' fill='#7a7a7a'>{t}s</text>")
    return ''.join(p)
def lighten(c, d):
    v = int(c[1:3], 16) + d; v = max(0, min(255, v)); return '#%02x%02x%02x' % (v, v, v)

DARK = ['#505050', '#464646', '#3d3d3d', '#343434', '#2c2c2c', '#252525', '#1e1e1e', '#191919']
TIMES = [2, 4, 6, 8, 10, 12, 14, 16]

def compose(head_tf, strip_tf, extra='', second=None, lang='en'):
    clip = f"<clipPath id='lit'><polygon points='{CONE}' transform='{head_tf}'/></clipPath>"
    cone = f"<polygon points='{CONE}' fill='#1b1b1b' transform='{head_tf}'/>"
    note = f"<text x='14' y='6.5' {FONT} font-size='3.2' fill='#717171'>gr 3 · f/8</text>"   # own grey value: brighter in the red theme
    s_dark = f"<g transform='{strip_tf}'>{bands(TIMES, DARK)}{note}</g>"
    s_lit = f"<g clip-path='url(#lit)'><g transform='{strip_tf}'>{bands(TIMES, [lighten(c, 22) for c in DARK])}{note}</g></g>"
    sec = ''
    if second:
        cols = ['#2e2e2e', '#2a2a2a', '#262626', '#222222', '#1f1f1f', '#1c1c1c', '#191919', '#161616']
        sec = (f"<g transform='{second}'>{bands(TIMES, cols, False)}</g>"
               f"<g clip-path='url(#lit)'><g transform='{second}'>{bands(TIMES, [lighten(c, 18) for c in cols], False)}</g></g>")
    return f"<defs>{clip}</defs>{cone}{extra}<!--S-->{sec}{s_dark}{s_lit}<!--/S--><g transform='{head_tf}'>{head()}</g>" + title(lang)


# ---------- FRONT COMPOSITIONS ----------
STRIP_ANGLE, STRIP_X0, STRIP_EXIT = 10, -8, 190.4      # strip: 10 deg, upper edge leaves right trim edge at y=190.4
_y0 = STRIP_EXIT - (151 - STRIP_X0) * math.tan(math.radians(STRIP_ANGLE))
HEAD_TF = "translate(93,76) rotate(17) scale(0.8)"     # enlarger head: position, 17 deg, scale
STRIP_TF = f"translate({STRIP_X0},{_y0:.1f}) rotate({STRIP_ANGLE}) translate(0,3.5)"
FRONT = {l: compose(HEAD_TF, STRIP_TF, lang=l) for l in LANGS}
# light theme: own test-strip palette, 8 s band kept at #999999, 16 steps apart
DARK_STRIP = DARK
DARK = ['#%02x%02x%02x' % ((153 + (3 - i) * 16,) * 3) for i in range(8)]
FRONT_LIGHT = {l: compose(HEAD_TF, STRIP_TF, lang=l) for l in LANGS}
DARK = DARK_STRIP

# ---------- BACK ----------
def back(pages, lang='en'):
    T = TEXT[lang]
    if T['pre']:   # two-part title: small pre-title, large title
        p = [f"<text x='77' y='24.5' {FONT} font-size='2.9' letter-spacing='0.5' text-anchor='middle' fill='#bdbdbd'>{T['pre']}</text>",
             f"<text x='77' y='31' {FONT} font-weight='bold' font-size='5.2' letter-spacing='0.8' text-anchor='middle' fill='{TXT}'>{T['title']}</text>",
             f"<text x='77' y='37' {FONT} font-size='2.9' letter-spacing='0.5' text-anchor='middle' fill='#bdbdbd'>{T['sub']}</text>"]
    else:
        p = [f"<text x='77' y='30' {FONT} font-weight='bold' font-size='5.2' letter-spacing='0.8' text-anchor='middle' fill='{TXT}'>{T['title']}</text>",
             f"<text x='77' y='36' {FONT} font-size='2.9' letter-spacing='0.5' text-anchor='middle' fill='#bdbdbd'>{T['sub']}</text>"]
    para = T['para']
    y = 48
    for l in para:
        if l: p.append(f"<text x='77' y='{y}' {FONT} font-size='3.2' text-anchor='middle' fill='#d6d6d6'>{l}</text>")
        y += 4.6
    # film strip, 35 mm to scale
    x0, yt, yb, end = 31, 146, 181, 170
    tongue, trans = 38, 16
    half = yt + 17
    d = (f"M {x0+3} {yb} L {end} {yb} L {end} {yt} L {x0+tongue+trans} {yt} "
         f"C {x0+tongue+6} {yt} {x0+tongue+8} {half} {x0+tongue} {half} L {x0+3} {half} "
         f"Q {x0} {half} {x0} {half+3} L {x0} {yb-3} Q {x0} {yb} {x0+3} {yb} Z")
    p.append(f"<path d='{d}' fill='#313131'/>")
    x = x0 + 4
    while x < end:
        p.append(f"<rect x='{x:.2f}' y='{yb-2.0-2.79:.2f}' width='1.98' height='2.79' rx='0.5' fill='{BG}'/>")
        if x > x0 + tongue + trans: p.append(f"<rect x='{x:.2f}' y='{yt+2.0:.2f}' width='1.98' height='2.79' rx='0.5' fill='{BG}'/>")
        x += 4.75
    fx, n = x0 + tongue + trans + 14, 1
    while fx < end:
        p.append(f"<rect x='{fx:.2f}' y='{yt+5.5}' width='36' height='24' rx='1.2' fill='{G2}' stroke='{G4}' stroke-width='0.4'/>")
        p.append(f"<text x='{fx+2}' y='{yb-0.55}' {FONT} font-size='1.5' fill='#666'>{n}</text>")
        p.append(f"<text x='{fx+21}' y='{yb-0.55}' {FONT} font-size='1.5' fill='#666'>{n}A</text>")
        fx += 38; n += 1
    p.append(f"<text x='77' y='196' {FONT} font-size='2.8' text-anchor='middle' fill='#9a9a9a'>{pages}</text>")
    p.append(f"<text x='77' y='201' {FONT} font-size='2.4' text-anchor='middle' fill='#7a7a7a'>{T['edition']}</text>")
    return ''.join(p)



def invert(seg):
    def rep(m):
        h = m.group(1)
        if len(h) == 3: h = ''.join(c*2 for c in h)
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        if not (r == g == b): return m.group(0)
        if r == 0x0e: return '#ffffff'            # background -> pure white
        if r == 0x31: return '#a0a0a0'            # film strip body on the back: a bit darker in light theme
        v = max(0, 255 - r - 18); return '#%02x%02x%02x' % (v, v, v)
    return re.sub(r"#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})\b", rep, seg)
def light(body):
    parts = re.split(r'(<!--S-->.*?<!--/S-->)', body, flags=re.S)
    def half(seg):   # test strip in light theme: half the darkness, labels unchanged
        def rep(m):
            v = int(m.group(1)[0:2], 16); v = (255 + v) // 2
            return "<rect" + m.group(0)[5:].replace(m.group(1), '%02x%02x%02x' % (v, v, v))
        return re.sub(r"<rect[^>]*fill='#([0-9a-fA-F]{6})'", rep, seg)
    return ''.join(p if p.startswith('<!--S-->') else invert(p) for p in parts)

def red(body):
    """Safelight theme: dark background, every grey becomes a safelight red of similar lightness."""
    def rep(m):
        h = m.group(1)
        if len(h) == 3: h = ''.join(c*2 for c in h)
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        if not (r == g == b) or r == 0x0e: return m.group(0)       # keep background
        if r == 0x71: return '#d02c1e'                             # pencil note on the test strip: brighter
        return '#%02x%02x%02x' % (min(255, round(r * 1.3)), round(r * 0.2), round(r * 0.14))
    return re.sub(r"#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})\b", rep, body)

def page(body, theme, size, bleed):
    s = 1.0 if size == 'A5' else 210 / 148
    bl = 3 if bleed else 0
    Wt, Ht = (148, 210) if size == 'A5' else (210, 297)
    W, H = Wt + 2 * bl, Ht + 2 * bl
    bg = '#ffffff' if theme == 'light' else '#0e0e0e'
    if theme == 'light': body = light(body)
    if theme == 'red': body = red(body)
    return (f"<svg xmlns='http://www.w3.org/2000/svg' width='{W}mm' height='{H}mm' viewBox='0 0 {W} {H}'>"
            f"<rect width='{W}' height='{H}' fill='{bg}'/><g transform='translate({bl},{bl}) scale({s:.6f}) translate(-3,-3)'>{body}</g></svg>")

# one-piece covers (epubli "Cover als Gesamtdatei"): back | spine | front in one file.
# Values from epubli's cover calculator: page size, bleed (Beschnitt) and spine (Rücken) in mm,
# n = page count for the line at the bottom of the back.
SPREADS = {
    ('A5', 'booklet40'):   dict(w=148, h=210, bleed=3,  spine=0,  n=40),
    ('A5', 'softcover80'): dict(w=148, h=210, bleed=3,  spine=5,  n=80),
    ('A5', 'hardcover80'): dict(w=157, h=228, bleed=11, spine=9,  n=80),
    ('A5', 'ring80'):      dict(w=148, h=210, bleed=3,  spine=10, n=80),
    ('A4', 'booklet40'):   dict(w=210, h=297, bleed=3,  spine=0,  n=40),
    ('A4', 'softcover80'): dict(w=210, h=297, bleed=3,  spine=5,  n=80),
    ('A4', 'hardcover80'): dict(w=219, h=315, bleed=11, spine=9,  n=80),
    ('A4', 'ring80'):      dict(w=210, h=297, bleed=3,  spine=10, n=80),
}

def spread(front_body, back_body, theme, f):
    Wt, Ht, bl, sp = f['w'], f['h'], f['bleed'], f['spine']
    s = Wt / 148                                    # design is made for a 148 mm wide page
    dy = (Ht - 210 * s) / 2                         # centre vertically if the page is relatively taller
    W, H = 2 * Wt + sp + 2 * bl, Ht + 2 * bl
    bg = '#ffffff' if theme == 'light' else '#0e0e0e'
    if theme == 'light': front_body, back_body = light(front_body), light(back_body)
    if theme == 'red': front_body, back_body = red(front_body), red(back_body)
    xb, xf = bl + Wt, bl + Wt + sp                  # spine from xb to xf
    return (f"<svg xmlns='http://www.w3.org/2000/svg' width='{W}mm' height='{H}mm' viewBox='0 0 {W} {H}'>"
            f"<defs><clipPath id='cb'><rect x='0' y='0' width='{xb}' height='{H}'/></clipPath>"
            f"<clipPath id='cf'><rect x='{xf}' y='0' width='{W - xf}' height='{H}'/></clipPath></defs>"
            f"<rect width='{W}' height='{H}' fill='{bg}'/>"
            f"<g clip-path='url(#cb)'><g transform='translate({bl},{bl + dy:.3f}) scale({s:.6f}) translate(-3,-3)'>{back_body}</g></g>"
            f"<g clip-path='url(#cf)'><g transform='translate({xf},{bl + dy:.3f}) scale({s:.6f}) translate(-3,-3)'>{front_body}</g></g></svg>")

# single front/back files: print (ring binding, 3 mm bleed) and online (no bleed)
backs = {'print': {'ring80': (80, 'ring80')}, 'online': {'': (40, 'online')}}
def pages_line(n, size, binding, lang):
    T = TEXT[lang]; return f"{n} {T['pages']} · {size} · {T['bind'][binding]}"

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'dist', 'covers')
os.makedirs(OUT, exist_ok=True)
# optional filter: only build covers whose name contains all given words, e.g.  make_covers.py A5_print_red
import sys
FILTER = sys.argv[1:]
out = []
for lang in LANGS:
  pre = '' if lang == 'en' else f'{lang}_'         # English names unchanged, other languages: cover_de_…
  for theme in ('dark', 'light', 'red'):
    front = FRONT_LIGHT[lang] if theme == 'light' else FRONT[lang]
    for size in ('A5', 'A4'):
        for kind, bleed in (('print', True), ('online', False)):
            base = f"cover_{pre}{size}_{kind}_{theme}"
            jobs = [(f"{base}_front", front)]
            for tag, (n, bk) in backs[kind].items():
                jobs.append((f"{base}_back" + (f"_{tag}" if tag else ''), back(pages_line(n, size, bk, lang), lang)))
            for name, body in jobs:
                if not all(f in name for f in FILTER): continue
                svgf = os.path.join(OUT, f"{name}.svg"); png = os.path.join(OUT, f"{name}.png")
                open(svgf, 'w').write(page(body, theme, size, bleed))
                subprocess.run(['rsvg-convert', '-d', '350', '-p', '350', '-f', 'png', '-o', png, svgf], check=True)
                out.append(name)
        # one-piece covers (back | spine | front) for every binding
        for (sz, binding), fmt in SPREADS.items():
            if sz != size: continue
            name = f"cover_{pre}{size}_{binding}_{theme}_spread"
            if not all(f in name for f in FILTER): continue
            svgf = os.path.join(OUT, f"{name}.svg"); png = os.path.join(OUT, f"{name}.png")
            open(svgf, 'w').write(spread(front, back(pages_line(fmt['n'], size, binding, lang), lang), theme, fmt))
            subprocess.run(['rsvg-convert', '-d', '350', '-p', '350', '-f', 'png', '-o', png, svgf], check=True)
            out.append(name)
from PIL import Image
for n in out:
    p = os.path.join(OUT, n); im = Image.open(p + ".png").convert('RGB'); im.save(p + ".jpg", quality=95, dpi=(350, 350)); os.remove(p + ".png"); os.remove(p + ".svg")
print(len(out)); print('\n'.join(out))
