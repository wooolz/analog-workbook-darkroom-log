// SPDX-License-Identifier: AGPL-3.0-or-later
// Darkroom Log – workbook interior generator (Word .docx).
// Usage: node build_workbook.js '<json config>'
//   config: { size: "A5"|"A4", binding: "ring"|"booklet"|"online", freeCol: "INT" (label of the free column, max. 6 chars),
//             printer: "open"|"epubli",
//             cover: "light"|"dark"|"red" (online only), kind: "workbook"|"howto"|"samples", out: "<file.docx>", assets: "<dir with torn_*.png>",
//             covers: "<dir with cover JPGs>" }
'use strict';
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  HeightRule, VerticalAlign, AlignmentType, Footer, PageNumber, TableLayoutType, LineRuleType, Tab,
  TabStopType, LeaderType, ImageRun, Bookmark, InternalHyperlink, ExternalHyperlink, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
} = require('docx');
const T = require('./text_en.js');

const cfg = JSON.parse(process.argv[2]);
const A4 = cfg.size === 'A4';
const ONLINE = cfg.binding === 'online';
const RING = cfg.binding === 'ring';

// ---------------------------------------------------------------- geometry
const TRIM = A4 ? [210, 297] : [148, 210];
const BLEED = ONLINE ? 0 : 3;
const PW = TRIM[0] + 2 * BLEED, PH = TRIM[1] + 2 * BLEED;
const M = ONLINE ? { top: 12, bottom: 12, left: 12, right: 12, footer: 6 }
  : RING ? { top: 10, bottom: 10, left: 20, right: 10, footer: 6.5 }      // left = inner (mirrored)
    : { top: 10, bottom: 10, left: 15, right: 15, footer: 6.5 };
const CW = PW - M.left - M.right;            // 124 (A5) / 186 (A4)
const AVAIL = PH - M.top - M.bottom;
const K = CW / 124;                          // width factor vs. A5 logical widths
const S0 = A4 ? 1.25 : 1;                    // type / fixed-height factor
let S = S0;                                  // temporarily reduced by fitPage()
const SAFE = 2.5;

const mm = v => Math.round(v * 56.6929);
const hp = v => Math.round(v * S);           // half-points scaled
const FONT = 'Courier Prime', HAND = 'Caveat';
const C = { line: '9A9A9A', frame: '6E6E6E', label: '555555', shade: 'EEEEEE', text: '333333', hint: '888888',
  faint: 'D2D2D2', hand: '5E5E5E', imprint: '777777' };
const thin = { style: BorderStyle.SINGLE, size: 2, color: C.line };
const frame = { style: BorderStyle.SINGLE, size: 6, color: C.frame };
const dotted = { style: BorderStyle.DOTTED, size: 6, color: 'A8A8A8' };   // 0.75 pt: renders as clear square dots
const faint = { style: BorderStyle.DOTTED, size: 6, color: 'D6D6D6' };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

// ---------------------------------------------------------------- text helpers
const run = (text, o = {}) => new TextRun({ text, font: o.font || FONT, size: o.size || hp(12), color: o.color || C.label,
  bold: o.bold, characterSpacing: o.spacing });
const exact = (h, after = 0, before = 0) => ({ before: mm(before), after: mm(after), line: mm(h), lineRule: LineRuleType.EXACT });
function charsPerLine(sizeHp, widthMm, font) {
  const em = (sizeHp / 2) * 0.35278;                     // mm
  const adv = font === HAND ? 0.44 : 0.6;
  return Math.floor(widthMm * 0.97 / (em * adv));
}
function wrapCount(text, sizeHp, widthMm, font) {
  const n = charsPerLine(sizeHp, widthMm, font);
  let lines = 1, cur = 0;
  for (const w of text.split(' ')) {
    const l = w.length;
    if (cur === 0) cur = l; else if (cur + 1 + l <= n) cur += 1 + l; else { lines++; cur = l; }
  }
  return lines;
}
// paragraph with exact line height; returns {el, h}
function textP(text, o = {}) {
  const size = o.size || hp(14), lh = o.lh || 4.8 * S, w = (o.width || CW) - (o.indent || 0) - (o.hang || 0);
  const runs = Array.isArray(text) ? text : [run(text, { size, color: o.color || C.text, bold: o.bold, font: o.font })];
  const plain = Array.isArray(text) ? (o.plain || '') : text;
  const lines = o.lines || wrapCount(plain, size, w, o.font);
  const ind = o.hang ? { left: mm((o.indent || 0) + o.hang), hanging: mm(o.hang) } : o.indent ? { left: mm(o.indent) } : undefined;
  return { el: new Paragraph({ alignment: o.align, pageBreakBefore: o.pb, indent: ind,
    spacing: exact(lh, o.after || 0, o.before || 0), keepLines: true, children: runs }),
  h: lines * lh + (o.after || 0) + (o.before || 0) };
}
const spacer = (h, pb = false) => ({ el: new Paragraph({ pageBreakBefore: pb, spacing: exact(Math.max(h, 0.3)), children: [new TextRun({ text: '', size: 2 })] }), h: Math.max(h, 0.3) });
function title(t, sub, o = {}) {
  const kids = [run(t, { size: hp(22), bold: true, color: C.text })];
  if (sub) kids.push(run('   ' + sub, { size: hp(12) }));
  const ch = o.bookmark ? [new Bookmark({ id: o.bookmark, children: kids })] : kids;
  return { el: new Paragraph({ pageBreakBefore: o.pb, spacing: exact(5 * S, 2.5 * S), keepNext: true, children: ch }), h: 7.5 * S };
}
// title with subtitle on its own line (for long titles)
function title2(t, sub, o = {}) {
  const kids = [run(t, { size: hp(22), bold: true, color: C.text })];
  return [{ el: new Paragraph({ pageBreakBefore: o.pb, spacing: exact(5 * S, 0.5 * S), children: o.bookmark ? [new Bookmark({ id: o.bookmark, children: kids })] : kids }), h: 5.5 * S },
    { el: new Paragraph({ spacing: exact(4 * S, 1 * S), children: [run(sub, { size: hp(12) })] }), h: 5 * S }];
}

// ---------------------------------------------------------------- table helpers
// rows: [{h, cells:[{w, x?, text, ...}]}]; builds a union grid so cells with different widths line up
function grid(rows, opts = {}) {
  const edges = new Set([0]);
  rows.forEach(r => { let x = 0; r.cells.forEach(c => { if (c.x !== undefined) x = c.x; c._x0 = x; x += c.w; c._x1 = x; edges.add(+x.toFixed(2)); edges.add(+c._x0.toFixed(2)); }); });
  const E = [...edges].sort((a, b) => a - b);
  const colW = E.slice(1).map((e, i) => e - E[i]);
  const idx = v => E.findIndex(e => Math.abs(e - v) < 0.011);
  const trs = rows.map(r => new TableRow({ cantSplit: true, height: { value: mm(r.h), rule: HeightRule.EXACT },
    children: r.cells.map(c => {
      const span = idx(+c._x1.toFixed(2)) - idx(+c._x0.toFixed(2));
      return new TableCell({
        columnSpan: span > 1 ? span : undefined, rowSpan: c.rowSpan,
        width: { size: mm(c.w), type: WidthType.DXA },
        verticalAlign: c.valign || VerticalAlign.TOP,
        shading: c.fill ? { type: ShadingType.CLEAR, fill: c.fill, color: 'auto' } : c.shade ? { type: ShadingType.CLEAR, fill: C.shade, color: 'auto' } : undefined,
        margins: c.margins || { top: mm(0.5 * S), bottom: 0, left: mm(0.8 * S), right: mm(0.5 * S) },
        borders: c.borders,
        children: c.paras || [new Paragraph({ alignment: c.align, spacing: c.lh ? exact(c.lh) : { before: 0, after: 0 },
          children: c.runs || (c.text ? [run(c.text, { size: c.size || hp(12), color: c.color || C.label, font: c.font, bold: c.bold })] : []) })],
      });
    }) }));
  const table = new Table({ width: { size: mm(sum(colW)), type: WidthType.DXA }, columnWidths: colW.map(mm), layout: TableLayoutType.FIXED,
    borders: opts.borders || { top: frame, bottom: frame, left: frame, right: frame, insideHorizontal: thin, insideVertical: thin }, rows: trs });
  return { el: table, h: sum(rows.map(r => r.h)) };
}
const sum = a => a.reduce((x, y) => x + y, 0);
const L = v => v * K;   // logical A5 width -> real

// list table: header + n rows (dotted verticals, solid outer frame)
function listTable(cols, n, rowH, o = {}) {
  const headH = o.headH || 6 * S;
  const rows = [{ h: headH, cells: cols.map(([t, w], i) => ({ w: L(w), text: t, shade: true, valign: VerticalAlign.CENTER,
    align: (o.headCenter || []).includes(i) ? AlignmentType.CENTER : undefined })) }];
  for (let r = 0; r < n; r++) rows.push({ h: rowH, cells: cols.map(([, w], i) => ({ w: L(w),
    text: o.fill ? o.fill(r, i) : undefined, size: o.cellSize, color: o.cellColor, fill: o.fillColor ? o.fillColor(r, i) : undefined,
    align: (o.center || []).includes(i) ? AlignmentType.CENTER : undefined, valign: VerticalAlign.CENTER,
    borders: { top: thin, bottom: r === n - 1 ? frame : thin, left: i === 0 ? frame : dotted, right: i === cols.length - 1 ? frame : dotted } })) });
  return grid(rows);
}
function dottedLines(n, h = 7 * S) {
  const rows = [];
  for (let i = 0; i < n; i++) rows.push({ h, cells: [{ w: CW, borders: { top: none, left: none, right: none, bottom: dotted } }] });
  return grid(rows, { borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: dotted, insideVertical: none } });
}

// ---------------------------------------------------------------- page builder
const pages = [];          // each: {items:[{el,h}], used}
let cur = null;
function newPage() { cur = { items: [], used: 0 }; pages.push(cur); if (pages.length > 1) add(spacer(0.4, true)); return cur; }
function add(x) { if (Array.isArray(x)) x.forEach(add); else { cur.items.push(x); cur.used += x.h; } }
let FITM = 0;                                 // extra reserve while fitting flowing pages
const remaining = () => AVAIL - cur.used - SAFE - FITM;
// build a page with shrinking scale until it fits (how-to pages with flowing content)
function fitPage(builder, maxF = 1) {
  FITM = 3;                                  // 3 mm reserve for borders / rounding in Word
  const steps = []; for (let f = maxF; f > 0.69; f -= 0.02) steps.push(+f.toFixed(2));
  for (let k = 0; k < steps.length; k++) {
    const n = pages.length; S = S0 * steps[k];
    builder();
    const over = pages.slice(n).some(p => p.used > AVAIL - FITM + 0.01);
    if (!over || k === steps.length - 1) {
      if (steps[k] !== 1) console.log(`  fitted page ${n + 1} at ${(steps[k] * 100).toFixed(0)} %`);
      S = S0; FITM = 0; return;
    }
    pages.length = n; cur = pages[n - 1];
  }
}
function fillLines(h = 7 * S, after = 0) { const n = Math.floor((remaining() - after) / h); if (n > 0) add(dottedLines(n, h)); }

// ---------------------------------------------------------------- page plan & TOC
const plan = RING
  ? [['title'], ['howto1'], ['howto2'], ['setup'], ['papers'], ['projidx'], ...Array(5).fill(['sessidx']), ...Array(3).fill(['learn']),
    ...Array(60).fill(['session']), ['qr1'], ['qr2'], ['qr3'], ...Array(3).fill(['notes'])]
  : [['title'], ['howto1'], ['howto2'], ['setup'], ['papers'], ['projidx'], ...Array(2).fill(['sessidx']), ...Array(3).fill(['learn']),
    ...Array(25).fill(['session']), ['qr1'], ['qr2'], ['qr3'], ['notes']];
const firstPage = k => plan.findIndex(p => p[0] === k) + 1;
const toc = [
  [T.toc.howto, RING ? 2 : 2, 'howto', 0], [T.toc.processes, firstPage('setup'), 'setup', 0], [T.toc.papers, firstPage('papers'), 'papers', 0],
  [T.toc.projidx, firstPage('projidx'), 'projidx', 0], [T.toc.sessidx, firstPage('sessidx'), 'sessidx', 0],
  [T.toc.learn, firstPage('learn'), 'learn', 0], [T.toc.sessions, firstPage('session'), 'session', 0],
  [T.toc.qr, firstPage('qr1'), 'qr1', 0], [T.toc.qr1, firstPage('qr1'), 'qr1', 1], [T.toc.qr2, firstPage('qr2'), 'qr2', 1], [T.toc.qr3, firstPage('qr3'), 'qr3', 1],
  [T.toc.notes, firstPage('notes'), 'notes', 0],
];
const bm = id => (ONLINE ? id : undefined);

// ---------------------------------------------------------------- imprint
function imprintItems(print) {
  const it = [];
  const w = CW * 0.4;
  it.push({ el: new Paragraph({ spacing: exact(2 * S, 2 * S), indent: { right: mm(CW - w) }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 1 } }, children: [] }), h: 4 * S });
  const sz = hp(11), lh = 3.4 * S, o = { size: sz, lh, color: C.imprint };
  T.imprint.lines.forEach(l => it.push(l ? textP(l, o) : spacer(1.6 * S)));
  it.push(spacer(1.6 * S));
  if (!print) it.push(textP(T.imprint.online, o));                       // online: repository link
  else if (cfg.printer === 'epubli') T.imprint.epubli.forEach(l => it.push(textP(l, o)));
  else T.imprint.printFields.forEach(l => {                              // open print version: fields to fill in
    it.push(textP(l, o));
    it.push(grid([{ h: 4.6 * S, cells: [{ w: CW, borders: { top: none, left: none, right: none, bottom: dotted } }] }],
      { borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none } }));
    it.push(spacer(1 * S));
  });
  return it;
}
const itemsH = it => sum(it.map(x => x.h));

// ---------------------------------------------------------------- torn snippets (how-to examples)
// torn-paper edge: rendered per snippet so column lines and header shade run into the tear
const TORN = new Map();                      // key -> png buffer (filled between pass 1 and 2)
const PLACEHOLDER = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=', 'base64');
function tornImg(spec) {
  const key = JSON.stringify(spec);
  if (!TORN.has(key)) TORN.set(key, null);
  // tiny paragraph-mark font: no descent below the image (otherwise a white gap shows under the top edge)
  // floating image anchored to an empty paragraph of exactly its height: no baseline/descent gap
  return { el: new Paragraph({ run: { size: 2 }, spacing: { before: 0, after: 0, line: mm(spec.h), lineRule: LineRuleType.EXACT }, children: [new ImageRun({ type: 'png', data: TORN.get(key) || PLACEHOLDER,
    transformation: { width: spec.W / 25.4 * 96, height: spec.h / 25.4 * 96 },
    floating: { horizontalPosition: { relative: HorizontalPositionRelativeFrom.COLUMN, offset: 0 }, verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
      wrap: { type: TextWrappingType.NONE }, allowOverlap: true, lockAnchor: true } })] }), h: spec.h };
}
function tornSVG(sp) {
  let seed = sp.seed; const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
  const pts = []; let x = 0;
  while (x < sp.W) { pts.push([x, sp.h * (0.3 + rnd() * 0.32)]); x += 0.6 + rnd() * 1.2; }
  pts.push([sp.W, sp.h * 0.45]);
  const yAt = xx => { for (let i = 1; i < pts.length; i++) if (pts[i][0] >= xx) { const [x0, y0] = pts[i - 1], [x1, y1] = pts[i]; return y0 + (y1 - y0) * (xx - x0) / Math.max(x1 - x0, 1e-6); } return pts[pts.length - 1][1]; };
  const P = a => a.map(([x, y]) => `${x.toFixed(3)},${y.toFixed(3)}`).join(' ');
  const st = { solid: `stroke="#${C.line}" stroke-width="0.09"`, dotted: 'stroke="#A8A8A8" stroke-width="0.26" stroke-dasharray="0.26 0.26"',
    faint: 'stroke="#D6D6D6" stroke-width="0.26" stroke-dasharray="0.26 0.26"', frame: `stroke="#${C.frame}" stroke-width="0.26"` };
  const line = (xx, s1) => sp.top ? `<line x1="${xx}" y1="${yAt(xx).toFixed(3)}" x2="${xx}" y2="${sp.h}" ${st[s1]}/>` : `<line x1="${xx}" y1="0" x2="${xx}" y2="${yAt(xx).toFixed(3)}" ${st[s1]}/>`;
  const OFF = 0.09, fl = OFF, fr = sp.W - 0.25 + OFF;      // table borders sit ~0.09 mm right of the nominal edges
  const paper = sp.top ? [...pts, [sp.W, sp.h], [0, sp.h]] : [[0, 0], [sp.W, 0], ...pts.slice().reverse()];
  const shadow = pts.map(([x, y]) => [x, y + (sp.top ? -0.28 : 0.28)]);
  const px = v => Math.round(v / 25.4 * 600);   // render at 600 dpi
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${px(sp.W)}" height="${px(sp.h)}" viewBox="0 0 ${sp.W} ${sp.h}" preserveAspectRatio="none">
<rect width="${sp.W}" height="${sp.h}" fill="#ffffff"/><polygon points="${P(paper)}" fill="${sp.fill || '#ffffff'}"/>
<polyline points="${P(shadow)}" fill="none" stroke="#E6E6E6" stroke-width="0.4"/>
${sp.divs.map(d => line(+(d.x + OFF).toFixed(3), d.s)).join('')}${line(fl, 'frame')}${line(+fr.toFixed(3), 'frame')}
${sp.topLine ? `<line x1="${fl}" y1="${sp.h - 0.045}" x2="${fr.toFixed(3)}" y2="${sp.h - 0.045}" ${st.solid}/>` : ''}
<polyline points="${P(pts)}" fill="none" stroke="#8C8C8C" stroke-width="0.15" stroke-linejoin="round"/></svg>`;
}
function circSVG(sp) {
  const px = v => Math.round(v / 25.4 * 600), fs = sp.size / 2 * 0.3528;   // font size in mm
  const cx = sp.W / 2, cy = sp.H / 2, rx = sp.W * 0.42, ry = sp.H * 0.44;
  // hand-drawn ellipse: slightly more than one turn, open overlap
  const pts = []; for (let a = -0.5; a <= 2 * Math.PI + 0.35; a += 0.12) pts.push([cx + rx * (1 + 0.04 * Math.sin(3 * a)) * Math.cos(a), cy + ry * (1 + 0.05 * Math.cos(2 * a)) * Math.sin(a)]);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${px(sp.W)}" height="${px(sp.H)}" viewBox="0 0 ${sp.W} ${sp.H}">
<rect width="${sp.W}" height="${sp.H}" fill="${sp.fill || '#ffffff'}"/>
<text x="${cx}" y="${cy + fs * 0.33}" font-family="Caveat" font-size="${fs}" fill="#${C.hand}" text-anchor="middle">${sp.t}</text>
<polyline points="${pts.map(([x, y]) => x.toFixed(3) + ',' + y.toFixed(3)).join(' ')}" fill="none" stroke="#${C.hand}" stroke-width="0.22" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
}
function circImg(t, sizeHp, W, H) {
  const spec = { kind: 'circ', t, size: sizeHp, W: +W.toFixed(2), H: +H.toFixed(2) }, key = JSON.stringify(spec);
  if (!TORN.has(key)) TORN.set(key, null);
  return new ImageRun({ type: 'png', data: TORN.get(key) || PLACEHOLDER, transformation: { width: spec.W / 25.4 * 96, height: spec.H / 25.4 * 96 } });
}
function qrImg(url, sizeMm) {
  const spec = { kind: 'qr', url, W: sizeMm }, key = JSON.stringify(spec);
  if (!TORN.has(key)) TORN.set(key, null);
  return new ImageRun({ type: 'png', data: TORN.get(key) || PLACEHOLDER, transformation: { width: sizeMm / 25.4 * 96, height: sizeMm / 25.4 * 96 } });
}
async function renderTorn() {
  const sharp = require('sharp');
  for (const key of TORN.keys()) if (!TORN.get(key)) {
    const sp = JSON.parse(key);
    // all images without alpha channel (print services reject transparency)
    const opaque = buf => sharp(buf).flatten({ background: '#ffffff' }).removeAlpha().png().toBuffer();
    if (sp.kind === 'qr') { TORN.set(key, await opaque(await require('qrcode').toBuffer(sp.url, { margin: 0, width: 600, errorCorrectionLevel: 'M', color: { dark: '#333333', light: '#ffffff' } }))); continue; }
    TORN.set(key, await opaque(Buffer.from(sp.kind === 'circ' ? circSVG(sp) : tornSVG(sp))));
  }
}
const handRun = (t, sizeHp) => run(t, { font: HAND, size: sizeHp, color: C.hand });
// handwriting with explicit line breaks (\n)
const handRuns = (t, sizeHp) => t.split('\n').map((l, i) => new TextRun({ text: l, break: i ? 1 : undefined, font: HAND, size: sizeHp, color: C.hand }));
// cols: [[label, logicalW]], rows: [[...values]] ; a value {line:"…"} = full-width frame line
function snippet(label, cols, rows, o = {}) {
  const rh = o.rh || 4.8 * S, headH = o.headH || 4.6 * S, hs = o.handSize || hp(19);
  const W = o.width || CW;
  const scale = W / sum(cols.map(c => c[1]));
  const items = [];
  if (label && !o.noLabel) {
    const lab = run(label, { size: hp(12), color: C.text, bold: true });
    items.push(o.expl ? textP([lab, run('  ' + o.expl, { size: hp(13), color: C.text })], { plain: label + '  ' + o.expl, size: hp(13), lh: 4 * S, after: 0.8 * S })
      : textP(label, { size: hp(11), lh: 3.6 * S, color: C.label, bold: true }));
  }
  const tb = [];
  if (o.header) tb.push(...o.header(scale, headH));
  else tb.push({ h: headH, cells: cols.map(([t, w], i) => ({ w: w * scale, text: t, shade: true, valign: VerticalAlign.CENTER, size: hp(11),
    align: (o.headCenter || []).includes(i) ? AlignmentType.CENTER : undefined,
    borders: { top: none, bottom: thin, left: i === 0 ? frame : thin, right: i === cols.length - 1 ? frame : thin } })) });
  const hlh = hs / 2 * 0.3528 * 1.2;          // line height of the handwriting (mm)
  rows.forEach((r, ri) => {
    const last = ri === rows.length - 1;
    if (r.line) tb.push({ h: rh, cells: [{ w: W, runs: [handRun(r.line, hs)], valign: VerticalAlign.CENTER,
      borders: { top: none, bottom: none, left: frame, right: frame } }] });
    else {
      const cells = []; let i = 0;
      r.forEach(v => {
        const span = (v && v.span) || 1, t = v && v.span !== undefined ? v.t : (v && v.circ ? null : v);
        const w = sum(cols.slice(i, i + span).map(c => c[1])) * scale, j = i + span - 1;
        const fl = o.faintCols || [];
        const circ = v && v.circ ? [circImg(v.circ, hs, Math.min(w - 1.2, 9 * S), Math.min(rh - 0.6, 5.2 * S))] : null;
        cells.push({ w, runs: circ || (t ? handRuns(t, hs) : []), valign: VerticalAlign.CENTER, lh: circ ? undefined : hlh, _t: t,
          align: span > 1 || (o.center || []).includes(i) ? AlignmentType.CENTER : undefined,
          borders: { top: none, bottom: none, left: i === 0 ? frame : (fl.includes(i) ? faint : dotted), right: j === cols.length - 1 ? frame : (fl.includes(j + 1) ? faint : dotted) } });
        i += span;
      });
      const lines = Math.max(1, ...cells.map(c => c._t ? sum(c._t.split('\n').map(l => wrapCount(l, hs, c.w - 1.4 * S, HAND))) : 1));
      tb.push({ h: Math.max(rh, lines * hlh + 1.4 * S), cells });
    }
  });
  const g = grid(tb, { borders: { top: none, bottom: none, left: frame, right: frame, insideHorizontal: none, insideVertical: dotted } });
  // divider positions: header (first row, solid) continues into the top tear, last row (dotted) into the bottom tear
  const divs = (cells, styleOf) => { let x = 0; return cells.slice(0, -1).map(c => { x += c.w; return { x: +x.toFixed(3), s: styleOf(c) }; }); };
  const topCells = tb[0].cells.filter(c => c.x === undefined);
  const lastCells = tb[tb.length - 1].cells;
  const th = 3 * S, seed = Math.round(W * 10) + rows.length * 7 + (label || '').length;
  const topDivs = o.topDivs ? o.topDivs.map(x => ({ x: +x.toFixed(3), s: 'solid' })) : divs(topCells, () => 'solid');
  const IW = +(W + 0.25).toFixed(3);
  items.push(tornImg({ W: IW, h: th, top: true, fill: '#' + C.shade, seed, divs: topDivs, topLine: !!o.topLine }), g,
    tornImg({ W: IW, h: th, top: false, seed: seed + 3, divs: divs(lastCells, c => c.borders && c.borders.right === faint ? 'faint' : 'dotted') }));
  return items;
}

// ---------------------------------------------------------------- session table definitions
const XCOL = (cfg.freeCol || '').slice(0, 6);    // pre-printed label of the free column after FILTER (empty by default)
const SESS_A4 = { head: [['#', 14], ['DATE', 26], ['PROJECT · SESSION', 70], ['ENLARGER/LENS', 44], ['CHEM', 32]],
    cols: [['ROLL/NEG', 18], ['PRINT', 13], ['C', 10.5], ['M', 10.5], ['Y', 10.5], [XCOL, 12], ['TIME', 13], ['f/', 11], ['PAPER', 14], ['SIZE', 16], ['DODGE · BURN · NOTES', 57.5]] };
const SESS_A5 = { head: [['#', 11], ['DATE', 21], ['PROJECT · SESSION', 48], ['ENLARGER/LENS', 26], ['CHEM', 18]],
    cols: [['PRINT', 11], ['C', 8.5], ['M', 8.5], ['Y', 8.5], [XCOL, 10], ['TIME', 11], ['f/', 9], ['DODGE · BURN · NOTES', 57.5]] };
const SESS = A4 ? SESS_A4 : SESS_A5;
const FI = SESS.cols.findIndex(c => c[0] === 'C');     // first filter column
// two-level column header: FILTER over C · M · Y
function colHeader(cols, scale, h1, h2, o = {}) {
  const FI = cols.findIndex(c => c[0] === 'C');
  let x = 0; const xs = cols.map(([, w]) => { const v = x; x += w * scale; return v; });
  const lastI = cols.length - 1;
  const r1 = [], r2 = [];
  cols.forEach(([t, w], i) => {
    const bL = i === 0 ? frame : (i === FI + 1 || i === FI + 2 ? faint : thin), bR = i === lastI ? frame : (i === FI || i === FI + 1 ? faint : thin);
    if (i === FI) r1.push({ w: 3 * w * scale, text: 'FILTER', shade: true, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, size: o.size,
      borders: { top: o.top || frame, bottom: none, left: thin, right: thin } });
    else if (i === FI + 1 || i === FI + 2) { /* covered by FILTER */ }
    else r1.push({ w: w * scale, rowSpan: 2, text: t, runs: (!t && o.xHand) ? [run(o.xHand, { font: HAND, size: o.xHandSize, color: C.hand })] : undefined, shade: true, valign: VerticalAlign.CENTER, size: o.size,
      align: t.length <= 6 ? AlignmentType.CENTER : undefined, borders: { top: o.top || frame, bottom: thin, left: i === 0 ? frame : thin, right: i === lastI ? frame : thin } });
    if (i >= FI && i <= FI + 2) r2.push({ x: xs[i], w: w * scale, text: t, shade: true, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, size: hp(10), color: C.hint,
      borders: { top: none, bottom: thin, left: i === FI ? thin : faint, right: i === FI + 2 ? thin : faint } });
  });
  return [{ h: h1, cells: r1 }, { h: h2, cells: r2 }];
}
function sessionBlock(n, blockH) {
  const hH = 9 * S, h1 = 5 * S, h2 = 3.5 * S, nH = 17 * S;
  const body = blockH - hH - h1 - h2 - nH;
  const rows = [];
  rows.push({ h: hH, cells: SESS.head.map(([t, w], i) => i === 0
    ? { w, runs: [run(String(n).padStart(3, '0'), { size: hp(18), bold: true, color: C.text })], shade: true, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 } }
    : { w, text: t, shade: true }) });
  rows.push(...colHeader(SESS.cols, 1, h1, h2, { top: thin }));
  const last = SESS.cols.length - 1;
  rows.push({ h: body, cells: SESS.cols.map(([, w], i) => ({ w,
    borders: { top: thin, bottom: thin, left: i === 0 ? frame : (i === FI + 1 || i === FI + 2 ? faint : dotted), right: i === last ? frame : (i === FI || i === FI + 1 ? faint : dotted) } })) });
  rows.push({ h: nH, cells: [{ w: CW, text: 'SESSION NOTES', shade: true }] });
  return grid(rows);
}

// ---------------------------------------------------------------- pages
function pageTitle() {
  newPage();
  add({ el: new Paragraph({ alignment: AlignmentType.RIGHT, spacing: exact(9 * S, 1.5 * S), children: [run('DARKROOM LOG', { size: hp(34), bold: true, color: C.text, spacing: Math.round(60 * S) })] }), h: 10.5 * S });
  add({ el: new Paragraph({ alignment: AlignmentType.RIGHT, spacing: exact(4.5 * S, 6 * S), children: [run('A PRINTING WORKBOOK', { size: hp(14), color: C.label, spacing: Math.round(40 * S) })] }), h: 10.5 * S });
  add(grid([{ h: 9 * S, cells: [['BOOK NO.', 24], ['FROM', 30], ['TO', 30], ['DARKROOM', 40]].map(([t, w]) => ({ w: L(w), text: t, shade: true })) },
    { h: 14 * S, cells: [{ w: CW, text: 'OWNER · IF FOUND, PLEASE RETURN TO', shade: true }] }]));
  // contents, bottom-aligned
  const items = [title('CONTENTS')];
  toc.forEach(([t, p, id, sub]) => {
    const lh = sub ? 4.8 * S : 6.0 * S;
    const kids = [run(t, { size: hp(sub ? 13 : 15), color: sub ? C.label : C.text }), new TextRun({ children: [new Tab()] }), run(String(p), { size: hp(sub ? 13 : 15), color: sub ? C.label : C.text })];
    items.push({ el: new Paragraph({ spacing: exact(lh), indent: sub ? { left: mm(6 * S) } : undefined,
      tabStops: [{ type: TabStopType.RIGHT, position: mm(CW), leader: LeaderType.DOT }],
      children: ONLINE ? [new InternalHyperlink({ anchor: id, children: kids })] : kids }), h: lh });
  });
  add(spacer(remaining() - itemsH(items)));
  add(items);
}

const SN = T.samples;
function setupSnippets(o) {
  const E = o.expl ? T.howto.setupExpl : {};
  add(snippet(T.labels.processes, [['ABBREV', 13], ['DEVELOPER', 28], ['STOP', 20], ['FIX / BLIX', 26], ['WASH', 15], ['NOTES', 22]], SN.processes, { ...o, expl: E.processes, headCenter: [0], center: [0] }));
  add(spacer(o.gap));
  add(snippet(T.labels.enlargers, [['ABBREV', 13], ['ENLARGER', 28], ['FORMAT / HEAD', 26], ['LENS', 37], ['NOTES', 20]], (o.ext ? SN.enlargersExt : SN.enlargers), { ...o, expl: E.enlargers, headCenter: [0], center: [0] }));
  add(spacer(o.gap));
  add(snippet(T.labels.papers, [['ABBREV', 12], ['PAPER  make · name · grade · surface', 52], ['SIZE', 14], ['BASE', 11], ['PROCESS', 13], ['NOTES', 22]], (o.ext ? SN.papersExt : SN.papers), { ...o, expl: o.expl ? (A4 ? E.papersA4 : E.papersA5) : undefined, headCenter: [0], center: [0] }));
}
function sessionSnippet(o) {
  const small = o.a5 || !A4;                 // A5 style: frame lines; A4 style: extra columns
  const def = small ? SESS_A5 : SESS_A4;
  const cols = def.cols.map(c => [c[0], c[1]]);
  const rows = small ? SN.sessionA5 : SN.sessionA4;
  const fi = cols.findIndex(c => c[0] === 'C');
  const W = o.width || CW, sc = W / sum(cols.map(c => c[1])), hs = o.handSize || hp(19);
  // torn edge runs through the session header row: its dividers continue into the tear
  let hx = 0; const topDivs = def.head.slice(0, -1).map(([, w]) => (hx += w) * sc);
  const H = SN.sessionHead, vals = [null, H.date, H.title, H.enl, H.chem];
  const lab = t => new Paragraph({ spacing: exact(2.6 * S), children: [run(t, { size: hp(11) })] });
  const val = t => new Paragraph({ spacing: exact(5 * S), indent: { left: mm(1 * S) }, children: [run(t, { font: HAND, size: hs, color: C.hand })] });
  const headRow = (scale) => ({ h: 9.5 * S, cells: def.head.map(([t, w], i) => i === 0
    ? { w: w * scale, shade: true, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 },
      runs: [run(H.no, { size: hp(18), bold: true, color: C.text })], borders: { top: none, bottom: thin, left: frame, right: thin } }
    : { w: w * scale, shade: true, paras: [lab(t), val(vals[i])], borders: { top: none, bottom: thin, left: thin, right: i === def.head.length - 1 ? frame : thin } }) });
  add(snippet(T.labels.session, cols, rows, { ...o, topDivs, faintCols: [fi + 1, fi + 2],
    center: cols.map((c, i) => i).filter(i => i < cols.length - 1),
    header: (scale, h) => [headRow(scale), ...colHeader(cols, scale, h, 3 * S, { top: thin, size: hp(11), xHand: SN.xHead, xHandSize: hs })] }));
}
function bullets(list, o) { list.forEach(b => add(textP(b, { size: o.size, lh: o.lh, color: C.text, indent: 2 * S, hang: 3.6 * S }))); }
const subHead = t => add(textP(t, { size: hp(13), lh: 4.4 * S, color: C.text, bold: true, before: 2.2 * S }));

function pageHowto1() {
  newPage();
  add(title('HOW TO USE THIS WORKBOOK', null, { bookmark: bm('howto') }));
  const ts = { size: hp(14), lh: 4.2 * S };
  add(textP(T.howto.intro, { ...ts, after: 2.5 * S }));
  setupSnippets({ rh: 5.2 * S, headH: 4.6 * S, gap: 2.5 * S, handSize: hp(20), expl: true });
  subHead(T.howto.workflowTitle);
  bullets(T.howto.workflow, ts);
  add(textP(T.howto.qr, { ...ts, before: 2 * S }));
  // link to the extended how-to: short text + QR code
  const q = 16 * S, tw = CW - q - 3 * S, ms = { top: 0, bottom: 0, left: 0, right: 0 };
  const head = run(T.howto.moreTitle + '  ', { size: hp(13), bold: true, color: C.text });
  const kids = ONLINE ? T.howto.moreOnline.map(x => typeof x === 'object'
    ? new ExternalHyperlink({ link: T.ext.howtoUrl, children: [new TextRun({ text: x.link, font: FONT, size: ts.size, color: '2B4C7E', underline: {} })] })
    : run(x, { size: ts.size, color: C.text })) : [run(T.howto.morePrint, { size: ts.size, color: C.text })];
  add(spacer(2.5 * S));
  add(grid([{ h: q, cells: [
    { w: tw, margins: ms, valign: VerticalAlign.CENTER, paras: [new Paragraph({ spacing: exact(ts.lh), children: [head, ...kids] })], borders: { top: none, bottom: none, left: none, right: none } },
    { w: 3 * S, margins: ms, borders: { top: none, bottom: none, left: none, right: none } },
    { w: q, margins: ms, runs: [qrImg(T.ext.howtoUrl, q)], borders: { top: none, bottom: none, left: none, right: none } }] }],
    { borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none } }));
  if (!ONLINE) { const imp = imprintItems(true); add(spacer(remaining() - itemsH(imp))); add(imp); }
}
function pageHowto2() {
  newPage();
  const ts = { size: hp(14), lh: 4.2 * S };
  add(textP(T.howto.sessionIntro, { ...ts, after: 1 * S }));
  add(textP(T.howto.sessionHead, { ...ts, after: 1.5 * S }));
  sessionSnippet({ rh: 5.6 * S, headH: 4.8 * S, handSize: hp(20) });
  subHead(T.howto.readTitle);
  bullets(T.howto.read, ts);
  subHead(T.howto.notationTitle);
  bullets(T.howto.bulletsShort.concat([XCOL === 'INT' ? T.howto.bulletINT : T.howto.bulletX], A4 ? [T.howto.bulletA4] : []), ts);
  subHead(T.howto.convTitle);
  add(textP(T.howto.convText, ts));
  add(dottedLines(A4 ? 6 : 4, 7 * S));
  if (ONLINE) { const imp = imprintItems(false); add(spacer(remaining() - itemsH(imp))); add(imp); }
}
function pageSetup() {
  newPage();
  add(title('PROCESSES', 'chem · dilution · time · °C', { bookmark: bm('setup') }));
  add(listTable([['ABBREV', 13], ['DEVELOPER', 28], ['STOP', 20], ['FIX / BLIX', 26], ['WASH', 15], ['NOTES', 22]], A4 ? 6 : 5, 14 * S, { headCenter: [0], center: [0] }));   // A4: one more process row
  add(spacer(6 * S));
  add(title('ENLARGERS & LENSES'));
  const n = Math.floor((remaining() - 6 * S) / (8.5 * S)), rh = (remaining() - 6 * S) / n;   // stretch rows to the bottom
  add(listTable([['ABBREV', 13], ['ENLARGER', 28], ['FORMAT / HEAD', 26], ['LENS', 37], ['NOTES', 20]], n, rh, { headCenter: [0] }));
}
function pagePapers() {
  newPage();
  add(title('PRINTING PAPERS', null, { bookmark: bm('papers') }));
  const n = Math.floor((remaining() - 6 * S) / (10 * S)), rh = (remaining() - 6 * S) / n;
  add(listTable([['ABBREV', 12], ['PAPER  make · name · grade · surface', 52], ['SIZE', 14], ['BASE', 11], ['PROCESS', 13], ['NOTES', 22]], n, rh, { headCenter: [0] }));
}
function pageProjIdx() {
  newPage();
  add(title('PROJECT INDEX', null, { bookmark: bm('projidx') }));
  const n = Math.floor((remaining() - 6 * S) / (7.4 * S)), rh = (remaining() - 6 * S) / n;
  add(listTable([['#', 10], ['PROJECT', 50], ['SESSIONS #', 32], ['NOTES', 32]], n, rh,
    { headCenter: [0], center: [0], fill: (r, i) => i === 0 ? String(r + 1) : undefined, cellColor: C.hint }));
}
var sessIdxStart = 0;
function pageSessIdx(perPage, first) {
  newPage();
  const a = sessIdxStart; sessIdxStart += perPage;
  add(title('SESSION INDEX', `sessions ${String(a + 1).padStart(3, '0')}–${String(a + perPage).padStart(3, '0')}`, { bookmark: first ? bm('sessidx') : undefined }));
  const rh = (remaining() - 6 * S) / perPage;
  add(listTable([['#', 10], ['DATE', 21], ['SESSION TITLE · PROJECT', 45], ['NOTES', 48]], perPage, rh,
    { headCenter: [0], center: [0], fill: (r, i) => i === 0 ? String(a + r + 1).padStart(3, '0') : undefined, cellColor: C.hint }));
}
function pageLines(t, sub, id) {
  newPage();
  add(title2(t, sub, { bookmark: id }));
  fillLines(7 * S);
}
var sessNo = 0;
function pageSession(first) {
  newPage();
  if (first && ONLINE) add({ el: new Paragraph({ spacing: exact(0.3), children: [new Bookmark({ id: 'session', children: [new TextRun({ text: '', size: 2 })] })] }), h: 0.3 });
  const gap = 3 * S;
  const blockH = (remaining() - gap) / 2;
  add(sessionBlock(++sessNo, blockH)); add(spacer(gap)); add(sessionBlock(++sessNo, blockH));
}
// quick reference pages (content as in the personal edition)
const fmt = v => v < 10 ? v.toFixed(1) : String(Math.round(v));
function lineT(t, o = {}) { add(textP(t, { size: hp(14), lh: 4.8 * S, color: C.text, ...o })); }
function pageQR1() {
  newPage();
  add(title('QUICK REFERENCE', 'exposure', { bookmark: bm('qr1') }));
  add(textP('f-stop timing (seconds) · 1 stop = ×2 · 1/2 = ×1.41 · 1/3 = ×1.26', { size: hp(12), lh: 5 * S, color: C.label, after: 1.5 * S }));
  const ks = [-1, -1 / 2, -1 / 3, 0, 1 / 3, 1 / 2, 1], kl = ['-1', '-1/2', '-1/3', 'BASE', '+1/3', '+1/2', '+1'];
  const bases = [4, 5, 6, 8, 10, 12, 15, 20, 25, 30, 40, 50, 60];
  add(listTable(kl.map((k, i) => [k, i === 3 ? 20 : 17.33]), bases.length, 5.5 * S, { headCenter: [0, 1, 2, 3, 4, 5, 6], center: [0, 1, 2, 3, 4, 5, 6],
    fill: (r, i) => fmt(bases[r] * Math.pow(2, ks[i])), cellSize: hp(14), cellColor: C.text }));
  add(spacer(6 * S));
  lineT('Print size change (same aperture)', { bold: true });
  ['  ▸ bigger image → same light spread over more area → more time', '  ▸ t new = t old × (new length ÷ old length)²',
    '  ▸ length = long side of the projected image', '  ▸ e.g. 18 → 24 cm: (24÷18)² = 1.78 → 10 s becomes 18 s'].forEach(l => lineT(l));
  add(spacer(2.5 * S));
  const sizes = [['9×13', 13], ['10×15', 15], ['13×18', 18], ['18×24', 24], ['24×30', 30], ['30×40', 40]];
  add(listTable([['FROM \\ TO', 19], ...sizes.map(([n]) => [n, 17.5])], sizes.length, 5.5 * S, { headCenter: [0, 1, 2, 3, 4, 5, 6], center: [0, 1, 2, 3, 4, 5, 6],
    fill: (r, i) => i === 0 ? sizes[r][0] : (r === i - 1 ? '–' : Math.pow(sizes[i - 1][1] / sizes[r][1], 2).toFixed(2)), cellSize: hp(14), cellColor: C.text }));
  add(spacer(4 * S)); fillLines(7 * S);
}
function pageQR2() {
  newPage();
  add(title('QUICK REFERENCE', 'color · contrast', { bookmark: bm('qr2') }));
  add(textP('RA-4 cast correction (Y/M filtration)', { size: hp(12), lh: 5 * S, color: C.label, after: 1.5 * S }));
  const casts = [['yellow', 'add Y'], ['blue', 'remove Y'], ['magenta', 'add M'], ['green', 'remove M'], ['red', 'add Y + M'], ['cyan', 'remove Y + M']];
  add(listTable([['PRINT HAS CAST', 55], ['CORRECTION', 69]], casts.length, 5.5 * S, { fill: (r, i) => casts[r][i], cellSize: hp(14), cellColor: C.text }));
  add(spacer(5 * S));
  add(textP('B/W variable contrast on color head: more Y = softer · more M = harder', { size: hp(12), lh: 5 * S, color: C.label, after: 1.5 * S }));
  const grades = ['00', '0', '1/2', '1', '1 1/2', '2', '2 1/2', '3', '3 1/2', '4', '4 1/2', '5'];
  add(listTable([['GRADE', 20], ['Y', 20], ['M', 20], ['EXP. FACTOR', 26], ['MY HEAD / PAPER', 38]], grades.length, 6 * S,
    { headCenter: [0, 1, 2], center: [0], fill: (r, i) => i === 0 ? grades[r] : undefined, cellSize: hp(14), cellColor: C.text }));
  add(spacer(4 * S)); fillLines(7 * S);
}
function pageQR3() {
  newPage();
  add(title('QUICK REFERENCE', 'zone system · printing', { bookmark: bm('qr3') }));
  add(textP('Zone system: 11 tones from black to white, 1 stop apart', { size: hp(12), lh: 5 * S, color: C.label, after: 1.5 * S }));
  const zones = [['0', 'max. black of the paper', 'deepest shadows, no detail'], ['I', 'near black, no texture', 'first step above max. black'],
    ['II', 'first hint of texture', 'deep shadows'], ['III', 'dark, full texture', 'dark materials, shadow detail'], ['IV', 'dark mid-tone', 'dark foliage, open shade'],
    ['V', 'middle grey (18 %, grey card)', 'dry old asphalt, weathered wood'], ['VI', 'light mid-tone', 'light skin in sun, light stone'],
    ['VII', 'light, full texture', 'very light skin, light grey'], ['VIII', 'white with texture', 'textured snow, light fabric'],
    ['IX', 'white, barely tone', 'snow in flat sun, glare'], ['X', 'paper white', 'light sources, specular reflections']];
  const grey = z => { const Lz = z * 10, Y = Lz > 8 ? Math.pow((Lz + 16) / 116, 3) : Lz / 903.3;
    const v = Y <= 0.0031308 ? 12.92 * Y : 1.055 * Math.pow(Y, 1 / 2.4) - 0.055; const h = Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase(); return h + h + h; };
  add(listTable([['', 6.5], ['ZONE', 12], ['PRINT VALUE', 48], ['TYPICAL SUBJECT', 57.5]], zones.length, 6 * S,
    { headCenter: [0, 1], center: [1], fill: (r, i) => i === 0 ? undefined : zones[r][i - 1], fillColor: (r, i) => i === 0 ? grey(r) : undefined, cellSize: hp(13), cellColor: C.text }));
  add(spacer(5 * S));
  lineT('Printing rules of thumb', { bold: true });
  ['  ▸ time sets the highlights → find time for zone VIII', '    (first texture in the whites) on the test strip',
    '  ▸ contrast sets the shadows → pick grade so zone II/III', '    keep detail, then re-check time',
    '  ▸ prints dry darker (esp. FB) → judge wet prints lighter'].forEach(l => lineT(l));
  add(spacer(3 * S)); fillLines(7 * S);
}

// ---------------------------------------------------------------- extended how-to (A4, 2 pages)
function buildHowtoExtended() {
  const ts = () => ({ size: hp(18), lh: 5.2 * S });
  const H = (t, o = {}) => add(textP(t, { size: hp(20), lh: 6 * S, color: C.text, bold: true, before: o.before === undefined ? 3 * S : o.before, after: 1 * S }));
  fitPage(() => {
    newPage();
    add({ el: new Paragraph({ spacing: exact(8 * S, 1 * S), children: [run('DARKROOM LOG', { size: hp(30), bold: true, color: C.text, spacing: 40 })] }), h: 9 * S });
    add(textP('How to use the workbook – extended guide', { size: hp(18), lh: 5.5 * S, color: C.label, after: 3 * S }));
    T.ext.intro.forEach(p => {
      if (!Array.isArray(p)) return add(textP(p, { ...ts(), after: 1.5 * S }));
      const o = ts(), kids = p.map(x => typeof x === 'object'   // note: strings have a built-in .link()
        ? new ExternalHyperlink({ link: T.ext.repoUrl, children: [new TextRun({ text: x.link, font: FONT, size: o.size, color: '2B4C7E', underline: {} })] })
        : run(x, { size: o.size, color: C.text }));
      add(textP(kids, { ...o, plain: p.map(x => typeof x === 'object' ? x.link : x).join(''), after: 1.5 * S }));
    });
    H('What is inside');
    T.ext.contents.forEach(b => add(textP('▸ ' + b, { ...ts(), indent: 2, hang: 4.5 })));
    H('Setup tables');
    add(textP(T.ext.setup, { ...ts(), after: 1.5 * S }));
    setupSnippets({ rh: 6 * S, headH: 5 * S, gap: 2 * S, handSize: hp(22), ext: true });
  });
  fitPage(() => {
    newPage();
    H('The session log', { before: 0 });
    T.ext.session.forEach(p => add(textP(p, { ...ts(), after: 1.2 * S })));
    sessionSnippet({ rh: 6 * S, headH: 5 * S, handSize: hp(22), a5: true });
    H('Notation – one way to do it');
    T.ext.ids.forEach(b => add(textP('▸ ' + b, { ...ts(), indent: 2, hang: 4.5 })));
    H('Tips');
    T.ext.tips.forEach(b => add(textP('▸ ' + b, { ...ts(), indent: 2, hang: 4.5 })));
  }, 1.15);
}

// ---------------------------------------------------------------- example overview (for reviewing the sample texts)
function buildSamples() {
  newPage();
  const cap = t => add(textP(t, { size: hp(16), lh: 6 * S, color: C.text, bold: true, before: 4 * S, after: 1 * S }));
  const o = { rh: 5.2 * S, headH: 4.6 * S, handSize: hp(20) };
  const P = [['ABBREV', 13], ['DEVELOPER', 28], ['STOP', 20], ['FIX / BLIX', 26], ['WASH', 15], ['NOTES', 22]];
  const E = [['ABBREV', 13], ['ENLARGER', 28], ['FORMAT / HEAD', 26], ['LENS', 37], ['NOTES', 20]];
  const Pa = [['ABBREV', 12], ['PAPER  make · name · grade · surface', 52], ['SIZE', 14], ['BASE', 11], ['PROCESS', 13], ['NOTES', 22]];
  if (!A4) {
    cap('#1  Processes (all versions)'); add(snippet(null, P, SN.processes, { ...o, headCenter: [0], center: [0] }));
    cap('#2  Enlargers & lenses (workbooks)'); add(snippet(null, E, SN.enlargers, { ...o, headCenter: [0], center: [0] }));
    cap('#3  Printing papers (workbooks)'); add(snippet(null, Pa, SN.papers, { ...o, headCenter: [0], center: [0] }));
    newPage();
    cap('#4  Enlargers & lenses (extended how-to)'); add(snippet(null, E, SN.enlargersExt, { ...o, headCenter: [0], center: [0] }));
    cap('#5  Printing papers (extended how-to)'); add(snippet(null, Pa, SN.papersExt, { ...o, headCenter: [0], center: [0] }));
    newPage();
    cap('#6  Session log, A5 style with frame lines (A5 workbooks + extended how-to)'); sessionSnippet({ ...o, rh: 5.6 * S });
  } else {
    cap('#7  Session log, A4 style with extra columns (A4 workbooks)'); sessionSnippet({ ...o, rh: 5.6 * S });
  }
}

// ---------------------------------------------------------------- assemble
function buildContent() {
  pages.length = 0; cur = null; sessNo = 0; sessIdxStart = 0; S = S0; FITM = 0;
if (cfg.kind === 'howto') buildHowtoExtended();
else if (cfg.kind === 'samples') buildSamples();
else {
  const only = cfg.only ? new Set(cfg.only) : null, seen = new Set();
  plan.forEach(([k], i) => {
    if (only && (!only.has(k) || seen.has(k))) return;
    seen.add(k);
    switch (k) {
      case 'title': pageTitle(); break;
      case 'howto1': fitPage(pageHowto1, A4 ? 1.3 : 1.15); break;
      case 'howto2': fitPage(pageHowto2, A4 ? 1.3 : 1.15); break;
      case 'setup': pageSetup(); break;
      case 'papers': pagePapers(); break;
      case 'projidx': pageProjIdx(); break;
      case 'sessidx': pageSessIdx(RING ? 24 : 25, plan[i - 1][0] !== 'sessidx'); break;
      case 'learn': pageLines('MY DARKROOM LEARNINGS & NOTES', 'papers · tools · techniques · experience', plan[i - 1][0] !== 'learn' ? bm('learn') : undefined); break;
      case 'session': pageSession(plan[i - 1][0] !== 'session'); break;
      case 'qr1': pageQR1(); break;
      case 'qr2': pageQR2(); break;
      case 'qr3': pageQR3(); break;
      case 'notes': pageLines('GENERAL KNOWLEDGE NOTES', 'chemistry · light · color · contrast · materials', plan[i - 1][0] !== 'notes' ? bm('notes') : undefined); break;
    }
  });
}
  // overflow report
  pages.forEach((p, i) => { if (p.used > AVAIL + 0.01) console.error(`WARN page ${i + 1}: ${p.used.toFixed(1)} > ${AVAIL}`); });
}

async function main() {
  const log = console.log; console.log = () => {};
  buildContent();                            // pass 1: collect torn-edge images
  await renderTorn();
  console.log = log;
  buildContent();                            // pass 2: final content

const pnFooter = align => new Footer({ children: [new Paragraph({ alignment: align, spacing: exact(3 * S),
  children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: hp(14), color: C.label })] })] });
const licFooter = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: exact(4),
  children: [run(T.ext.footer, { size: 16, color: C.imprint })] })] });
const pageProps = (margins) => ({ page: { size: { width: mm(PW), height: mm(PH) }, margin: margins } });
const content = pages.flatMap(p => p.items.map(x => x.el));
const mainMargins = { top: mm(M.top), bottom: mm(M.bottom), left: mm(M.left), right: mm(M.right), header: mm(5), footer: mm(M.footer) };
const sections = [];
function coverSection(file) {
  const img = fs.readFileSync(file);
  return { properties: { ...pageProps({ top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 }) },
    footers: { default: new Footer({ children: [new Paragraph({ children: [] })] }), even: new Footer({ children: [new Paragraph({ children: [] })] }) },
    // full-page image, floating at the page origin (no white strip, no overflow page)
    children: [new Paragraph({ spacing: exact(0.4), children: [new ImageRun({ type: 'jpg', data: img,
      transformation: { width: PW / 25.4 * 96, height: PH / 25.4 * 96 },
      floating: { horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 }, verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
        wrap: { type: TextWrappingType.NONE }, behindDocument: true, allowOverlap: true } })] })] };
}
if (cfg.kind === 'howto') {
  sections.push({ properties: pageProps(mainMargins), footers: { default: licFooter(), even: licFooter() }, children: content });
} else {
  const theme = cfg.cover || 'light';         // online versions: cover theme light | dark | red
  if (ONLINE) sections.push(coverSection(path.join(cfg.covers, `cover_${cfg.size}_online_${theme}_front.jpg`)));
  sections.push({ properties: { ...pageProps(mainMargins), page: { ...pageProps(mainMargins).page, pageNumbers: { start: 1 } } },
    footers: { default: pnFooter(AlignmentType.RIGHT), even: pnFooter(AlignmentType.LEFT) }, children: content });
  if (ONLINE) sections.push(coverSection(path.join(cfg.covers, `cover_${cfg.size}_online_${theme}_back.jpg`)));
}
const doc = new Document({
  creator: 'Darkroom Log', title: 'Darkroom Log – A Printing Workbook',
  evenAndOddHeaderAndFooters: true,
  styles: { default: { document: { run: { font: FONT, size: hp(16) } } } },
  sections,
});
const b = await Packer.toBuffer(doc); fs.writeFileSync(cfg.out, b); console.log(`${path.basename(cfg.out)}: ${pages.length} content pages`);
}
main();
