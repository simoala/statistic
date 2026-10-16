/*
 * TaRL one-page report generator.
 * Reads a TaRL placement export (.xlsx), compares the pre- and post-placement
 * levels for Arabic, French and Maths, and builds a one-page Word report.
 * Everything runs in the browser: the file never leaves the user's computer.
 */
'use strict';

// ---------------------------------------------------------------------------
// 1. Configuration: subjects and placement levels (lowest → highest).
//    To support a new label, add it to the matching `aliases` list.
// ---------------------------------------------------------------------------
const SUBJECTS = [
  { key: 'ar', title: 'اللغة العربية', match: /العربية/, color: '0F7A55', fill: 'E3F5EE' },
  { key: 'fr', title: 'اللغة الفرنسية', match: /الفرنسية/, color: '5B45C9', fill: 'ECE9FB' },
  { key: 'ma', title: 'الرياضيات', match: /الرياضيات/, color: 'B7471B', fill: 'FCEBE3' },
];
const LEVELS = {
  ar: [
    { label: 'مبتدئ', aliases: ['مبتدئ'] },
    { label: 'حرف', aliases: ['حرف', 'حروف'] },
    { label: 'كلمة', aliases: ['كلمة', 'كلمات'] },
    { label: 'فقرة', aliases: ['فقرة', 'جملة', 'جملة/فقرة', 'جملة / فقرة'] },
    { label: 'أقصوصة', aliases: ['أقصوصة', 'قصة'] },
    { label: 'تحدي', aliases: ['تحدي', 'تحد'] },
  ],
  fr: [
    { label: 'Débutant', aliases: ['débutant'] },
    { label: 'Lettre', aliases: ['lettre', 'lettres'] },
    { label: 'Mot simple', short: 'M. simple', aliases: ['mot simple', 'mots simples'] },
    { label: 'Mot avancé', short: 'M. avancé', aliases: ['mot avancé', 'mots avancés'] },
    { label: 'Graphème complexe', short: 'G. complexe', aliases: ['graphème complexe', 'graphèmes complexes', 'mot avec g.c', 'mot avec gc', 'mot avec graphème complexe'] },
    { label: 'Phrase D.P', aliases: ['phrase (d.p)', 'phrase d.p', 'phrase avec d.p', 'phrase (dp)', 'phrase dp', 'phrase'] },
    { label: 'Texte', aliases: ['texte'] },
  ],
  ma: [
    { label: 'مبتدئ', aliases: ['مبتدئ'] },
    { label: 'رقم', aliases: ['رقم', 'رقم واحد'] },
    { label: 'رقمان', aliases: ['رقمان', 'رقمين'] },
    { label: 'الجمع', aliases: ['جمع'] },
    { label: 'الطرح', aliases: ['طرح'] },
    { label: 'الضرب', aliases: ['ضرب'] },
    { label: 'القسمة', aliases: ['قسمة'] },
    { label: 'تحدي', aliases: ['تحدي', 'تحد'] },
  ],
};
const COLOR_PRE = '#eb6834';
const COLOR_POST = '#2a78d6';

// Normalise a label so that small spelling differences still match
// (spaces, case, French accents, Arabic hamza/diacritics, leading "ال").
function norm(s) {
  let t = String(s ?? '').normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[ً-ٰٟـ]/g, '')
    .replace(/ى/g, 'ي')
    .toLowerCase().replace(/\s+/g, ' ').trim();
  if (/^ال[؀-ۿ]{2,}/.test(t)) t = t.slice(2);
  return t;
}
const LEVEL_INDEX = {};
for (const [k, list] of Object.entries(LEVELS)) {
  LEVEL_INDEX[k] = new Map();
  list.forEach((lv, i) => { LEVEL_INDEX[k].set(norm(lv.label), i); lv.aliases.forEach(a => LEVEL_INDEX[k].set(norm(a), i)); });
}

// ---------------------------------------------------------------------------
// 2. Minimal .xlsx reader (JSZip + DOMParser).
// ---------------------------------------------------------------------------
function colIndex(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function textOf(node) {
  let s = '';
  for (const t of node.getElementsByTagName('t')) if (!t.parentNode || t.parentNode.nodeName !== 'rPh') s += t.textContent;
  return s;
}
async function readXlsx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = async p => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
  const wb = await xml('xl/workbook.xml');
  if (!wb) throw new Error('الملف ليس ملف Excel صالحاً (.xlsx).');
  const rels = await xml('xl/_rels/workbook.xml.rels');
  const relMap = {};
  for (const r of rels.getElementsByTagName('Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');
  const sst = await xml('xl/sharedStrings.xml');
  const shared = sst ? Array.from(sst.getElementsByTagName('si'), textOf) : [];
  const sheets = [];
  for (const sh of wb.getElementsByTagName('sheet')) {
    const rid = sh.getAttribute('r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let target = relMap[rid];
    if (!target) continue;
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const doc = await xml(target);
    if (!doc) continue;
    const rows = [];
    for (const rowEl of doc.getElementsByTagName('row')) {
      const r = (+rowEl.getAttribute('r') || rows.length + 1) - 1;
      let nextCol = 0;
      for (const c of rowEl.getElementsByTagName('c')) {
        const ref = c.getAttribute('r');
        const col = ref ? colIndex(ref.replace(/\d+/g, '')) : nextCol;
        nextCol = col + 1;
        const t = c.getAttribute('t');
        let v = null;
        if (t === 's') { const vv = c.getElementsByTagName('v')[0]; v = vv ? shared[+vv.textContent] : null; }
        else if (t === 'inlineStr') { const is = c.getElementsByTagName('is')[0]; v = is ? textOf(is) : null; }
        else { const vv = c.getElementsByTagName('v')[0]; v = vv ? vv.textContent : null; }
        if (v == null || String(v).trim() === '') continue;
        (rows[r] ||= [])[col] = String(v);
      }
    }
    sheets.push({ name: sh.getAttribute('name'), hidden: (sh.getAttribute('state') || 'visible') !== 'visible', rows });
  }
  return sheets;
}

// ---------------------------------------------------------------------------
// 3. Understand the TaRL layout of each sheet.
// ---------------------------------------------------------------------------
function findCell(rows, test, fromRow = 0, toRow = rows.length) {
  for (let r = fromRow; r < Math.min(toRow, rows.length); r++) {
    const row = rows[r]; if (!row) continue;
    for (let c = 0; c < row.length; c++) if (row[c] != null && test(row[c])) return { r, c };
  }
  return null;
}
function valueRightOf(rows, label) {
  const hit = findCell(rows, v => norm(v) === norm(label));
  if (!hit) return '';
  const row = rows[hit.r];
  for (let c = hit.c + 1; c < row.length; c++) if (row[c] != null) return String(row[c]).trim();
  return '';
}
function parseSheet(sheet) {
  const rows = sheet.rows;
  const hdr = findCell(rows, v => norm(v) === norm('رقم التلميذ'));
  if (!hdr) return null;
  const cols = {};
  for (const s of SUBJECTS) {
    let pos = null;
    for (const r of [hdr.r - 1, hdr.r - 2]) {
      const row = rows[r]; if (!row) continue;
      const c = row.findIndex(v => v != null && s.match.test(v));
      if (c >= 0) { pos = c; break; }
    }
    if (pos == null) pos = { ar: hdr.c + 2, fr: hdr.c + 4, ma: hdr.c + 6 }[s.key];
    let pre = pos, post = pos + 1;
    const h = rows[hdr.r] || [];
    if (/البعدي/.test(h[pre] || '') && /القبلي/.test(h[post] || '')) [pre, post] = [post, pre];
    cols[s.key] = { pre, post };
  }
  const expected = {};
  const er = findCell(rows, v => /المستوى المتوقع/.test(v), 0, hdr.r);
  if (er) {
    const row = rows[er.r] || [], next = rows[er.r + 1] || [];
    for (const s of SUBJECTS) {
      const c = row.findIndex(v => v != null && s.match.test(v));
      if (c >= 0 && next[c] != null) expected[s.key] = next[c];
    }
  }
  const students = [];
  for (let r = hdr.r + 1; r < rows.length; r++) {
    const row = rows[r]; if (!row) continue;
    const st = { sheet: sheet.name, id: row[hdr.c] || '', name: (row[hdr.c + 1] || '').replace(/\s+/g, ' ').trim() };
    let any = false;
    for (const s of SUBJECTS) {
      st[s.key + '_pre'] = row[cols[s.key].pre] ?? null;
      st[s.key + '_post'] = row[cols[s.key].post] ?? null;
      if (st[s.key + '_pre'] != null || st[s.key + '_post'] != null) any = true;
    }
    if (any) students.push(st);
  }
  const meta = {
    academy: valueRightOf(rows, 'الأكاديمية'),
    direction: valueRightOf(rows, 'المديرية الإقليمية'),
    school: valueRightOf(rows, 'المؤسسة'),
    year: valueRightOf(rows, 'السنة الدراسية'),
    level: valueRightOf(rows, 'المستوى'),
  };
  return { name: sheet.name, students, expected, meta };
}

// ---------------------------------------------------------------------------
// 4. Statistics (same definitions as the original report).
// ---------------------------------------------------------------------------
function analyse(students, key, expectedRaw, warnings) {
  const list = LEVELS[key], idx = LEVEL_INDEX[key];
  const unknown = new Map();
  const rank = v => {
    if (v == null) return null;
    const i = idx.get(norm(v));
    if (i === undefined) { unknown.set(String(v).trim(), (unknown.get(String(v).trim()) || 0) + 1); return null; }
    return i;
  };
  const pre = students.map(s => rank(s[key + '_pre']));
  const post = students.map(s => rank(s[key + '_post']));
  const title = SUBJECTS.find(s => s.key === key).title;
  for (const [v, n] of unknown) warnings.push(`${title}: المستوى «${v}» غير معروف (${n} مرة) ولم يُحتسب.`);

  const present = new Set([...pre, ...post].filter(v => v != null));
  let target = expectedRaw != null ? idx.get(norm(expectedRaw)) : undefined;
  if (target === undefined) {
    target = present.size ? Math.max(...present) : list.length - 1;
    warnings.push(`${title}: لم يُعثر على المستوى المتوقع في الملف، فاعتُمد أعلى مستوى موجود («${list[target].label}»).`);
  }
  present.add(target);
  const shown = [...present].sort((a, b) => a - b);

  const nPre = pre.filter(v => v != null).length;
  const nPost = post.filter(v => v != null).length;
  const dist = (arr, n) => shown.map(i => { const c = arr.filter(v => v === i).length; return { count: c, pct: n ? 100 * c / n : 0 }; });
  const pairs = pre.map((a, i) => [a, post[i]]).filter(([a, b]) => a != null && b != null);
  const below = pairs.filter(([a]) => a < target);
  const upBelow = below.filter(([a, b]) => b > a).length;
  const jump2 = pairs.filter(([a, b]) => b - a >= 2).length;
  const rows = [];
  for (const i of shown) {
    if (i >= target) continue;
    const from = pairs.filter(([a]) => a === i);
    if (!from.length) continue;
    const up = from.filter(([, b]) => b > i);
    const dest = new Map();
    up.forEach(([, b]) => dest.set(b, (dest.get(b) || 0) + 1));
    rows.push({ level: i, n: from.length, up: up.length, dest: [...dest.entries()].sort((x, y) => x[0] - y[0]) });
  }
  return {
    key, title, list, shown, target,
    nPre, nPost,
    preAt: pre.filter(v => v === target).length, postAt: post.filter(v => v === target).length,
    distPre: dist(pre, nPre), distPost: dist(post, nPost),
    below: below.length, upBelow, jump2, rows,
  };
}

// ---------------------------------------------------------------------------
// 5. Charts drawn on a canvas (pre = orange, post = blue).
// ---------------------------------------------------------------------------
const FONT_STACK = 'Arial, "Segoe UI", Tahoma, "Noto Sans Arabic", "DejaVu Sans", sans-serif';
const isArabic = s => /[؀-ۿ]/.test(s);
function chartLabel(an, i) {
  const l = an.list[i].label;
  if (an.key !== 'fr') return [l];
  const parts = l.split(' ');
  return parts.length > 1 ? [parts[0], parts.slice(1).join(' ')] : [l];
}
function drawChart(an) {
  const W = 1125, H = an.key === 'fr' ? 540 : 486;
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const g = cv.getContext('2d');
  g.fillStyle = '#ffffff'; g.fillRect(0, 0, W, H);
  const text = (s, x, y, size, color, align = 'center', weight = 'normal') => {
    g.font = `${weight} ${size}px ${FONT_STACK}`; g.fillStyle = color; g.textAlign = align; g.textBaseline = 'alphabetic';
    g.direction = isArabic(s) ? 'rtl' : 'ltr';
    g.fillText(s, x, y);
  };
  const n = an.shown.length;
  const left = 22, right = W - 10, bandTop = 66, top = 142, base = H - (an.key === 'fr' ? 110 : 72);
  const slot = (right - left) / n, bw = slot * 0.4;
  const maxV = Math.max(1, ...an.distPre.map(d => d.pct), ...an.distPost.map(d => d.pct));
  const scale = v => (base - top) * v / maxV;
  // expected-level band
  const ti = an.shown.indexOf(an.target);
  g.fillStyle = '#f0efec'; g.fillRect(left + slot * ti, bandTop, slot, base - bandTop);
  text('المستوى المتوقع', left + slot * (ti + 0.5), bandTop + 28, 23, '#52514e');
  const lab = an.key === 'fr' && n > 5 ? 22 : 25;
  an.shown.forEach((lvl, k) => {
    const cx = left + slot * (k + 0.5);
    [[an.distPre[k].pct, COLOR_PRE, -1], [an.distPost[k].pct, COLOR_POST, 1]].forEach(([v, col, side]) => {
      const x = side < 0 ? cx - bw : cx, h = scale(v);
      g.fillStyle = col; g.fillRect(x + 1.5, base - h, bw - 3, h);
      text(`${Math.round(v)}%`, x + bw / 2, base - h - 9, lab, '#1f1f1e');
    });
    chartLabel(an, lvl).forEach((line, j) => text(line, cx, base + 34 + j * 30, an.key === 'fr' ? 23 : 26, '#52514e'));
  });
  g.strokeStyle = '#c9c8c2'; g.lineWidth = 2; g.beginPath(); g.moveTo(left, base + 0.5); g.lineTo(right, base + 0.5); g.stroke();
  // legend (marker on the right of each label, as read in Arabic)
  const legend = [['مستوى التموضع القبلي', COLOR_PRE], ['مستوى التموضع البعدي', COLOR_POST]];
  let x = 30;
  g.font = `normal 25px ${FONT_STACK}`;
  for (const [label, col] of legend) {
    const w = g.measureText(label).width;
    text(label, x + w, 40, 25, '#1f1f1e', 'right');
    g.fillStyle = col; g.fillRect(x + w + 10, 24, 20, 20);
    x += w + 80;
  }
  return cv;
}
const canvasPng = cv => new Promise(res => cv.toBlob(b => b.arrayBuffer().then(a => res(new Uint8Array(a))), 'image/png'));

// ---------------------------------------------------------------------------
// 6. Word document (docx library).
// ---------------------------------------------------------------------------
const FONT = 'Arial';
const LTR_RE = /(?:[A-Za-zÀ-ÿ0-9(+]|-(?=\d))[A-Za-zÀ-ÿ0-9.,%()+\-\/:'’ ]*[A-Za-zÀ-ÿ0-9%)]|[A-Za-zÀ-ÿ0-9%]/g;
const LRM = '‎';
function balance(seg) {
  let lead = '', trail = '';
  const cnt = (t, ch) => t.split(ch).length - 1;
  while (seg.endsWith(')') && cnt(seg, '(') < cnt(seg, ')')) { trail = ')' + trail; seg = seg.slice(0, -1).trimEnd(); }
  while (seg.startsWith('(') && cnt(seg, '(') > cnt(seg, ')')) { lead += '('; seg = seg.slice(1).trimStart(); }
  return [lead, seg, trail];
}
// Split mixed Arabic / Latin text into runs with the right direction.
function segRuns(text, o = {}) {
  const { TextRun } = docx;
  const mk = (t, rtl) => new TextRun({ text: t, rightToLeft: rtl, bold: o.bold, color: o.color || '1F1F1E', size: o.size || 16, font: FONT });
  const runs = []; let last = 0, m;
  LTR_RE.lastIndex = 0;
  while ((m = LTR_RE.exec(text)) !== null) {
    if (m.index > last) runs.push(mk(text.slice(last, m.index), true));
    const [lead, core, trail] = balance(m[0]);
    if (lead) runs.push(mk(lead, true));
    if (core) runs.push(mk(/^[(+\-]|[)]$/.test(core) ? LRM + core + LRM : core, false));
    if (trail) runs.push(mk(trail, true));
    last = m.index + m[0].length;
  }
  if (last < text.length) runs.push(mk(text.slice(last), true));
  return runs;
}

async function buildDocx(meta, analyses, nPre, nPost, classesNote) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
    AlignmentType, VerticalAlign, ImageRun, HeightRule } = docx;
  const HEAD = '1C5CAB';
  const CW = 10466;
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noB = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
  const thin = c => ({ style: BorderStyle.SINGLE, size: 4, color: c || 'C9C8C2' });
  const B = c => ({ top: thin(c), bottom: thin(c), left: thin(c), right: thin(c) });
  const pct = (a, b) => `${b ? (100 * a / b).toFixed(1) : '0.0'}%`;
  const greens = ['F1F8F1', 'DDF0DD', 'C3E6C3', 'A6DAA6'];
  const shade = p => greens[p >= 75 ? 3 : p >= 60 ? 2 : p >= 45 ? 1 : 0];

  const para = (text, o = {}) => new Paragraph({
    bidirectional: !o.ltr, alignment: o.align ?? AlignmentType.CENTER,
    spacing: { before: o.before || 0, after: o.after || 0, line: o.line || 240 },
    children: o.ltr ? [new TextRun({ text, font: FONT, size: o.size || 16, bold: o.bold, color: o.color || '1F1F1E' })]
                    : segRuns(text, { size: o.size || 16, bold: o.bold, color: o.color }),
  });
  const cell = (text, w, o = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: B(), verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 30, bottom: 30, left: 60, right: 60 },
    children: [para(String(text), o)],
  });
  const grid = (header, rows, widths, o = {}) => new Table({
    visuallyRightToLeft: true, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { fill: o.head || HEAD, color: 'FFFFFF', bold: true, size: 14 })) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((x, i) => {
        const c = typeof x === 'object' ? x : { text: x };
        return cell(c.text, widths[i], { size: o.size || 15, fill: c.fill || (ri % 2 ? 'F6F5F2' : undefined), bold: c.bold, ltr: c.ltr, align: c.align });
      }) })),
    ],
  });
  const tiny = () => new Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: 'exact' }, children: [new TextRun({ text: '', size: 2 })] });

  const body = [];
  body.push(para(`نتائج روائز التموضع TaRL${meta.level ? ' – المستوى ' + meta.level : ''}`, { size: 28, bold: true, color: '104281', after: 10 }));
  const city = (meta.direction.split(/[:：]/)[1] || '').trim();
  const sub = [
    [meta.school, city].filter(Boolean).join(' – '),
    meta.year ? `السنة الدراسية ${meta.year}` : '',
    `التموضع القبلي: ${nPre} تلميذاً`, `التموضع البعدي: ${nPost} تلميذاً`,
  ].filter(Boolean).join('  |  ');
  body.push(para(sub, { size: 16, color: '52514E', after: classesNote ? 0 : 100 }));
  if (classesNote) body.push(para(classesNote, { size: 15, color: '52514E', after: 100 }));

  // summary table
  const sumRows = analyses.map(an => {
    const S = SUBJECTS.find(s => s.key === an.key);
    const rate = an.below ? 100 * an.upBelow / an.below : 0;
    return [
      { text: an.title, bold: true, fill: S.fill },
      { text: an.list[an.target].label, ltr: an.key === 'fr' },
      { text: `${an.preAt} (${pct(an.preAt, an.nPre)})`, ltr: true },
      { text: `${an.postAt} (${pct(an.postAt, an.nPost)})`, ltr: true, bold: true },
      { text: pct(an.upBelow, an.below), ltr: true, bold: true, fill: shade(rate) },
      { text: String(an.jump2), ltr: true },
    ];
  });
  body.push(grid(
    ['المادة', 'المستوى المتوقع', 'عدد ونسبة التلاميذ في المستوى المتوقع (التموضع القبلي)', 'عدد ونسبة التلاميذ في المستوى المتوقع (التموضع البعدي)', 'نسبة الانتقال لدى من كانوا دون المستوى المتوقع', 'انتقلوا بمستويين فأكثر'],
    sumRows, [1300, 1100, 2200, 2200, 2266, 1400]));

  // one band per subject
  const CHART_W = 4800, TAB_W = CW - CHART_W;
  const dot = { style: BorderStyle.DOTTED, size: 6, color: '9A9994' };
  for (const an of analyses) {
    const S = SUBJECTS.find(s => s.key === an.key);
    const fr = an.key === 'fr';
    const rows = an.rows.map(r => {
      const rate = 100 * r.up / r.n;
      const dest = r.dest.map(([lv, c]) => `${fr ? (an.list[lv].short || an.list[lv].label) : an.list[lv].label} (${c})`).join(fr ? ', ' : '، ');
      return [
        { text: an.list[r.level].label, ltr: fr, bold: true },
        { text: String(r.n), ltr: true }, { text: String(r.up), ltr: true },
        { text: rate.toFixed(1) + '%', ltr: true, bold: true, fill: shade(rate) },
        { text: dest || '—', ltr: fr, align: fr ? AlignmentType.LEFT : AlignmentType.START },
      ];
    });
    rows.push([{ text: 'المجموع', bold: true }, { text: String(an.below), ltr: true, bold: true }, { text: String(an.upBelow), ltr: true, bold: true },
      { text: pct(an.upBelow, an.below), ltr: true, bold: true, fill: shade(an.below ? 100 * an.upBelow / an.below : 0) }, { text: '' }]);
    const tw = [1150, 520, 900, 700, TAB_W - 3270 - 140];
    const tab = grid(['مستوى التموضع القبلي', 'العدد', 'انتقلوا إلى مستوى أعلى', 'النسبة', 'مستوى التموضع البعدي لمن انتقلوا'], rows, tw, { size: fr ? 13 : 14, head: S.color });

    body.push(new Paragraph({
      bidirectional: true, spacing: { before: 60, after: 30 }, keepNext: true,
      shading: { type: ShadingType.CLEAR, fill: S.fill, color: 'auto' },
      border: { right: { style: BorderStyle.SINGLE, size: 24, color: S.color, space: 6 } },
      children: segRuns(an.title, { size: 21, bold: true, color: S.color }),
    }));
    body.push(new Table({
      visuallyRightToLeft: true, width: { size: CW, type: WidthType.DXA }, columnWidths: [CHART_W, TAB_W], borders: noB,
      rows: [new TableRow({ cantSplit: true, children: [
        new TableCell({ width: { size: CHART_W, type: WidthType.DXA }, borders: noB, verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: an.png, transformation: { width: 314, height: fr ? 151 : 136 } })] })] }),
        new TableCell({ width: { size: TAB_W, type: WidthType.DXA }, borders: noB, verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 140 },
          children: [tab, tiny()] }),
      ] })],
    }));
    body.push(new Table({
      visuallyRightToLeft: true, width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], borders: noB,
      rows: [0, 1, 2, 3].map(() => new TableRow({ height: { value: 285, rule: HeightRule.EXACT }, children: [new TableCell({
        width: { size: CW, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: dot }, children: [new Paragraph({ children: [] })],
      })] })),
    }));
  }

  // general summary with writing lines
  body.push(new Paragraph({ spacing: { before: 60, after: 20 }, bidirectional: true, keepNext: true, children: segRuns('خلاصة عامة:', { size: 18, bold: true, color: '104281' }) }));
  const dash = { style: BorderStyle.DASHED, size: 4, color: HEAD };
  body.push(new Table({
    visuallyRightToLeft: true, width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: [0, 1, 2].map(i => new TableRow({ height: { value: 285, rule: HeightRule.EXACT }, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA }, borders: { top: i === 0 ? dash : dot, bottom: i === 2 ? dash : dot, left: dash, right: dash },
      children: [new Paragraph({ children: [] })],
    })] })),
  }));
  body.push(tiny());

  const doc = new Document({
    creator: meta.school || 'TaRL', title: 'نتائج روائز التموضع TaRL',
    styles: { default: { document: { run: { font: FONT, size: 16 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 380, bottom: 340, left: 720, right: 720 } } }, children: body }],
  });
  return Packer.toBlob(doc);
}

// ---------------------------------------------------------------------------
// 7. Page wiring.
// ---------------------------------------------------------------------------
const $ = s => document.querySelector(s);
let parsed = null;     // [{name, students, expected, meta}]
let lastBlob = null, lastName = 'TaRL.docx';

function setStatus(msg, kind = '') { const el = $('#status'); el.textContent = msg; el.className = 'status ' + kind; }
function escapeHtml(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

async function onFile(file) {
  if (!file) return;
  $('#results').hidden = true; $('#download').disabled = true; lastBlob = null;
  if (!/\.xlsx$/i.test(file.name)) { setStatus('يرجى اختيار ملف Excel بصيغة xlsx (إذا كان بصيغة xls افتحه في Excel واحفظه بصيغة xlsx).', 'error'); return; }
  setStatus('جاري قراءة الملف…');
  try {
    const sheets = await readXlsx(await file.arrayBuffer());
    parsed = sheets.filter(s => !s.hidden).map(parseSheet).filter(p => p && p.students.length);
    if (!parsed.length) throw new Error('لم يُعثر على جدول نتائج TaRL في الملف (عمود «رقم التلميذ»).');
    const box = $('#classes'); box.innerHTML = '';
    parsed.forEach((p, i) => {
      const id = 'cls' + i;
      box.insertAdjacentHTML('beforeend', `<label for="${id}"><input type="checkbox" id="${id}" data-i="${i}" checked> ${escapeHtml(p.name)} <span class="muted">(${p.students.length})</span></label>`);
    });
    $('#step2').hidden = false;
    setStatus(`تمت قراءة ${parsed.length} قسم/أقسام. اختر الأقسام ثم اضغط «إنشاء التقرير».`, 'ok');
  } catch (e) {
    console.error(e); setStatus('تعذّرت قراءة الملف: ' + e.message, 'error');
  }
}

async function generate() {
  const chosen = [...document.querySelectorAll('#classes input:checked')].map(i => parsed[+i.dataset.i]);
  if (!chosen.length) { setStatus('اختر قسماً واحداً على الأقل.', 'error'); return; }
  setStatus('جاري التحليل وإنشاء ملف Word…');
  try {
    const students = chosen.flatMap(p => p.students);
    const meta = chosen[0].meta;
    const expected = chosen.find(p => Object.keys(p.expected).length)?.expected || {};
    const warnings = [];
    const analyses = SUBJECTS.map(s => analyse(students, s.key, expected[s.key], warnings));
    const charts = $('#charts'); charts.innerHTML = '';
    for (const an of analyses) {
      const cv = drawChart(an);
      an.png = await canvasPng(cv);
      const fig = document.createElement('figure');
      fig.innerHTML = `<figcaption>${escapeHtml(an.title)}</figcaption>`; fig.appendChild(cv);
      charts.appendChild(fig);
    }
    const nPre = students.filter(s => SUBJECTS.some(x => s[x.key + '_pre'] != null)).length;
    const nPost = students.filter(s => SUBJECTS.some(x => s[x.key + '_post'] != null)).length;
    const classesNote = chosen.length < parsed.length ? 'الأقسام: ' + chosen.map(p => p.name).join('، ') : '';
    renderSummary(analyses);
    $('#warnings').innerHTML = warnings.map(w => `<li>${escapeHtml(w)}</li>`).join('');
    $('#warnbox').hidden = !warnings.length;
    lastBlob = await buildDocx(meta, analyses, nPre, nPost, classesNote);
    const tag = (meta.level || 'TaRL').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, '_');
    lastName = `ملخص_TaRL_${tag}${chosen.length < parsed.length ? '_' + chosen.map(p => p.name).join('_') : ''}.docx`;
    $('#results').hidden = false; $('#download').disabled = false;
    setStatus('التقرير جاهز. اضغط «تحميل ملف Word».', 'ok');
  } catch (e) {
    console.error(e); setStatus('حدث خطأ أثناء إنشاء التقرير: ' + e.message, 'error');
  }
}

function renderSummary(analyses) {
  const pct = (a, b) => `${b ? (100 * a / b).toFixed(1) : '0.0'}%`;
  $('#summary tbody').innerHTML = analyses.map(an => `<tr>
    <th>${escapeHtml(an.title)}</th><td>${escapeHtml(an.list[an.target].label)}</td>
    <td>${an.preAt} (${pct(an.preAt, an.nPre)})</td><td><b>${an.postAt} (${pct(an.postAt, an.nPost)})</b></td>
    <td><b>${pct(an.upBelow, an.below)}</b> <span class="muted">(${an.upBelow} من ${an.below})</span></td><td>${an.jump2}</td></tr>`).join('');
}

function download() {
  if (!lastBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(lastBlob); a.download = lastName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

window.addEventListener('DOMContentLoaded', () => {
  const input = $('#file'), drop = $('#drop');
  input.addEventListener('change', () => onFile(input.files[0]));
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => onFile(e.dataTransfer.files[0]));
  $('#generate').addEventListener('click', generate);
  $('#download').addEventListener('click', download);
});
