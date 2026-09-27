// Ережелердің 4.5-тармағына сай ортақ ресімдеу: A4, Times New Roman 14, 1,5 интервал,
// абзац шегінісі 1,25 см, жиектер: сол 3 см, қалғандары 2 см, бет нөмірі төменде ортада.
const fs = require("fs");
const path = require("path");
const D = require("docx");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType,
  WidthType, BorderStyle, ShadingType, Footer, PageNumber, PageBreak, HeadingLevel, TabStopType,
  VerticalAlign, PageOrientation, LevelFormat,
} = D;

const FONT = "Times New Roman";
const CM = 567;
const SZ = 28;      // 14 pt
const SZ_T = 24;    // 12 pt кестелерде (кемінде 10 pt)
const A4 = { w: 11906, h: 16838 };
const MARGIN = { left: 3 * CM, right: 2 * CM, top: 2 * CM, bottom: 2 * CM };
const LINE = 360;   // 1,5 интервал
const FIG = path.join(__dirname, "figs");

/* Жолішілік белгілеу:
   **қалың**, _курсив_, [___] – оқушы толтыратын орын (сары фон),
   ==мәтін== – оқушы өз сөзімен тексеретін/түзететін орын (сары фон). */
function runs(text, o = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|==[^=]+==|\[[^\]]*\]|_(?=[^\s_])[^_]+?(?<=[^\s_])_)/g;
  let last = 0, m;
  const base = { font: FONT, size: o.size || SZ, bold: o.bold, italics: o.italics, color: o.color };
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ ...base, text: text.slice(last, m.index) }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ ...base, text: t.slice(2, -2), bold: true }));
    else if (t.startsWith("==")) out.push(new TextRun({ ...base, text: t.slice(2, -2), highlight: "yellow" }));
    else if (t.startsWith("[")) {
      // [1], [5, 7] – дереккөзге сілтеме; қалғаны – толтырылатын орын
      // тек «_» бар жақша – толтырылатын орын; [1], [Электрондық ресурс] – қарапайым мәтін
      out.push(new TextRun({ ...base, text: t, highlight: t.includes("_") ? "yellow" : undefined }));
    } else out.push(new TextRun({ ...base, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ ...base, text: text.slice(last) }));
  return out;
}

const P = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  indent: o.noIndent || o.align === AlignmentType.CENTER ? undefined : { firstLine: Math.round(1.25 * CM) },
  spacing: { line: o.line || LINE, before: o.before || 0, after: o.after || 0 },
  keepNext: o.keepNext,
  children: runs(text, o),
});
const C = (text, o = {}) => P(text, { ...o, align: AlignmentType.CENTER });
const L = (text, o = {}) => P(text, { ...o, align: AlignmentType.LEFT, noIndent: true });
const R = (text, o = {}) => P(text, { ...o, align: AlignmentType.RIGHT, noIndent: true });
const BR = () => new Paragraph({ children: [new PageBreak()] });
const EMPTY = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { line: LINE }, children: [new TextRun({ text: "", font: FONT, size: SZ })] }));

// Тізім: «–» белгісімен немесе нөмірмен
const bullet = (text, o = {}) => new Paragraph({
  numbering: { reference: o.num ? "num" + o.num : "dash", level: 0 },
  alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE },
  children: runs(text, o),
});
const NUMBERING = {
  config: [
    { reference: "dash", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: Math.round(1.25 * CM) + 360, hanging: 360 } } } }] },
    ...Array.from({ length: 12 }, (_, i) => ({ reference: "num" + (i + 1), levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(1.25 * CM) + 400, hanging: 400 } } } }] })),
  ],
};

const H1 = (text, o = {}) => new Paragraph({
  heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, pageBreakBefore: o.newPage !== false,
  spacing: { line: LINE, after: 240 }, keepNext: true,
  children: [new TextRun({ text, font: FONT, size: SZ, bold: true, allCaps: o.caps !== false })],
});
const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2, alignment: AlignmentType.LEFT, indent: { firstLine: Math.round(1.25 * CM) },
  spacing: { line: LINE, before: 240, after: 120 }, keepNext: true,
  children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
});

// Кесте
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const BORDERS = { top: border, bottom: border, left: border, right: border };
function table(widthsPct, header, rows, o = {}) {
  const total = o.width || (A4.w - MARGIN.left - MARGIN.right);
  const widths = widthsPct.map(p => Math.round(total * p / 100));
  widths[widths.length - 1] = total - widths.slice(0, -1).reduce((a, b) => a + b, 0);
  const size = o.size || SZ_T;
  const cell = (txt, i, head, shade) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders: BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    shading: head || shade ? { type: ShadingType.CLEAR, color: "auto", fill: head ? "E7EEF3" : shade } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: String(txt).split("\n").map(line => new Paragraph({
      alignment: head || (o.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { line: o.line || 240 },
      children: runs(line, { size, bold: head || undefined }),
    })),
  });
  const trs = [];
  if (header) trs.push(new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, i, true)) }));
  rows.forEach(r => trs.push(new TableRow({ cantSplit: true, height: o.rowHeight ? { value: o.rowHeight, rule: "atLeast" } : undefined,
    children: r.map((c, i) => cell(c, i, false, o.shadeFirst && i === 0 ? "F4F7F9" : undefined)) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
const tcap = (n, text) => new Paragraph({ alignment: AlignmentType.LEFT, keepNext: true, spacing: { line: LINE, before: 120, after: 60 },
  children: runs(`${n}-кесте – ${text}`, {}) });
const src = (text = "Дереккөзі: автордың жұмысы") => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 276, before: 60, after: 160 },
  children: runs(text, { size: 24, italics: true }) });

// Сурет
function img(file, widthCm, o = {}) {
  const buf = fs.readFileSync(path.join(FIG, file));
  const { w, h } = pngSize(buf);
  const wpx = Math.round(widthCm / 2.54 * 96);
  return new ImageRun({ type: "png", data: buf, transformation: { width: wpx, height: Math.round(wpx * h / w) },
    altText: o.alt ? { title: o.alt, description: o.alt, name: file } : undefined });
}
function pngSize(buf) { return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; }
const figure = (file, widthCm, n, caption, source) => [
  new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120 }, children: [img(file, widthCm, { alt: caption })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 0 }, keepNext: !!source,
    children: runs(`${n}-сурет – ${caption}`, {}) }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 276, after: 200 },
    children: runs(source || "Дереккөзі: автордың жұмысы, «Сөздер лабиринті» ойынының скриншоты", { size: 24, italics: true }) }),
];
// Бірнеше скриншот бір қатарда
function figRow(files, widthCm, n, caption, source) {
  const total = A4.w - MARGIN.left - MARGIN.right, w = Math.floor(total / files.length);
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const nb = { top: none, bottom: none, left: none, right: none };
  return [
    new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: files.map(() => w), rows: [new TableRow({ cantSplit: true,
      children: files.map(([f, sub]) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: nb, children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [img(f, widthCm, { alt: sub })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: runs(sub, { size: 24 }) }),
      ] })) })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, before: 120 }, keepNext: true, children: runs(`${n}-сурет – ${caption}`, {}) }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 276, after: 200 },
      children: runs(source || "Дереккөзі: автордың жұмысы, «Сөздер лабиринті» ойынының скриншоттары", { size: 24, italics: true }) }),
  ];
}

// Мазмұны: нүктелі толтырғыш + бет нөмірі
const tocLine = (text, page, level = 1) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: A4.w - MARGIN.left - MARGIN.right, leader: "dot" }],
  indent: level === 2 ? { left: 400 } : undefined, spacing: { line: LINE },
  children: [new TextRun({ text, font: FONT, size: SZ, bold: level === 1 && /^[А-ЯӘҒҚҢӨҰҮҺІ\s]+$/.test(text) }),
    new TextRun({ text: "\t" + (page ?? ""), font: FONT, size: SZ })],
});

// Құжатты жинау: titlePage – бірінші бетте нөмір жоқ
function doc(sections, o = {}) {
  const landscape = !!o.landscape;
  const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] });
  return new Document({
    creator: "Сөздер лабиринті жобасы", title: o.title || "Сөздер лабиринті",
    styles: { default: { document: { run: { font: FONT, size: SZ } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 1 } },
      ] },
    numbering: NUMBERING,
    sections: [{
      properties: {
        titlePage: o.titlePage !== false,
        page: { size: landscape ? { width: A4.w, height: A4.h, orientation: PageOrientation.LANDSCAPE } : { width: A4.w, height: A4.h },
          margin: MARGIN },
      },
      footers: { default: footer },
      children: sections,
    }],
  });
}
async function save(d, file) {
  const buf = await Packer.toBuffer(d);
  fs.writeFileSync(file, buf);
  console.log("wrote", file);
}

// Қазылар алқасына берілетін данадағы бас бөлік (4-, 6-, 7-қосымшалар үлгісі)
const PROJECT = {
  title: "СӨЗДЕР ЛАБИРИНТІ: қазақ тілінен интерактивті ойын құрастыру",
  direction: "Қоғамдық-гуманитарлық бағыт (ҚГБ)",
  section: "Қазақ тілі және әдебиеті",
  category: "☒ Бастауыш буын (2–4 сынып)   ☐ Орта буын (5–8 сынып)",
  type: "☒ Жеке   ☐ Командалық",
};

module.exports = { D, FONT, CM, SZ, SZ_T, A4, MARGIN, LINE, runs, P, C, L, R, BR, EMPTY, bullet, H1, H2, table, tcap, src,
  img, figure, figRow, tocLine, doc, save, PROJECT, AlignmentType };
