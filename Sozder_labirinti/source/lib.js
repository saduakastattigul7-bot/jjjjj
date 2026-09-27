// Барлық құжатқа ортақ пішімдеу функциялары («Зерде» ережелерінің 4.5-тармағы бойынша)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType,
  BorderStyle, ShadingType, VerticalAlign, Footer, PageNumber, HeadingLevel, LevelFormat, ImageRun,
  TabStopType,
} = require("docx");

const F = "Times New Roman";
const S = 28;            // 14 кегль
const ST = 24;           // кестеде 12 кегль (≥10)
const PAGE_W = 11906, ML = 1701, MR = 1134;   // сол жақ 3 см, оң жақ 2 см
const TW = PAGE_W - ML - MR;
const C = AlignmentType.CENTER, J = AlignmentType.JUSTIFIED, LEFT = AlignmentType.LEFT;
const LINE = 360;        // 1,5 жоларалық интервал
const IND = 709;         // абзацтық шегініс 1,25 см

// {{...}} — оқушы/жетекші өзі толтыратын немесе тексеретін жер (сары түспен белгіленеді)
const HL = /(\{\{[^}]*\}\})/;
function runs(t, o = {}) {
  return String(t).split(HL).filter(p => p !== "").map(p => {
    const hl = HL.test(p);
    return new TextRun({
      text: hl ? p.slice(2, -2) : p, font: F, size: o.size || S, bold: o.bold, italics: o.italics,
      highlight: hl || o.hl ? "yellow" : undefined, underline: o.underline ? {} : undefined,
    });
  });
}
function P(t, o = {}) {
  return new Paragraph({
    alignment: o.align ?? J,
    indent: o.noInd ? o.indent : { firstLine: IND, ...(o.indent || {}) },
    spacing: { line: o.line ?? LINE, before: o.before || 0, after: o.after || 0 },
    keepNext: o.keepNext, keepLines: o.keepLines, pageBreakBefore: o.pb,
    children: o.children || runs(t, o),
  });
}
const Pc = (t, o = {}) => P(t, { align: C, noInd: true, ...o });
const Pl = (t, o = {}) => P(t, { align: LEFT, noInd: true, ...o });
function H1(t, o = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, alignment: C, pageBreakBefore: o.pb ?? true,
    spacing: { line: LINE, after: 240 }, keepNext: true,
    children: [new TextRun({ text: t, font: F, size: S, bold: true, color: "000000" })],
  });
}
function H2(t) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, alignment: LEFT, indent: { firstLine: IND },
    spacing: { line: LINE, before: 240, after: 120 }, keepNext: true,
    children: [new TextRun({ text: t, font: F, size: S, bold: true, color: "000000" })],
  });
}
function bullet(t, o = {}) {
  return new Paragraph({
    numbering: { reference: o.num ? "num" + o.num : "dash", level: 0 }, alignment: J,
    spacing: { line: o.line ?? LINE }, children: runs(t, o),
  });
}
const numbering = {
  config: [
    { reference: "dash", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: LEFT,
      style: { paragraph: { indent: { left: IND + 360, hanging: 360 } } } }] },
    ...[1, 2, 3, 4, 5, 6, 7, 8].map(i => ({ reference: "num" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL,
      text: "%1)", alignment: LEFT, style: { paragraph: { indent: { left: IND + 360, hanging: 360 } } } }] })),
  ],
};

const b = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: b, bottom: b, left: b, right: b };
function table(pct, header, rows, o = {}) {
  const tw = o.tw || TW;
  const w = pct.map(p => Math.floor(tw * p / 100));
  w[w.length - 1] += tw - w.reduce((a, c) => a + c, 0);
  const size = o.size || ST;
  const cell = (t, i, head) => new TableCell({
    borders, width: { size: w[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: head ? { fill: "E3EAF3", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: String(t).split("\n").map(line => new Paragraph({
      alignment: head || o.center?.includes(i) ? C : LEFT,
      spacing: { line: 240 }, children: runs(line, { size, bold: head || (o.boldCol === i) }) })),
  });
  const rowsOut = [];
  if (header) rowsOut.push(new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }));
  rows.forEach(r => rowsOut.push(new TableRow({ cantSplit: true,
    height: o.minH ? { value: o.minH, rule: "atLeast" } : undefined, children: r.map((c, i) => cell(c, i, false)) })));
  return new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: w, rows: rowsOut });
}
function caption(t) { return P(t, { align: C, noInd: true, size: 24, after: 200, italics: false }); }
function tableTitle(t) { return P(t, { align: LEFT, noInd: true, keepNext: true, before: 120, size: 28 }); }

function pngSize(file) {
  const buf = fs.readFileSync(file);
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20), buf };
}
function image(file, widthPx = 560) {
  const { w, h, buf } = pngSize(file);
  const height = Math.round(widthPx * h / w);
  return new Paragraph({ alignment: C, keepNext: true, spacing: { before: 120, after: 60 },
    children: [new ImageRun({ type: "png", data: buf, transformation: { width: widthPx, height } })] });
}
const gap = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

// Мазмұны: ұпай жолы (dot leader) және бет нөмірі
function tocLine(t, page, level = 1) {
  return new Paragraph({
    indent: { left: level === 2 ? 400 : 0 }, spacing: { line: LINE },
    tabStops: [{ type: TabStopType.RIGHT, position: TW, leader: "dot" }],
    children: [...runs(t), new TextRun({ text: "\t" + String(page ?? ""), font: F, size: S })],
  });
}

function footer() {
  return new Footer({ children: [new Paragraph({ alignment: C,
    children: [new TextRun({ font: F, size: 24, children: [PageNumber.CURRENT] })] })] });
}
function doc(children, o = {}) {
  return new Document({
    creator: "Сөздер лабиринті жобасы", title: o.title || "",
    styles: { default: { document: { run: { font: F, size: S } } } },
    numbering,
    sections: [{
      properties: {
        titlePage: !!o.titlePage,
        page: { size: { width: PAGE_W, height: 16838, orientation: o.landscape ? require("docx").PageOrientation.LANDSCAPE : undefined }, margin: { top: 1134, bottom: 1134, left: ML, right: MR } },
      },
      footers: { default: footer(), first: new Footer({ children: [new Paragraph({ children: [] })] }) },
      children,
    }],
  });
}
async function write(d, file) {
  const buf = await Packer.toBuffer(d);
  fs.writeFileSync(file, buf);
  return file;
}

// Ресми бланк басы (қосымшалардағыдай)
function formHeader(o) {
  const out = [];
  if (o.cipher !== false) out.push(Pl("Тіркеу коды (шифр): ______________________________"),
    Pl("(«Дарын» РҒПО қызметкері толтырады)", { size: 20, indent: { left: 3400 } }));
  if (o.direction) out.push(Pl("Бағыт: қоғамдық-гуманитарлық"));
  if (o.section) out.push(Pl("Секцияның атауы: Қазақ тілі және әдебиеті"));
  if (o.category) out.push(Pl("Қатысушылар санаты:"), Pl("☑ Бастауыш буын (2-4 сынып)"), Pl("☐ Орта буын (5-8 сынып)"));
  if (o.type) out.push(Pl("Жобаның түрі:"), Pl("☑ Жеке"), Pl("☐ Командалық"));
  out.push(Pl("Жобаның тақырыбы: «Сөздер лабиринті: қазақ тілінен интерактивті ойын құрастыру»", { after: 120 }));
  return out;
}

const TW_LAND = 16838 - ML - MR;
module.exports = { TW_LAND, F, S, ST, TW, C, J, LEFT, LINE, IND, runs, P, Pc, Pl, H1, H2, bullet, table, caption, tableTitle,
  image, gap, tocLine, doc, write, formHeader, TextRun, Paragraph, PageBreak: require("docx").PageBreak };
