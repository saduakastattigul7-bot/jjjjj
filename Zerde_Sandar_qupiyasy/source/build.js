// Негізгі жұмысты құрастырады: node build.js out.docx [--jury]
// --jury: қазылар алқасына арналған дана (титулда аты-жөндер жоқ, тек шифр).
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, Footer, PageNumber,
  HeadingLevel, TabStopType, LevelFormat, VerticalAlign,
} = require("docx");

const DIR = __dirname;
const OUT = process.argv[2] || path.join(DIR, "out.docx");
const JURY = process.argv.includes("--jury");
const D = JSON.parse(fs.readFileSync(path.join(DIR, "data.json"), "utf8"));
const N = D.natije || {};
const HAS = Object.keys(N).length > 0;
const TOC_PAGES = fs.existsSync(path.join(DIR, "toc_pages.json"))
  ? JSON.parse(fs.readFileSync(path.join(DIR, "toc_pages.json"), "utf8")) : {};

const FONT = "Times New Roman";
const SZ = 28, SZ_T = 24, SZ_A = 22;   // 14 pt; кестеде 12 pt; қосымшада 11 pt
const LINE = 360;                        // 1,5 жоларалық интервал
const CM = 567;
const PAGE_W = 11906, ML = 3 * CM, MR = 2 * CM, MT = 2 * CM, MB = 2 * CM;
const TEXT_W = PAGE_W - ML - MR;
const TITLE = "Қазақ ертегілеріндегі сандар құпиясы";
const PH = "[___]";

// ---------- Дереккөздер ----------
const REFS = {
  zerde: "2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсын ұйымдастыру және өткізу ережелері. – «Дарын» РҒПО директорының 2026 жылғы 9 қыркүйектегі № 153 бұйрығымен бекітілген.",
  birtutas: "Оқушыларды тәрбиелеудің «Біртұтас тәрбие» бағдарламасы. – Астана: ҚР Оқу-ағарту министрлігі, 2023.",
  ifrah: "Ifrah G. The Universal History of Numbers: From Prehistory to the Invention of the Computer. – New York: Wiley, 2000.",
  toporov: "Топоров В.Н. Числа // Мифы народов мира: энциклопедия: в 2 т. / гл. ред. С.А. Токарев. – М.: Советская энциклопедия, 1982. – Т. 2.",
  kaidar: "Қайдар Ә. Қазақтар ана тілі әлемінде: этнолингвистикалық сөздік: 3 томдық. – Алматы: Дайк-Пресс, 2009.",
  zhanpeisov: "Жанпейісов Е.Н. Этнокультурная лексика казахского языка. – Алма-Ата: Наука, 1989.",
  kaskabasov: "Қасқабасов С. Қазақтың халық прозасы. – Алматы: Ғылым, 1984.",
  olrik: "Olrik A. Epische Gesetze der Volksdichtung // Zeitschrift für deutsches Altertum und deutsche Literatur. – 1909. – Bd. 51. – S. 1–12.",
  propp: "Пропп В.Я. Морфология сказки. – Л.: Academia, 1928.",
};

// ---------- Деректерден алынатын мәндер ----------
const f1 = v => (v === null || v === undefined || v === "") ? PH : String(Math.round(v * 10) / 10).replace(".", ",");
const V = {};
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];
const timesWord = n => `${n} рет`;
if (HAS) {
  for (const k of ["tales_read", "total", "distinct", "sacred", "triad"]) V[k] = N[k];
  for (const k of ["per_tale", "sacred_pct", "F_pct", "sacred_F_pct", "nonsacred_F_pct"]) V[k] = f1(N[k]);
  const top = N.top.slice(0, 3);
  V.top_text = `Ең жиі кездескен сандар: ${listJoin(top.map(([n, c]) => `${n} (${timesWord(c)})`))}.`;
  const zero = D.numbers.filter(n => N.by_num[String(n)].all === 0);
  V.other_text = (zero.length ? `Ал ${listJoin(zero.map(String))} сандары бірде-бір ертегіде кездеспеді.` : "") +
    (N.other_list.length ? ` Кестеде бөлек көрсетілмеген сандар да кездесті: ${listJoin(N.other_list.map(String))}.` : "");
  const sacTop = top.filter(([n]) => D.sacred.includes(Number(n))).length;
  V.sacred_text = N.sacred_pct >= 50 ? `Яғни ертегілердегі әрбір екі санның бірінен көбі – «киелі» сан. Ал «киелі» сандар бес қана сан екенін ескерсек, бұл өте жоғары көрсеткіш.`
    : `Бұл «киелі» сандардың жиі кездесетінін көрсетеді, бірақ ертегілерде басқа сандар да аз емес.`;
  V.sacred_text += ` Ең жиі кездескен үш санның ${sacTop}-і «киелі» сандар қатарына жатады.`;
  V.func_text = N.sacred_F_pct > N.nonsacred_F_pct
    ? "Демек, «киелі» сандар басқа сандарға қарағанда символдық мағынада әлдеқайда жиі қолданылады: олар нақты мөлшерді емес, ерекше мағынаны береді."
    : "Күткенімізге қарамастан, «киелі» сандар басқа сандарға қарағанда символдық мағынада жиірек қолданылмады.";
  const cats = Object.entries(N.by_cat).sort((a, b) => b[1] - a[1]);
  V.cat_text = `Сандар көбіне ${cats[0][0].toLowerCase()} (${cats[0][1]}) және ${cats[1][0].toLowerCase()} (${cats[1][1]}) ұғымдарын білдіреді.`;
  const types = Object.entries(N.by_type).filter(([, o]) => o.tales > 0).map(([t, o]) => [t, o.n / o.tales, o.n ? o.F / o.n * 100 : 0]);
  const tMost = [...types].sort((a, b) => b[1] - a[1])[0], tF = [...types].sort((a, b) => b[2] - a[2])[0];
  const LOC = { "Қиял-ғажайып": "қиял-ғажайып ертегілерде", "Хайуанаттар": "хайуанаттар туралы ертегілерде", "Тұрмыс-салт": "тұрмыс-салт ертегілерінде", "Аңыз-ертегі": "аңыз-ертегіде" };
  V.type_text = `Бір ертегіге шаққанда сандар ең көп ${LOC[tMost[0]]} кездеседі (орта есеппен ${f1(tMost[1])}), ал символдық сандардың үлесі ${LOC[tF[0]]} ең жоғары (${f1(tF[2])}%).`;
  V.triad_text = `Үштік құрылым ${N.triad} ертегіде анықталды, оның ішінде ${N.triad_magic} – қиял-ғажайып ертегі (барлығы ${N.magic_read} қиял-ғажайып ертегі оқылды).`;
  const okA = N.top.slice(0, 5).map(([n]) => Number(n)).filter(n => [3, 7, 40].includes(n)).length >= 2;
  const okB = N.sacred_F_pct > N.nonsacred_F_pct;
  const okC = N.triad_magic > N.magic_read / 2;
  const nOk = [okA, okB, okC].filter(Boolean).length;
  V.hyp_short = nOk === 3 ? "Болжамымыз расталды." : nOk === 0 ? "Болжамымыз расталмады." : "Болжамымыз ішінара расталды.";
  V.hyp_text = `Болжамымызды үш бөлік бойынша тексердік. Біріншіден, 3, 7 және 40 сандары ең жиі кездесетін бес санның қатарына ${okA ? "кірді – бұл бөлігі расталды" : "толық кірмеді – бұл бөлігі расталмады"}. Екіншіден, «киелі» сандардың ${f1(N.sacred_F_pct)}%-ы, басқа сандардың ${f1(N.nonsacred_F_pct)}%-ы символдық қызмет атқарды – ${okB ? "бұл да расталды" : "бұл расталмады"}. Үшіншіден, ${N.magic_read} қиял-ғажайып ертегінің ${N.triad_magic}-інде үштік құрылым бар – ${okC ? "яғни олардың көбі үштік заңына бағынады" : "яғни олардың көбі үштік заңына бағынбайды"}. ${V.hyp_short}`;
  V.concl_top = V.top_text;
  V.concl_func = `Символдық қызметтегі сандардың үлесі «киелі» сандарда ${f1(N.sacred_F_pct)}%, басқа сандарда ${f1(N.nonsacred_F_pct)}% болды.`;
  V.concl_triad = V.triad_text;
}
for (const [k, txt] of Object.entries({
  top_text: "[Ең жиі кездескен сандарды атаңыз.]", other_text: "", sacred_text: "[«Киелі» сандардың үлесі туралы қорытынды жазыңыз.]",
  func_text: "[«Киелі» және басқа сандардың қызметін салыстырыңыз.]", cat_text: "[Сандар көбіне нені білдіреді?]",
  type_text: "[Ертегі түрлерін салыстырыңыз.]", triad_text: "[Үштік құрылым неше ертегіде кездесті?]",
  hyp_text: "[Болжамның расталғанын не расталмағанын деректермен көрсетіңіз.]", hyp_short: "[Болжам расталды / ішінара расталды / расталмады.]",
  concl_top: "[Ең жиі кездескен сандар.]", concl_func: "[Сандардың қызметі бойынша қорытынды.]", concl_triad: "[Үштік құрылым бойынша қорытынды.]",
})) if (V[k] === undefined) V[k] = txt;

const fill = s => s.replace(/\{\{([^}]+)\}\}/g, (_, k) => (V[k] !== undefined && V[k] !== null) ? String(V[k]) : PH);

// ---------- Дереккөздерді нөмірлеу ----------
const content = fs.readFileSync(path.join(DIR, "content.txt"), "utf8").split("\n").filter(l => l.trim() !== "");
const order = [];
for (const line of content) for (const m of line.matchAll(/\[@([a-z_]+)\]/g)) if (!order.includes(m[1])) order.push(m[1]);
for (const k of order) if (!REFS[k]) throw new Error("Missing ref " + k);
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) console.error("Unused refs:", unused);
const cite = s => s.replace(/\[@([a-z_]+)\]/g, (_, k) => `[${order.indexOf(k) + 1}]`);

// ---------- Көмекші функциялар ----------
const HL_RE = /(\[___\]|\[(?!Мәтін\])[А-ЯӘҒҚҢӨҰҮҺІа-яәғқңөұүһі][^\]]*\])/;
function runs(text, o = {}) {
  text = cite(fill(String(text)));
  const out = [];
  text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).forEach(seg => {
    const b = seg.startsWith("**");
    if (b) seg = seg.slice(2, -2);
    seg.split(HL_RE).filter(p => p !== "").forEach(p => out.push(new TextRun({
      text: p, font: FONT, size: o.size || SZ, bold: o.bold || b, italics: o.italics,
      highlight: HL_RE.test(p) || o.hl ? "yellow" : undefined, color: o.color,
    })));
  });
  return out;
}
function para(text, o = {}) {
  return new Paragraph({
    alignment: o.align ?? AlignmentType.JUSTIFIED,
    indent: o.noIndent ? o.indent : { firstLine: Math.round(1.25 * CM) },
    spacing: { line: o.line || LINE, before: o.before || 0, after: o.after || 0 },
    keepNext: o.keepNext, pageBreakBefore: o.pb,
    children: runs(text, o),
  });
}
const C = AlignmentType.CENTER, L = AlignmentType.LEFT;
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, w, o = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA }, columnSpan: o.span,
    shading: o.head || o.shade ? { fill: o.head ? "E7EEF5" : "F3F3F0", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 30, bottom: 30, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: String(text).split("\n").map(t => new Paragraph({
      alignment: o.center ? C : L, spacing: { line: 240 },
      children: runs(t, { size: o.size || SZ_T, bold: o.bold || o.head, italics: o.italics }),
    })),
  });
}
function mkTable(widthsPct, header, rows, o = {}) {
  const widths = widthsPct.map(p => Math.floor(TEXT_W * p / 100));
  widths[widths.length - 1] += TEXT_W - widths.reduce((a, b) => a + b, 0);
  const trs = [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { head: true, center: true, size: o.size })) })];
  rows.forEach(r => {
    if (r.section) {
      trs.push(new TableRow({ cantSplit: true, children: [cell(r.section, TEXT_W, { span: header.length, bold: true, shade: true, size: o.size })] }));
      return;
    }
    trs.push(new TableRow({
      cantSplit: true,
      children: r.map((c, i) => cell(c, widths[i], { center: o.centerCols?.includes(i), bold: (o.boldCol0 && i === 0) || o.boldRow?.(r), size: o.size })),
    }));
  });
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
let tNo = 0, fNo = 0;
const caption = t => para(`${++tNo}-кесте – ${t}`, { noIndent: true, keepNext: true, before: 120, after: 60, align: L, line: 276 });
const gap = () => new Paragraph({ spacing: { after: 120, line: 240 }, children: [] });

// ---------- Кестелер ----------
const TABLES = {
  kieli: () => [caption("Қазақ мәдениетіндегі «киелі» сандар"), mkTable([10, 45, 45], ["Сан", "Мәдениеттегі мысалдар", "Мағынасы"], [
    ["3", "Үш жүз (Ұлы, Орта, Кіші жүз); ертегідегі үш ағайынды, үш сынақ", "Толықтық, бүтіндік; оқиғаның басы, ортасы, аяғы"],
    ["7", "«Жеті атасын білу» дәстүрі; жеті қазына; марқұмның жетісі; Жеті қарақшы шоқжұлдызы", "Ұрпақ сабақтастығы, толық жиынтық, қасиеттілік"],
    ["9", "Тоғыздап сыйлау, тоғыз түрлі сый; тоғызқұмалақ ойыны", "Молшылық, құрмет, сыйдың ең жоғары өлшемі"],
    ["12", "Мүшел – 12 жылдық жыл санау; он екі жануардың аты берілген жылдар", "Уақыттың толық айналымы"],
    ["40", "Баланы қырқынан шығару; марқұмның қырқы; «қырық» сөзінің «көп» мағынасы", "Өтпелі кезеңнің аяқталуы, көптік"],
  ], { centerCols: [0], boldCol0: true }), gap()],
  korpus: () => {
    const t = HAS ? N.tales : D.corpus.map(([n, g]) => ({ name: n, type: g, edition: "" }));
    return [caption("Зерттеу материалы: ертегілер"), mkTable([6, 30, 16, 48], ["№", "Ертегі", "Түрі", "Оқыған басылым немесе сайт"],
      t.map((x, i) => [String(i + 1), `«${x.name}»`, x.type, x.edition || "[кітап, баспа, жылы]"]), { centerCols: [0] }), gap()];
  },
  san_sany: () => {
    const rows = D.numbers.map(n => { const o = HAS ? N.by_num[String(n)] : null; return [String(n), o ? String(o.all) : PH, o ? String(o["Н"]) : PH, o ? String(o["Ф"]) : PH]; });
    const oF = HAS ? N.rows.filter(x => !D.numbers.includes(x.num) && x.func === "Ф").length : 0;
    rows.push(["басқа", HAS ? String(N.other) : PH, HAS ? String(N.other - oF) : PH, HAS ? String(oF) : PH]);
    rows.push(["Барлығы", HAS ? String(N.total) : PH, HAS ? String(N.total - N.F) : PH, HAS ? String(N.F) : PH]);
    return [caption("Ертегілердегі сандардың кездесуі"), mkTable([25, 25, 25, 25], ["Сан", "Барлығы", "Нақты сан (Н)", "Формула, символ (Ф)"], rows,
      { centerCols: [0, 1, 2, 3], boldRow: r => r[0] === "Барлығы" || D.sacred.includes(Number(r[0])) }), gap()];
  },
  formulalar: () => {
    const F = [
      ["отыз күн ойын, қырық күн той", /отыз күн ойын|қырық күн той/, [30, 40]], ["қырық құлаш (белбеу, қазан, қармақ)", /қырық құлаш/, [40]],
      ["алты айлық жол, жеті айлық жер", /айлық (жол|жер)/, [6, 7]], ["алты таудың, алты қырдың ар жағы; жеті көлдің желкесі", /таудың ар|қырдың ар|көлдің желке/, [6, 7]],
      ["алты аяқты ала ат, жеті аяқты жирен ат", /аяқты (ала|жирен)/, [6, 7]], ["жеті басты дәу", /жеті бас/, [7]],
      ["жүз шақырым", /жүз шақырым/, [100]], ["үш күн, үш түн; үш жұма, үш ай", /үш (күн|жұма)/, [3]], ["мың қоян, мың құлан", /мың (қоян|құлан)/, [1000]],
    ];
    const rows = F.map(([lab, re, nums]) => {
      const m = HAS ? N.rows.filter(x => nums.includes(x.num) && re.test(x.phrase.toLowerCase())) : [];
      return [`«${lab}»`, HAS ? String(m.length) : PH, HAS ? [...new Set(m.map(x => x.tale))].map(t => `«${t}»`).join(", ") : PH];
    });
    return [caption("Ертегілерде қайталанған сан формулалары"), mkTable([42, 12, 46], ["Формула", "Саны", "Қай ертегілерде"], rows, { centerCols: [1] }), gap()];
  },
  turleri: () => {
    const rows = [];
    D.tale_types.forEach(t => {
      const tales = HAS ? N.tales.filter(x => x.type === t) : D.corpus.filter(c => c[1] === t).map(c => ({ name: c[0] }));
      rows.push({ section: t });
      tales.forEach(x => {
        const r = HAS ? N.rows.filter(y => y.tale === x.name) : [];
        rows.push([`«${x.name}»`, HAS ? String(r.length) : PH, HAS ? String(r.filter(y => y.func === "Ф").length) : PH, HAS && r.length ? f1(r.filter(y => y.func === "Ф").length / r.length * 100) : PH]);
      });
      if (HAS) {
        const o = N.by_type[t];
        rows.push([`Барлығы: ${o.tales} ертегі, бір ертегіге ${f1(o.n / o.tales)} сан`, String(o.n), String(o.F), f1(o.F / o.n * 100)]);
      }
    });
    if (HAS) rows.push(["Барлық 7 ертегі", String(N.total), String(N.F), f1(N.F_pct)]);
    return [caption("Ертегілер мен ертегі түрлері бойынша сандар"), mkTable([46, 18, 18, 18], ["Ертегі", "Сандар саны", "Символ (Ф)", "Символ, %"], rows,
      { centerCols: [1, 2, 3], boldRow: r => r[0].startsWith("Барлы") }), gap()];
  },
};

function figBlock(name, cap) {
  const buf = fs.readFileSync(path.join(DIR, "figs", name + ".png"));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const maxW = 600, maxH = name === "zhadynama" ? 620 : 520;
  let dw = maxW, dh = Math.round(maxW * h / w);
  if (dh > maxH) { dh = maxH; dw = Math.round(maxH * w / h); }
  return [
    new Paragraph({ alignment: C, keepNext: true, spacing: { before: 120 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: dw, height: dh },
        altText: { title: cap, description: cap, name } })] }),
    para(`${++fNo}-сурет – ${cap}`, { noIndent: true, align: C, after: 160, before: 60, line: 276 }),
  ];
}
function h1(text, pb = true) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, alignment: C, pageBreakBefore: pb,
    spacing: { after: 240, line: LINE }, keepNext: true,
    children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, alignment: L,
    indent: { firstLine: Math.round(1.25 * CM) }, spacing: { before: 240, after: 120, line: LINE }, keepNext: true,
    children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
  });
}

// ---------- Титул парағы (8-қосымша) ----------
const line = (t, o = {}) => new Paragraph({ alignment: o.align ?? C, spacing: { line: 276, before: o.before || 0, after: o.after || 0 },
  indent: o.indent, children: runs(t, o) });
const front = [
  line("Қазақстан Республикасы Оқу-ағарту министрлігі"),
  line("«Дарын» республикалық ғылыми-практикалық орталығы", { after: 360 }),
  line("2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсы", { after: 900 }),
  line("Тіркеу коды (шифр) ____________________________"),
  line("(«Дарын» РҒПО қызметкері толтырады)", { size: 22, italics: true, after: 600 }),
  line("Бағыт: қоғамдық-гуманитарлық бағыт (ҚГБ)", { align: L, after: 200 }),
  line("Секцияның атауы: Қазақ тілі және әдебиеті", { align: L, after: 200 }),
  line("Жас санаты:", { align: L }),
  line("☐ «Бастауыш буын» (2–4 сынып)", { align: L, indent: { left: 567 } }),
  line("☒ «Орта буын» (5–8 сынып)", { align: L, indent: { left: 567 }, after: 700 }),
  line("Жобаның тақырыбы", { bold: true }),
  line(`«${TITLE}»`, { bold: true, size: 32, after: 900 }),
  ...(JURY ? [line("", { after: 1400 })] : [
    line("Қатысушының Т.А.Ә.: [Оқушының аты-жөні], 7-сынып", { align: L, after: 300 }),
    line("Ғылыми жетекшінің Т.А.Ә.: [Жетекшінің аты-жөні]", { align: L, after: 1000 }),
  ]),
  line("2026–2027 оқу жылы"),
];

// ---------- Аннотация ----------
const ANNOT = "Жұмыстың мақсаты – қазақ ертегілеріндегі сандарды санап, олардың жиілігі мен қызметін анықтау. Өзектілігі: ертегілерде кейбір сандар жиі қайталанады, ал олардың мағынасы халықтың дүниетанымын көрсетеді. Зерттеу сұрағы: қазақ ертегілерінде қандай сандар жиі кездеседі және олар нақты мөлшерді білдіре ме, әлде символ ма? Болжам: «киелі» 3, 7 және 40 сандары жиі кездеседі және көбіне символ ретінде қолданылады. Әдістер: мәтінді талдау және кодтау, санау, пайызды есептеу, салыстыру. 7 ертегіде 91 сан табылды, бір ертегіге орта есеппен 13 сан келеді. Ең жиі сан – 40 (15 рет), одан кейін 1 мен 6 (14 реттен). Сандардың 76,9%-ы символ ретінде қолданылады; 3 пен 7 әрқашан символ болды, ал 9 бен 40 көбіне нақты топты санады. «Отыз күн ойын, қырық күн той» формуласы үш ертегіде қайталанды. Болжам ішінара расталды. Нәтиже бойынша математика мен әдебиетті ұштастыратын «Ертегі-есептер» жинағы жасалды.";
const annotWords = fill(ANNOT).split(/\s+/).filter(w => /[\p{L}\d]/u.test(w)).length;
if (annotWords > 200) throw new Error("Аннотация 200 сөзден асты: " + annotWords);
console.log("Аннотация сөз саны:", annotWords);
const annot = [h1("АННОТАЦИЯ"), para(ANNOT), para("Кілт сөздер: қазақ ертегілері, сан, киелі сандар, символ, ертегі формуласы, контент-талдау.", { before: 120 })];

// ---------- Қосымшалар ----------
const APPX = [
  { id: "А", title: "Ертегілердегі сандар кестесі" },
  { id: "Б", title: "«Ертегі-есептер» жинағы" },
];
const ESEPTER = [
  ["Той неше апта?", "Ертегілер көбіне «отыз күн ойын, қырық күн тойын жасапты» деп аяқталады. Ойын мен той бірінен соң бірі өтті делік. Той-думан барлығы неше аптаға созылды?", "(30 + 40) : 7 = 10 апта."],
  ["Ер Төстіктің өсуі", "«Ер Төстік» ертегісінде Төстік «бір айда бір жастағы баладай, екі айда екі жастағы баладай» өседі. Осылай өссе, 9 айда ол неше жастағы баладай болады? Ол қарапайым баладан неше есе тез өседі?", "9 айда 9 жастағы баладай болады. Қарапайым бала бір жаста 12 ай өседі, Төстік – 1 ай, яғни 12 есе тез."],
  ["Керқұла аттың жылдамдығы", "«Керқұла атты Кендебай» ертегісінде Керқұла ат «алты сағатта алты айлық жол» жүреді. Қарапайым ат алты айлық жолды 180 тәулік жүреді делік. Бір тәулікте 24 сағат. Керқұла қарапайым аттан неше есе жылдам?", "180 · 24 = 4320 сағат; 4320 : 6 = 720 есе жылдам."],
  ["Отыз тоғыз ат", "«Аяз би» ертегісінде қырық аттың отыз тоғызы сойылып, біреуі қалдырылады. Аттардың неше пайызы қалды?", "1 : 40 · 100% = 2,5%."],
  ["Жеті буын", "Әр адамның екі ата-анасы бар, олардың әрқайсысының да екі ата-анасы бар. Әр буында ата-бабалар саны екі есе артады. Жетінші буында неше ата-баба болады? Жауабын дәреже түрінде жаз.", "2⁷ = 128 ата-баба. (Дәстүрдегі «жеті ата» әке жағынан саналады, ал есепте барлық ата-бабалар есептелді.)"],
  ["Мүшел жыл", "«Жыл басына таласқан хайуанаттар» ертегісінде 12 жануар жыл басы болуға таласады. 2026 жыл – жылқы жылы. 2050 жыл қай жануардың жылы болады?", "2050 − 2026 = 24 = 12 · 2, яғни 2050 жыл да – жылқы жылы."],
  ["Жүз жылда неше мүшел?", "Мүшел 12 жылда бір рет айналып келеді. 100 жылда неше толық мүшел өтеді және неше жыл артылып қалады?", "100 = 12 · 8 + 4: 8 толық мүшел, 4 жыл артылады."],
  ["Тоғызқұмалақ", "Тоғызқұмалақ тақтасында әр ойыншының 9 отауы бар, әр отауда басында 9 құмалақтан болады. Тақтада барлығы неше құмалақ бар? Жеңу үшін жартысынан көп құмалақ жинау керек. Ең кемі неше құмалақ жинау қажет?", "2 · 9 · 9 = 162 құмалақ; жартысы – 81, жеңу үшін ең кемі 82 құмалақ керек."],
  ["Қырқынан шығару", "Қазақ салтында нәрестені туылғаннан кейін 40 күн өткенде қырқынан шығарады. Бала 1 наурызда туылса, 40 күн қай күні толады?", "Наурызда 31 күн: 1 наурыз + 30 күн = 31 наурыз, тағы 10 күн – 10 сәуір."],
  ["Тазшаның өтірігі", "«Тазшаның қырық өтірігі» ертегісінде Тазша 40 өтірік айтуы керек. Ол әр кеште 3 өтірік айтса, 40 өтірікті неше кеште айтып бітіреді?", "40 : 3 = 13 (қалдық 1), сондықтан 14 кеш керек."],
];
function appendix() {
  const out = [];
  const head = (a) => { out.push(h1(`${a.id} ҚОСЫМШАСЫ`)); out.push(para(a.title, { noIndent: true, align: C, bold: true, after: 200 })); };
  head(APPX[0]);
  out.push(para("Кестеде 7 ертегіде кездескен барлық 91 сан берілген. Қызметі: Н – нақты сан, Ф – тұрақты формула немесе символ.", { after: 120 }));
  const rows = HAS ? N.rows : [];
  out.push(mkTable([6, 26, 9, 33, 16, 10], ["№", "Ертегі", "Сан", "Мәтіндегі тіркес", "Нені білдіреді", "Қызметі"],
    rows.length ? rows.map((x, i) => [String(i + 1), `«${x.tale}»`, String(x.num), x.phrase + (x.page ? ` (${x.page}-б.)` : ""), x.cat, x.func])
      : [1, 2, 3].map(i => [String(i), "[ертегі]", "[сан]", "[тіркес, беті]", "[нені білдіреді]", "[Н/Ф]"]), { size: 20, centerCols: [0, 2, 5] }));
  head(APPX[1]);
  out.push(para("Есептер біз оқыған ертегілердегі және қазақ салт-дәстүріндегі сандарға негізделген. Олар 7-сынып математикасының тақырыптарына сай келеді: дәреже, пайыз, бөлу, қалдықпен бөлу, күнтізбе бойынша есептеу.", { after: 120 }));
  ESEPTER.forEach(([t, q, a], i) => {
    out.push(mkTable([22, 78], [`${i + 1}-есеп`, t], [["Шарты", q], ["Жауабы", a]], { size: 22, boldCol0: true }));
    out.push(gap());
  });
  return out;
}

// ---------- Мазмұны ----------
const tocEntries = [{ t: "АННОТАЦИЯ", lvl: 1 }];
for (const l of content) {
  let m;
  if ((m = l.match(/^#\*? (.*)$/))) tocEntries.push({ t: m[1], lvl: 1 });
  else if ((m = l.match(/^## (.*)$/))) tocEntries.push({ t: m[1], lvl: 2 });
}
APPX.forEach(a => tocEntries.push({ t: `${a.id} ҚОСЫМШАСЫ`, lvl: 1, label: `${a.id} қосымшасы. ${a.title}` }));
const toc = [h1("МАЗМҰНЫ")];
for (const e of tocEntries) {
  toc.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W, leader: "dot" }],
    indent: e.lvl === 2 ? { left: 400 } : undefined,
    spacing: { line: 300, after: 20 },
    children: [new TextRun({ text: e.label || e.t, font: FONT, size: SZ }), new TextRun({ text: `\t${TOC_PAGES[e.t] ?? ""}`, font: FONT, size: SZ })],
  }));
}

// ---------- Негізгі мәтін ----------
const body = [];
for (const l of content) {
  let m;
  if ((m = l.match(/^#\*? (.*)$/))) body.push(h1(m[1]));
  else if ((m = l.match(/^## (.*)$/))) body.push(h2(m[1]));
  else if ((m = l.match(/^-- (.*)$/))) body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, numbering: { reference: "dash", level: 0 },
    spacing: { line: LINE }, children: runs(m[1]) }));
  else if ((m = l.match(/^!fig (\w+) \| (.*)$/))) body.push(...figBlock(m[1], m[2]));
  else if ((m = l.match(/^!table (\w+)/))) body.push(...TABLES[m[1]]());
  else if ((m = l.match(/^!note (.*)$/))) body.push(para("НҰСҚАУ (тапсыру алдында өшіріледі): " + m[1], { hl: true, size: 24, line: 276, before: 60, after: 120 }));
  else if (l.startsWith("!references")) order.forEach((k, i) => body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 }, spacing: { line: LINE },
    tabStops: [{ type: TabStopType.LEFT, position: 567 }],
    children: [new TextRun({ text: `${i + 1}\t${REFS[k]}`, font: FONT, size: SZ })] })));
  else body.push(para(l));
}
const app = appendix();
console.log("Кестелер:", tNo, "Суреттер:", fNo, "Дереккөздер:", order.length);

const doc = new Document({
  creator: "", title: TITLE,
  styles: {
    default: { document: { run: { font: FONT, size: SZ } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "dash", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–",
    alignment: L, style: { paragraph: { indent: { left: Math.round(1.25 * CM) + 360, hanging: 360 } } } }] }] },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MT, bottom: MB, left: ML, right: MR } },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: C, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }),
      first: new Footer({ children: [] }),
    },
    children: [...front, ...annot, ...toc, ...body, ...app],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("written", OUT); });
