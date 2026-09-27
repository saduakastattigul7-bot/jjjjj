// Builds the main project document from content.txt.
// Usage: node build.js out.docx [--jury]   (--jury: anonymised copy without names)
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
const TOC_PAGES = fs.existsSync(path.join(DIR, "toc_pages.json"))
  ? JSON.parse(fs.readFileSync(path.join(DIR, "toc_pages.json"), "utf8")) : {};

const FONT = "Times New Roman";
const SZ = 28;          // 14 pt
const SZ_T = 24;        // 12 pt in tables (rule: not less than 10 pt)
const LINE = 360;       // 1.5 line spacing
const CM = 567;         // dxa per cm
const PAGE_W = 11906, ML = 3 * CM, MR = 2 * CM, MT = 2 * CM, MB = 2 * CM;
const TEXT_W = PAGE_W - ML - MR;
const C = AlignmentType.CENTER;

// ---------- References ----------
const REFS = {
  propp: "Пропп В.Я. Морфология сказки. – 2-е изд. – М.: Наука, 1969.",
  potanin: "Потанин Г.Н. Очерки Северо-Западной Монголии. Вып. IV: Материалы этнографические. – СПб., 1883.",
  oqulyq: "«Ер Төстік» ертегісі // Қазақ әдебиеті: жалпы білім беретін мектептің 5-сыныбына арналған оқулық. – ⟦Қала: Баспа, жылы. Авторларын өз оқулығыңның титул бетінен көшіріп жаз⟧.",
  wiki: "Ер Төстік // Уикипедия – ашық энциклопедия. – URL: https://kk.wikipedia.org/wiki/Ер_Төстік (қаралған күні: ⟦__.__.2026⟧).",
  zan: "Қазақстан Республикасының 2003 жылғы 19 желтоқсандағы № 508-II «Жарнама туралы» Заңы // «Әділет» ақпараттық-құқықтық жүйесі. – URL: https://adilet.zan.kz/kaz/docs/Z030000508_ (қаралған күні: ⟦__.__.2026⟧).",
  ogilvy: "Огилви Д. Огилви о рекламе / пер. с англ. – М.: Эксмо, 2007.",
  strong: "Strong E.K. The Psychology of Selling and Advertising. – New York: McGraw-Hill, 1925.",
};

// ---------- Tables ----------
const PH = "⟦__⟧";
const TABLES = {
  plan: {
    caption: "Жұмыс жоспары", src: "автор құрастырған",
    widths: [8, 62, 30], header: ["№", "Не істедім", "Мерзімі"], center: [0, 2],
    rows: [
      ["1", "Тақырып таңдау, зерттеу сұрағы мен болжамды жазу", "қыркүйек"],
      ["2", "Ертегі мен жарнама туралы дереккөздерді оқу", "қыркүйек"],
      ["3", "Сауалнама құрастыру, ата-аналардан келісім алу, сауалнама жүргізу", "қазан"],
      ["4", "«Ер Төстікті» қайта оқып, қызықты тұстарды іріктеу", "қазан"],
      ["5", "Слоган, плакат, бейнеролик сценарийін жасау", "қазан"],
      ["6", "Жарнаманы сыныптастарыма көрсету, нәтижені салыстыру", "қараша"],
      ["7", "Нәтижелерді санау, кесте мен диаграмма жасау, қорытынды жазу", "қараша"],
      ["8", "Жұмысты рәсімдеу, стенд жасау, қорғауға дайындалу", "желтоқсан"],
    ],
  },
  aida: {
    caption: "Жақсы жарнаманың төрт қадамы және оның менің жарнамамдағы көрінісі", src: "[@strong] бойынша автор құрастырған",
    widths: [24, 36, 40], header: ["Қадам", "Мағынасы", "Менің жарнамамда"], bold0: true,
    rows: [
      ["A – Attention (назар аудару)", "Адамды бірден тоқтатып, қаратып алу", "Сұрақ-тақырып: «Сөйлейтін атты көрдің бе?»; жарқын түстер"],
      ["I – Interest (қызығушылық)", "Ары қарай білгісі келетіндей ету", "Ғажайып серіктер, жер асты елі, құпия сұрақ"],
      ["D – Desire (қалау)", "«Маған да керек!» деген ой тудыру", "«Сен де Төстіктей батыр болғың келе ме?»"],
      ["A – Action (әрекет)", "Енді не істеу керегін айту", "«Мектеп кітапханасынан ал да, оқы!»"],
    ],
  },
  tusy: {
    caption: "«Ер Төстік» ертегісінен жарнамаға іріктелген қызықты тұстар", src: "автордың ертегіні талдауы [@oqulyq]",
    widths: [6, 26, 30, 38], header: ["№", "Ертегідегі қызықты тұс", "Неге қызықты?", "Жарнамада қалай пайдаландым"], center: [0],
    rows: [
      ["1", "Батырдың аты неге «Төстік»?", "Аты ерекше, құпиясы бар", "Бейнероликтегі сұрақ: «Батырдың аты неге Төстік? Жауабы – кітапта!»"],
      ["2", "Сөйлейтін ақылды тұлпар – Шалқұйрық", "Ат сөйлейді, иесіне ақыл береді", "Плакаттың ортасындағы басты сурет; бейнеролик соның дауысымен басталады"],
      ["3", "Ғажайып серіктер: Желаяқ, Саққұлақ, Көлтауысар, Таусоғар", "Әрқайсысының ерекше күші бар, қазіргі суперқаһармандарға ұқсайды", "«Қазақтың нағыз суперқаһармандары!» деген кадр"],
      ["4", "Жер асты елі", "Құпия әлем, не болатыны белгісіз", "Плакаттағы сұрақ: «Жер астында Төстікті не күтіп тұр?»"],
      ["5", "Алып Самұрық құс", "Үлкен, ғажайып құс, батырды құтқарады", "Плакаттың жоғарғы бұрышындағы сурет"],
      ["6", "Жалмауыз кемпір", "Қорқынышты, сондықтан қызық", "Суретін салмадым, тек құпия ретінде айттым"],
      ["7", "Бауырларын іздеп табуы, Кенжекейдің адалдығы", "Отбасы мен адалдық туралы", "Слоганда «батыр бол» деген сөз, радио жарнамасы"],
    ],
  },
  slogan: {
    caption: "Слоган нұсқалары және сыныптастарымның таңдауы", src: "автордың сауалнамасы",
    widths: [7, 57, 18, 18], header: ["№", "Слоган", "Таңдағандар саны", "%"], center: [0, 2, 3],
    rows: [
      ["1", "«Жер астына түсіп, жеңіп шыққан батыр – Ер Төстік!»", PH, PH],
      ["2", "«Шалқұйрық шапса – ертегі басталады!»", PH, PH],
      ["3", "«Бір кітап – мың шытырман оқиға!»", PH, PH],
      ["4", "«Ер Төстік – оқы да, батыр бол!»", PH, PH],
      ["", "Барлығы", PH, "100"],
    ],
  },
  survey: {
    caption: "Сауалнама нәтижелері", src: "автордың сауалнамасы, қатысқандар саны – ⟦__⟧",
    widths: [34, 38, 14, 14], header: ["Сұрақ", "Жауап", "Саны", "%"], center: [2, 3],
    rows: [
      ["1. Қазақ ертегілерін оқығанды ұнатасың ба?", "Иә", PH, PH], ["", "Кейде", PH, PH], ["", "Жоқ", PH, PH],
      ["2. Соңғы бір айда кітаптан ертегі оқыдың ба?", "Иә", PH, PH], ["", "Жоқ", PH, PH],
      ["3. «Ер Төстік» ертегісін қалай білесің?", "Кітаптан оқыдым", PH, PH], ["", "Мультфильмін көрдім", PH, PH],
      ["", "Атын ғана естідім", PH, PH], ["", "Білмеймін", PH, PH],
      ["4. Кітапты оқуға сені не қызықтырады?*", "Мұқабасы мен суреті", PH, PH], ["", "Досымның кеңесі", PH, PH],
      ["", "Мультфильм немесе бейне", PH, PH], ["", "Мұғалімнің тапсырмасы", PH, PH],
      ["5. Жарнаманы көбіне қай жерден көресің?", "Телефоннан", PH, PH], ["", "Теледидардан", PH, PH],
      ["", "Көшедегі плакаттан", PH, PH], ["", "Мектептегі хабарландырудан", PH, PH],
    ],
    note: "* 4-сұрақта бірнеше жауап таңдауға болатындықтан, пайыздардың қосындысы 100%-дан асуы мүмкін.",
  },
  compare: {
    caption: "А және Б плакаттарын салыстыру", src: "автордың тәжірибесі, қатысқандар саны – ⟦__⟧",
    widths: [52, 24, 24], header: ["Көрсеткіш", "А плакаты (жай хабарландыру)", "Б плакаты (менің жарнамам)"], center: [1, 2],
    rows: [
      ["«Осыны көргенде ертегіні оқығым келеді» деп таңдағандар саны", PH, PH],
      ["Сол, %", PH, PH],
      ["10 минуттан кейін кейіпкердің атын есте сақтағандар саны", PH, PH],
      ["10 минуттан кейін ертегінің атын есте сақтағандар саны", PH, PH],
    ],
  },
  beforeafter: {
    caption: "«Ер Төстікті оқығың келе ме?» сұрағына жарнамаға дейін және кейін берілген жауаптар", src: "автордың тәжірибесі",
    widths: [46, 18, 18, 18], header: ["Көрсеткіш", "Жарнамаға дейін", "Жарнамадан кейін", "Өзгеріс"], center: [1, 2, 3],
    rows: [
      ["Қатысушылар саны", PH, PH, "–"],
      ["Орташа балл (1-ден 5-ке дейін)", PH, PH, PH],
      ["«5» қойғандар саны", PH, PH, PH],
      ["«4» немесе «5» қойғандар, %", PH, PH, PH],
      ["Бір аптада ертегіні оқығандар немесе оқи бастағандар саны", "–", PH, "–"],
    ],
    note: "Орташа балды былай есептедім: барлық қойылған балдарды қосып, қатысушылар санына бөлдім.",
  },
  video: {
    caption: "Бейнероликтің кадрлары", src: "автордың сценарийі",
    widths: [10, 14, 36, 40], header: ["Кадр", "Уақыты", "Экранда не көрінеді", "Не айтылады (дыбыс)"], center: [0, 1],
    rows: [
      ["1", "0–5 сек", "Қараңғы экран, ат дүбірі естіледі, Шалқұйрықтың суреті пайда болады", "Шалқұйрық: «Мен сөйлейтін атпын! Менің иемді білесің бе?»"],
      ["2", "5–10 сек", "Шалқұйрыққа мінген Төстіктің суреті", "«Бұл – Ер Төстік! Ол ағаларын іздеп, алыс сапарға шықты»"],
      ["3", "10–15 сек", "Желаяқ, Саққұлақ, Көлтауысар, Таусоғардың суреттері кезекпен шығады", "«Оның серіктері – қазақтың нағыз суперқаһармандары!»"],
      ["4", "15–20 сек", "Жердегі үлкен қараңғы үңгір, сұрақ белгісі", "«Ал жер астында оны не күтіп тұр?»"],
      ["5", "20–25 сек", "Кітаптың мұқабасы, оны бала қолына алады", "«Жауабы – осы кітапта! Батырдың аты неге Төстік екенін де білесің»"],
      ["6", "25–30 сек", "Слоган жазылған экран, мектеп кітапханасының суреті", "«Ер Төстік – оқы да, батыр бол! Мектеп кітапханасында күтеміз!»"],
    ],
  },
};

// ---------- Citations ----------
const content = fs.readFileSync(path.join(DIR, "content.txt"), "utf8").split("\n").filter(l => l.trim() !== "");
const CITE_RE = /\[(@[a-z_]+(?:;\s*@[a-z_]+)*)\]/g;
const order = [];
const note = s => { for (const m of s.matchAll(CITE_RE)) for (const k of m[1].split(/;\s*/)) { const key = k.slice(1); if (!order.includes(key)) order.push(key); } };
for (const l of content) {
  note(l);
  const t = l.match(/^!table (\w+)/);
  if (t) { const T = TABLES[t[1]]; note(T.src || ""); T.rows.forEach(r => r.forEach(note)); }
}
for (const k of order) if (!REFS[k]) throw new Error("Missing ref " + k);
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) console.error("Unused refs:", unused);
const cite = s => s.replace(CITE_RE, (_, g) => "[" + g.split(/;\s*/).map(k => order.indexOf(k.slice(1)) + 1).join("; ") + "]");

// ---------- Helpers ----------
const HL_RE = /(⟦[^⟧]*⟧)/;
function runs(text, o = {}) {
  return cite(text).split(HL_RE).filter(p => p !== "").map(p => {
    const hl = HL_RE.test(p);
    return new TextRun({
      text: hl ? "[" + p.slice(1, -1) + "]" : p, font: FONT, size: o.size || SZ,
      bold: o.bold, italics: o.italics, highlight: hl ? "yellow" : undefined,
    });
  });
}
function para(text, o = {}) {
  return new Paragraph({
    alignment: o.align ?? AlignmentType.JUSTIFIED,
    indent: o.noIndent ? undefined : { firstLine: Math.round(1.25 * CM) },
    spacing: { line: o.line || LINE, before: o.before || 0, after: o.after || 0 },
    keepNext: o.keepNext,
    children: runs(text, o),
  });
}
let tNo = 0, fNo = 0;
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, w, o = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: o.head ? { fill: "E7EEF5", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 90, right: 90 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: o.center ? C : AlignmentType.LEFT, spacing: { line: 240 },
      children: runs(text, { size: SZ_T, bold: o.bold || o.head }),
    })],
  });
}
function mkTable(T) {
  const widths = T.widths.map(p => Math.floor(TEXT_W * p / 100));
  widths[widths.length - 1] += TEXT_W - widths.reduce((a, b) => a + b, 0);
  const center = T.center || [];
  const trs = [new TableRow({ tableHeader: true, children: T.header.map((h, i) => cell(h, widths[i], { head: true, center: true })) })];
  T.rows.forEach(r => trs.push(new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], { center: center.includes(i), bold: (T.bold0 && i === 0) || r[1] === "Барлығы" })),
  })));
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
const small = (t, o = {}) => para(t, { noIndent: true, size: SZ_T, line: 240, align: AlignmentType.LEFT, ...o });
function tableBlock(name) {
  const T = TABLES[name];
  tNo++;
  const out = [para(`${tNo}-кесте – ${T.caption}`, { noIndent: true, keepNext: true, before: 120, after: 60, align: AlignmentType.LEFT, line: 240 })];
  out.push(mkTable(T));
  if (T.note) out.push(small(T.note, { before: 40 }));
  out.push(small(`Дереккөзі: ${T.src}.`, { before: 40, after: 160 }));
  return out;
}
function figBlock(name, caption, src) {
  fNo++;
  const buf = fs.readFileSync(path.join(DIR, "figs", name + ".png"));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const maxW = 560, maxH = name === "poster" || name === "photo" ? 420 : 380;
  let dw = maxW, dh = Math.round(maxW * h / w);
  if (dh > maxH) { dh = maxH; dw = Math.round(maxH * w / h); }
  return [
    new Paragraph({ alignment: C, keepNext: true, spacing: { before: 120 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: dw, height: dh },
        altText: { title: caption, description: caption, name } })] }),
    para(`${fNo}-сурет – ${caption}`, { noIndent: true, align: C, before: 60, line: 240, keepNext: true }),
    para(`Дереккөзі: ${src}.`, { noIndent: true, align: C, after: 160, line: 240, size: SZ_T }),
  ];
}
const h1 = text => new Paragraph({
  heading: HeadingLevel.HEADING_1, alignment: C, pageBreakBefore: true, spacing: { after: 240, line: LINE }, keepNext: true,
  children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
});
const h2 = text => new Paragraph({
  heading: HeadingLevel.HEADING_2, alignment: AlignmentType.LEFT, indent: { firstLine: Math.round(1.25 * CM) },
  spacing: { before: 240, after: 120, line: LINE }, keepNext: true,
  children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
});
const h3 = text => new Paragraph({
  alignment: AlignmentType.LEFT, indent: { firstLine: Math.round(1.25 * CM) },
  spacing: { before: 120, line: LINE }, keepNext: true,
  children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
});

// ---------- Title page (Appendix 8) ----------
const line = (t, o = {}) => new Paragraph({ alignment: o.align ?? C, spacing: { line: 276, before: o.before || 0, after: o.after || 0 },
  indent: o.indent, children: runs(t, o) });
const L = AlignmentType.LEFT;
const title = [
  line("Қазақстан Республикасы Оқу-ағарту министрлігі"),
  line("«Дарын» республикалық ғылыми-практикалық орталығы", { after: 480 }),
  line("2-8 сынып оқушыларына арналған «Зерде» республикалық зерттеу"),
  line("жобалары мен шығармашылық жұмыстарының конкурсы", { after: 720 }),
  line("Тіркеу коды (шифр) ____________________________"),
  line("(«Дарын» РҒПО қызметкері толтырады)", { size: 20, after: 480 }),
  line("Бағыт: қоғамдық-гуманитарлық бағыт (ҚГБ)", { align: L, after: 200 }),
  line("Секцияның атауы: Қазақ тілі және әдебиеті", { align: L, after: 200 }),
  line("Жас санаты:", { align: L }),
  line("□ «Бастауыш буын» (2-4 сынып)", { align: L }),
  line("☒ «Орта буын» (5-8 сынып)", { align: L, after: 200 }),
  line("Жобаның түрі: жеке, шығармашылық-зерттеу жұмысы", { align: L, after: 600 }),
  line("Жобаның тақырыбы", { bold: true }),
  line("«ҚАЗАҚ ЕРТЕГІСІНІҢ ЖАРНАМАСЫ»", { bold: true, size: 32 }),
  line("(«Ер Төстік» ертегісі мысалында)", { after: JURY ? 2400 : 900 }),
];
if (!JURY) title.push(
  line("Қатысушының Т.А.Ә.: ⟦_____________________⟧, 5-сынып", { align: L, after: 360 }),
  line("Ғылыми жетекшінің Т.А.Ә.: ⟦_____________________⟧", { align: L, after: 1100 }),
);
title.push(line("2026-2027 оқу жылы"));

// ---------- TOC ----------
const toc = [new Paragraph({ alignment: C, pageBreakBefore: true, spacing: { after: 240 },
  children: [new TextRun({ text: "МАЗМҰНЫ", font: FONT, size: SZ, bold: true })] })];
for (const l of content) {
  const m = l.match(/^(#{1,2}) (.*)$/);
  if (!m) continue;
  const lvl = m[1].length, t = m[2];
  toc.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W, leader: "dot" }],
    indent: lvl === 2 ? { left: 400 } : undefined, spacing: { line: 300, after: 20 },
    children: [new TextRun({ text: t, font: FONT, size: SZ }),
      new TextRun({ text: `\t${TOC_PAGES[t] ?? ""}`, font: FONT, size: SZ })],
  }));
}

// ---------- Body ----------
const body = [];
for (const l of content) {
  let m;
  if ((m = l.match(/^# (.*)$/))) body.push(h1(m[1]));
  else if ((m = l.match(/^## (.*)$/))) body.push(h2(m[1]));
  else if ((m = l.match(/^### (.*)$/))) body.push(h3(m[1]));
  else if ((m = l.match(/^> (.*)$/))) body.push(para(m[1], { italics: true }));
  else if ((m = l.match(/^-- (.*)$/))) body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, numbering: { reference: "dash", level: 0 },
    spacing: { line: LINE }, children: runs(m[1]) }));
  else if ((m = l.match(/^!(fig|ph) (\w+) \| (.*?) \| (.*)$/))) body.push(...figBlock(m[2], m[3], m[4]));
  else if ((m = l.match(/^!table (\w+)/))) body.push(...tableBlock(m[1]));
  else if ((m = l.match(/^!note (.*)$/))) body.push(para(m[1], { italics: true }));
  else if (l.startsWith("!references")) order.forEach((k, i) => body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 }, spacing: { line: LINE, after: 60 },
    tabStops: [{ type: TabStopType.LEFT, position: 567 }],
    children: runs(`${i + 1}.\t${REFS[k]}`) })));
  else body.push(para(l));
}

const doc = new Document({
  creator: "", title: "Қазақ ертегісінің жарнамасы",
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
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(1.25 * CM) + 360, hanging: 360 } } } }] }] },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MT, bottom: MB, left: ML, right: MR } },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: C, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ })] })] }),
      first: new Footer({ children: [] }),
    },
    children: [...title, ...toc, ...body],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("written", OUT, "refs:", order.length); });
