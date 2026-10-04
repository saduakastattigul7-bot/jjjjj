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
const TITLE = "Ертегілер мен жырлардағы ұлттық ойындар";
const PH = "[___]";

// ---------- Дереккөздер ----------
const REFS = {
  zerde: "2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсын ұйымдастыру және өткізу ережелері. – «Дарын» РҒПО директорының 2026 жылғы 9 қыркүйектегі № 153 бұйрығымен бекітілген.",
  unesco_asyk: "UNESCO. Kazakh traditional Assyk games. – Decision of the Intergovernmental Committee 12.COM 11.B.18. – 2017. – URL: https://ich.unesco.org/en/RL/kazakh-traditional-assyk-games-01086",
  nomad: "V Дүниежүзілік көшпенділер ойындары. – Астана, 2024 жылғы 8–13 қыркүйек.",
  birtutas: "Оқушыларды тәрбиелеудің «Біртұтас тәрбие» бағдарламасы. – Астана: ҚР Оқу-ағарту министрлігі, 2023.",
  huizinga: "Хейзинга Й. Homo Ludens. В тени завтрашнего дня. – М.: Прогресс, 1992. – 464 с.",
  vygotsky: "Выготский Л.С. Игра и её роль в психическом развитии ребёнка // Вопросы психологии. – 1966. – № 6. – С. 62–76.",
  sagyndykov: "Сағындықов Е. Қазақтың ұлттық ойындары. – Алматы: Рауан, 1991.",
  totenaev: "Төтенаев Б. Қазақтың ұлттық ойындары. – Алматы: Қайнар, 1994.",
  kaskabasov: "Қасқабасов С. Қазақтың халық прозасы. – Алматы: Ғылым, 1984.",
  konyratbaev: "Қоңыратбаев Ә. Қазақ эпосы және түркология. – Алматы: Ғылым, 1987.",
  auezov: "Әуезов М. Әдебиет тарихы. – Алматы: Ана тілі, 1991.",
  propp: "Пропп В.Я. Морфология сказки. – Л.: Academia, 1928.",
};

// ---------- Деректерден алынатын мәндер ----------
const f1 = v => (v === null || v === undefined || v === "") ? PH : String(Math.round(v * 10) / 10).replace(".", ",");
const TYPE_NAME = Object.fromEntries(D.types);
const FUNC_NAME = Object.fromEntries(D.funcs.map(f => [f[0], f[1]]));
const V = {};
const sumG = o => o["Ертегі"] + o["Жыр"];
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];
if (HAS) {
  for (const k of ["works_read", "ep_total", "ep_Ертегі", "ep_Жыр", "works_with_games", "games_found"]) V[k] = N[k];
  if (N.n_resp) V.n_resp = N.n_resp;
  const games = Object.entries(N.by_game).sort((a, b) => sumG(b[1]) - sumG(a[1]));
  const top = games.filter(([, o]) => sumG(o) > 0).slice(0, 3);
  const zero = games.filter(([, o]) => sumG(o) === 0).map(([g]) => `«${g.toLowerCase()}»`);
  V.top_games = top.length ? `Ең жиі кездескен ойындар: ${listJoin(top.map(([g, o]) => `«${g.toLowerCase()}» (${sumG(o)} эпизод)`))}.` + (zero.length ? ` Ал ${listJoin(zero)} бірде-бір шығармада кездеспеді.` : "") : "";
  V.other_games = N.other_games.length ? `Тізімде жоқ ойындар да кездесті: ${listJoin(N.other_games.map(g => "«" + g.toLowerCase() + "»"))}.` : "";
  const tE = Object.entries(N.by_type).sort((a, b) => b[1]["Ертегі"] - a[1]["Ертегі"])[0];
  const tJ = Object.entries(N.by_type).sort((a, b) => b[1]["Жыр"] - a[1]["Жыр"])[0];
  V.type_compare = `Ертегілерде ең жиі кездескені – ${TYPE_NAME[tE[0]].toLowerCase()} (${tE[1]["Ертегі"]} эпизод), ал жырларда – ${TYPE_NAME[tJ[0]].toLowerCase()} (${tJ[1]["Жыр"]} эпизод).`;
  const fE = Object.entries(N.by_func).sort((a, b) => b[1]["Ертегі"] - a[1]["Ертегі"])[0];
  const fJ = Object.entries(N.by_func).sort((a, b) => b[1]["Жыр"] - a[1]["Жыр"])[0];
  V.func_compare = `Ертегілерде ойын көбіне «${FUNC_NAME[fE[0]].toLowerCase()}» қызметін атқарады (${fE[1]["Ертегі"]} эпизод), ал жырларда – «${FUNC_NAME[fJ[0]].toLowerCase()}» қызметін (${fJ[1]["Жыр"]} эпизод).`;
  V.concl_games = V.top_games;
  V.concl_types = `Ойын түрлері бойынша: ${V.type_compare.charAt(0).toLowerCase()}${V.type_compare.slice(1)}`;
  V.concl_func = `Ойынның сюжеттегі қызметі бойынша: ${V.func_compare.charAt(0).toLowerCase()}${V.func_compare.slice(1)}`;
  const bala = N.by_type["БАЛА"], at = N.by_type["АТ"], kush = N.by_type["КҮШ"], test = N.by_func["С"];
  const okA = bala["Ертегі"] >= bala["Жыр"], okB = (at["Жыр"] + kush["Жыр"]) >= (at["Ертегі"] + kush["Ертегі"]), okC = fJ[0] === "С";
  const nOk = [okA, okB, okC].filter(Boolean).length;
  V.hyp_short = nOk === 3 ? "Болжамымыз расталды." : nOk === 0 ? "Болжамымыз расталмады." : "Болжамымыз ішінара расталды.";
  V.hyp_text = `Болжамымызды үш бөлік бойынша тексердік. Біріншіден, балалар ойындары ертегіде ${bala["Ертегі"]}, жырда ${bala["Жыр"]} рет кездесті – ${okA ? "болжамның бұл бөлігі расталды" : "болжамның бұл бөлігі расталмады"}. Екіншіден, ат ойындары мен күш сынасу ойындары жырда ${at["Жыр"] + kush["Жыр"]}, ертегіде ${at["Ертегі"] + kush["Ертегі"]} рет кездесті – ${okB ? "бұл да расталды" : "бұл расталмады"}. Үшіншіден, жырда ойынның ең жиі қызметі «${FUNC_NAME[fJ[0]].toLowerCase()}» болды – ${okC ? "яғни жырда ойын шынымен де көбіне қаһарманды сынаудың құралы" : "яғни жырда ойын көбіне сынақ емес, басқа қызмет атқарады"}. ${V.hyp_short}`;
  if (N.n_resp) {
    const sv = Object.entries(N.survey).filter(([, o]) => o.know !== null).sort((a, b) => b[1].know - a[1].know);
    const hi = sv[0], lo = sv[sv.length - 1];
    const pl = Object.entries(N.survey).filter(([, o]) => o.played !== null).sort((a, b) => b[1].played - a[1].played)[0];
    V.survey_text = `Оқушылар ең жақсы білетін ойын – «${hi[0].toLowerCase()}» (${f1(hi[1].know)}%), ең аз білетіні – «${lo[0].toLowerCase()}» (${f1(lo[1].know)}%). Ең көп ойнаған ойыны – «${pl[0].toLowerCase()}» (${f1(pl[1].played)}%).`;
    V.concl_survey = `Сауалнамаға ${N.n_resp} оқушы қатысты. ${V.survey_text}`;
  }
}
for (const [k, txt] of Object.entries({
  top_games: "[Ең жиі және ең сирек кездескен ойындарды атаңыз.]", other_games: "", type_compare: "[Ертегі мен жырдағы ойын түрлерін салыстырыңыз.]",
  func_compare: "[Ойынның ертегі мен жырдағы қызметін салыстырыңыз.]", hyp_text: "[Болжамның расталғанын не расталмағанын деректермен көрсетіңіз.]",
  hyp_short: "[Болжам расталды / ішінара расталды / расталмады.]", survey_text: "[Ең жақсы және ең аз белгілі ойындарды атаңыз.]",
  concl_games: "[Ең жиі кездескен ойындар.]", concl_types: "[Ойын түрлері бойынша қорытынды.]", concl_func: "[Ойынның қызметі бойынша қорытынды.]",
  concl_survey: "[Сауалнама бойынша қорытынды.]",
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
const HL_RE = /(\[___\]|\[(?!Мәтін\]|баспаға)[А-ЯӘҒҚҢӨҰҮҺІа-яәғқңөұүһі][^\]]*\])/;
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
  oiyndar: () => [caption("Зерттеуге алынған ұлттық ойындар"), mkTable([5, 17, 22, 56], ["№", "Ойын", "Түрі", "Қысқаша ережесі"],
    D.games.map(([g, t, r], i) => [String(i + 1), g, TYPE_NAME[t], r]), { centerCols: [0], size: 22 }), gap()],
  funcs: () => [caption("Ойынның шығармадағы қызметін кодтау жүйесі"), mkTable([10, 25, 65], ["Код", "Қызметі", "Сипаттамасы"], D.funcs, { centerCols: [0], boldCol0: true }), gap()],
  korpus: () => {
    const w = HAS ? N.works : D.corpus.map(([n, g]) => ({ name: n, genre: g, edition: "", read: false }));
    return [caption("Зерттеу материалы: шығармалар"), mkTable([6, 40, 14, 40], ["№", "Шығарма", "Жанры", "Оқыған басылым"],
      w.map((x, i) => [String(i + 1), `«${x.name}»`, x.genre, x.edition || "[кітап, баспа, жылы]"]), { centerCols: [0, 2] }), gap()];
  },
  oyn_sany: () => [caption("Ойындардың ертегі мен жырда кездесуі, эпизод саны"), mkTable([40, 20, 20, 20], ["Ойын", "Ертегіде", "Жырда", "Барлығы"],
    [...D.games.map(([g]) => { const o = HAS ? N.by_game[g] : null; return [g, o ? String(o["Ертегі"]) : PH, o ? String(o["Жыр"]) : PH, o ? String(sumG(o)) : PH]; }),
      ["Барлығы", HAS ? String(N["ep_Ертегі"]) : PH, HAS ? String(N["ep_Жыр"]) : PH, HAS ? String(N.ep_total) : PH]],
    { centerCols: [1, 2, 3], boldRow: r => r[0] === "Барлығы" }), gap()],
  turi: () => [caption("Ойын түрлерінің ертегі мен жырда кездесуі, эпизод саны"), mkTable([40, 20, 20, 20], ["Ойын түрі", "Ертегіде", "Жырда", "Барлығы"],
    D.types.map(([c, n]) => { const o = HAS ? N.by_type[c] : null; return [n, o ? String(o["Ертегі"]) : PH, o ? String(o["Жыр"]) : PH, o ? String(sumG(o)) : PH]; }),
    { centerCols: [1, 2, 3] }), gap()],
  qyzmet: () => [caption("Ойынның сюжеттегі қызметі, эпизод саны"), mkTable([40, 20, 20, 20], ["Қызметі", "Ертегіде", "Жырда", "Барлығы"],
    D.funcs.map(([c, n]) => { const o = HAS ? N.by_func[c] : null; return [`${c} – ${n}`, o ? String(o["Ертегі"]) : PH, o ? String(o["Жыр"]) : PH, o ? String(sumG(o)) : PH]; }),
    { centerCols: [1, 2, 3] }), gap()],
  mysal: () => {
    const ex = HAS ? D.funcs.map(([c]) => N.episodes.find(e => e.func === c)).filter(Boolean) : [];
    const short = q => { q = q.replace(/\s+/g, " ").trim(); return q.length > 140 ? q.slice(0, 140).replace(/\s\S*$/, "") + "…" : q; };
    const rows = ex.length ? ex.map(e => [`«${e.work}»`, e.game, e.episode + (e.quote ? ` Үзінді: «${short(e.quote)}»` : ""), `${e.func} – ${FUNC_NAME[e.func] || ""}`])
      : [1, 2, 3].map(() => ["[шығарма]", "[ойын]", "[эпизод, үзінді]", "[қызметі]"]);
    return [caption("Ойынның сюжеттегі қызметін көрсететін мысалдар"), mkTable([20, 14, 48, 18], ["Шығарма", "Ойын", "Эпизод және үзінді", "Қызметі"], rows), gap()];
  },
  saualnama: () => [caption("Сауалнама нәтижесі, оқушылар үлесі, %"), mkTable([34, 22, 22, 22], ["Ойын", "Біледі", "Ойнаған", "Ертегіден оқыған"],
    D.games.map(([g]) => { const o = HAS && N.n_resp ? N.survey[g] : null; return [g, o ? f1(o.know) : PH, o ? f1(o.played) : PH, o ? f1(o.folk) : PH]; }),
    { centerCols: [1, 2, 3] }), gap()],
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
    line("Қатысушының Т.А.Ә.: [Оқушының аты-жөні], [__] сынып", { align: L, after: 300 }),
    line("Ғылыми жетекшінің Т.А.Ә.: [Жетекшінің аты-жөні]", { align: L, after: 1000 }),
  ]),
  line("2026–2027 оқу жылы"),
];

// ---------- Аннотация ----------
const ANNOT = "Жұмыстың мақсаты – қазақ ертегілері мен жырларындағы ұлттық ойындарды анықтап, олардың шығармадағы қызметін талдау және сыныптастарымның бұл ойындарды білу деңгейін зерттеу. Өзектілігі: ұлттық ойындар – халықтың рухани мұрасы, ал бүгінгі оқушылар оларды аз ойнайды. Зерттеу сұрағы: ертегі мен жырда қандай ұлттық ойындар кездеседі, олар оқиғада қандай рөл атқарады және оқушылар оларды біле ме? Болжам: ертегілерде балалар ойындары, жырларда ат ойындары мен күш сынасу ойындары жиі кездеседі, жырда ойын көбіне қаһарманды сынайды. Әдістер: мәтінді талдау және кодтау, салыстыру, сауалнама. 11 шығарма (4 ертегі, 7 жыр) оқылып, ойын кездесетін 16 эпизод табылды, сауалнамаға {{n_resp}} оқушы қатысты. Ең жиі ойындар – бәйге мен асық, ал 12 ойынның 6-уы кездеспеді. Жырларда ат ойындары басым, ертегіде ойын көбіне сынақ, жырда батырдың қасиетін көрсетеді; асық оқиғаны бастаушы қызмет атқарады. Болжам ішінара расталды. Нәтижесінде «Ертегі-ойын» карточкалары жасалды.";
const annotWords = fill(ANNOT).split(/\s+/).filter(w => /[\p{L}\d]/u.test(w)).length;
if (annotWords > 200) throw new Error("Аннотация 200 сөзден асты: " + annotWords);
console.log("Аннотация сөз саны:", annotWords);
const annot = [h1("АННОТАЦИЯ"), para(ANNOT), para("Кілт сөздер: ұлттық ойындар, ертегі, батырлар жыры, асық, бәйге, фольклор, сауалнама.", { before: 120 })];

// ---------- Қосымшалар ----------
const APPX = [
  { id: "А", title: "Мәтін талдауының кестесі" },
  { id: "Ә", title: "Сауалнама парағы" },
  { id: "Б", title: "Ата-ананың келісім парағы" },
  { id: "В", title: "«Ертегі-ойын» карточкалары" },
];
function appendix() {
  const out = [];
  const head = (a) => { out.push(h1(`${a.id} ҚОСЫМШАСЫ`)); out.push(para(a.title, { noIndent: true, align: C, bold: true, after: 200 })); };
  // А
  head(APPX[0]);
  out.push(para("Кестеде шығармаларда ұлттық ойын кездесетін барлық эпизод берілген. Қызмет кодтары: С – сынақ, Б – сюжетті бастау, Қ – қасиетін көрсету, Т – той-думан көрінісі, Ш – шешуші сәт.", { after: 120 }));
  const eps = HAS ? N.episodes : [];
  out.push(mkTable([5, 18, 13, 14, 42, 8], ["№", "Шығарма", "Ойын", "Кейіпкер", "Эпизод және үзінді", "Код"],
    eps.length ? eps.map((e, i) => [String(i + 1), `«${e.work}» (${e.genre.toLowerCase()})`, e.game, e.hero, e.episode + (e.quote ? `\nҮзінді: «${e.quote.replace(/\s+/g, " ").trim().slice(0, 400)}${e.quote.length > 400 ? "…" : ""}»` : ""), e.func])
      : [1, 2, 3].map(i => [String(i), "[шығарма]", "[ойын]", "[кейіпкер]", "[эпизод, үзінді, беті]", "[код]"]), { size: 20, centerCols: [0, 5] }));
  // Ә
  head(APPX[1]);
  out.push(para("Құрметті оқушы! Бұл сауалнама қазақтың ұлттық ойындарын білу туралы. Атыңды жазба. Әр ойын бойынша екі сұраққа жауап бер.", { after: 60 }));
  out.push(para("Сыныбың: ______        Код (сауалнама жүргізуші толтырады): ______", { noIndent: true, after: 120 }));
  out.push(mkTable([5, 25, 40, 30], ["№", "Ойын", "Бұл ойынды білесің бе? (біреуін белгіле)", "Ол туралы ертегіден не жырдан оқыдың ба?"],
    D.games.map(([g], i) => [String(i + 1), g, "☐ 0 – білмеймін\n☐ 1 – естігенмін\n☐ 2 – ойнағанмын / көргенмін", "☐ иә     ☐ жоқ"]), { size: 22, centerCols: [0] }));
  out.push(para("Рахмет!", { noIndent: true, align: C, before: 120 }));
  // Б
  head(APPX[2]);
  [
    "Құрметті ата-ана!",
    `Мектебіміздің [__] сынып оқушысы «${TITLE}» тақырыбында зерттеу жобасын орындап жатыр. Зерттеу аясында сыныптастар арасында қазақтың ұлттық ойындарын білу туралы қысқа сауалнама (5–7 минут) жүргізіледі.`,
    "Сауалнамада баланың аты-жөні жазылмайды, жауаптар тек жалпы сан түрінде қолданылады және үшінші тұлғаларға берілмейді. Қатысу ерікті, бала кез келген уақытта бас тарта алады.",
    "Балаңыздың сауалнамаға қатысуына келісім беруіңізді сұраймыз.",
  ].forEach((t, i) => out.push(para(t, { noIndent: i === 0, after: 120 })));
  ["Баланың аты-жөні: ________________________________   Сыныбы: ______",
    "☐ Келісемін        ☐ Келіспеймін",
    "Ата-ананың (заңды өкілінің) аты-жөні: ________________________________",
    "Қолы: ______________          Күні: «___» __________ 2026 ж."].forEach(t => out.push(para(t, { noIndent: true, align: L, before: 200 })));
  // В
  head(APPX[3]);
  out.push(para("Әр карточкада ойынның ережесі және ол кездесетін ертегі немесе жыр көрсетілген. Карточкаларды қиып алып, сабақта не үзілісте пайдалануға болады.", { after: 120 }));
  D.games.forEach(([g, t, r]) => {
    const ep = HAS ? N.episodes.filter(e => e.game === g) : [];
    const folk = ep.length ? ep.slice(0, 2).map(e => `«${e.work}»: ${e.episode}`).join(" ") : (HAS ? "Зерттелген шығармаларда кездеспеді." : "[ертегі немесе жырдан эпизод]");
    out.push(mkTable([28, 72], [g.toUpperCase(), TYPE_NAME[t]], [["Ережесі", r], ["Ертегі мен жырда", folk]], { size: 22, boldCol0: true }));
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
