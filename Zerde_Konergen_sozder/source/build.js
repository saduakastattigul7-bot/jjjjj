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
const TITLE = "Қазақ тіліндегі көнерген сөздердің қазіргі қолданыстағы орны";
const PH = "[___]";

// ---------- Дереккөздер ----------
const REFS = {
  zerde: "2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсын ұйымдастыру және өткізу ережелері. – «Дарын» РҒПО директорының 2026 жылғы 9 қыркүйектегі № 153 бұйрығымен бекітілген.",
  til_zany: "Қазақстан Республикасындағы тіл туралы: Қазақстан Республикасының 1997 жылғы 11 шілдедегі № 151-I Заңы.",
  bolganbai: "Болғанбайұлы Ә., Қалиұлы Ғ. Қазіргі қазақ тілінің лексикологиясы мен фразеологиясы. – Алматы: Санат, 1997.",
  kenesbaev: "Кеңесбаев І., Мусабаев Ғ. Қазіргі қазақ тілі: лексика, фонетика. – Алматы: Мектеп, 1975.",
  shansky: "Шанский Н.М. Лексикология современного русского языка. – М.: Просвещение, 1972.",
  syzdykova: "Сыздықова Р. Сөздер сөйлейді. – Алматы: Жалын, 1980.",
  kaidar: "Қайдар Ә. Қазақтар ана тілі әлемінде: этнолингвистикалық сөздік. 1-том: Адам. – Алматы: Дайк-Пресс, 2009.",
  qats: "Қазақ әдеби тілінің сөздігі. Он бес томдық. – Алматы: А. Байтұрсынұлы атындағы Тіл білімі институты, 2006–2011.",
  termincom: "Қазақстан Республикасы Үкіметі жанындағы Республикалық терминология комиссиясы бекіткен терминдер [Электрондық ресурс]. – URL: https://termincom.kz",
  sozdikqor: "Сөздікқор: қазақ тілінің бірыңғай электрондық сөздік қоры [Электрондық ресурс]. – URL: https://sozdikqor.kz",
  egemen: "Egemen Qazaqstan: республикалық газет [Электрондық ресурс]. – URL: https://egemen.kz",
};

// ---------- Деректерден алынатын мәндер ----------
const f1 = v => (v === null || v === undefined || v === "") ? PH : String(Math.round(v * 10) / 10).replace(".", ",");
const GNAME = Object.fromEntries(D.groups);
const KNAME = Object.fromEntries(D.kinds.map(k => [k[0], k[1]]));
const STNAME = Object.fromEntries(D.status.map(s => [s[0], s[1]]));
const CTXNAME = Object.fromEntries(D.contexts.map(s => [s[0], s[1]]));
const GENNAME = Object.fromEntries(D.gens);
const GEN_SHORT = { "О": "оқушылар", "Ә": "ата-аналар", "Ү": "ата-әжелер" };
const V = { n_per_gen: D.n_per_gen, n_words: D.words.length };
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];
const q = w => `«${w}»`;
const W = HAS ? N.words : [];
const WB = Object.fromEntries(W.map(x => [x.word, x]));
if (HAS) {
  const m = N.marks || {};
  const noMark = W.filter(x => x.mark === "белгі жоқ").map(x => q(x.word));
  V.marks_text = `Мен 36 сөзді «Қазақ әдеби тілінің сөздігінен» тексердім. Олардың ${m["көн."] || 0}-інің жанында «көн.» (көнерген), ${m["тар."] || 0}-інің жанында «тар.» (тарихи) деген белгі бар, ${m["белгі жоқ"] || 0} сөзде белгі жоқ, ал ${m["сөздікте жоқ"] || 0} сөз сөздікте мүлде табылмады.` +
    (noMark.length ? ` Белгісі жоқ сөздер: ${listJoin(noMark)}. Яғни сөздік бұл сөздерді көнерген деп емес, жалпы қолданыстағы сөз деп санайды.` : "");
  const top = W.filter(x => x.h1 !== null).sort((a, b) => b.hits - a.hits).slice(0, 5);
  const FULL = N.media_filled === D.words.length;  // мәртебе мен болжам тек барлық сөз тексерілгенде шығарылады
  V.media_text = `Мен ${N.media_filled} сөзді egemen.kz сайтынан тексердім, барлығы шамамен ${N.hits_total} нәтиже шықты. Ең жиі кездескен сөздер: ${listJoin(top.map(x => `${q(x.word)} (${x.hits})`))}.` +
    (N.zero_hits.length ? ` Ал ${listJoin(N.zero_hits.map(q))} соңғы бір жылда бірде-бір рет кездеспеді.` : "");
  const c = N.ctx || {};
  V.ctx_text = `Ең жаңа мысалдардың мәнмәтінін талдағанда, сөз ${c["Ж"] || 0} мысалда жаңа мағынада, ${c["А"] || 0} мысалда атау ішінде, ${c["Т"] || 0} мысалда тарихи-мәдени мәтінде және ${c["Б"] || 0} мысалда бейнелі мағынада қолданылғанын анықтадым.`;
  if (N.n_people) {
    const ga = N.gen_avg, gens = Object.keys(ga).filter(g => ga[g] !== null).sort((a, b) => ga[b] - ga[a]);
    const oRank = D.test_words.map(w => [w, N.test[w]["О"].ok]).filter(([, v]) => v !== null).sort((a, b) => b[1] - a[1]);
    V.test_text = `Тестке ${N.n_people} адам қатысты: ${D.gens.map(([g]) => `${GEN_SHORT[g]} – ${N.n_by_gen[g]}`).join(", ")}. ` +
      `Көне сөздерді ең жақсы ${GEN_SHORT[gens[0]]} білді: олар сөздердің орта есеппен ${f1(ga[gens[0]])}%-ын дұрыс түсіндірді. Бұл көрсеткіш оқушыларда – ${f1(ga["О"])}%, ата-аналарда – ${f1(ga["Ә"])}%, ата-әжелерде – ${f1(ga["Ү"])}% (5-кесте, 6-сурет).` +
      (oRank.length ? ` Оқушылар ең жақсы білетін сөз – ${q(oRank[0][0])} (${f1(oRank[0][1])}%), ал ең қиыны – ${q(oRank[oRank.length - 1][0])} (${f1(oRank[oRank.length - 1][1])}%).` : "");
    V.concl_test = `Үш буынның ішінде көне сөздерді ең жақсы ${GEN_SHORT[gens[0]]} біледі (${f1(ga[gens[0]])}%), оқушыларда бұл көрсеткіш – ${f1(ga["О"])}%.`;
  }
  const byS = Object.fromEntries(D.status.map(([s]) => [s, W.filter(x => x.status === s).map(x => x.word)]));
  if (FULL) V.status_text = D.status.map(([s, n]) => `${n} сөздер – ${byS[s].length}${byS[s].length ? ": " + listJoin(byS[s].map(q)) : ""}.`).join(" ");
  V.concl_media = `«Egemen Qazaqstan» сайтында ${N.media_filled} сөз соңғы бір жылда шамамен ${N.hits_total} рет кездесті; ең жиісі – ${listJoin(top.slice(0, 3).map(x => q(x.word)))}` + (N.zero_hits.length ? `, ал ${N.zero_hits.length} сөз мүлде кездеспеді.` : ".");
  if (FULL) V.concl_status = `Мәртебе бойынша: ${D.status.map(([s, n]) => `${n.toLowerCase()} – ${byS[s].length}`).join(", ")} сөз.`;
  // Болжамды тексеру
  const gs = N.by_group_status;
  const jA = gs["БИЛ"]["Ж"] + gs["ҚОҒ"]["Ж"], tB = gs["ТҰР"]["Т"] + gs["ТҰР"]["Ұ"];
  const okA = FULL ? jA >= 7 : null, okB = FULL ? tB >= 4 : null;
  const ga = N.gen_avg || {};
  const hasT = N.n_people && ga["О"] !== null && ga["Ә"] !== null && ga["Ү"] !== null;
  const okC = hasT ? (ga["Ү"] > ga["Ә"] && ga["Ә"] > ga["О"]) : null;
  const parts = [okA, okB, okC].filter(v => v !== null), nOk = parts.filter(Boolean).length;
  if (FULL) V.hyp_short = nOk === parts.length ? "Болжамым расталды." : nOk === 0 ? "Болжамым расталмады." : "Болжамым ішінара расталды.";
  if (FULL) V.hyp_text = `Болжамымды үш бөлік бойынша тексердім. Біріншіден, билік пен қоғамдық өмірге қатысты 12 сөздің ${jA}-і қайта жанданған болып шықты – ${okA ? "болжамның бұл бөлігі расталды" : "яғни олардың көбі қайта жанданды деген болжамым расталмады"}. ` +
    `Екіншіден, киім мен тұрмысқа қатысты 6 сөздің ${tB}-і тарихи мәтінде ғана қалған немесе ұмыт болып барады – ${okB ? "бұл бөлігі де расталды" : "бұл бөлігі расталмады"}. ` +
    (hasT ? `Үшіншіден, көне сөздерді дұрыс түсіндіру үлесі ата-әжелерде ${f1(ga["Ү"])}%, ата-аналарда ${f1(ga["Ә"])}%, оқушыларда ${f1(ga["О"])}% болды – ${okC ? "яғни буын үлкен болған сайын көне сөзді жақсы біледі деген болжамым расталды" : "яғни буындар арасындағы айырмашылық мен күткендей болмады"}. ` : "") + V.hyp_short;
  if (FULL) V.annot_results = `Түсіндірме сөздікте 36 сөздің 19-ында «көнерген» не «тарихи» белгісі жоқ. egemen.kz сайтында сөздер шамамен ${N.hits_total} рет кездесті, тестке ${N.n_people || PH} адам қатысты. ${V.concl_status} ${V.hyp_short}`;
}
for (const [k, txt] of Object.entries({
  marks_text: "[Сөздіктегі белгілер бойынша нәтижені жазыңыз: неше сөз «көн.», неше сөз «тар.» деп белгіленген.]",
  media_text: "[egemen.kz нәтижесін жазыңыз: ең жиі және мүлде кездеспеген сөздер.]", ctx_text: "[Мысалдардың мәнмәтіні бойынша нәтиже.]",
  test_text: "[Үш буынның нәтижесін салыстырыңыз.]",
  status_text: "[Әр мәртебеге қандай сөздер кірді?]", hyp_text: "[Болжамның расталғанын не расталмағанын деректермен көрсетіңіз.]",
  hyp_short: "[Болжам расталды / ішінара расталды / расталмады.]", concl_media: "[БАҚ мониторингі бойынша қорытынды.]",
  concl_test: "[Тест бойынша қорытынды.]", concl_status: "[Мәртебе бойынша қорытынды.]",
  annot_results: "[Негізгі нәтижелерді 2–3 сөйлеммен жазыңыз.]",
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
const HL_RE = /(\[___\]|\[(?!Мәтін\]|Электрондық ресурс\]|баспаға)[А-ЯӘҒҚҢӨҰҮҺІа-яәғқңөұүһі][^\]]*\])/;
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
const pc = v => v === null || v === undefined ? PH : f1(v);
const TABLES = {
  sozder: () => [caption("Зерттеуге алынған көнерген сөздер"), mkTable([30, 35, 35], ["Тақырыптық топ", "Тарихи сөздер", "Архаизмдер"],
    D.groups.map(([g, n]) => [n, D.words.filter(w => w[1] === g && w[2] === "Т").map(w => w[0]).join(", ") || "–",
      D.words.filter(w => w[1] === g && w[2] === "А").map(w => w[0] + (g === "СИН" ? ` (${w[3]})` : "")).join(", ") || "–"])), gap()],
  status: () => [caption("Көне сөздің бүгінгі «тірлік мәртебесі»"), mkTable([10, 28, 62], ["Код", "Мәртебе", "Сипаттамасы"], D.status, { centerCols: [0], boldCol0: true }), gap()],
  ctx: () => [caption("БАҚ-тағы мысалдың мәнмәтінін кодтау"), mkTable([10, 28, 62], ["Код", "Мәнмәтін", "Сипаттамасы және мысал"], D.contexts, { centerCols: [0], boldCol0: true }), gap()],
  media_grp: () => [caption("Көне сөздердің egemen.kz сайтында кездесуі, тақырыптық топтар бойынша"),
    mkTable([34, 18, 26, 22], ["Тақырыптық топ", "Нәтиже саны (барлығы)", "Ең жиі сөз", "«Жаңа мағынада» мысал"],
      D.groups.map(([g, n]) => {
        if (!HAS) return [n, PH, PH, PH];
        const ws = W.filter(x => x.group === g && x.h1 !== null), t = [...ws].sort((a, b) => b.hits - a.hits)[0];
        if (!t) return [n, PH, PH, PH];
        return [n, String(ws.reduce((a, x) => a + x.hits, 0)), `${t.word} (${t.hits})`, String(ws.filter(x => x.c1 === "Ж").length)];
      }), { centerCols: [1, 3] }), gap()],
  test: () => [caption("Көне сөздерді дұрыс түсіндіргендер үлесі, %"), mkTable([31, 23, 23, 23], ["Сөз", ...D.gens.map(([g]) => GEN_SHORT[g][0].toUpperCase() + GEN_SHORT[g].slice(1))],
    [...D.test_words.map(w => [w, ...D.gens.map(([g]) => HAS && N.n_people ? pc(N.test[w][g].ok) : PH)]),
      ["Орташа", ...D.gens.map(([g]) => HAS && N.n_people ? pc(N.gen_avg[g]) : PH)]], { centerCols: [1, 2, 3], boldRow: r => r[0] === "Орташа" }), gap()],
  status_list: () => [caption("Көне сөздердің бүгінгі мәртебесі"), mkTable([26, 10, 64], ["Мәртебе", "Саны", "Сөздер"],
    D.status.map(([s, n]) => { const ws = W.filter(x => x.status === s).map(x => x.word); return [`${s} – ${n}`, HAS ? String(ws.length) : PH, HAS ? (ws.join(", ") || "–") : PH]; }),
    { centerCols: [1] }), gap()],
};


function figBlock(name, cap) {
  const buf = fs.readFileSync(path.join(DIR, "figs", name + ".png"));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const maxW = 600, maxH = 520;
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
const ANNOT = "Жұмыстың мақсаты – қазақ тіліндегі көнерген сөздердің түсіндірме сөздіктегі белгісін, бүгінгі баспасөзде және үш буынның тілінде қолданылуын зерттеп, әр сөздің бүгінгі мәртебесін анықтау. Өзектілігі: тәуелсіздік жылдары кейбір көне сөздер («әкім», «теңге», «Мәжіліс») тілге қайта оралды, ал кейбірі ұмыт болып барады. Зерттеу сұрағы: көнерген сөздер бүгін қайда және қандай мағынада қолданылады және оларды әр буын қаншалықты біледі? Болжам: билік пен қоғамдық өмірге қатысты көне сөздердің көбі қайта жанданған, тұрмыстық тарихи сөздер тек тарихи мәтінде қалған, ал көне сөздерді ата-әжелер оқушылардан жақсы біледі. Әдістер: сөздікпен тексеру, БАҚ мониторингі, үш буынға сөз тану тесті, салыстыру. Алты тақырыптық топтан алынған 36 көнерген сөз зерттелді. {{annot_results}} Нәтижесінде «Көне сөз – жаңа тыныс» мини-сөздігі жасалды.";
const annotWords = fill(ANNOT).split(/\s+/).filter(w => /[\p{L}\d]/u.test(w)).length;
if (annotWords > 200) throw new Error("Аннотация 200 сөзден асты: " + annotWords);
console.log("Аннотация сөз саны:", annotWords);
const annot = [h1("АННОТАЦИЯ"), para(ANNOT), para("Кілт сөздер: көнерген сөздер, тарихи сөздер, архаизмдер, қайта жандану, түсіндірме сөздік, БАҚ мониторингі.", { before: 120 })];

// ---------- Қосымшалар ----------
const APPX = [
  { id: "А", title: "БАҚ мониторингінің кестесі" },
  { id: "Ә", title: "Сөз тану тестінің парағы" },
  { id: "Б", title: "Келісім парақтары" },
  { id: "В", title: "«Көне сөз – жаңа тыныс» мини-сөздігі" },
];
function appendix() {
  const out = [];
  const head = (a) => { out.push(h1(`${a.id} ҚОСЫМШАСЫ`)); out.push(para(a.title, { noIndent: true, align: C, bold: true, after: 200 })); };
  // А
  head(APPX[0]);
  out.push(para(`Нәтиже саны – Google іздеу жүйесінде egemen.kz сайты бойынша соңғы бір жылдағы нәтиже (шамамен). Мәнмәтін: Ж – жаңа мағынада, А – атау ішінде, Т – тарихи-мәдени мәтінде, Б – бейнелі мағынада. Мәртебе: Ж – қайта жанданған, А – атауда сақталған, Т – тарихи мәтінде ғана, Ұ – ұмыт болып барады.`, { after: 120 }));
  const nz = v => v === null || v === undefined ? "–" : String(v);
  out.push(mkTable([5, 13, 12, 9, 50, 11], ["№", "Сөз", "Нәтиже саны", "Код", "Ең жаңа мысал", "Мәртебе"],
    D.words.map(([w], i) => { const x = WB[w]; return x ? [String(i + 1), w, nz(x.h1), x.c1 || "–", x.example ? `«${x.example.slice(0, 220)}${x.example.length > 220 ? "…" : ""}»` : "–", x.status || "–"]
      : [String(i + 1), w, PH, PH, PH, PH]; }), { size: 20, centerCols: [0, 2, 3, 5] }));
  // Ә
  head(APPX[1]);
  out.push(para("Құрметті қатысушы! Мен қазақ тіліндегі көне сөздер туралы зерттеу жүргізіп жатырмын. Сізге 12 сөз айтамын, әрқайсысының мағынасын өз сөзіңізбен түсіндіріп беріңізші. Білмесеңіз, «білмеймін» деп айта беріңіз – бұл емтихан емес. Есіміңіз жазылмайды.", { after: 60 }));
  out.push(para("Қатысушының коды: ______     Буыны: ☐ оқушы   ☐ ата-ана   ☐ ата-әже", { noIndent: true, after: 120 }));
  out.push(mkTable([5, 17, 50, 28], ["№", "Сөз", "Мағынасын түсіндіруі (зерттеуші толтырады)", "Соңғы бір жылда естідіңіз бе / қолдандыңыз ба?"],
    D.test_words.map((w, i) => [String(i + 1), w, "☐ 2 – дұрыс   ☐ 1 – шамамен   ☐ 0 – білмейді", "☐ иә     ☐ жоқ"]), { size: 22, centerCols: [0] }));
  out.push(para("Рахмет!", { noIndent: true, align: C, before: 120 }));
  // Б
  head(APPX[2]);
  out.push(para("1. Ата-ананың (заңды өкілінің) келісімі", { noIndent: true, bold: true, after: 120 }));
  [
    "Құрметті ата-ана!",
    `Мектебіміздің [__] сынып оқушысы «${TITLE}» тақырыбында зерттеу жобасын орындап жатыр. Зерттеу аясында қатысушыларға 12 көне сөздің мағынасын түсіндіру ұсынылады (5–7 минут).`,
    "Баланың аты-жөні жазылмайды, жауаптар тек жалпы сан түрінде қолданылады және үшінші тұлғаларға берілмейді. Қатысу ерікті, бала кез келген уақытта бас тарта алады.",
  ].forEach((t, i) => out.push(para(t, { noIndent: i === 0, after: 100 })));
  ["Баланың аты-жөні: ______________________________   Сыныбы: ______", "☐ Келісемін        ☐ Келіспеймін",
    "Ата-ананың (заңды өкілінің) аты-жөні: ______________________________", "Қолы: ______________          Күні: «___» __________ 2026 ж."]
    .forEach(t => out.push(para(t, { noIndent: true, align: L, before: 160 })));
  out.push(para("2. Ересек қатысушының келісімі", { noIndent: true, bold: true, before: 400, after: 120 }));
  out.push(para(`Мен «${TITLE}» тақырыбындағы оқушы зерттеуіне ерікті түрде қатысуға келісемін. Менің аты-жөнім жазылмайтыны, жауаптарым тек жалпы сан түрінде қолданылатыны және кез келген уақытта бас тарта алатыным маған түсіндірілді.`, { after: 100 }));
  ["Аты-жөні: ______________________________", "Қолы: ______________          Күні: «___» __________ 2026 ж."]
    .forEach(t => out.push(para(t, { noIndent: true, align: L, before: 160 })));
  // В
  head(APPX[3]);
  out.push(para("Мини-сөздікте әр сөздің мағынасы, түрі, бүгінгі мәртебесі және бүгінгі баспасөзден алынған мысалы берілген.", { after: 120 }));
  D.groups.forEach(([g, gname]) => {
    out.push(para(gname, { noIndent: true, bold: true, before: 160, after: 60, keepNext: true }));
    D.words.filter(w => w[1] === g).forEach(([w, , k, m]) => {
      const x = WB[w];
      const st = x && x.status ? ` Мәртебесі: ${STNAME[x.status].toLowerCase()}.` : (HAS ? "" : " Мәртебесі: [___].");
      const ex = x && x.example ? ` Мысал: «${x.example.slice(0, 200)}${x.example.length > 200 ? "…" : ""}»` : "";
      out.push(para(`**${w}** (${KNAME[k].toLowerCase()}) – ${((x && x.meaning) || m).replace(/[.\s]+$/, "")}.${st}${ex}`, { size: 24, line: 276, after: 60 }));
    });
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