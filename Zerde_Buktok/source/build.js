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
const TITLE = "Буктюб пен буктокерлер: әлеуметтік желі жастардың кітап оқуына әсер ете ме?";
const PH = "[___]";

// ---------- Дереккөздер ----------
const REFS = {
  zerde: "2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсын ұйымдастыру және өткізу ережелері. – «Дарын» РҒПО директорының 2026 жылғы 9 қыркүйектегі № 153 бұйрығымен бекітілген.",
  birtutas: "Оқушыларды тәрбиелеудің «Біртұтас тәрбие» бағдарламасы. – Астана: ҚР Оқу-ағарту министрлігі, 2023.",
  forbes: "The Power of BookTok: Why TikTok’s Book Community Is Driving a New Era in Publishing [Электрондық ресурс] // Forbes / FirstWriter. – 21.04.2025. – URL: https://www.firstwriter.com/articles/?GUID=4079",
  tiktok_eu: "#BookTok community helps sell more than 50 million books across Europe [Электрондық ресурс] // TikTok Newsroom. – 19.03.2026. – URL: https://newsroom.tiktok.com/booktok-community-50-million-books",
  insidehook: "TikTok Shop Wants to Be Your New Favorite Bookstore [Электрондық ресурс] // InsideHook. – 2025. – URL: https://www.insidehook.com/?p=942920",
  pw2023: "BookTok Helped Book Sales Soar. How Long Will That Last? [Электрондық ресурс] // Publishers Weekly. – 21.08.2023. – URL: https://www.publishersweekly.com/pw/print/20230821/93014-booktok-helped-book-sales-soar-how-long-will-that-last.html",
  whitenights: "White Nights Resurges [Электрондық ресурс] // Russian Life. – 23.12.2024. – URL: https://www.russianlife.com/the-russia-file/white-nights-resurges",
  jerasa: "Jerasa S., Boffone T. BookTok 101: TikTok, Digital Literacies, and Out-of-School Reading Practices // Journal of Adolescent & Adult Literacy. – 2021. – Vol. 65, № 3. – P. 219–226.",
};

// ---------- Деректерден алынатын мәндер ----------
const f1 = v => (v === null || v === undefined || v === "") ? PH : String(Math.round(v * 10) / 10).replace(".", ",");
const pcs = v => (v === null || v === undefined) ? PH : f1(v) + "%";
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];
const q = w => `«${w}»`;
// Санға жалғау: барыс (10-ға, 13-ке) және шығыс (11-ден, 10-нан) септіктері
const NUMW = n => { n = Math.abs(Math.round(n)); return n === 0 ? "нөл" : n % 10 ? ["", "бір", "екі", "үш", "төрт", "бес", "алты", "жеті", "сегіз", "тоғыз"][n % 10]
  : n % 100 ? ["", "он", "жиырма", "отыз", "қырық", "елу", "алпыс", "жетпіс", "сексен", "тоқсан"][(n % 100) / 10] : n % 1000 ? "жүз" : "мың"; };
const backV = w => /[аоұы]/.test(w.replace(/.*[әеіөүи]/, "")) || (!/[әеіөү]/.test(w) && /[аоұы]/.test(w));
const lastCh = w => w[w.length - 1];
const dat = n => { const w = NUMW(n), b = /[аоұы][^аоұыәеіөүи]*$/.test(w), c = lastCh(w);
  return /[кқпстшфх]/.test(c) ? (b ? "қа" : "ке") : (b ? "ға" : "ге"); };
const abl = n => { const w = NUMW(n), b = /[аоұы][^аоұыәеіөүи]*$/.test(w), c = lastCh(w);
  return /[мнң]/.test(c) ? (b ? "нан" : "нен") : /[кқпстшфх]/.test(c) ? (b ? "тан" : "тен") : (b ? "дан" : "ден"); };

const V = { exp_name: (HAS && N.exp_name) || D.exp_default[0], ctrl_name: (HAS && N.ctrl_name) || D.ctrl_default[0] };
const SV = HAS ? N.survey : null;
if (HAS && N.n_resp) {
  const bc = N.by_class, src = Object.entries(N.src).sort((a, b) => (b[1] || 0) - (a[1] || 0));
  V.survey_text = `Сауалнамаға ${N.n_resp} оқушы қатысты: 7-сыныптан ${bc["7"]}, 8-сыныптан ${bc["8"]}, 9-сыныптан ${bc["9"]} оқушы. ` +
    `Оқушылардың ${pcs(SV.TIME["3 сағаттан көп"])}-ы әлеуметтік желіде күніне 3 сағаттан көп, ${pcs(SV.TIME["1–3 сағат"])}-ы 1–3 сағат уақыт өткізеді. ` +
    `Жаңа кітап туралы оқушылар көбіне «${src[0][0].toLowerCase()}» арқылы біледі (${pcs(src[0][1])}), одан кейін – «${src[1][0].toLowerCase()}» (${pcs(src[1][1])}). ` +
    `Мұғалімнің ұсынысын ${pcs(N.src["Мұғалім"])}, әлеуметтік желіні ${pcs(N.src["Әлеуметтік желі"])} оқушы атады (3-кесте, 2-сурет).`;
  V.survey_text2 = `Оқушылардың ${pcs(SV.HEARD["Иә"])}-ы «буктокер» не «буктюбер» сөзін естіген, ал ${pcs(SV.TRIED["Иә"])}-ы блогер ұсынған кітапты оқып көрген. ` +
    `Солардың ${pcs(SV.FINISH["Иә"])}-ы кітапты соңына дейін оқыған, ${pcs(SV.FINISH["Жартылай"])}-ы жартылай оқыған, ${pcs(SV.FINISH["Жоқ"])}-ы аяқтамаған (3-сурет). ` +
    `Қазақ тіліндегі кітап блогерін оқушылардың ${pcs(SV.KAZ["Иә"])}-ы ғана көреді, бірақ ${pcs(SV.VIDEO["Иә"])}-ы өзі кітап туралы бейне түсіріп көрер еді. ` +
    `Әлеуметтік желіні кітап оқуға «көмектеседі» деп ${pcs(SV.OPIN["Көмектеседі"])}, «кедергі жасайды» деп ${pcs(SV.OPIN["Кедергі жасайды"])}, «екеуі де» деп ${pcs(SV.OPIN["Екеуі де"])} оқушы жауап берді.`;
  V.concl_survey = `Сауалнамаға қатысқан ${N.n_resp} оқушының ${pcs(N.src["Әлеуметтік желі"])}-ы жаңа кітап туралы әлеуметтік желіден біледі, ${pcs(SV.TRIED["Иә"])}-ы блогер ұсынған кітапты оқып көрген, бірақ қазақ тіліндегі кітап блогерін ${pcs(SV.KAZ["Иә"])}-ы ғана көреді.`;
}
if (HAS && N.exp_filled) {
  const E = N.exp, sh = (a, b) => b ? pcs(a / b * 100) : PH;
  V.exp_text = `Тәжірибеге дейінгі 4 аптада эксперименттік сынып (${V.exp_name}) кітапханадан ${E.e_take.before} кітап, бақылау сыныбы (${V.ctrl_name}) ${E.c_take.before} кітап алды. ` +
    `Тәжірибе кезіндегі 4 аптада бұл сан эксперименттік сыныпта ${E.e_take.during}-${dat(E.e_take.during)}, бақылау сыныбында ${E.c_take.during}-${dat(E.c_take.during)} жетті (4-кесте, 4-сурет). ` +
    (N.growth_e !== null && N.growth_c !== null ? `Яғни эксперименттік сыныпта алынған кітап саны ${f1(N.growth_e)} есе, бақылау сыныбында ${f1(N.growth_c)} есе өзгерді. ` : "") +
    `Тәжірибе кезінде эксперименттік сыныпта ${E.e_fin.during} кітап соңына дейін оқылды (алынғанның ${sh(E.e_fin.during, E.e_take.during)}), бақылау сыныбында – ${E.c_fin.during} (${sh(E.c_fin.during, E.c_take.during)}). ` +
    `Өз пікірін жазған немесе кітап туралы бейне түсірген оқушылар эксперименттік сыныпта ${E.e_rev.during}, бақылау сыныбында ${E.c_rev.during} болды.`;
  V.concl_exp = N.growth_e !== null && N.growth_c !== null
    ? `«Буктрейлер» тәжірибесінде эксперименттік сыныптың кітапханадан алған кітап саны ${E.e_take.before}-${abl(E.e_take.before)} ${E.e_take.during}-${dat(E.e_take.during)} (${f1(N.growth_e)} есе), бақылау сыныбында ${E.c_take.before}-${abl(E.c_take.before)} ${E.c_take.during}-${dat(E.c_take.during)} (${f1(N.growth_c)} есе) өзгерді.`
    : `«Буктрейлер» тәжірибесінде эксперименттік сынып ${E.e_take.during}, бақылау сыныбы ${E.c_take.during} кітап алды.`;
}
if (HAS && N.trailers.length) {
  const t = N.trailers, top = [...t].filter(x => x.asked !== null).sort((a, b) => b.asked - a.asked)[0];
  V.trailer_text = `Тәжірибеде ${t.length} буктрейлер көрсеттім: ${listJoin(t.map(x => q(x.book) + (x.author ? ` (${x.author})` : "")))}.` +
    (top ? ` Эксперименттік сыныпта ең көп сұралған кітап – ${q(top.book)}: оны ${top.asked} оқушы сұрады.` : "");
}
if (HAS && N.interviews.length) {
  V.interview_text = "Тәжірибе аяқталған соң сұхбат жүргіздім. " + N.interviews.map(x => `${x.who || "Қатысушы"} (${x.code}): ${x.idea}${x.quote ? ` «${x.quote}»` : ""}`.replace(/\.?$/, ".")).join(" ");
}
if (HAS && N.bloggers.length) {
  const b = N.bloggers, nets = {}; b.forEach(x => nets[x.net || "Басқа"] = (nets[x.net || "Басқа"] || 0) + 1);
  const lit = b.filter(x => x.kaz_lit === "Иә").length;
  V.bloggers_text = `Мен қазақ тіліндегі ${b.length} кітап блогерін таптым: ${Object.entries(nets).map(([k, v]) => `${k} – ${v}`).join(", ")} (Б қосымшасы). ` +
    `Олардың ${lit}-і ғана ұлттық әдебиет туралы бейне не жазба жасайды.` + (SV ? ` Сауалнама бойынша қазақ тіліндегі кітап блогерін оқушылардың ${pcs(SV.KAZ["Иә"])}-ы ғана көреді.` : "");
  V.concl_kaz = `Қазақ тіліндегі кітап блогерлері аз: мен ${b.length} блогер таптым, олардың ${lit}-і ұлттық әдебиетке көңіл бөледі. Бұл – бос тұрған мүмкіндік.`;
}
// Болжамды тексеру: (1) эксперименттік сыныпта өсім бақылау сыныбынан жоғары; (2) кітапты аяқтамайтындар едәуір (кемінде 25%)
const hasA = HAS && N.growth_e !== null && N.growth_e !== undefined && N.growth_c !== null && N.growth_c !== undefined;
const okA = hasA ? N.growth_e > N.growth_c : null;
const notFin = SV && SV.FINISH.n ? (SV.FINISH["Жартылай"] || 0) + (SV.FINISH["Жоқ"] || 0) : null;
const okB = notFin !== null ? notFin >= 25 : null;
if (okA !== null || okB !== null) {
  V.hyp_text = "Болжамымды екі бөлік бойынша тексердім. " +
    (okA !== null ? `Біріншіден, эксперименттік сыныпта алынған кітап саны ${f1(N.growth_e)} есе, бақылау сыныбында ${f1(N.growth_c)} есе өзгерді – ${okA ? "яғни құрдасы жасаған буктрейлер кітапты қолға алуға түрткі болды, болжамның бұл бөлігі расталды" : "яғни буктрейлердің әсері байқалмады, болжамның бұл бөлігі расталмады"}. ` : "[Бірінші бөлігі тәжірибе аяқталғаннан кейін жазылады.] ") +
    (okB !== null ? `Екіншіден, блогер ұсынған кітапты оқығандардың ${pcs(notFin)}-ы оны соңына дейін оқымаған – ${okB ? `демек кітапты аяқтау үшін қосымша қолдау қажет, бұл бөлігі ${okA ? "де " : ""}расталды` : `демек оқушылардың көбі кітапты аяқтайды, бұл бөлігі ${okA === false ? "де " : ""}расталмады`}.` : "[Екінші бөлігі сауалнама нәтижесі енгізілгеннен кейін жазылады.]");
  if (okA !== null && okB !== null) {
    V.hyp_short = okA && okB ? "Болжамым расталды." : !okA && !okB ? "Болжамым расталмады." : "Болжамым ішінара расталды.";
    V.hyp_text += " " + V.hyp_short;
    V.annot_results = `${V.concl_survey || ""} ${V.concl_exp || ""} ${V.hyp_short}`.replace(/\s+/g, " ").trim();
  }
}
if (okA !== null) V.discuss_text = okA
  ? "Менің тәжірибемде де құрдасының бейнесі мен сыныптағы әңгіме оқушыларды кітапқа жетеледі."
  : "Ал менің тәжірибемде қысқа буктрейлердің кітап алуға әсері айқын байқалмады, сондықтан бейненің өзі жеткіліксіз, оған талқылау мен қолдау қосу керек деп ойлаймын.";
for (const [k, txt] of Object.entries({
  survey_text: "[Сауалнама нәтижесі енгізілгеннен кейін жазылады: қатысушылар саны, әлеуметтік желідегі уақыт, жаңа кітап туралы ақпарат көзі.]",
  survey_text2: "[Блогер ұсынған кітапты оқығандар, оны соңына дейін оқығандар, қазақ тіліндегі блогерді көретіндер.]",
  exp_text: "[Тәжірибе аяқталғаннан кейін жазылады: екі сыныптың тәжірибеге дейінгі және тәжірибе кезіндегі кітап саны.]",
  trailer_text: "[Көрсетілген буктрейлерлер тізімі және ең көп сұралған кітап.]",
  interview_text: "[Кітапханашы мен оқушылардан алынған сұхбаттың негізгі ойлары.]",
  bloggers_text: "[Табылған қазақ тіліндегі кітап блогерлерінің саны, желілері және ұлттық әдебиетке көңіл бөлетіні.]",
  hyp_text: "[Болжамның расталғанын не расталмағанын деректермен көрсетіңіз.]", hyp_short: "[Болжам расталды / ішінара расталды / расталмады.]",
  concl_survey: "[Сауалнама бойынша қорытынды.]", concl_exp: "[Тәжірибе бойынша қорытынды.]", concl_kaz: "[Қазақ тіліндегі кітап блогерлері бойынша қорытынды.]",
  annot_results: "[Негізгі нәтижелерді 2–3 сөйлеммен жазыңыз.]", discuss_text: "[Өз тәжірибеңіздің нәтижесін осы пікірмен салыстырыңыз.]",
})) if (V[k] === undefined) V[k] = txt;

const fill = s => s.replace(/\{\{([^}]+)\}\}/g, (_, k) => (V[k] !== undefined && V[k] !== null) ? String(V[k]) : PH);

// ---------- Дереккөздерді нөмірлеу ----------
const content = fs.readFileSync(path.join(DIR, "content.txt"), "utf8").split("\n").filter(l => l.trim() !== "");
const order = [];
for (const line of content) for (const m of line.matchAll(/\[@([a-z_0-9]+)\]/g)) if (!order.includes(m[1])) order.push(m[1]);
for (const k of order) if (!REFS[k]) throw new Error("Missing ref " + k);
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) console.error("Unused refs:", unused);
const cite = s => s.replace(/\[@([a-z_0-9]+)\]/g, (_, k) => `[${order.indexOf(k) + 1}]`);

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
const TABLES = {
  concepts: () => [caption("Зерттеудің негізгі ұғымдары"), mkTable([24, 32, 28, 16], ["Ұғым", "Мағынасы", "Пішіні", "Ұзақтығы"], D.concepts, { boldCol0: true, size: 22 }), gap()],
  proscons: () => [caption("Әлеуметтік желідегі кітап бейнелерінің пайдасы мен қаупі"), mkTable([50, 50], ["Пайдасы", "Қаупі"], D.pros_cons, { size: 22 }), gap()],
  survey: () => [caption(`Сауалнама нәтижесі${HAS && N.n_resp ? ` (${N.n_resp} оқушы)` : ""}, %`), mkTable([46, 54], ["Сұрақ", "Нәтиже"],
    D.survey.slice(1).map(([code, qq, opts]) => [qq + (code === "FINISH" ? " (блогер ұсынған кітапты оқығандар)" : ""),
      SV ? opts.map(o => `${o} – ${pcs(SV[code][o])}`).join("; ") : PH]), { size: 22 }), gap()],
  exp: () => {
    const E = HAS && N.exp_filled ? N.exp : null, c = (k, p) => E ? String(E[k][p]) : PH;
    return [caption("«Буктрейлер» тәжірибесінің нәтижесі (4 апта, кітапхана формулярлары бойынша)"),
      mkTable([34, 16.5, 16.5, 16.5, 16.5], ["Көрсеткіш", `${V.exp_name}: дейін`, `${V.exp_name}: кезінде`, `${V.ctrl_name}: дейін`, `${V.ctrl_name}: кезінде`],
        [["Алынған кітап", "e_take", "c_take"], ["Соңына дейін оқылған кітап", "e_fin", "c_fin"], ["Пікір жазған / бейне түсірген оқушы", "e_rev", "c_rev"]]
          .map(([n, ke, kc]) => [n, c(ke, "before"), c(ke, "during"), c(kc, "before"), c(kc, "during")]), { centerCols: [1, 2, 3, 4] }), gap()];
  },
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
    line("Қатысушының Т.А.Ә.: [Оқушының аты-жөні], 8 «[__]» сынып", { align: L, after: 300 }),
    line("Ғылыми жетекшінің Т.А.Ә.: [Жетекшінің аты-жөні]", { align: L, after: 1000 }),
  ]),
  line("2026–2027 оқу жылы"),
];

// ---------- Аннотация ----------
const ANNOT = "Жұмыс әлеуметтік желідегі кітап блогерлерінің (буктюберлер мен буктокерлердің) 7–9 сынып оқушыларының кітап оқуына әсерін зерттеуге арналған. Өзектілігі: жасөспірімдер жаңа кітап туралы көбіне әлеуметтік желіден біледі, ал BookTok қауымдастығы ұсынған кітаптар әлемде миллиондап сатылады. Зерттеу сұрағы: кітап бейнелері оқушылардың кітап таңдауы мен оқуына әсер ете ме? Болжам: құрдасы жасаған қысқа, эмоциялы бейнешолу оқушыны кітапты қолына алуға жетелейді, бірақ кітапты соңына дейін оқу үшін талқылау мен қолдау қажет. Әдістер: әдебиетті талдау, сауалнама, эксперименттік және бақылау сыныптарымен 4 апталық «Буктрейлер» тәжірибесі, кітапхана формулярларын талдау, сұхбат. Тәжірибе үшін қазақ тіліндегі қысқа буктрейлерлер жасалды. {{annot_results}} Нәтижесінде «Оқу шынжыры» моделі ұсынылды: бейнешолу → кітап → өз пікірің → жаңа бейнешолу.";
const annotWords = fill(ANNOT).split(/\s+/).filter(w => /[\p{L}\d]/u.test(w)).length;
if (annotWords > 200) throw new Error("Аннотация 200 сөзден асты: " + annotWords);
console.log("Аннотация сөз саны:", annotWords);
const annot = [h1("АННОТАЦИЯ"), para(ANNOT), para("Кілт сөздер: буктюб, буктокер, BookTok, буктрейлер, әлеуметтік желі, кітап оқу, оқу сауаттылығы.", { before: 120 })];

// ---------- Қосымшалар ----------
const APPX = [
  { id: "А", title: "Сауалнама сұрақтары" },
  { id: "Ә", title: "Буктрейлер сценарийінің үлгісі және көрсетілген буктрейлерлер" },
  { id: "Б", title: "Қазақ тіліндегі кітап блогерлерін талдау кестесі" },
  { id: "В", title: "Ата-ананың келісім парағы" },
];
function appendix() {
  const out = [];
  const head = (a) => { out.push(h1(`${a.id} ҚОСЫМШАСЫ`)); out.push(para(a.title, { noIndent: true, align: C, bold: true, after: 200 })); };
  // А
  head(APPX[0]);
  out.push(para("Сауалнама Google Forms арқылы жүргізілді. Есім мен электрондық пошта сұралмады.", { after: 120 }));
  const qs = [...D.survey.slice(0, 5).map(([, qq, o]) => `${qq} (${o.join(" / ")})`), D.src_q + ` (${D.sources.join(" / ")})`,
    ...D.survey.slice(5).map(([, qq, o]) => `${qq} (${o.join(" / ")})`), ...D.text_q.map(([, qq]) => qq)];
  qs.forEach((t, i) => out.push(para(`${i + 1}. ${t}`, { noIndent: true, line: 300, after: 40 })));
  // Ә
  head(APPX[1]);
  out.push(para("Үлгі: М. Жұмабаев, «Шолпанның күнәсі» (60 секунд).", { noIndent: true, bold: true, after: 100 }));
  out.push(mkTable([14, 32, 54], ["Уақыт", "Кадр", "Мәтін (дауыс)"], [
    ["0–5 с", "Кітап мұқабасы, қараңғы фон", "«Бір сәттік қателік бүкіл өмірді өзгерте ала ма?»"],
    ["5–20 с", "Кейіпкер туралы сурет / коллаж", "Шолпанның тағдыры туралы 2 сөйлем, сюжетті ашпай"],
    ["20–40 с", "Кітаптан қысқа үзінді оқу", "Ең әсерлі 1–2 сөйлем"],
    ["40–55 с", "Оқушының өз жүзі", "«Мен бұл кітапты бір кеште оқып шықтым, өйткені…»"],
    ["55–60 с", "Жазу: «Кітапханада бар»", "«Сен Шолпанды ақтар ма едің? Пікіріңді жаз»"],
  ], { size: 22, centerCols: [0] }));
  out.push(para("Тәжірибеде көрсетілген буктрейлерлер", { noIndent: true, bold: true, before: 240, after: 100 }));
  const tr = HAS ? N.trailers : [];
  out.push(mkTable([6, 34, 24, 14, 10, 12], ["№", "Кітап", "Автор", "Көрсетілген күні", "Ұзақтығы, с", "Сұрағандар"],
    tr.length ? tr.map((x, i) => [String(i + 1), x.book, x.author, x.date, x.sec ?? "–", x.asked ?? "–"]) : [1, 2, 3].map(i => [String(i), "[кітап]", "[автор]", "[күні]", PH, PH]),
    { size: 22, centerCols: [0, 3, 4, 5] }));
  // Б
  head(APPX[2]);
  const bl = HAS ? N.bloggers : [];
  out.push(mkTable([6, 28, 14, 14, 26, 12], ["№", "Блогер (парақша)", "Желі", "Жазылушы", "Негізгі тақырып / жанр", "Ұлттық әдебиет"],
    bl.length ? bl.map((x, i) => [String(i + 1), x.name, x.net, x.subs, x.topic, x.kaz_lit]) : [1, 2, 3].map(i => [String(i), "[парақша]", "[желі]", PH, "[тақырып]", "[иә/жоқ]"]),
    { size: 22, centerCols: [0, 2, 3, 5] }));
  // В
  head(APPX[3]);
  [
    "Құрметті ата-ана!",
    `Мектебіміздің 8-сынып оқушысы «${TITLE}» тақырыбында зерттеу жобасын орындап жатыр. Зерттеу аясында оқушылар арасында әлеуметтік желі мен кітап оқу туралы қысқа анонимді сауалнама (5–7 минут) және сабақта қазақ тіліндегі қысқа буктрейлерлер көрсетіледі. Кейбір оқушылармен қысқа сұхбат жүргізілуі мүмкін.`,
    "Баланың аты-жөні жазылмайды, жауаптар тек жалпы сан түрінде қолданылады және үшінші тұлғаларға берілмейді. Мектеп кітапханасының деректері тек сынып бойынша жалпы сан түрінде алынады. Қатысу ерікті, бала кез келген уақытта бас тарта алады.",
    "Балаңыздың зерттеуге қатысуына келісім беруіңізді сұраймыз.",
  ].forEach((t, i) => out.push(para(t, { noIndent: i === 0, after: 120 })));
  ["Баланың аты-жөні: ________________________________   Сыныбы: ______",
    "☐ Келісемін        ☐ Келіспеймін",
    "Ата-ананың (заңды өкілінің) аты-жөні: ________________________________",
    "Қолы: ______________          Күні: «___» __________ 2026 ж."].forEach(t => out.push(para(t, { noIndent: true, align: L, before: 200 })));
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