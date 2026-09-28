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
const TITLE = "ЖИ қазақтың тұрақты тіркестерін қалай түсіндіреді?";
const AIS = ["ЖИ-1", "ЖИ-2", "ЖИ-3"];
const PH = "[___]";

// ---------- Дереккөздер ----------
const REFS = {
  zerde: "2–8 сынып оқушыларына арналған «Зерде» республикалық зерттеу жобалары мен шығармашылық жұмыстарының конкурсын ұйымдастыру және өткізу ережелері. – «Дарын» РҒПО директорының 2026 жылғы 9 қыркүйектегі № 153 бұйрығымен бекітілген.",
  zhi_concept: "Қазақстан Республикасында жасанды интеллектіні дамытудың 2024–2029 жылдарға арналған тұжырымдамасы. – ҚР Үкіметінің 2024 жылғы 24 шілдедегі № 592 қаулысымен бекітілген.",
  kenesbaev: "Кеңесбаев І. Қазақ тілінің фразеологиялық сөздігі. – Алматы: Ғылым, 1977. – 712 б.",
  bolganbaev: "Болғанбайұлы Ә., Қалиұлы Ғ. Қазіргі қазақ тілінің лексикологиясы мен фразеологиясы. – Алматы: Санат, 1997. – 256 б.",
  smagulova: "Смағұлова Г. Мағыналас фразеологизмдердің ұлттық-мәдени аспектілері. – Алматы: Ғылым, 1998.",
  kaidar: "Қайдар Ә. Қазақтар ана тілі әлемінде: этнолингвистикалық сөздік. Т. 1: Адам. – Алматы: Дайк-Пресс, 2009.",
  tusindirme: "Қазақ тілінің түсіндірме сөздігі / жалпы ред. Т. Жанұзақов. – Алматы: Дайк-Пресс, 2008. – 968 б.",
  vinogradov: "Виноградов В.В. Об основных типах фразеологических единиц в русском языке // Виноградов В.В. Избранные труды: Лексикология и лексикография. – М.: Наука, 1977. – С. 140–161.",
  vaswani: "Vaswani A., Shazeer N., Parmar N. et al. Attention Is All You Need // Advances in Neural Information Processing Systems 30 (NeurIPS 2017). – 2017. – P. 5998–6008.",
  bender: "Bender E.M., Gebru T., McMillan-Major A., Shmitchell S. On the Dangers of Stochastic Parrots: Can Language Models Be Too Big? // Proceedings of the 2021 ACM Conference on Fairness, Accountability, and Transparency (FAccT ’21). – 2021. – P. 610–623.",
  liu: "Liu E., Cui C., Zheng K., Neubig G. Testing the Ability of Language Models to Interpret Figurative Language // Proceedings of the 2022 Conference of the North American Chapter of the Association for Computational Linguistics (NAACL 2022). – Seattle, 2022.",
  dankers: "Dankers V., Lucas C., Titov I. Can Transformer be Too Compositional? Analysing Idiom Processing in Neural Machine Translation // Proceedings of the 60th Annual Meeting of the Association for Computational Linguistics (ACL 2022). – Dublin, 2022.",
  joshi: "Joshi P., Santy S., Budhiraja A., Bali K., Choudhury M. The State and Fate of Linguistic Diversity and Inclusion in the NLP World // Proceedings of the 58th Annual Meeting of the Association for Computational Linguistics (ACL 2020). – 2020. – P. 6282–6293.",
  ji: "Ji Z., Lee N., Frieske R. et al. Survey of Hallucination in Natural Language Generation // ACM Computing Surveys. – 2023. – Vol. 55, No. 12. – Article 248.",
  unesco: "Miao F., Holmes W. Guidance for Generative AI in Education and Research. – Paris: UNESCO, 2023.",
};

// ---------- Деректерден алынатын мәндер ----------
const f1 = v => (v === null || v === undefined || v === "") ? PH : String(Math.round(v * 10) / 10).replace(".", ",");
const pct = (a, b) => f1(a / b * 100);
const T = D.taldau;
const V = {
  total: T.total, any_body: T.any_body, any_animal: T.any_animal, groups_total: T.groups_total,
  any_body_pct: pct(T.any_body, T.total), any_animal_pct: pct(T.any_animal, T.total),
};
for (const [k, v] of Object.entries(T.body)) V["b_" + k] = v;
for (const [k, v] of Object.entries(T.animal)) V["a_" + k] = v;
const names = N.ai_names || [];
AIS.forEach((_, i) => { V["ai" + (i + 1)] = names[i] && !/^ЖИ-\d$/.test(names[i]) ? names[i] : "[атауы, нұсқасы]"; });
const ERRN = Object.fromEntries(D.err.map(([c, n]) => [c, n]));
if (HAS) {
  for (const [k, v] of Object.entries(N)) if (typeof v === "number") V[k] = f1(v);
  const sum = p => [1, 2, 3].reduce((a, k) => a + (N[`${p}_${k}`] || 0), 0);
  V.e1_ok_sum = sum("e1_ok"); V.e1_half_sum = sum("e1_half"); V.e1_zero_sum = sum("e1_zero");
  const A = N["e1_А_avg"], Ae = N["e1_Ә_avg"], B = N["e1_Б_avg"], all = N.e1_all_avg;
  const gname = { "А": "мөлдір бейнелі (А)", "Ә": "астарлы бейнелі (Ә)", "Б": "ұлттық-мәдени (Б)" };
  const minG = Object.entries({ "А": A, "Ә": Ae, "Б": B }).sort((a, b) => a[1] - b[1])[0][0];
  V.e1_trend = (A > Ae && Ae > B) ? "Демек, тіркестің бейнесі ұлттық мәдениетке неғұрлым тереңірек байланған сайын, ЖИ-дің қателесуі соғұрлым жиілей түседі."
    : minG === "Б" ? "Ұлттық-мәдени тіркестер ЖИ үшін ең қиын топ болып шықты."
    : `Күткенімізге қарамастан, ұлттық-мәдени тіркестер ЖИ үшін ең қиын топ болмады: ең төмен көрсеткіш ${gname[minG]} тіркестерде байқалды.`;
  const byAi = [1, 2, 3].map(k => [k, N[`e1_all_${k}`] || 0]).sort((a, b) => b[1] - a[1]);
  V.e1_best = `Ең жоғары жалпы көрсеткішті ЖИ-${byAi[0][0]} (${f1(byAi[0][1])}%), ең төменін ЖИ-${byAi[2][0]} (${f1(byAi[2][1])}%) көрсетті.`;
  const e2avg = [1, 2, 3].reduce((a, k) => a + N[`e2_all_${k}`], 0) / 3, e2pct = e2avg / 12 * 100;
  V.e2_trend = `Үш ЖИ-дің орташа нәтижесі – ${f1(e2avg)} балл, яғни ${f1(e2pct)}%. ` + (e2pct < all - 5
    ? `Бұл 1-тәжірибедегі көрсеткіштен (${f1(all)}%) төмен: ЖИ тіркестің мағынасын түсіндіре алғанымен, оған мағыналас немесе мағынасы қарама-қарсы тұрақты тіркесті табу оған қиынырақ.`
    : e2pct > all + 5 ? `Бұл 1-тәжірибедегі көрсеткіштен (${f1(all)}%) жоғары: берілген жұптар ЖИ-ге таныс болып шықты.`
    : `Бұл 1-тәжірибедегі көрсеткішке (${f1(all)}%) жуық.`);
  const n3 = D.e3_n * D.e3.length * 3, good = N["e3_Н_sum"], bad = N["e3_Ж_sum"] + N["e3_Т_sum"];
  V.e3_trend = bad > good ? "Яғни ЖИ ұсынған тіркестердің көбі сенімсіз болды: оларды сөздіктен тексермей қолдануға болмайды."
    : `Яғни ЖИ ұсынған тіркестердің ${f1(good / n3 * 100)}%-ы ғана сенімді болды, қалғанын сөздіктен тексеру қажет.`;
  const errTot = D.err.map(([c, n]) => [c, n, N[`e1_err_${c}_sum`] + (N[`e2_err_${c}_sum`] || 0)]).sort((a, b) => b[2] - a[2]);
  V.err_top = `«${errTot[0][1].toLowerCase()}» (${errTot[0][2]} жағдай)`;
  V.err_trend = `1- және 2-тәжірибелерде ең көп кездескен қате түрі – «${errTot[0][1].toLowerCase()}» (${errTot[0][2]} жағдай), ең азы – «${errTot[3][1].toLowerCase()}» (${errTot[3][2]} жағдай).`;
  const okRate = V.e1_ok_sum / 90;
  V.lit_compare = okRate < 1 ? `Біздің тәжірибемізде де ЖИ жауаптарының ${f1(okRate * 100)}%-ы ғана толық дұрыс болды, яғни шамамен әрбір ${Math.max(2, Math.round(1 / (1 - okRate)))}-ші тіркесте ЖИ дәл жауап бере алмады. Бұл ЖИ-дің бейнелі тілді түсінудегі қиындығы қазақ тілінде де байқалатынын көрсетеді.`
    : "Біздің тәжірибемізде ЖИ барлық тіркесті дұрыс түсіндірді, бұл шетелдік зерттеулердің нәтижесінен өзгеше.";
  const d = A - B;
  V.hyp = d >= 10 ? "Бұл болжамымыздың расталғанын көрсетеді: ЖИ ұлттық-мәдени тіркестерде әлдеқайда жиі қателеседі."
    : d > 0 ? "Болжамымыз ішінара расталды: Б тобының көрсеткіші төмен болғанымен, топтар арасындағы айырмашылық үлкен емес."
    : "Болжамымыз расталмады: ұлттық-мәдени тіркестер ЖИ үшін күткеніміздей қиын болмады.";
  V.hyp_short = d >= 10 ? "Болжам расталды." : d > 0 ? "Болжам ішінара расталды." : "Болжам расталмады.";
  const e3pct = good / n3 * 100;
  V.e23_concl = (e2pct < all - 5 || e3pct < all - 5)
    ? "ЖИ-ге тіркестің мағынасын түсіндіруден гөрі, берілген мағынадағы басқа тіркесті табу қиынырақ екенін көрсетті: ол жиі жай сөзбен жауап берді немесе жоқ тіркесті ойдан шығарды."
    : "ЖИ-дің берілген мағынадағы тіркестерді табуда да түсіндірудегідей нәтиже көрсететінін анықтады.";
} else {
  Object.assign(V, {
    e1_trend: "[Топтардың көрсеткішін салыстырып, қорытынды сөйлем жазыңыз.]",
    e1_best: "[Қай ЖИ ең жақсы, қайсысы ең төмен нәтиже көрсетті?]",
    e2_trend: "[2-тәжірибе нәтижесін 1-тәжірибемен салыстырыңыз.]",
    e3_trend: "[ЖИ ұсынған тіркестердің қаншасы сенімді болды?]",
    err_trend: "[Ең көп және ең аз кездескен қате түрін атаңыз.]",
    err_top: PH, lit_compare: "[Өз нәтижеңізді осы тұжырыммен салыстырыңыз.]",
    hyp: "[Болжам расталды ма, ішінара расталды ма, әлде расталмады ма?]", hyp_short: "[Болжам расталды / ішінара расталды / расталмады.]",
    e23_concl: "[2- және 3-тәжірибе бойынша қорытынды жазыңыз.]",
  });
}
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
const HL_RE = /(\[___\]|\[[А-ЯӘҒҚҢӨҰҮҺІа-яәғқңөұүһі][^\]]*\])/;
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
    children: [new Paragraph({
      alignment: o.center ? C : L, spacing: { line: 240 },
      children: runs(text, { size: o.size || SZ_T, bold: o.bold || o.head, italics: o.italics }),
    })],
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
const v = (k) => HAS && N[k] !== undefined && N[k] !== null ? f1(N[k]) : PH;
const TABLES = {
  erkin: () => [caption("Еркін тіркес пен тұрақты тіркестің айырмашылығы"), mkTable([24, 38, 38], ["Белгісі", "Еркін тіркес", "Тұрақты тіркес"], [
    ["Жасалуы", "Сөйлеу кезінде жаңадан құралады", "Тілде дайын күйінде сақталады, бүтін қалпында қолданылады"],
    ["Сөзді алмастыру", "Болады: «тез жүрді» – «жылдам жүрді»", "Болмайды: «аяғы аяғына жұқпады» тіркесін өзгертуге келмейді"],
    ["Мағынасы", "Құрамындағы сөздердің мағынасынан тікелей шығады", "Бүтін тіркеске тән, көбіне бейнелі"],
    ["Сөйлемдегі қызметі", "Әр сөз жеке сөйлем мүшесі бола алады", "Бүтін тіркес бір сөйлем мүшесі болады"],
    ["Мысал", "«Ол таңертең көзін ашты» (оянды)", "«Оқып жүріп көзі ашылды» (санасы оянды)"],
  ], { boldCol0: true }), gap()],
  sinant: () => [caption("Мағыналас және мағынасы қарама-қарсы тұрақты тіркестер"), mkTable([15, 30, 30, 25], ["Түрі", "1-тіркес", "2-тіркес", "Мағынасы"],
    D.e2.map(([k, x, a, m]) => [k === "антоним" ? "Антоним" : "Синоним", x, a.split(";")[0], m]), { boldCol0: true }), gap()],
  toptar: () => {
    const g = [...D.toptar].sort((a, b) => b.items.length - a.items.length).slice(0, 8);
    return [caption("Сөздік тізбесіндегі ең үлкен мағыналық топтар"), mkTable([26, 10, 64], ["Мағыналық топ", "Саны", "Мысалдар"],
      g.map(t => [t.name.charAt(0) + t.name.slice(1).toLowerCase(), String(t.items.length), t.items.slice(0, 4).map(s => s.replace("*", "")).join("; ") + "…"]),
      { boldCol0: true, centerCols: [1] }), gap()];
  },
  e1list: () => [caption("Тәжірибеге іріктелген 30 тұрақты тіркес"), mkTable([7, 48, 45], ["№", "Тұрақты тіркес", "Сөздіктегі мағынасы"],
    ["А", "Ә", "Б"].flatMap(g => [{ section: `${g} тобы – ${D.groups[g]}` }, ...D.e1.map((r, i) => [r, i]).filter(([r]) => r[0] === g).map(([r, i]) => [String(i + 1), r[1], r[2]])]),
    { centerCols: [0] }), gap()],
  shkala: () => [caption("Жауаптарды бағалау шкаласы және қате түрлері"), mkTable([13, 27, 60], ["Балл / код", "Атауы", "Сипаттамасы"], [
    { section: "1- және 2-тәжірибе: жауапқа қойылатын балл" }, ...D.score,
    { section: "Қате түрлері (0 немесе 1 балл қойылғанда белгіленеді)" }, ...D.err,
    { section: "3-тәжірибе: ЖИ атаған тіркестің коды" }, ...D.e3_codes,
  ], { centerCols: [0], boldCol0: true }), gap()],
  zhospar: () => [caption("Зерттеудің жұмыс жоспары"), mkTable([11, 21, 40, 28], ["Кезең", "Мерзімі", "Орындалатын жұмыс", "Нәтижесі"], [
    ["1", "[қыркүйек, 1–2-апта]", "Тақырып таңдау, зерттеу сұрағы мен болжамды тұжырымдау, әдебиетті оқу", "Зерттеу жоспары, дереккөздер тізімі"],
    ["2", "[қыркүйек, 3-апта – қазан, 1-апта]", "Сөздік тізбесін жасау, мағыналарды сөздіктен тексеру, тірек сөздерді санау", `${T.total} тіркестен тұратын тізбе, 3-сурет`],
    ["3", "[қазан, 2-апта]", "30 тіркес іріктеу, үш топқа бөлу, сұраныстар мен бағалау шкаласын дайындау", "4- және 5-кестелер, тәжірибе кестесі"],
    ["4", "[қазан, 3–4-апта]", "Үш ЖИ-ге сұраныстар беру, скриншоттарды сақтау, балл қою", "183 жауап, Б қосымшасы"],
    ["5", "[қараша, 1–3-апта]", "Нәтижелерді есептеу, диаграмма салу, жадынама жасау, қорытынды мен рефлексия жазу", "2-бөлім, жадынама, стенд"],
  ], { centerCols: [0] }), gap()],
  e1res: () => [caption("1-тәжірибе: топтар бойынша сапа көрсеткіші, %"), mkTable([40, 15, 15, 15, 15], ["Тіркестер тобы", ...AIS, "Орташа"],
    [...["А", "Ә", "Б"].map(g => [`${g} – ${D.groups[g].split(" (")[0].toLowerCase()}`, ...[1, 2, 3].map(k => v(`e1_${g}_${k}`)), v(`e1_${g}_avg`)]),
      ["Барлық 30 тіркес", ...[1, 2, 3].map(k => v(`e1_all_${k}`)), v("e1_all_avg")]],
    { centerCols: [1, 2, 3, 4], boldRow: r => r[0].startsWith("Барлық") }), gap()],
  e1ex: () => [caption("ЖИ жауаптарының сипатты мысалдары"), mkTable([22, 20, 40, 18], ["Тіркес", "Сөздікте", "ЖИ жауабы (қысқаша)", "Баға, қате түрі"],
    [1, 2, 3, 4].map(() => ["[тіркес]", "[мағынасы]", "[ЖИ-? жауабы]", "[балл, түрі]"])), gap()],
  e2res: () => [caption("2-тәжірибе: синоним мен антонимді табу нәтижесі, балл"), mkTable([40, 15, 15, 15, 15], ["Тапсырма", ...AIS, "Орташа"], [
    ["Антоним табу (ең жоғары – 6)", ...[1, 2, 3].map(k => v(`e2_ant_${k}`)), HAS ? f1([1, 2, 3].reduce((a, k) => a + N[`e2_ant_${k}`], 0) / 3) : PH],
    ["Синоним табу (ең жоғары – 6)", ...[1, 2, 3].map(k => v(`e2_syn_${k}`)), HAS ? f1([1, 2, 3].reduce((a, k) => a + N[`e2_syn_${k}`], 0) / 3) : PH],
    ["Барлығы (ең жоғары – 12)", ...[1, 2, 3].map(k => v(`e2_all_${k}`)), HAS ? f1([1, 2, 3].reduce((a, k) => a + N[`e2_all_${k}`], 0) / 3) : PH],
  ], { centerCols: [1, 2, 3, 4], boldRow: r => r[0].startsWith("Барлығы") }), gap()],
  e3res: () => [caption("3-тәжірибе: ЖИ атаған тіркестердің құрамы, саны"), mkTable([40, 15, 15, 15, 15], ["Код", ...AIS, "Барлығы"],
    [...D.e3_codes.map(([c, n]) => [`${c} – ${n}`, ...[1, 2, 3].map(k => v(`e3_${c}_${k}`)), v(`e3_${c}_sum`)]),
      ["Барлығы", "25", "25", "25", "75"]], { centerCols: [1, 2, 3, 4], boldRow: r => r[0] === "Барлығы" }), gap()],
  err: () => [caption("1- және 2-тәжірибедегі қате түрлерінің саны"), mkTable([40, 15, 15, 15, 15], ["Қате түрі", ...AIS, "Барлығы"],
    D.err.map(([c, n]) => [`${c} – ${n}`, ...[1, 2, 3].map(k => HAS ? String(N[`e1_err_${c}_${k}`] + (N[`e2_err_${c}_${k}`] || 0)) : PH),
      HAS ? String(N[`e1_err_${c}_sum`] + (N[`e2_err_${c}_sum`] || 0)) : PH]), { centerCols: [1, 2, 3, 4] }), gap()],
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
const ANNOT = "Зерттеудің мақсаты – танымал жасанды интеллект (ЖИ) құралдарының қазақ тіліндегі тұрақты тіркестерді түсіндіру сапасын бағалау. Өзектілігі: оқушылар сөздің мағынасын ЖИ-ден жиі сұрайды, ал бейнелі тіркестер ЖИ үшін қиын. Зерттеу сұрағы: ЖИ қазақ фразеологизмдерін қаншалықты дұрыс түсіндіреді және қандай қателер жібереді? Болжам: ЖИ мөлдір бейнелі тіркестерді жақсы түсіндіреді, ал ұлттық-мәдени тіркестерде жиі қателеседі. Әдістер: сөздік тізбесін жасау және санау, эксперимент, салыстыру, бағалау шкаласы. {{total}} тіркестен тұратын тізбеден 30 тіркес іріктеліп, үш ЖИ-ге мағынасын түсіндіру, синоним мен антоним табу және мағыналық топ бойынша тіркес атау тапсырмалары берілді; барлығы 183 жауап сөздікпен салыстырылды. Нәтижесінде орташа сапа көрсеткіші мөлдір бейнелі тіркестерде {{e1_А_avg}}%, ұлттық-мәдени тіркестерде {{e1_Б_avg}}% болды, ал ЖИ атаған 75 тіркестің {{e3_Ж_sum}}-і сөздіктерде мүлде кездеспеді. ЖИ қателерінің төрт түрі анықталды: сөзбе-сөз түсіндіру, басқа мағына беру, ойдан шығару және жалпылама жауап. {{hyp_short}} Қорытынды: ЖИ фразеологизмді түсінуге көмектесе алады, бірақ оның жауабын сөздікпен тексеру қажет. Нәтижелер негізінде оқушыларға арналған «ЖИ-ге сен, бірақ тексер» жадынамасы ұсынылды.";
const annotWords = fill(ANNOT).split(/\s+/).filter(w => /[\p{L}\d]/u.test(w)).length;
if (annotWords > 200) throw new Error("Аннотация 200 сөзден асты: " + annotWords);
console.log("Аннотация сөз саны:", annotWords);
const annot = [h1("АННОТАЦИЯ"), para(ANNOT), para(`Кілт сөздер: тұрақты тіркес, фразеологизм, жасанды интеллект, чат-бот, сөзбе-сөз түсіндіру, ұлттық-мәдени мағына.`, { before: 120 })];

// ---------- Қосымшалар ----------
const APPX = [
  { id: "А", title: "Фразеологизмдердің сөздік тізбесі" },
  { id: "Ә", title: "Тұрақты тіркестердің мағыналық топтары" },
  { id: "Б", title: "Тәжірибе хаттамасы" },
  { id: "В", title: "«ЖИ-ге сен, бірақ тексер» жадынамасы" },
];
function appendix() {
  const out = [];
  const head = (a) => {
    out.push(h1(`${a.id} ҚОСЫМШАСЫ`));
    out.push(para(a.title, { noIndent: true, align: C, bold: true, after: 200 }));
  };
  // А
  head(APPX[0]);
  out.push(para(`Тізбеге ${T.total} тұрақты тіркес енді, олар әліпби ретімен берілген. Мағыналары І. Кеңесбаевтың фразеологиялық сөздігі [@kenesbaev] және қазақ тілінің түсіндірме сөздігі [@tusindirme] бойынша тексерілді.`, { after: 120 }));
  out.push(mkTable([8, 46, 46], ["№", "Тұрақты тіркес", "Мағынасы"], D.sozdik.map(([x, m], i) => [String(i + 1), x, m]), { size: SZ_A, centerCols: [0] }));
  // Ә
  head(APPX[1]);
  out.push(para(`Тізбеде 22 мағыналық топқа кіретін ${T.groups_total} тіркес бар. Жұлдызшамен (*) белгіленген тіркестер басқа мағынада да қолданылады.`, { after: 120 }));
  out.push(mkTable([28, 72], ["Мағыналық топ", "Тұрақты тіркестер"], D.toptar.map(t => [`${t.name} (${t.items.length})\n`, t.items.join("; ")]).map(([a, b]) => [a.trim(), b]), { size: SZ_A, boldCol0: true }));
  // Б
  head(APPX[2]);
  out.push(para("**Сұраныстар мәтіні.** Сұраныстар барлық ЖИ-ге сөзбе-сөз бірдей, әр жолы жаңа чатта берілді:", { after: 60 }));
  ["1-тәжірибе", "2-тәжірибе", "3-тәжірибе"].forEach((n, i) => out.push(para(`${n}: ${D.prompts[i].replace("{x}", "…").replace("{rel}", "жақын (қарама-қарсы)")}`, { after: 60 })));
  out.push(para(`**ЖИ құралдары:** ЖИ-1 – {{ai1}}; ЖИ-2 – {{ai2}}; ЖИ-3 – {{ai3}}. Тәжірибе күндері: [күні].`, { before: 60, after: 120 }));
  out.push(para("**1-тәжірибенің нәтижелері.** Кестеде әр тіркеске қойылған балл (2 – дұрыс, 1 – жартылай дұрыс, 0 – қате) және қате түрі (С, Б, О, Ж) берілген. Толық деректер (ЖИ жауаптарының қысқаша мазмұны) «Tazhiribe_kestesi.xlsx» файлында сақталған.", { after: 120 }));
  const sc = k => N[`e1_scores_${k}`] || [];
  out.push(mkTable([7, 45, 16, 16, 16], ["№", "Тұрақты тіркес", ...AIS], D.e1.map((r, i) => [String(i + 1), `${r[1]} (${r[0]})`,
    ...[1, 2, 3].map(k => (sc(k)[i] ?? null) === null ? PH : String(sc(k)[i]))]), { size: SZ_A, centerCols: [0, 2, 3, 4] }));
  out.push(para("**Скриншоттар.** ЖИ жауаптарының скриншоттары «ЖИ1_Т01.png» үлгісімен атаулы электрондық бумада сақталған. Жұмыста талданған мысалдардың скриншоттары төменде берілген.", { before: 120 }));
  out.push(para("[Осы жерге 8-кестедегі мысалдардың скриншоттарын қойыңыз]", { noIndent: true, align: C, before: 120 }));
  // В
  head(APPX[3]);
  out.push(...figBlock("zhadynama", "«ЖИ-ге сен, бірақ тексер» жадынамасы (басып шығаруға арналған нұсқа)"));
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
