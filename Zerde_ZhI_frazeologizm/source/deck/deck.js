// Қорғауға арналған презентация: node deck.js ../../Qorgau_prezentatsiyasy.pptx
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const D = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "data.json"), "utf8"));
const N = D.natije, T = D.taldau;
const f1 = v => String(Math.round(v * 10) / 10).replace(".", ",");
const AIS = N.ai_names; // ChatGPT, DeepSeek, Gemini

const THEME = {
  name: "Zerde ZhI frazeologizm",
  headFontFace: "Times New Roman",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1B2A30", lt1: "FFFFFF", dk2: "0B3C49", lt2: "EAF5F6",
    accent1: "00A3B4", accent2: "E8A900", accent3: "D64545", accent4: "2E8B57",
    accent5: "5B6B73", accent6: "7A5195", hlink: "00A3B4", folHlink: "7A5195",
  },
};
const HEX = THEME.colors;
const AI_COL = [HEX.accent1, HEX.accent2, HEX.dk2];

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "ЖИ қазақтың тұрақты тіркестерін қалай түсіндіреді?";
const C = pres.SchemeColor;
const W = 13.333;

// ---------- Макеттер (layouts) ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.55, w: 7.4, h: 2.6, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.3, w: 7.4, h: 0.9, fontSize: 18, italic: true, color: C.accent2, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.85, fontSize: 32, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "«Зерде» · Қазақ тілі және әдебиеті", options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

// ---------- Иконкалар ----------
async function icon(Comp, color = "FFFFFF") {
  const svg = renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function iconCircle(s, Comp, x, y, d, fill, name) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + "-circle" });
  const pad = d * 0.25;
  s.addImage({ data: await icon(Comp), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: name + "-icon" });
}

// ---------- Көмекші элементтер ----------
const txt = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, ...o });
function card(s, x, y, w, h, fill, name, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill },
    line: line ? { color: line, width: 1.25 } : { type: "none" }, objectName: name });
}
function idiom(s, x, y, w, h, text, name, size = 20) { // мотив: тұрақты тіркес – алтын «» ішіндегі карточка
  card(s, x, y, w, h, C.background1, name, C.accent2);
  txt(s, [{ text: "«", options: { color: C.accent2, bold: true } }, { text, options: { color: C.text2, bold: true } }, { text: "»", options: { color: C.accent2, bold: true } }],
    { x: x + 0.2, y, w: w - 0.4, h, fontFace: THEME.headFontFace, fontSize: size, align: "center", valign: "middle", objectName: name + "-text" });
}
function bigStat(s, x, y, w, value, label, color, name, size = 54) {
  txt(s, value, { x, y, w, h: 0.95, fontSize: size, bold: true, color, fontFace: THEME.headFontFace, objectName: name + "-value" });
  txt(s, label, { x, y: y + 0.95, w, h: 0.6, fontSize: 14, color: C.text1, valign: "top", objectName: name + "-label" });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };
const chartQuiet = () => ({
  ...chartText, catAxisLabelColor: HEX.dk1, valAxisLabelColor: HEX.accent5, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  dataLabelFontSize: 12, dataLabelColor: HEX.dk1, valGridLine: { color: "D9E2E4", size: 0.75 }, catGridLine: { style: "none" },
  legendFontSize: 12, legendColor: HEX.dk1,
});

(async () => {
  let s;
  // ================= 1. Титул =================
  pres.addSection({ title: "Кіріспе" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Кіріспе" });
  txt(s, "«Зерде» республикалық конкурсы  ·  Қазақ тілі және әдебиеті  ·  «Орта буын»", { x: 0.8, y: 0.6, w: 11.5, h: 0.4, fontSize: 14, color: C.accent1, objectName: "kicker" });
  s.addText("ЖИ қазақтың тұрақты тіркестерін қалай түсіндіреді?", { placeholder: "title" });
  s.addText("ChatGPT, DeepSeek және Gemini-ді 182 жауап арқылы сынадық", { placeholder: "body" });
  txt(s, "Тіркеу коды (шифр): ______________", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.background2, objectName: "cipher" });
  ["Ит өлген жер", "Асығы алшысынан түсті", "Жер қаптыру"].forEach((t, i) =>
    idiom(s, 8.75, 1.7 + i * 1.35, 3.85, 1.0, t, "title-idiom-" + i, 20));
  s.addNotes("Құрметті қазылар алқасы! Менің жобамның тақырыбы – «ЖИ қазақтың тұрақты тіркестерін қалай түсіндіреді?». Мен үш танымал чат-ботты – ChatGPT, DeepSeek және Gemini-ді – қазақтың тұрақты тіркестерімен сынап, олардың 182 жауабын сөздікпен салыстырдым. (15 секунд)");

  // ================= 2. Мәселе: бір тіркес – үш жауап =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Бір тіркес – үш түрлі жауап", { placeholder: "title" });
  idiom(s, 0.6, 1.7, 4.6, 1.5, "Жер қаптыру", "hook-idiom", 30);
  txt(s, "Сөздікте:", { x: 0.6, y: 3.5, w: 4.6, h: 0.4, fontSize: 14, color: C.accent5, objectName: "hook-dict-label" });
  txt(s, "алдау", { x: 0.6, y: 3.9, w: 4.6, h: 0.7, fontSize: 30, bold: true, color: C.accent4, fontFace: THEME.headFontFace, objectName: "hook-dict" });
  const hook = [
    ["ChatGPT", "«жеңу, жеңіліске ұшырату»", false, fa.FaTimes],
    ["DeepSeek", "«алдау, алаяқтық жасау»", true, fa.FaCheck],
    ["Gemini", "«біреуді ойсырата жеңу»", false, fa.FaTimes],
  ];
  for (let i = 0; i < 3; i++) {
    const [ai, ans, ok, Ic] = hook[i], y = 1.7 + i * 1.05;
    card(s, 5.7, y, 7.0, 0.85, C.background2, "hook-row-" + i);
    txt(s, ai, { x: 5.95, y, w: 1.6, h: 0.85, fontSize: 16, bold: true, color: C.text2, valign: "middle", objectName: "hook-ai-" + i });
    txt(s, ans, { x: 7.6, y, w: 4.2, h: 0.85, fontSize: 18, color: C.text1, valign: "middle", objectName: "hook-ans-" + i });
    await iconCircle(s, Ic, 11.95, y + 0.17, 0.52, ok ? C.accent4 : C.accent3, "hook-mark-" + i);
  }
  card(s, 0.6, 5.2, 12.1, 1.3, C.text2, "hook-question");
  txt(s, "ЖИ әрқашан сенімді жазады. Бірақ оның жауабына сенуге бола ма?", { x: 1.0, y: 5.2, w: 11.3, h: 1.3, fontSize: 22, bold: true, color: C.background1, valign: "middle", objectName: "hook-question-text" });
  s.addNotes("Бәрі осы тіркестен басталды. «Жер қаптыру» сөздікте «алдау» деген мағынаны білдіреді. Ал оны мен үш ЖИ-ден сұрағанда, екеуі «жеңу» деп жауап берді, тек DeepSeek дұрыс айтты. Ең қызығы – қате жауаптар да өте сенімді, сауатты жазылған. Сондықтан мен ЖИ-дің жауабына қаншалықты сенуге болатынын тексеруді шештім. (30 секунд)");

  // ================= 3. Сұрақ, мақсат, болжам =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеу сұрағы мен болжам", { placeholder: "title" });
  card(s, 0.6, 1.55, 12.1, 1.45, C.background2, "q-card");
  await iconCircle(s, fa.FaQuestion, 0.95, 1.85, 0.85, C.accent1, "q");
  txt(s, "ЖИ чат-боттары қазақ тіліндегі тұрақты тіркестердің мағынасын қаншалықты дұрыс түсіндіреді және қандай қателер жібереді?",
    { x: 2.1, y: 1.55, w: 10.3, h: 1.45, fontSize: 21, bold: true, color: C.text2, valign: "middle", objectName: "q-text" });
  const qa = [
    [fa.FaBullseye, "Мақсаты", "Үш ЖИ-дің тұрақты тіркестерді түсіндіруін сөздікпен салыстырып бағалау, қателерін анықтау және оқушыларға ЖИ жауабын тексеру жадынамасын жасау", C.accent1],
    [fa.FaLightbulb, "Болжам", "ЖИ мағынасы оңай түсінілетін тіркестерді жақсы түсіндіреді, ал ұлттық-мәдени тіркестерде жиі қателеседі: сөзбе-сөз түсіндіреді немесе мағынасын ойдан шығарады", C.accent2],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25;
    card(s, x, 3.35, 5.85, 3.3, C.background1, "aim-card-" + i, "C9DDE0");
    await iconCircle(s, qa[i][0], x + 0.35, 3.65, 0.75, qa[i][3], "aim-" + i);
    txt(s, qa[i][1], { x: x + 1.3, y: 3.65, w: 4.3, h: 0.75, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: "aim-title-" + i });
    txt(s, qa[i][2], { x: x + 0.35, y: 4.6, w: 5.15, h: 1.9, fontSize: 18, color: C.text1, valign: "top", objectName: "aim-text-" + i });
  }
  s.addNotes("Менің зерттеу сұрағым: ЖИ қазақ тіліндегі тұрақты тіркестерді қаншалықты дұрыс түсіндіреді және қандай қателер жібереді? Мақсатым – ЖИ жауаптарын сөздікпен салыстырып бағалау және оқушыларға жадынама жасау. Болжамым: ЖИ оңай тіркестерді жақсы түсіндіреді, ал қазақ тұрмысына байланысты ұлттық-мәдени тіркестерде қателеседі. (30 секунд)");

  // ================= 4. Неге қиын =================
  pres.addSection({ title: "Зерттеу" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Зерттеу" });
  s.addText("Неге тұрақты тіркес ЖИ-ге қиын болуы мүмкін?", { placeholder: "title" });
  idiom(s, 0.6, 1.6, 5.6, 1.0, "Ит өлген жер", "why-idiom", 24);
  card(s, 0.6, 2.95, 5.6, 1.35, "FBE9E9", "why-literal");
  txt(s, [{ text: "ит + өлген + жер", options: { bold: true, breakLine: true } }, { text: "= «иттің өлген жері»  ✗", options: {} }],
    { x: 0.9, y: 2.95, w: 5.0, h: 1.35, fontSize: 18, color: C.accent3, valign: "middle", objectName: "why-literal-text" });
  card(s, 0.6, 4.55, 5.6, 1.35, "E6F3EC", "why-figurative");
  txt(s, [{ text: "тұтас бейнелі мағына", options: { bold: true, breakLine: true } }, { text: "= «өте алыс жер»  ✓", options: {} }],
    { x: 0.9, y: 4.55, w: 5.0, h: 1.35, fontSize: 18, color: C.accent4, valign: "middle", objectName: "why-fig-text" });
  const reasons = [
    [fa.FaPuzzlePiece, "Мағынасы сөздерден шықпайды", "ЖИ тіркесті сөзбе-сөз «жинап» түсіндіруі мүмкін"],
    [fa.FaGlobeAsia, "Қазақша мәтін аз", "Ғаламторда қазақ тіліндегі мәтін ағылшын тіліне қарағанда әлдеқайда аз"],
    [fa.FaRobot, "ЖИ ойдан шығара алады", "Білмегенін «білмеймін» демей, сенімді түрде жалған жауап береді («галлюцинация»)"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.5;
    await iconCircle(s, reasons[i][0], 6.8, y + 0.1, 0.8, C.accent1, "reason-" + i);
    txt(s, reasons[i][1], { x: 7.85, y, w: 4.85, h: 0.5, fontSize: 18, bold: true, color: C.text2, objectName: "reason-title-" + i });
    txt(s, reasons[i][2], { x: 7.85, y: y + 0.5, w: 4.85, h: 0.75, fontSize: 14, color: C.text1, valign: "top", objectName: "reason-text-" + i });
  }
  s.addNotes("Тұрақты тіркестің мағынасы ішіндегі сөздерден шықпайды. «Ит өлген жер» – «иттің өлген жері» емес, «өте алыс жер». Ғылыми еңбектерді оқып, ЖИ-ге бұл неге қиын болатынының үш себебін таптым: мағынасы сөздердің қосындысы емес, қазақша мәтін аз, және ЖИ білмегенін ойдан шығаруы мүмкін. (35 секунд)");

  // ================= 5. Сөздік тізбесі =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Зерттеу" });
  s.addText(`Алдымен ${T.total} тіркесті жинап, санадым`, { placeholder: "title" });
  bigStat(s, 0.6, 1.6, 4.4, String(T.total), "тұрақты тіркес жинадым, 22 мағыналық топқа бөлдім", C.text2, "st-total");
  bigStat(s, 0.6, 3.35, 4.4, f1(T.any_body / T.total * 100) + "%", "тіркесте дене мүшесінің атауы бар", C.accent1, "st-body");
  bigStat(s, 0.6, 5.1, 4.4, f1(T.any_animal / T.total * 100) + "%", "тіркесте жануар атауы бар (ит, ат, қой)", C.accent2, "st-animal");
  const top = Object.entries(T.body).sort((a, b) => b[1] - a[1]).slice(0, 6).reverse();
  s.addChart(pres.charts.BAR, [{ name: "Тіркес саны", labels: top.map(x => x[0]), values: top.map(x => x[1]) }], {
    x: 5.4, y: 1.5, w: 7.3, h: 5.2, barDir: "bar", chartColors: [HEX.accent1], showValue: true, dataLabelPosition: "outEnd",
    showTitle: true, title: "Ең жиі кездесетін дене мүшелері (тіркес саны)", titleFontSize: 14, titleColor: HEX.dk1,
    showLegend: false, valAxisHidden: true, barGapWidthPct: 45, objectName: "chart-body", ...chartQuiet(), valGridLine: { style: "none" },
  });
  s.addNotes(`Тәжірибеден бұрын ${T.total} тұрақты тіркесті жинап, санадым. Олардың ${f1(T.any_body / T.total * 100)} пайызында дене мүшесінің атауы бар, ең жиісі – «көз», ${T.body["көз"]} тіркесте кездеседі. Ал ${f1(T.any_animal / T.total * 100)} пайызында ит, ат, қой сияқты жануарлар бар – бұл мал баққан қазақ өмірінің тілдегі ізі. Осы санау тәжірибеге тіркес таңдауға көмектесті. (35 секунд)`);

  // ================= 6. Тәжірибе барысы =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Зерттеу" });
  s.addText("Тәжірибе қалай жүрді", { placeholder: "title" });
  const steps = [
    [fa.FaListOl, "30 тіркес", "3 топ:\nА – мөлдір бейнелі\nӘ – астарлы бейнелі\nБ – ұлттық-мәдени"],
    [fa.FaRobot, "3 ЖИ", "ChatGPT, DeepSeek, Gemini\n\nәр сұраныс – жаңа чатта, мәтіні бірдей"],
    [fa.FaTasks, "3 тапсырма", "мағынасын түсіндіру – 90\nсиноним, антоним – 18\nтоп бойынша атау – 74"],
    [fa.FaBook, "Сөздікпен тексеру", "Кеңесбаев сөздігі\n\n2 – дұрыс\n1 – жартылай\n0 – қате"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.1;
    card(s, x, 1.6, 2.75, 4.25, C.background2, "step-card-" + i);
    await iconCircle(s, steps[i][0], x + 0.95, 1.85, 0.85, i === 3 ? C.accent2 : C.accent1, "step-" + i);
    txt(s, steps[i][1], { x: x + 0.2, y: 2.85, w: 2.35, h: 0.6, fontSize: 20, bold: true, color: C.text2, align: "center", objectName: "step-title-" + i });
    txt(s, steps[i][2], { x: x + 0.25, y: 3.5, w: 2.25, h: 2.2, fontSize: 14, color: C.text1, valign: "top", objectName: "step-text-" + i });
    if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.8, y: 3.55, w: 0.25, h: 0.35, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "step-arrow-" + i });
  }
  txt(s, [{ text: "182", options: { bold: true, color: C.accent1, fontSize: 28, fontFace: THEME.headFontFace } }, { text: "  жауапты сөздікпен салыстырып, баллды өзім қойдым", options: { color: C.text1, fontSize: 18 } }],
    { x: 0.6, y: 6.05, w: 12.1, h: 0.7, valign: "middle", objectName: "steps-total" });
  s.addNotes("Тәжірибеге 30 тіркес алып, оларды үш топқа бөлдім: мағынасы оңай түсінілетін, астарлы және ұлттық-мәдени тіркестер. Үш ЖИ-ге бірдей сұраныстарды әр жолы жаңа чатта бердім. Тапсырма үшеу болды: мағынасын түсіндіру, синоним мен антоним табу және мағыналық топ бойынша тіркес атау. Барлығы 182 жауапты Кеңесбаевтың сөздігімен салыстырып, 0, 1, 2 балл қойдым. (40 секунд)");

  // ================= 7. Нәтиже 1 =================
  pres.addSection({ title: "Нәтижелер" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("ЖИ тіркестердің 94%-ын дұрыс түсіндірді", { placeholder: "title" });
  bigStat(s, 0.6, 1.6, 3.9, f1(N.e1_all_avg) + "%", "1-тәжірибедегі орташа сапа көрсеткіші (90 жауап)", C.accent1, "r1-avg", 66);
  AIS.forEach((ai, k) => {
    const y = 3.55 + k * 0.95;
    card(s, 0.6, y, 3.9, 0.75, C.background2, "r1-ai-card-" + k);
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.24, w: 0.27, h: 0.27, fill: { color: AI_COL[k] }, line: { type: "none" }, objectName: "r1-ai-dot-" + k });
    txt(s, ai, { x: 1.25, y, w: 1.8, h: 0.75, fontSize: 16, color: C.text1, valign: "middle", objectName: "r1-ai-name-" + k });
    txt(s, f1(N[`e1_all_${k + 1}`]) + "%", { x: 3.0, y, w: 1.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, align: "right", valign: "middle", objectName: "r1-ai-val-" + k });
  });
  s.addChart(pres.charts.BAR, AIS.map((ai, k) => ({ name: ai, labels: ["А – мөлдір", "Ә – астарлы", "Б – ұлттық-мәдени"], values: ["А", "Ә", "Б"].map(g => N[`e1_${g}_${k + 1}`]) })), {
    x: 4.9, y: 1.5, w: 7.8, h: 5.3, barDir: "col", barGrouping: "clustered", chartColors: AI_COL, showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0", valAxisMinVal: 0, valAxisMaxVal: 110, valAxisMajorUnit: 20, valAxisLabelFormatCode: "0", showLegend: true, legendPos: "b",
    showTitle: true, title: "Топтар бойынша сапа көрсеткіші, %", titleFontSize: 14, titleColor: HEX.dk1, barGapWidthPct: 60, objectName: "chart-groups", ...chartQuiet(),
  });
  s.addNotes(`Бірінші нәтиже мені таңғалдырды: үш ЖИ тіркестердің басым бөлігін дұрыс түсіндірді. Орташа сапа көрсеткіші – ${f1(N.e1_all_avg)} пайыз. Ең жоғары нәтижені ChatGPT көрсетті – ${f1(N.e1_all_1)} пайыз. Астарлы тіркестерде үшеуі де 100 пайыз алды. Бірақ диаграммада бір топтың төмен екені көрініп тұр – ол ұлттық-мәдени тіркестер. (35 секунд)`);

  // ================= 8. Әлсіз тұс =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Әлсіз тұс – ұлттық-мәдени тіркестер", { placeholder: "title" });
  bigStat(s, 0.6, 1.6, 4.2, f1(N["e1_Б_avg"]) + "%", `ұлттық-мәдени тіркестерде (А тобында – ${f1(N["e1_А_avg"])}%, Ә тобында – ${f1(N["e1_Ә_avg"])}%)`, C.accent3, "weak-stat", 66);
  card(s, 0.6, 3.9, 4.2, 2.6, C.text2, "weak-callout");
  txt(s, [{ text: "4 / 4", options: { fontSize: 44, bold: true, color: C.accent2, fontFace: THEME.headFontFace, breakLine: true } },
    { text: "0 балл алған жауаптың бәрі осы топта", options: { fontSize: 17, color: C.background1 } }],
    { x: 0.95, y: 3.9, w: 3.6, h: 2.6, valign: "middle", objectName: "weak-callout-text" });
  const errs = [
    ["Есек құрты мұртына түсу", "баю", "Gemini", "«ашуға міну, шамдану»"],
    ["Айдарынан жел есу", "асқақтап, дәурен сүру", "DeepSeek", "«мақтанып, тәкаппарлану»"],
    ["Жер қаптыру", "алдау", "ChatGPT, Gemini", "«жеңу»"],
  ];
  errs.forEach(([id, dict, ai, ans], i) => {
    const y = 1.6 + i * 1.65;
    card(s, 5.3, y, 7.4, 1.45, C.background2, "err-card-" + i);
    txt(s, [{ text: "«" + id + "»", options: { bold: true, color: C.text2, fontFace: THEME.headFontFace, fontSize: 20, breakLine: true } },
      { text: "сөздікте: ", options: { color: C.accent5, fontSize: 14 } }, { text: dict, options: { color: C.accent4, bold: true, fontSize: 14 } }],
      { x: 5.6, y, w: 3.7, h: 1.45, valign: "middle", objectName: "err-idiom-" + i });
    txt(s, [{ text: ai + ":", options: { color: C.accent5, fontSize: 14, breakLine: true } }, { text: ans, options: { color: C.accent3, bold: true, fontSize: 17 } }],
      { x: 9.4, y, w: 3.1, h: 1.45, valign: "middle", objectName: "err-ai-" + i });
  });
  s.addNotes(`Ұлттық-мәдени тіркестерде көрсеткіш ${f1(N["e1_Б_avg"])} пайызға дейін түсті. Ең маңыздысы – 0 балл алған төрт жауаптың төртеуі де осы топта. Мысалы, Gemini «есек құрты мұртына түсу» – «баю» тіркесін «ашуға міну» деп түсіндірді, ал DeepSeek «айдарынан жел есу» тіркесін «мақтану» деп берді. Бұл тіркестер қазақтың тұрмысымен байланысты және сирек айтылады. (35 секунд)`);

  // ================= 9. Күткеніміз бен шыққаны =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Күткенім мен шыққан нәтиже", { placeholder: "title" });
  const cols = [
    ["Күттім", C.accent5, [["Сөзбе-сөз түсіндіреді", "«иттің өлген жері»"], ["Жоқ тіркестерді ойлап табады", "сөздікте жоқ тіркестер"]]],
    ["Шықты", C.accent1, [["Сөзбе-сөз түсіндіру – 0", "ЖИ тіркестің бейнелі екенін әрқашан «сезді»"], ["Жоқ тіркес – 0", "атаған тіркестердің бәрі тілде бар"]]],
  ];
  for (let c = 0; c < 2; c++) {
    const x = 0.6 + c * 4.3;
    txt(s, cols[c][0], { x, y: 1.55, w: 3.9, h: 0.55, fontSize: 22, bold: true, color: cols[c][1], objectName: "exp-col-" + c });
    cols[c][2].forEach(([h, d], i) => {
      const y = 2.25 + i * 1.6;
      card(s, x, y, 3.9, 1.4, c ? C.background2 : "F2F4F5", `exp-card-${c}-${i}`);
      txt(s, [{ text: h, options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 13, color: C.text1 } }],
        { x: x + 0.25, y, w: 3.4, h: 1.4, valign: "middle", objectName: `exp-text-${c}-${i}` });
    });
  }
  card(s, 9.2, 1.55, 3.5, 3.65, C.text2, "real-errors");
  txt(s, [{ text: "Нақты қателер", options: { bold: true, fontSize: 18, color: C.accent2, breakLine: true } },
    { text: "4", options: { bold: true, fontSize: 40, color: C.background1, fontFace: THEME.headFontFace, breakLine: true } },
    { text: "сенімді ойдан шығару", options: { fontSize: 14, color: C.background1, breakLine: true } },
    { text: "6", options: { bold: true, fontSize: 40, color: C.background1, fontFace: THEME.headFontFace, breakLine: true } },
    { text: "сөздікте жоқ артық мағына қосу", options: { fontSize: 14, color: C.background1 } }],
    { x: 9.5, y: 1.55, w: 3.0, h: 3.65, valign: "middle", objectName: "real-errors-text" });
  card(s, 0.6, 5.55, 12.1, 1.1, "FFF6DC", "verdict");
  await iconCircle(s, fa.FaBalanceScale, 0.85, 5.72, 0.75, C.accent2, "verdict");
  txt(s, "Болжам ішінара расталды: ЖИ ұлттық-мәдени тіркестерде қателесті, бірақ сөзбе-сөз емес – сенімді, сауатты, бірақ қате мағына берді",
    { x: 1.85, y: 5.55, w: 10.6, h: 1.1, fontSize: 16, bold: true, color: C.text2, valign: "middle", objectName: "verdict-text" });
  s.addNotes("Менің болжамым ішінара ғана расталды. Мен ЖИ тіркестерді сөзбе-сөз түсіндіреді деп ойлаған едім, бірақ бірде-бір сөзбе-сөз түсіндіру кездеспеді. ЖИ жоқ тіркестерді де ойлап таппады. Оның қателері басқаша болды: сенімді түрде ойдан шығарылған мағына және сөздікте жоқ артық мағына. Менің ойымша, мұндай қате қауіптірек, өйткені оны байқау қиын. (35 секунд)");

  // ================= 10. 2- және 3-тәжірибе =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("ЖИ өзі тіркес таба ала ма?", { placeholder: "title" });
  const e2pct = (N.e2_all_1 + N.e2_all_2 + N.e2_all_3) / 36 * 100, e3pct = N["e3_Н_sum"] / N.e3_named_sum * 100;
  bigStat(s, 0.6, 1.55, 3.0, f1(e2pct) + "%", "синоним мен антонимді дұрыс тапты (2-тәжірибе)", C.accent1, "e2-stat", 46);
  bigStat(s, 3.8, 1.55, 3.0, f1(e3pct) + "%", `атаған ${N.e3_named_sum} тіркесі сөздікте бар әрі дұрыс (3-тәжірибе)`, C.accent1, "e3-stat", 46);
  const codes = [["Н", "Дұрыс тіркес", HEX.accent1], ["М", "Басқа мағынадағы тіркес", HEX.accent2], ["Т", "Тұрақты тіркес емес", HEX.accent5]];
  s.addChart(pres.charts.BAR, codes.map(([c, n]) => ({ name: n, labels: AIS, values: [1, 2, 3].map(k => N[`e3_${c}_${k}`]) })), {
    x: 0.5, y: 3.4, w: 6.5, h: 3.4, barDir: "bar", barGrouping: "stacked", chartColors: codes.map(c => c[2]), showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: "0;-0;;",
    dataLabelColor: "FFFFFF", showLegend: true, legendPos: "b", valAxisHidden: true, barGapWidthPct: 50, objectName: "chart-e3",
    showTitle: true, title: "3-тәжірибе: ЖИ атаған тіркестер", titleFontSize: 14, titleColor: HEX.dk1, ...chartQuiet(), dataLabelColor: "FFFFFF", valGridLine: { style: "none" },
  });
  card(s, 7.4, 1.55, 5.3, 5.2, C.background2, "near-card");
  await iconCircle(s, fa.FaRegLightbulb, 7.7, 1.8, 0.75, C.accent2, "near");
  txt(s, "«Жақын» сөзі ЖИ-ді шатастырды", { x: 8.6, y: 1.8, w: 3.9, h: 0.75, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "near-title" });
  txt(s, [{ text: "«Жақын жер» мағынасындағы тіркестерді сұрағанда, DeepSeek былай деп жазды:", options: { breakLine: true } }],
    { x: 7.7, y: 2.75, w: 4.75, h: 0.8, fontSize: 14, color: C.text1, valign: "top", objectName: "near-intro" });
  idiom(s, 7.7, 3.6, 4.75, 0.75, "үш қайнаса сорпасы қосылмайтын", "near-i1", 15);
  idiom(s, 7.7, 4.5, 4.75, 0.75, "ат құлағында ойнайтын", "near-i2", 15);
  txt(s, "ЖИ «жақын» сөзін қашықтық емес, туыстық мағынада түсінді. Бұл топта атаған 14 тіркестің тек 9-ы дұрыс болды.",
    { x: 7.7, y: 5.4, w: 4.75, h: 1.2, fontSize: 14, color: C.text1, valign: "top", objectName: "near-text" });
  s.addNotes(`Синоним мен антонимді ЖИ ${f1(e2pct)} пайыз жағдайда дұрыс тапты. Ал мағыналық топ бойынша атаған ${N.e3_named_sum} тіркестің ${f1(e3pct)} пайызы дұрыс болды, ${N["e3_Т_sum"]} жағдайда ЖИ тұрақты тіркестің орнына жай сөз жазды. Ең қызығы – «жақын жер» тобы. DeepSeek «үш қайнаса сорпасы қосылмайтын» тіркесін атады. Яғни ол «жақын» сөзін қашықтық емес, туыстық мағынада түсінді. (35 секунд)`);

  // ================= 11. Өнім =================
  pres.addSection({ title: "Қорытынды" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Нәтижеден – өнімге: «ЖИ-ге сен, бірақ тексер»", { placeholder: "title" });
  const memo = [
    [fa.FaBook, "Сөздікпен салыстыр", "ЖИ сенімді жазса да, мағынасын фразеологиялық сөздіктен тексер"],
    [fa.FaFlag, "Ұлттық тіркеске сақ бол", "Барлық қате жауап қазақ тұрмысы мен салтына байланысты тіркестерде болды"],
    [fa.FaSearchPlus, "Артық мағынаны байқа", "ЖИ дұрыс мағынаға сөздікте жоқ мағына қосуы мүмкін"],
    [fa.FaUserFriends, "Үлкендерден сұра", "Күмәнің қалса – мұғалімнен, ата-әжеңнен сұра"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + (i % 2) * 6.25, y = 1.6 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 5.85, 2.25, C.background2, "memo-card-" + i);
    await iconCircle(s, memo[i][0], x + 0.35, y + 0.35, 0.9, [C.accent1, C.accent2, C.accent1, C.accent2][i], "memo-" + i);
    txt(s, `${i + 1}. ${memo[i][1]}`, { x: x + 1.5, y: y + 0.3, w: 4.1, h: 0.6, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "memo-title-" + i });
    txt(s, memo[i][2], { x: x + 1.5, y: y + 0.95, w: 4.1, h: 1.1, fontSize: 15, color: C.text1, valign: "top", objectName: "memo-text-" + i });
  }
  s.addNotes("Нәтижелер бойынша оқушыларға арналған «ЖИ-ге сен, бірақ тексер» жадынамасын жасадым. Ол төрт қадамнан тұрады: сөздікпен салыстыр, ұлттық тіркеске сақ бол, артық мағынаны байқа және күмәнің қалса, үлкендерден сұра. Жадынаманы қазақ тілі сабағында және сынып сағатында қолдануға болады. (30 секунд)");

  // ================= 12. Қорытынды =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Қорытынды", { placeholder: "title" });
  const concl = [
    `ЖИ тұрақты тіркестерді жалпы жақсы түсіндіреді – орташа ${f1(N.e1_all_avg)}%`,
    `Ең әлсіз тұсы – ұлттық-мәдени тіркестер (${f1(N["e1_Б_avg"])}%): қате жауаптың бәрі осы топта`,
    "ЖИ сөзбе-сөз түсіндірмейді, бірақ қате мағынаны сенімді жазады – сондықтан оны байқау қиын",
    "ЖИ – жақсы көмекші, бірақ шешімді сөздікке қарап адам қабылдауы керек",
  ];
  concl.forEach((t, i) => {
    const y = 1.55 + i * 1.02;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "concl-num-bg-" + i });
    txt(s, String(i + 1), { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "concl-num-" + i });
    txt(s, t, { x: 1.5, y, w: 11.2, h: 0.8, fontSize: 19, color: C.text1, valign: "middle", objectName: "concl-text-" + i });
  });
  card(s, 0.6, 5.75, 5.85, 0.95, C.background2, "limits");
  txt(s, [{ text: "Шектеулер: ", options: { bold: true, color: C.text2 } }, { text: "30 тіркес, 3 ЖИ-дің тегін нұсқалары", options: { color: C.text1 } }],
    { x: 0.85, y: 5.75, w: 5.4, h: 0.95, fontSize: 14, valign: "middle", objectName: "limits-text" });
  card(s, 6.85, 5.75, 5.85, 0.95, C.background2, "next");
  txt(s, [{ text: "Келешекте: ", options: { bold: true, color: C.text2 } }, { text: "мақал-мәтелдер, сөйлем ішіндегі тіркестер", options: { color: C.text1 } }],
    { x: 7.1, y: 5.75, w: 5.4, h: 0.95, fontSize: 14, valign: "middle", objectName: "next-text" });
  s.addNotes(`Қорытындылай келе: бірінші – ЖИ тұрақты тіркестерді жалпы жақсы түсіндіреді. Екінші – оның әлсіз тұсы ұлттық-мәдени тіркестер. Үшінші – ЖИ сөзбе-сөз аудармайды, бірақ қате мағынаны сенімді жазады. Төртінші – ЖИ жақсы көмекші, бірақ шешімді адам қабылдауы керек. Жұмыстың шектеулері де бар: мен 30 тіркес пен үш ЖИ-дің тегін нұсқаларын ғана тексердім. Келешекте мақал-мәтелдерді зерттегім келеді. (40 секунд)`);

  // ================= 13. Рахмет =================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Қорытынды" });
  s.addText("Шешімді адам қабылдайды – ЖИ тек көмектеседі", { placeholder: "title" });
  s.addText("Назарларыңызға рахмет! Сұрақтарыңызға жауап беруге дайынмын", { placeholder: "body" });
  [["Тәжірибе", "182 жауап"], ["Орташа сапа", f1(N.e1_all_avg) + "%"], ["Қате жауаптар", "4 / 90"]].forEach(([l, vv], i) => {
    const y = 1.6 + i * 1.5;
    card(s, 8.9, y, 3.7, 1.25, C.background1, "end-stat-" + i);
    txt(s, [{ text: vv, options: { fontSize: 30, bold: true, color: C.text2, fontFace: THEME.headFontFace, breakLine: true } }, { text: l, options: { fontSize: 14, color: C.accent5 } }],
      { x: 9.2, y, w: 3.2, h: 1.25, valign: "middle", objectName: "end-stat-text-" + i });
  });
  s.addNotes(`Менің басты түйінім: шешімді адам қабылдайды, ЖИ тек көмектеседі. Назарларыңызға рахмет!

ЫҚТИМАЛ СҰРАҚТАР:
1) Неге осы үш ЖИ? – Тегін, оқушыларға қолжетімді және қазақша жауап береді.
2) Баллды кім қойды, әділ ме? – Өзім, алдын ала жазылған шкала бойынша (2, 1, 0) және сөздікке қарап. Күмәнді жауаптарды жетекшіммен қайта қарадым.
3) Неге әр сұранысты жаңа чатта бердің? – Алдыңғы жауап келесісіне әсер етпеуі үшін.
4) Нәтиже ертең өзгере ме? – Мүмкін, ЖИ жаңарып тұрады. Сондықтан тексеру уақытын (2026 ж. қыркүйек) көрсеттім.
5) Сөзбе-сөз қате неге болмады? – Бүгінгі ЖИ өте көп мәтін оқыған, белгілі тіркестерді жақсы біледі. Қиындық сирек, ұлттық-мәдени тіркестерде.
6) Өзің не үйрендің? – Сөздікпен жұмыс істеуді, деректерді кестеге жинап есептеуді, ЖИ жауабын тексеруді.
7) ЖИ-ді жұмысты жазуға пайдаландың ба? – Иә, декларацияда ашық көрсеттім: құрылымын жоспарлау мен мәтінді редакциялауға. Тәжірибені өзім жүргіздім, баллды өзім қойдым.`);

  // ================= Қосымша слайдтар (сұрақтарға) =================
  pres.addSection({ title: "Сұрақтарға арналған қосымша" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Сұрақтарға арналған қосымша" });
  s.addText("Қосымша: бағалау шкаласы және қате түрлері", { placeholder: "title" });
  const hdr = o => ({ text: o, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } });
  const rows = [[hdr("Белгі"), hdr("Атауы"), hdr("Сипаттамасы"), hdr("Саны")],
    ...D.score.map(([a, b, c]) => [a, b, c, String([1, 2, 3].reduce((x, k) => x + N[`e1_${a === "2" ? "ok" : a === "1" ? "half" : "zero"}_${k}`], 0)) + " (1-тәж.)"]),
    ...D.err.map(([a, b, c]) => [a, b, c, String(N[`e1_err_${a}_sum`] + N[`e2_err_${a}_sum`])])];
  s.addTable(rows, { x: 0.6, y: 1.55, w: 12.1, colW: [1.1, 2.6, 6.6, 1.8], fontSize: 14, fontFace: "+mn-lt", color: HEX.dk1,
    border: { type: "solid", pt: 0.75, color: "C9DDE0" }, valign: "middle", rowH: 0.62, objectName: "table-scale" });
  s.addNotes("Бұл слайд сұрақ қойылса көрсетуге арналған: балл қою ережесі мен қате түрлерінің саны.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Сұрақтарға арналған қосымша" });
  s.addText("Қосымша: Б тобындағы 10 тіркестің нәтижесі", { placeholder: "title" });
  const scoreCell = v => ({ text: v === null || v === undefined ? "–" : String(v), options: { align: "center", bold: true,
    color: v === 2 ? HEX.accent4 : v === 1 ? "B07D00" : HEX.accent3 } });
  const bRows = D.e1.map((r, i) => [r, i]).filter(([r]) => r[0] === "Б").map(([r, i]) => [r[1], r[2], ...[1, 2, 3].map(k => scoreCell(N[`e1_scores_${k}`][i]))]);
  s.addTable([[hdr("Тіркес"), hdr("Сөздікте"), ...AIS.map(a => ({ text: a, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: "center" } }))], ...bRows],
    { x: 0.6, y: 1.5, w: 12.1, colW: [4.0, 4.0, 1.37, 1.37, 1.36], fontSize: 14, fontFace: "+mn-lt", color: HEX.dk1,
      border: { type: "solid", pt: 0.75, color: "C9DDE0" }, valign: "middle", rowH: 0.47, objectName: "table-groupB" });
  s.addNotes("Сұрақ қойылса: ұлттық-мәдени тіркестердің әрқайсысы бойынша үш ЖИ-дің балы. 2 – дұрыс, 0 – қате.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });
