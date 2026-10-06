// Қорғауға арналған презентация: PPTX_SKILL=<pptx skill> node deck.js ../../Qorgau_prezentatsiyasy.pptx
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const D = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "data.json"), "utf8"));
const N = D.natije;
const f1 = v => String(Math.round(v * 10) / 10).replace(".", ",");
const W = N.words, WB = Object.fromEntries(W.map(x => [x.word, x]));
const GA = N.gen_avg, T = N.test;
const PILOT = N.n_complete !== D.gens.length * D.n_per_gen;
const q = w => `«${w}»`;
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];

// Ескі қолжазба: сия-күлгін, алтын, тот-қызыл; архаизм – көк, тарихи сөз – қызыл
const THEME = {
  name: "Zerde Konergen sozder",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "22202E", lt1: "FFFFFF", dk2: "3A2E5C", lt2: "F0EEF6",
    accent1: "2A78D6", accent2: "D99A1E", accent3: "B5462F", accent4: "2E8B57",
    accent5: "5F5B6E", accent6: "1B9E9E", hlink: "2A78D6", folHlink: "1B9E9E",
  },
};
const HEX = THEME.colors;
const ARX = HEX.accent1, TAR = HEX.accent3; // архаизм – көк, тарихи сөз – қызыл
const GEN_COL = { "О": HEX.accent1, "Ә": HEX.accent2, "Ү": HEX.accent4 };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Қазақ тіліндегі көнерген сөздердің қазіргі қолданыстағы орны";
const C = pres.SchemeColor;

// ---------- Макеттер ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.55, w: 7.3, h: 2.7, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.4, w: 7.3, h: 0.9, fontSize: 18, italic: true, color: C.accent2, valign: "top", align: "left", margin: 0 }, text: "" } },
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
  const pad = d * 0.24;
  s.addImage({ data: await icon(Comp), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: name + "-icon" });
}

// ---------- Көмекші элементтер ----------
const txt = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, ...o });
function card(s, x, y, w, h, fill, name, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill },
    line: line ? { color: line, width: 1.25 } : { type: "none" }, objectName: name });
}
function bigStat(s, x, y, w, value, label, color, name, size = 48) {
  txt(s, value, { x, y, w, h: 0.85, fontSize: size, bold: true, color, fontFace: THEME.headFontFace, objectName: name + "-value" });
  txt(s, label, { x, y: y + 0.88, w, h: 0.75, fontSize: 14, color: C.text1, valign: "top", objectName: name + "-label" });
}
// Мотив: ойын «карточкасы» – алтын жиекті дөңгелек бұрышты тақта
function gameChip(s, x, y, w, h, text, name, o = {}) {
  card(s, x, y, w, h, o.fill || C.background1, name, o.line || C.accent2);
  txt(s, text, { x: x + 0.12, y, w: w - 0.24, h, fontSize: o.size || 16, bold: true, color: o.color || C.text2,
    align: "center", valign: "middle", fontFace: THEME.headFontFace, objectName: name + "-text" });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };
const chartQuiet = () => ({
  ...chartText, catAxisLabelColor: HEX.dk1, valAxisLabelColor: HEX.accent5, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  dataLabelFontSize: 12, dataLabelColor: HEX.dk1, valGridLine: { color: "D9E4E7", size: 0.75 }, catGridLine: { style: "none" },
  legendFontSize: 13, legendColor: HEX.dk1,
});
const legendDot = (s, x, y, color, label, name) => {
  s.addShape(pres.shapes.OVAL, { x, y: y + 0.08, w: 0.22, h: 0.22, fill: { color }, line: { type: "none" }, objectName: name + "-dot" });
  txt(s, label, { x: x + 0.32, y, w: 1.6, h: 0.38, fontSize: 14, color: C.text1, valign: "middle", objectName: name + "-label" });
};
const wordChip = gameChip; // мотив: көне сөз – алтын жиекті карточка
const ok = v => v === null || v === undefined ? "–" : f1(v) + "%";
const nA = D.words.filter(w => w[2] === "А").length, nT = D.words.filter(w => w[2] === "Т").length;
const MK = N.marks_by_kind;
const aNo = (MK["А"] || {})["белгі жоқ"] || 0, tMarked = ((MK["Т"] || {})["көн."] || 0) + ((MK["Т"] || {})["тар."] || 0);
const noMarkA = W.filter(x => x.kind === "А" && x.mark === "белгі жоқ").map(x => x.word);
const noMarkT = W.filter(x => x.kind === "Т" && x.mark === "белгі жоқ").map(x => x.word);
const allKnow = D.test_words.filter(w => D.gens.every(([g]) => T[w][g].ok === 100));
const oZero = D.test_words.filter(w => T[w]["О"].ok === 0 && Math.max(T[w]["Ә"].ok || 0, T[w]["Ү"].ok || 0) > 0);
const who = D.gens.filter(([g]) => N.n_by_gen[g]).map(([g]) => `${N.n_by_gen[g]} ${{ "О": "оқушы", "Ә": "ата-ана", "Ү": "ата-әже" }[g]}`);
const okA = aNo > nA / 2, okB = tMarked > nT / 2;
const okOrder = GA["Ү"] > GA["Ә"] && GA["Ә"] > GA["О"], okLow = GA["О"] < GA["Ә"] && GA["О"] < GA["Ү"];
const STN = Object.fromEntries(D.status.map(s => [s[0], s[1]]));
const ST_COL = { "Ж": HEX.accent4, "А": HEX.accent1, "Т": HEX.accent2, "Ұ": "8E8A99" };
const mon = W.filter(x => x.h1 !== null);
const dul = WB["дулыға"];

(async () => {
  let s;
  // ================= 1. Титул =================
  pres.addSection({ title: "Кіріспе" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Кіріспе" });
  txt(s, "«Зерде» республикалық конкурсы  ·  Қазақ тілі және әдебиеті  ·  «Орта буын»", { x: 0.8, y: 0.6, w: 11.5, h: 0.4, fontSize: 14, color: C.accent2, objectName: "kicker" });
  s.addText("Қазақ тіліндегі көнерген сөздердің қазіргі қолданыстағы орны", { placeholder: "title" });
  s.addText(`36 сөз · түсіндірме сөздік · баспасөз · ${N.n_people} адамға тест`, { placeholder: "body" });
  txt(s, "Тіркеу коды (шифр): ______________", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.background2, objectName: "cipher" });
  ["сауыт", "теңге", "тамұқ", "дулыға", "болыс", "тұлпар"].forEach((t, i) =>
    wordChip(s, 8.8 + (i % 2) * 2.0, 1.55 + Math.floor(i / 2) * 1.35, 1.8, 1.05, t, "title-chip-" + i, { size: 20 }));
  s.addNotes("Құрметті қазылар алқасы! Менің жобамның тақырыбы – «Қазақ тіліндегі көнерген сөздердің қазіргі қолданыстағы орны». Мен 36 көне сөзді түсіндірме сөздіктен тексеріп, олардың бүгінгі баспасөзде қалай қолданылатынын қарадым және үш буынға тест жүргіздім. (15 секунд)");

  // ================= 2. Мәселе: дулыға =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Дулыға: соғыстан – спортқа", { placeholder: "title" });
  card(s, 0.6, 1.55, 5.85, 3.3, C.background2, "hook-old");
  await iconCircle(s, fa.FaBook, 0.95, 1.85, 0.75, C.accent3, "hook-old-ic");
  txt(s, "Сөздікте", { x: 1.9, y: 1.85, w: 4.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "hook-old-t" });
  txt(s, `«${dul.meaning}»`, { x: 0.95, y: 2.8, w: 5.2, h: 1.6, fontSize: 17, italic: true, color: C.text1, valign: "top", objectName: "hook-old-text" });
  txt(s, `Белгісі: «${dul.mark}» – көнерген`, { x: 0.95, y: 4.3, w: 5.2, h: 0.4, fontSize: 14, bold: true, color: C.accent3, objectName: "hook-old-mark" });
  card(s, 6.85, 1.55, 5.85, 3.3, C.background2, "hook-new");
  await iconCircle(s, fa.FaNewspaper, 7.2, 1.85, 0.75, C.accent1, "hook-new-ic");
  txt(s, "Бүгінгі газетте", { x: 8.15, y: 1.85, w: 4.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "hook-new-t" });
  txt(s, `«${dul.example}»`, { x: 7.2, y: 2.8, w: 5.2, h: 1.6, fontSize: 17, italic: true, color: C.text1, valign: "top", objectName: "hook-new-text" });
  txt(s, "egemen.kz", { x: 7.2, y: 4.3, w: 5.2, h: 0.4, fontSize: 14, bold: true, color: C.accent1, objectName: "hook-new-src" });
  card(s, 0.6, 5.2, 12.1, 1.35, C.text2, "hook-q");
  txt(s, "Көнерген сөздер шынымен тілден шығып кетті ме, әлде олар жаңа жерде өмір сүре ме?",
    { x: 1.0, y: 5.2, w: 11.3, h: 1.35, fontSize: 22, bold: true, color: C.background1, valign: "middle", objectName: "hook-q-text" });
  s.addNotes("Бәрі «дулыға» сөзінен басталды. Сөздікте ол «көн.» деп белгіленген: соғысқа шыққан батырлардың темір бас киімі. Ал «Egemen Qazaqstan» сайтынан осы сөзді іздегенде, мен оны дзюдодан өтетін «Үлкен дулыға» турнирінің атауынан таптым. Сонда менде сұрақ туды: көнерген сөздер шынымен тілден шығып кетті ме, әлде олар жаңа жерде өмір сүре ме? (30 секунд)");

  // ================= 3. Сұрақ, мақсат, болжам =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеу сұрағы мен болжам", { placeholder: "title" });
  card(s, 0.6, 1.55, 12.1, 1.45, C.background2, "q-card");
  await iconCircle(s, fa.FaQuestion, 0.95, 1.85, 0.85, C.accent1, "q");
  txt(s, "Көнерген сөздер бүгін қайда және қандай мағынада қолданылады және оларды әр буын қаншалықты біледі?",
    { x: 2.1, y: 1.55, w: 10.3, h: 1.45, fontSize: 21, bold: true, color: C.text2, valign: "middle", objectName: "q-text" });
  const qa = [
    [fa.FaBullseye, "Мақсаты", "Көне сөздердің сөздіктегі белгісін, бүгінгі баспасөздегі қолданысын және үш буынның оларды білуін зерттеп, мини-сөздік жасау", C.accent1],
    [fa.FaLightbulb, "Болжам", "Қайта қолданысқа енген архаизмдер сөздікте көнерген деп белгіленбеген, тарихи сөздердің көбі белгіленген; көне сөздерді ата-әжелер жақсы, оқушылар нашар біледі", C.accent2],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25;
    card(s, x, 3.35, 5.85, 3.3, C.background1, "aim-card-" + i, "D5D0E3");
    await iconCircle(s, qa[i][0], x + 0.35, 3.65, 0.75, qa[i][3], "aim-" + i);
    txt(s, qa[i][1], { x: x + 1.3, y: 3.65, w: 4.3, h: 0.75, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: "aim-title-" + i });
    txt(s, qa[i][2], { x: x + 0.35, y: 4.6, w: 5.2, h: 1.9, fontSize: 17, color: C.text1, valign: "top", objectName: "aim-text-" + i });
  }
  s.addNotes("Менің зерттеу сұрағым: көнерген сөздер бүгін қайда және қандай мағынада қолданылады және оларды әр буын қаншалықты біледі? Болжамым үш бөліктен тұрды: бүгін қайта қолданысқа енген архаизмдер сөздікте көнерген деп белгіленбеген, тарихи сөздердің көбі белгіленген, ал көне сөздерді ата-әжелер жақсы, оқушылар нашар біледі. (30 секунд)");

  // ================= 4. Теория =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Көнерген сөздің екі түрі және қайта жандану", { placeholder: "title" });
  const kinds = [
    ["Тарихи сөздер", "Заты да, атауы да қолданыстан шыққан", ["сауыт", "болыс", "батпан"], TAR],
    ["Архаизмдер", "Заты бар, атауы ескірген не ауысқан", ["ләшкер → әскер", "тамұқ → тозақ", "мейман → қонақ"], ARX],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25, [h, d, ex, col] = kinds[i];
    card(s, x, 1.55, 5.85, 3.2, C.background2, "kind-card-" + i);
    txt(s, h, { x: x + 0.35, y: 1.75, w: 5.2, h: 0.6, fontSize: 24, bold: true, color: col, fontFace: THEME.headFontFace, objectName: "kind-h-" + i });
    txt(s, d, { x: x + 0.35, y: 2.35, w: 5.2, h: 0.5, fontSize: 16, color: C.text1, objectName: "kind-d-" + i });
    ex.forEach((e, k) => wordChip(s, x + 0.35, 3.0 + k * 0.55, 3.6, 0.45, e, `kind-ex-${i}-${k}`, { size: 15 }));
  }
  card(s, 0.6, 5.0, 12.1, 1.6, C.background1, "revive", C.accent4);
  await iconCircle(s, fa.FaRedo, 0.9, 5.35, 0.85, C.accent4, "revive-ic");
  txt(s, [{ text: "Қайта жандану: ", options: { bold: true, color: HEX.accent4 } },
    { text: "тәуелсіздік жылдары «әкім» (1990-жылдар), «теңге» мен «тиын» (1993), «Мәжіліс», «Ұлттық құрылтай» (2022) сияқты көне сөздер мемлекеттік атауға айналды", options: { color: HEX.dk1 } }],
    { x: 2.0, y: 5.0, w: 10.5, h: 1.6, fontSize: 17, valign: "middle", objectName: "revive-text" });
  s.addNotes("Ғалымдар көнерген сөздерді екі түрге бөледі. Тарихи сөздердің заты да, атауы да қолданыстан шыққан: сауыт, болыс, батпан. Архаизмдердің заты бар, бірақ атауы ескірген: ләшкер орнына әскер, тамұқ орнына тозақ дейміз. Бірақ кейбір көне сөздер тілге қайта оралады. Тәуелсіздік жылдары «әкім», «теңге», «тиын», «Мәжіліс», «Ұлттық құрылтай» сияқты сөздер мемлекеттік атауға айналды. (40 секунд)");

  // ================= 5. Әдістеме =================
  pres.addSection({ title: "Әдістеме" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Әдістеме" });
  s.addText("Қалай зерттедім: үш әдіс", { placeholder: "title" });
  const steps = [
    [fa.FaBookOpen, "Сөздік", "36 сөзді «Қазақ әдеби тілінің сөздігінен» (15 томдық) тауып, «көн.» не «тар.» белгісін жаздым", C.accent3],
    [fa.FaNewspaper, "Баспасөз", "Әскери топтағы 4 сөзді egemen.kz сайтынан соңғы бір жыл бойынша іздеп, мысал сөйлемін талдадым", C.accent2],
    [fa.FaUsers, "Үш буын тесті", `12 сөздің мағынасын оқушылардан, ата-аналардан, ата-әжелерден сұрадым: ${listJoin(who)}`, C.accent1],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.6, 3.75, 4.3, C.background2, "step-card-" + i);
    await iconCircle(s, steps[i][0], x + 0.35, 1.9, 0.95, steps[i][3], "step-" + i);
    txt(s, `${i + 1}`, { x: x + 2.6, y: 1.85, w: 0.9, h: 1.0, fontSize: 44, bold: true, color: "CFC8E0", fontFace: THEME.headFontFace, align: "right", objectName: "step-num-" + i });
    txt(s, steps[i][1], { x: x + 0.35, y: 3.05, w: 3.1, h: 0.6, fontSize: 22, bold: true, color: C.text2, objectName: "step-title-" + i });
    txt(s, steps[i][2], { x: x + 0.35, y: 3.7, w: 3.1, h: 2.1, fontSize: 15, color: C.text1, valign: "top", objectName: "step-text-" + i });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 3.55, w: 0.3, h: 0.4, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "step-arrow-" + i });
  }
  txt(s, "Тест: 2 – дұрыс түсіндірді, 1 – шамамен, 0 – білмейді. Есімдер жазылмады, келісім алынды.", { x: 0.6, y: 6.15, w: 12.1, h: 0.45, fontSize: 14, italic: true, color: C.accent5, objectName: "method-note" });
  s.addNotes(`Зерттеуде үш әдіс қолдандым. Біріншісі – сөздікпен тексеру: 36 сөздің бәрін 15 томдық «Қазақ әдеби тілінің сөздігінен» тауып, жанындағы белгісін жаздым. Екіншісі – баспасөз: әскери топтағы 4 сөзді «Egemen Qazaqstan» сайтынан соңғы бір жыл бойынша іздедім. Үшіншісі – үш буынға тест: 12 сөздің мағынасын түсіндіріп беруді сұрадым. Әзірге тестке ${N.n_people} адам қатысты. Қатысушылардың есімін жазбадым, келісімін алдым. (45 секунд)`);

  // ================= 6. Сөздік =================
  pres.addSection({ title: "Нәтижелер" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText(`Сөздікте ${N.marks["белгі жоқ"]} көне сөзде белгі жоқ`, { placeholder: "title" });
  const marks = ["көн.", "тар.", "белгі жоқ", "сөздікте жоқ"];
  s.addChart(pres.charts.BAR, marks.map(m => ({ name: m, labels: ["Архаизмдер (19)", "Тарихи сөздер (17)"], values: [(MK["А"] || {})[m] || 0, (MK["Т"] || {})[m] || 0] })),
    { x: 0.6, y: 1.45, w: 7.6, h: 4.2, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.accent2, HEX.accent3, HEX.accent1, "A7A3B3"], ...chartQuiet(),
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: "0;;;", showLegend: true, legendPos: "b",
      showTitle: true, title: "Сөздіктегі белгі, сөз саны", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-marks" });
  bigStat(s, 8.7, 1.5, 4.0, `${N.marks["белгі жоқ"]} / 36`, "сөзде сөздікте ешқандай белгі жоқ", C.accent1, "marks-none");
  bigStat(s, 8.7, 3.15, 4.0, `${aNo} / ${nA}`, "архаизмде белгі жоқ – олардың көбі бүгін күнде айтылады", C.text2, "marks-arx");
  card(s, 0.6, 5.85, 12.1, 0.85, C.background2, "marks-list");
  txt(s, [{ text: "Белгісі жоқ архаизмдер: ", options: { bold: true, color: HEX.dk2 } }, { text: noMarkA.join(", "), options: { color: HEX.dk1 } }],
    { x: 0.85, y: 5.85, w: 11.6, h: 0.85, fontSize: 15, valign: "middle", objectName: "marks-list-text" });
  s.addNotes(`Сөздікпен тексеру мені таңғалдырды. Көнерген деп алған 36 сөздің ${N.marks["белгі жоқ"]}-ында сөздікте ешқандай белгі жоқ. ${nA} архаизмнің ${aNo}-інде белгі жоқ: сарбаз, әкім, елші, мәжіліс, кеден, теңге, тиын, тұлпар және басқалар. Олардың көбі – бүгін мемлекеттік өмірде күнделікті айтылатын сөздер. Ал ${nT} тарихи сөздің ${tMarked}-ы «көн.» не «тар.» деп белгіленген. (40 секунд)`);

  // ================= 7. Сөздік артта қалды =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Сөздік тілдің өзгерісінен артта қалады", { placeholder: "title" });
  const lag = [
    ["жарлық", WB["жарлық"], "Мемлекет басшысының Жарлығы"],
    ["құрылтай", WB["құрылтай"], "Ұлттық құрылтай (2022)"],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.55 + i * 2.1, [w, x, now] = lag[i];
    wordChip(s, 0.6, y + 0.35, 2.3, 1.1, w, "lag-chip-" + i, { size: 22 });
    card(s, 3.2, y, 4.6, 1.8, C.background2, "lag-old-" + i);
    txt(s, [{ text: `Сөздікте: «${x.mark}»\n`, options: { bold: true, color: HEX.accent3, breakLine: true } }, { text: x.meaning, options: { color: HEX.dk1, italic: true } }],
      { x: 3.4, y, w: 4.2, h: 1.8, fontSize: 14, valign: "middle", objectName: "lag-old-text-" + i });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 7.95, y: y + 0.7, w: 0.45, h: 0.4, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "lag-arrow-" + i });
    card(s, 8.55, y, 4.15, 1.8, C.background1, "lag-new-" + i, C.accent4);
    txt(s, [{ text: "Бүгін:\n", options: { bold: true, color: HEX.accent4, breakLine: true } }, { text: now, options: { color: HEX.dk1, bold: true } }],
      { x: 8.8, y, w: 3.7, h: 1.8, fontSize: 17, valign: "middle", objectName: "lag-new-text-" + i });
  }
  card(s, 0.6, 5.85, 12.1, 0.85, C.background2, "lag-note");
  txt(s, [{ text: "Белгісі жоқ тарихи сөздер: ", options: { bold: true, color: HEX.dk2 } }, { text: `${noMarkT.join(", ")} – олардың заты (той, қымыз, жүзу) бүгін де бар`, options: { color: HEX.dk1 } }],
    { x: 0.85, y: 5.85, w: 11.6, h: 0.85, fontSize: 15, valign: "middle", objectName: "lag-note-text" });
  s.addNotes("Ең қызық қарама-қайшылықты «жарлық» пен «құрылтай» сөздерінен көрдім. Сөздікте олар «тар.» деп, тек ескі мағынасымен берілген. Ал бүгін «Мемлекет басшысының Жарлығы», «Ұлттық құрылтай» деп жаңалықтан күн сайын естиміз. Яғни сөздік тілдің өзгерісінен артта қалып отыр. Керісінше, сауыт, сәукеле, саба, құлаш сияқты тарихи сөздерде белгі жоқ – олардың заты тойда, қымыз ашытқанда бүгін де бар. (40 секунд)");

  // ================= 8. Баспасөз =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Әскери сөздер газетте әлі бар", { placeholder: "title" });
  const ctxNote = { "сауыт": "Мұражайдағы «Атадан қалған ақ сауыт» көрмесі", "дулыға": "Дзюдодан «Үлкен дулыға» турнирі" };
  mon.forEach((x, i) => {
    const y = 1.55 + i * 1.15;
    wordChip(s, 0.6, y, 2.4, 0.95, x.word, "mon-chip-" + i, { size: 20 });
    txt(s, [{ text: String(x.h1), options: { fontSize: 30, bold: true, color: HEX.dk2, fontFace: THEME.headFontFace } }, { text: "  рет", options: { fontSize: 14, color: HEX.dk1 } }],
      { x: 3.3, y, w: 1.8, h: 0.95, valign: "middle", objectName: "mon-n-" + i });
    card(s, 5.3, y, 7.4, 0.95, C.background2, "mon-card-" + i);
    s.addShape(pres.shapes.OVAL, { x: 5.5, y: y + 0.3, w: 0.35, h: 0.35, fill: { color: ST_COL[x.status] }, line: { type: "none" }, objectName: "mon-dot-" + i });
    txt(s, [{ text: STN[x.status], options: { bold: true, color: HEX.dk2 } }, { text: ctxNote[x.word] ? ` – ${ctxNote[x.word]}` : " – тарихи-мәдени мәтінде", options: { color: HEX.dk1 } }],
      { x: 6.05, y, w: 6.5, h: 0.95, fontSize: 15, valign: "middle", objectName: "mon-text-" + i });
  });
  txt(s, "egemen.kz, соңғы бір жыл, Google іздеу нәтижесі (шамамен). Мониторинг әскери топтағы 4 сөзге жүргізілді.", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 13, italic: true, color: C.accent5, objectName: "mon-note" });
  s.addNotes("Әскери топтағы 4 сөзді «Egemen Qazaqstan» сайтынан соңғы бір жыл бойынша іздедім. Төртеуі де газетте әлі бар: садақ пен қорамсақ – шамамен 10 реттен, сауыт – 8, дулыға – 5 рет. Бірақ «сауыт», «садақ», «қорамсақ» тарихи-мәдени мәтінде, мысалы мұражай көрмесі туралы жаңалықта кездесті. Ал «дулыға» спорт жарысының атауына айналған, сондықтан оған «атауда сақталған» мәртебесін бердім. (35 секунд)");

  // ================= 9. Тест =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Көне сөздерді жастар нашар біледі", { placeholder: "title" });
  s.addChart(pres.charts.BAR, D.gens.map(([g, n]) => ({ name: n.split(" (")[0], labels: D.test_words, values: D.test_words.map(w => T[w][g].ok || 0) })),
    { x: 0.6, y: 1.4, w: 8.2, h: 5.35, barDir: "col", barGrouping: "clustered", barGapWidthPct: 50, chartColors: D.gens.map(([g]) => GEN_COL[g]), ...chartQuiet(),
      catAxisLabelFontSize: 12, catAxisLabelRotate: -35, valAxisMaxVal: 100, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"", showLegend: true, legendPos: "b",
      showTitle: true, title: "Дұрыс түсіндіргендер үлесі, %", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-test" });
  D.gens.forEach(([g, n], i) => {
    const y = 1.5 + i * 1.3;
    txt(s, [{ text: ok(GA[g]), options: { fontSize: 36, bold: true, color: GEN_COL[g], fontFace: THEME.headFontFace, breakLine: true } }, { text: n.split(" (")[0].toLowerCase(), options: { fontSize: 15, color: HEX.dk1 } }],
      { x: 9.2, y, w: 3.5, h: 1.15, valign: "middle", objectName: "gen-stat-" + i });
  });
  card(s, 9.2, 5.45, 3.5, 1.3, C.background2, "pilot-note");
  txt(s, PILOT ? `Алдын ала нәтиже: ${N.n_people} адам (${listJoin(who)})` : `${N.n_people} адам, әр буыннан ${D.n_per_gen}`,
    { x: 9.4, y: 5.45, w: 3.1, h: 1.3, fontSize: 14, italic: true, color: C.text1, valign: "middle", objectName: "pilot-note-text" });
  s.addNotes(`Тест нәтижесі: оқушылар көне сөздердің ${ok(GA["О"])}-ын ғана дұрыс түсіндірді, ата-аналар ${ok(GA["Ә"])}, ата-әжелер ${ok(GA["Ү"])}. ${PILOT ? `Бұл – алдын ала нәтиже, тестке әзірге ${N.n_people} адам қатысты: ${listJoin(who)}. ` : ""}Яғни көне сөздерді жастар ең нашар біледі. (35 секунд)`);

  // ================= 10. Тест: қай сөздер =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Тірі сөзді бәрі біледі", { placeholder: "title" });
  const cols2 = [
    [fa.FaCheckCircle, "Барлық буын білді", allKnow, "Бүгін жиі айтылады, сөздікте белгісі жоқ", HEX.accent4],
    [fa.FaQuestionCircle, "Бірде-бір оқушы білмеді", oZero, "Бірақ үлкендердің ішінде білетіндер болды", HEX.accent3],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25, [Ic, h, ws, d, col] = cols2[i];
    card(s, x, 1.55, 5.85, 4.6, C.background2, "tw-card-" + i);
    await iconCircle(s, Ic, x + 0.35, 1.85, 0.8, col, "tw-ic-" + i);
    txt(s, h, { x: x + 1.35, y: 1.85, w: 4.3, h: 0.8, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: "tw-h-" + i });
    ws.forEach((w, k) => wordChip(s, x + 0.35 + (k % 2) * 2.65, 2.95 + Math.floor(k / 2) * 0.95, 2.45, 0.75, w, `tw-chip-${i}-${k}`, { size: 18 }));
    txt(s, d, { x: x + 0.35, y: 5.0, w: 5.2, h: 0.9, fontSize: 15, italic: true, color: C.text1, valign: "top", objectName: "tw-d-" + i });
  }
  s.addNotes(`Қай сөздерді білетіні де қызық. ${listJoin(allKnow.map(q))} сөздерін барлық буын дұрыс түсіндірді. Бұл сөздер бүгін де жиі айтылады және сөздікте көнерген деп белгіленбеген. Ал ${listJoin(oZero.map(q))} сөздерін бірде-бір оқушы дұрыс түсіндірмеді, бірақ үлкендердің ішінде оларды білетіндер болды. Яғни сөз қолданыстан шыққан сайын, оны жастар ұмытады. (35 секунд)`);

  // ================= 11. Болжам =================
  pres.addSection({ title: "Қорытынды" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  const allOk = okA && okB && okOrder;
  s.addText(allOk ? "Болжамым расталды" : "Болжамым ішінара расталды", { placeholder: "title" });
  const hyp = [
    ["Қайта қолданысқа енген архаизмдер сөздікте белгіленбеген", `${okA ? "Расталды" : "Расталмады"}: 19 архаизмнің ${aNo}-інде белгі жоқ`, okA ? "y" : "n"],
    ["Тарихи сөздердің көбі «көн.» не «тар.» деп белгіленген", `${okB ? "Расталды" : "Расталмады"}: 17 тарихи сөздің ${tMarked}-ы белгіленген`, okB ? "y" : "n"],
    ["Ата-әжелер ең жақсы, оқушылар ең нашар біледі", okOrder ? `Расталды: ${ok(GA["Ү"])} > ${ok(GA["Ә"])} > ${ok(GA["О"])}` : okLow ? `Ішінара: оқушылар ең төмен (${ok(GA["О"])}), бірақ ата-әже (${ok(GA["Ү"])}) ата-анадан (${ok(GA["Ә"])}) төмен` : "Расталмады", okOrder ? "y" : okLow ? "h" : "n"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.55 + i * 1.55, [h, r, st] = hyp[i];
    card(s, 0.6, y, 12.1, 1.35, st === "y" ? "E3F2EA" : st === "h" ? "FBF0DA" : "F8E6E1", "hyp-row-" + i);
    await iconCircle(s, st === "y" ? fa.FaCheck : st === "h" ? fa.FaAdjust : fa.FaTimes, 0.9, y + 0.3, 0.75, st === "y" ? C.accent4 : st === "h" ? C.accent2 : C.accent3, "hyp-mark-" + i);
    txt(s, h, { x: 1.95, y, w: 5.0, h: 1.35, fontSize: 17, bold: true, color: C.text2, valign: "middle", objectName: "hyp-h-" + i });
    txt(s, r, { x: 7.1, y, w: 5.4, h: 1.35, fontSize: 15, color: C.text1, valign: "middle", objectName: "hyp-r-" + i });
  }
  if (PILOT) txt(s, "Үшінші бөлік – алдын ала тест нәтижесі: ата-ана мен ата-әжеден аз адам қатысты", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: C.accent5, objectName: "hyp-note" });
  s.addNotes(`Болжамымды үш бөлік бойынша тексердім. Бірінші бөлігі расталды: архаизмдердің көбінде сөздікте белгі жоқ. Екінші бөлігі де расталды: тарихи сөздердің көбі белгіленген. Үшінші бөлігі ${okOrder ? "расталды" : "ішінара расталды: оқушылар шынымен ең нашар білді, бірақ ата-әже ата-анадан төмен нәтиже көрсетті"}. ${PILOT ? "Бұл бөлігі алдын ала нәтиже, өйткені ата-ана мен ата-әжеден аз адам қатысты." : ""} (35 секунд)`);

  // ================= 12. Өнім =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Өнім: «Көне сөз – жаңа тыныс» мини-сөздігі", { placeholder: "title" });
  const img = path.join(__dirname, "..", "figs", "pasport.png");
  const buf = fs.readFileSync(img), iw = buf.readUInt32BE(16), ih = buf.readUInt32BE(20);
  s.addImage({ path: img, x: 0.6, y: 1.5, w: 7.0, h: 7.0 * ih / iw, objectName: "pasport-image" });
  const use = [
    [fa.FaChalkboardTeacher, "Қазақ тілі сабағында", "«Көнерген сөздер» тақырыбын өткенде"],
    [fa.FaBookReader, "Әдебиет сабағында", "Батырлар жыры мен тарихи шығармаларды оқығанда"],
    [fa.FaTags, "Атау таңдағанда", "Дүкенге, жобаға, командаға ұлттық атау іздегенде"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.55 + i * 1.75;
    await iconCircle(s, use[i][0], 8.1, y + 0.1, 0.8, [C.accent1, C.accent2, C.accent3][i], "use-" + i);
    txt(s, use[i][1], { x: 9.1, y, w: 3.6, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "use-title-" + i });
    txt(s, use[i][2], { x: 9.1, y: y + 0.55, w: 3.6, h: 0.9, fontSize: 15, color: C.text1, valign: "top", objectName: "use-text-" + i });
  }
  s.addNotes("Нәтиже бойынша «Көне сөз – жаңа тыныс» мини-сөздігін жасадым. Онда 36 сөздің әрқайсысына «сөз паспорты» бар: сөздік бойынша мағынасы, түрі, сөздіктегі белгісі және тексерілген сөздерге – бүгінгі газеттен мысал. Мысалы, мына «дулыға» сөзінің паспорты. Мини-сөздікті қазақ тілі мен әдебиет сабағында және атау таңдағанда пайдалануға болады. (30 секунд)");

  // ================= 13. Қорытынды =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Қорытынды", { placeholder: "title" });
  const concl = [
    `Сөздікте 36 көне сөздің ${N.marks["белгі жоқ"]}-ында белгі жоқ: қайта жанданған архаизмдер жалпы сөз ретінде берілген`,
    "Сөздік тілден артта қалады: «жарлық», «құрылтай» тек ескі мағынасымен берілген",
    "Әскери тарихи сөздер газетте әлі бар; «дулыға» спорт турнирінің атауына айналған",
    `Оқушылар көне сөздердің ${ok(GA["О"])}-ын ғана біледі, үлкендер әлдеқайда жақсы біледі`,
  ];
  concl.forEach((t, i) => {
    const y = 1.55 + i * 1.02;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "concl-num-bg-" + i });
    txt(s, String(i + 1), { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "concl-num-" + i });
    txt(s, t, { x: 1.5, y, w: 11.2, h: 0.8, fontSize: 18, color: C.text1, valign: "middle", objectName: "concl-text-" + i });
  });
  card(s, 0.6, 5.75, 5.85, 0.95, C.background2, "limits");
  txt(s, [{ text: "Шектеулер: ", options: { bold: true, color: C.text2 } }, { text: `баспасөзде 4 сөз, тестте ${N.n_people} адам`, options: { color: C.text1 } }],
    { x: 0.85, y: 5.75, w: 5.4, h: 0.95, fontSize: 15, valign: "middle", objectName: "limits-text" });
  card(s, 6.85, 5.75, 5.85, 0.95, C.background2, "next");
  txt(s, [{ text: "Келешекте: ", options: { bold: true, color: C.text2 } }, { text: "барлық сөзге мониторинг, көбірек қатысушы", options: { color: C.text1 } }],
    { x: 7.1, y: 5.75, w: 5.4, h: 0.95, fontSize: 15, valign: "middle", objectName: "next-text" });
  s.addNotes(`Қорытындылай келе: бірінші – сөздікте 36 көне сөздің ${N.marks["белгі жоқ"]}-ында белгі жоқ, қайта жанданған архаизмдер жалпы сөз ретінде берілген. Екінші – сөздік тілдің өзгерісінен артта қалады. Үшінші – әскери тарихи сөздер газетте әлі бар, ал «дулыға» спорт турнирінің атауына айналған. Төртінші – оқушылар көне сөздерді үлкендерден әлдеқайда нашар біледі. Шектеулері: баспасөзде тек 4 сөзді, тестте ${N.n_people} адамды зерттедім. Келешекте барлық сөзге мониторинг жүргізіп, көбірек адам қатыстырғым келеді. (45 секунд)`);

  // ================= 14. Рахмет =================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Қорытынды" });
  s.addText("Көне сөз өлмейді – ол жаңа орын табады", { placeholder: "title" });
  s.addText("Назарларыңызға рахмет! Сұрақтарыңызға жауап беруге дайынмын", { placeholder: "body" });
  [["Сөз", "36"], ["Сөздікте белгісі жоқ", String(N.marks["белгі жоқ"])], ["Оқушылардың білуі", ok(GA["О"])]].forEach(([l, vv], i) => {
    const y = 1.6 + i * 1.5;
    card(s, 8.9, y, 3.7, 1.25, C.background1, "end-stat-" + i);
    txt(s, [{ text: vv, options: { fontSize: 30, bold: true, color: HEX.dk2, fontFace: THEME.headFontFace, breakLine: true } }, { text: l, options: { fontSize: 14, color: HEX.accent5 } }],
      { x: 9.2, y, w: 3.2, h: 1.25, valign: "middle", objectName: "end-stat-text-" + i });
  });
  s.addNotes(`Менің басты түйінім: көне сөз өлмейді – ол жаңа орын табады: мемлекеттік атауда, спортта, мұражайда. Назарларыңызға рахмет!

ЫҚТИМАЛ СҰРАҚТАР:
1) Сөздерді қалай таңдадың? – Жетекшім екеуміз алты тақырыптық топтан 6 сөзден алдық: оқулықтағы мысалдарға және Ә. Қайдардың сөздігіне сүйендік.
2) Неге баспасөзде тек 4 сөзді тексердің? – Әр сөзге іздеу, мақаланы оқу және мәнмәтінін талдау уақыт алады. Мен әскери топты мысал ретінде алдым. Бұл – жұмыстың шектеуі, келесі жолы барлық сөзді тексеремін.
3) Google нәтиже саны дәл ме? – Жоқ, шамамен ғана. Сонымен қатар іздеу сөздің тек бір тұлғасын табады. Сондықтан мен санға ғана емес, мысал сөйлемнің мағынасына да қарадым.
4) Тестке неше адам қатысты, неге аз? – ${N.n_people} адам: ${listJoin(who)}. Нәтижені алдын ала деп атадым, оны барлық адамға жалпылауға болмайды.
5) Неге ата-әже ата-анадан төмен шықты? – Ата-әжеден бір ғана адам қатысты, сондықтан бұл кездейсоқ болуы мүмкін.
6) Этика қалай сақталды? – Есімдер жазылмады, әр адамға код берілді, ересектерден келісім, оқушылардан ата-анасының келісімі алынды.
7) ЖИ пайдаландың ба? – Иә, декларацияда ашық көрсеттім: жұмыстың құрылымын, кесте мен диаграмманы, мәтіннің жобасын жасауға. Сөздікті, газетті және тестті өзім жүргіздім.`);

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });
