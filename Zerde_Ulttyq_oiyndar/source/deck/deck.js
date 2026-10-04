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
const SV = N.survey;
const GAMES = D.games.map(g => g[0]);
const sumG = g => N.by_game[g]["Ертегі"] + N.by_game[g]["Жыр"];

// Ұлттық түстер: көк аспан мен алтын күн, ертегіге – қызғылт-қоңыр
const THEME = {
  name: "Zerde Ulttyq oiyndar",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1C2B33", lt1: "FFFFFF", dk2: "0D4A5E", lt2: "EAF4F6",
    accent1: "0092B0", accent2: "E3A400", accent3: "B5532F", accent4: "2E8B57",
    accent5: "5E6E75", accent6: "7A5195", hlink: "0092B0", folHlink: "7A5195",
  },
};
const HEX = THEME.colors;
const ERT = HEX.accent3, ZHYR = HEX.accent1; // ертегі – қоңыр, жыр – көк

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Ертегілер мен жырлардағы ұлттық ойындар";
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

(async () => {
  let s;
  // ================= 1. Титул =================
  pres.addSection({ title: "Кіріспе" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Кіріспе" });
  txt(s, "«Зерде» республикалық конкурсы  ·  Қазақ тілі және әдебиеті  ·  «Орта буын»", { x: 0.8, y: 0.6, w: 11.5, h: 0.4, fontSize: 14, color: C.accent1, objectName: "kicker" });
  s.addText("Ертегілер мен жырлардағы ұлттық ойындар", { placeholder: "title" });
  s.addText(`11 шығарма, ${N.ep_total} эпизод және ${N.n_resp} сыныптасымның жауабы`, { placeholder: "body" });
  txt(s, "Тіркеу коды (шифр): ______________", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.background2, objectName: "cipher" });
  ["Бәйге", "Асық", "Күрес", "Аударыспақ"].forEach((t, i) =>
    gameChip(s, 8.9 + (i % 2) * 1.95, 1.7 + Math.floor(i / 2) * 1.45, 1.75, 1.15, t, "title-chip-" + i, { size: i === 3 ? 15 : 20 }));
  await iconCircle(s, gi.GiHorseHead, 9.55, 4.75, 1.15, C.accent2, "title-horse");
  await iconCircle(s, gi.GiBookCover, 11.0, 4.75, 1.15, C.accent1, "title-book");
  s.addNotes("Құрметті қазылар алқасы! Менің жобамның тақырыбы – «Ертегілер мен жырлардағы ұлттық ойындар». Мен 11 ертегі мен жырды оқып, ұлттық ойын кездесетін 16 эпизод таптым, ал сыныптастарымның ойындарды қаншалықты білетінін сауалнама арқылы анықтадым. (15 секунд)");

  // ================= 2. Мәселе =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Төстік асық ойнамаса, ертегі басталмас еді", { placeholder: "title" });
  card(s, 0.6, 1.55, 7.2, 3.2, C.background2, "hook-quote-card");
  await iconCircle(s, fa.FaQuoteLeft, 0.95, 1.85, 0.75, C.accent3, "hook-quote");
  txt(s, "Төстік асық ойнап жүріп, кемпірдің баласын ұрып қалады. Кемпірдің ызалы сөзінен кейін ол ағаларын іздеп жолға шығады.",
    { x: 1.95, y: 1.8, w: 5.6, h: 2.0, fontSize: 20, italic: true, color: C.text1, valign: "top", objectName: "hook-quote-text" });
  txt(s, "«Ер Төстік» ертегісі", { x: 1.95, y: 3.95, w: 5.6, h: 0.45, fontSize: 15, bold: true, color: C.accent3, objectName: "hook-quote-src" });
  bigStat(s, 8.3, 1.6, 4.4, "12", "ұлттық ойынды зерттедім – ат ойындарынан бастап асық пен алтыбақанға дейін", C.accent2, "hook-12", 54);
  bigStat(s, 8.3, 3.25, 4.4, "63%", "сыныптастарымның ғана асық ойнағаны бар", C.accent1, "hook-63", 54);
  card(s, 0.6, 5.15, 12.1, 1.4, C.text2, "hook-question");
  txt(s, "Ертегі мен жырда ойындар не үшін керек – жай көрініс пе, әлде оқиғаның бір бөлігі ме?",
    { x: 1.0, y: 5.15, w: 11.3, h: 1.4, fontSize: 22, bold: true, color: C.background1, valign: "middle", objectName: "hook-question-text" });
  s.addNotes("Бәрі «Ер Төстік» ертегісінен басталды. Төстік асық ойнап жүріп кемпірдің баласын ұрып қалады да, сол оқиғадан кейін ағаларын іздеп жолға шығады. Яғни асық ойнамаса, ертегі басталмас еді. Мен ертегі мен жырларда ойындар жай ғана көрініс пе, әлде оқиғаның маңызды бөлігі ме – соны білгім келді. Сонымен қатар бүгінгі оқушылар бұл ойындарды біле ме – соны да тексердім. (30 секунд)");

  // ================= 3. Сұрақ, мақсат, болжам =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеу сұрағы мен болжам", { placeholder: "title" });
  card(s, 0.6, 1.55, 12.1, 1.45, C.background2, "q-card");
  await iconCircle(s, fa.FaQuestion, 0.95, 1.85, 0.85, C.accent1, "q");
  txt(s, "Ертегі мен жырда қандай ұлттық ойындар кездеседі, олар оқиғада қандай рөл атқарады және сыныптастарым оларды біле ме?",
    { x: 2.1, y: 1.55, w: 10.3, h: 1.45, fontSize: 21, bold: true, color: C.text2, valign: "middle", objectName: "q-text" });
  const qa = [
    [fa.FaBullseye, "Мақсаты", "Ертегі мен жырлардағы ұлттық ойындарды анықтап, олардың шығармадағы қызметін талдау және сыныптастарымның оларды білу деңгейін зерттеу", C.accent1],
    [fa.FaLightbulb, "Болжам", "Ертегіде балалар ойындары, жырда ат ойындары мен күш сынасу ойындары жиі кездеседі; жырда ойын көбіне қаһарманды сынайды", C.accent2],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25;
    card(s, x, 3.35, 5.85, 3.3, C.background1, "aim-card-" + i, "C9DDE2");
    await iconCircle(s, qa[i][0], x + 0.35, 3.65, 0.75, qa[i][3], "aim-" + i);
    txt(s, qa[i][1], { x: x + 1.3, y: 3.65, w: 4.3, h: 0.75, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: "aim-title-" + i });
    txt(s, qa[i][2], { x: x + 0.35, y: 4.6, w: 5.2, h: 1.9, fontSize: 17, color: C.text1, valign: "top", objectName: "aim-text-" + i });
  }
  s.addNotes("Менің зерттеу сұрағым: ертегі мен жырда қандай ұлттық ойындар кездеседі, олар оқиғада қандай рөл атқарады және сыныптастарым оларды біле ме? Болжамым мынадай болды: ертегіде балалар ойындары, ал жырда ат ойындары мен күш сынасу ойындары жиі кездеседі, ал жырда ойын көбіне батырды сынайды. (30 секунд)");

  // ================= 4. Әдістеме =================
  pres.addSection({ title: "Әдістеме" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Әдістеме" });
  s.addText("Қалай зерттедім: үш қадам", { placeholder: "title" });
  const steps = [
    [fa.FaBookOpen, "Оқыдым", "11 шығарма: 4 ертегі және 7 батырлар мен ғашықтық жыры. Ойын кездесетін әр эпизодты кестеге жаздым", C.accent3],
    [fa.FaTags, "Кодтадым", "12 ойынды 5 түрге бөлдім. Ойынның оқиғадағы қызметін 5 кодпен белгіледім: С, Б, Қ, Т, Ш", C.accent2],
    [fa.FaPoll, "Сұрадым", `Жетекшім екеуміз Google Forms-та сауалнама жасадық. Ата-ана келісімімен ${N.n_resp} сыныптасым қатысты`, C.accent1],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.6, 3.75, 4.0, C.background2, "step-card-" + i);
    await iconCircle(s, steps[i][0], x + 0.35, 1.9, 0.95, steps[i][3], "step-" + i);
    txt(s, `${i + 1}`, { x: x + 2.6, y: 1.85, w: 0.9, h: 1.0, fontSize: 44, bold: true, color: "B9D3D9", fontFace: THEME.headFontFace, align: "right", objectName: "step-num-" + i });
    txt(s, steps[i][1], { x: x + 0.35, y: 3.05, w: 3.1, h: 0.6, fontSize: 22, bold: true, color: C.text2, objectName: "step-title-" + i });
    txt(s, steps[i][2], { x: x + 0.35, y: 3.7, w: 3.1, h: 1.8, fontSize: 15, color: C.text1, valign: "top", objectName: "step-text-" + i });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 3.4, w: 0.3, h: 0.4, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "step-arrow-" + i });
  }
  const codes = [["С", "Сынақ"], ["Б", "Сюжетті бастау"], ["Қ", "Қасиетін көрсету"], ["Т", "Той-думан"], ["Ш", "Шешуші сәт"]];
  codes.forEach(([c, n], i) => {
    const x = 0.6 + i * 2.45;
    s.addShape(pres.shapes.OVAL, { x, y: 5.95, w: 0.55, h: 0.55, fill: { color: C.text2 }, line: { type: "none" }, objectName: "code-bg-" + i });
    txt(s, c, { x, y: 5.95, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "code-" + i });
    txt(s, n, { x: x + 0.65, y: 5.95, w: 1.75, h: 0.55, fontSize: 14, color: C.text1, valign: "middle", objectName: "code-name-" + i });
  });
  s.addNotes(`Зерттеу үш қадамнан тұрды. Бірінші – оқыдым: 4 ертегі мен 7 жырды басынан аяғына дейін оқып, ойын кездескен әр эпизодты кестеге жаздым – шығарманы, ойынды, кейіпкерді және мәтіннен үзіндіні. Екінші – кодтадым: 12 ойынды бес түрге бөлдім, ал әр эпизодтағы ойынның қызметін бес кодпен белгіледім. Мысалы, С – сынақ, Б – сюжетті бастау. Күмәнді эпизодтарды жетекшім екеуміз бірге қарадық. Үшінші – сұрадым: жетекшім екеуміз Google Forms-та сауалнама жасадық, оған ата-анасы келісім берген ${N.n_resp} сыныптасым қатысты. (45 секунд)`);

  // ================= 5. Қандай ойындар кездеседі =================
  pres.addSection({ title: "Нәтижелер" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Ертегі мен жырда қандай ойындар бар?", { placeholder: "title" });
  const found = Object.keys(N.by_game).filter(g => sumG(g) > 0).sort((a, b) => sumG(a) - sumG(b));
  const labels = [...found, "Жаяу бәйге"];
  s.addChart(pres.charts.BAR, [
    { name: "Ертегі", labels, values: [...found.map(g => N.by_game[g]["Ертегі"]), 1] },
    { name: "Жыр", labels, values: [...found.map(g => N.by_game[g]["Жыр"]), 0] },
  ], { x: 0.6, y: 1.45, w: 7.4, h: 5.3, barDir: "bar", barGrouping: "stacked", chartColors: [ERT, ZHYR], ...chartQuiet(),
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: "0;;;", valAxisMaxVal: 4, valAxisMajorUnit: 1,
    showLegend: true, legendPos: "b", showTitle: true, title: "Эпизод саны", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-games" });
  bigStat(s, 8.5, 1.5, 4.2, `${N.ep_total}`, "эпизод 11 шығарманың бәрінен табылды", C.accent2, "games-ep");
  bigStat(s, 8.5, 3.15, 4.2, "6 / 12", "ойын бірде-бір рет кездеспеді", C.accent3, "games-zero");
  const zero = Object.keys(N.by_game).filter(g => sumG(g) === 0);
  txt(s, zero.join(", ").toLowerCase(), { x: 8.5, y: 4.75, w: 4.2, h: 0.8, fontSize: 14, italic: true, color: C.accent5, valign: "top", objectName: "games-zero-list" });
  card(s, 8.5, 5.65, 4.2, 1.05, C.background2, "games-new");
  txt(s, [{ text: "Жаңа ойын: ", options: { bold: true, color: C.text2 } }, { text: "«Күнікей қызда» жаяу бәйге бар – ол тізімде жоқ еді", options: { color: C.text1 } }],
    { x: 8.7, y: 5.65, w: 3.85, h: 1.05, fontSize: 14, valign: "middle", objectName: "games-new-text" });
  s.addNotes(`11 шығарманың бәрінде кемінде бір ойын кездесті, барлығы ${N.ep_total} эпизод. Ең жиі кездескен ойындар – бәйге мен асық, әрқайсысы 4 реттен. Одан кейін күрес – 3 рет. Ал мені таңғалдырғаны: 12 ойынның жартысы – тоғызқұмалақ, алтыбақан, қыз қуу, теңге алу, жұмбақ айтысу және ақсүйек – бірде-бір шығармада кездеспеді. Есесіне «Күн астындағы Күнікей қыз» ертегісінен тізімде жоқ ойынды таптым – жаяу бәйгені. (40 секунд)`);

  // ================= 6. Ертегі мен жыр: ойын түрлері =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Жырда – ат, ертегіде – күш", { placeholder: "title" });
  const side = [
    ["Жырларда", `${N.by_type["АТ"]["Жыр"]} / ${N.ep_Жыр}`, "эпизод – ат ойындары: бәйге, аударыспақ, көкпар", "Батырдың ең жақын серігі – тұлпары: Тайбурыл, Байшұбар, Тарлан", gi.GiHorseHead, ZHYR],
    ["Ертегілерде", `${N.by_type["КҮШ"]["Ертегі"]} / ${N.ep_Ертегі}`, "эпизод – күш сынасу: күрес", "Балуандар мен дәулер күш сынасады, жеңіс көбіне ғажайып көмекшінің арқасында келеді", fa.FaFistRaised, ERT],
  ];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 6.25, [h, v, l, d, Ic, col] = side[i];
    card(s, x, 1.55, 5.85, 4.1, C.background2, "type-card-" + i);
    await iconCircle(s, Ic, x + 0.35, 1.85, 0.9, col, "type-" + i);
    txt(s, h, { x: x + 1.45, y: 1.85, w: 4.1, h: 0.9, fontSize: 24, bold: true, color: C.text2, valign: "middle", objectName: "type-title-" + i });
    txt(s, v, { x: x + 0.35, y: 3.0, w: 2.4, h: 1.0, fontSize: 48, bold: true, color: col, fontFace: THEME.headFontFace, objectName: "type-value-" + i });
    txt(s, l, { x: x + 2.75, y: 3.0, w: 2.8, h: 1.0, fontSize: 16, color: C.text1, valign: "middle", objectName: "type-label-" + i });
    txt(s, d, { x: x + 0.35, y: 4.2, w: 5.2, h: 1.2, fontSize: 15, italic: true, color: C.accent5, valign: "top", objectName: "type-desc-" + i });
  }
  card(s, 0.6, 5.95, 12.1, 0.8, C.background1, "type-asyk", C.accent2);
  txt(s, [{ text: "Асық екеуінде де бар: ", options: { bold: true, color: C.text2 } }, { text: "ертегіде 2, жырда 2 эпизод. Зияткерлік ойындар мен ойын-сауық мүлде кездеспеді", options: { color: C.text1 } }],
    { x: 0.9, y: 5.95, w: 11.6, h: 0.8, fontSize: 16, valign: "middle", objectName: "type-asyk-text" });
  s.addNotes(`Ертегі мен жырды салыстырғанда айқын айырмашылық көрдім. Жырларда ${N.ep_Жыр} эпизодтың ${N.by_type["АТ"]["Жыр"]}-уы – ат ойындары: бәйге, аударыспақ, көкпар. Бұл заңды, өйткені батырдың ең жақын серігі – оның тұлпары. Ал ертегілерде ең жиі кездескені – күрес: балуандар мен дәулер күш сынасады. Асық ертегіде де, жырда да екі реттен кездесті. (35 секунд)`);

  // ================= 7. Ойынның қызметі =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Ертегіде – сынақ, жырда – батырдың қасиеті", { placeholder: "title" });
  const FC = ["С", "Б", "Қ"], FL = ["Сынақ", "Сюжетті бастау", "Қасиетін көрсету"];
  s.addChart(pres.charts.BAR, [
    { name: "Ертегі", labels: FL, values: FC.map(f => N.by_func[f]["Ертегі"]) },
    { name: "Жыр", labels: FL, values: FC.map(f => N.by_func[f]["Жыр"]) },
  ], { x: 0.6, y: 1.45, w: 7.0, h: 5.3, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60, chartColors: [ERT, ZHYR], ...chartQuiet(),
    showValue: true, dataLabelPosition: "outEnd", valAxisMaxVal: 6, valAxisMajorUnit: 1, showLegend: true, legendPos: "b",
    showTitle: true, title: "Эпизод саны (Т және Ш қызметі кездеспеді)", titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-func" });
  const fx = [
    [`${N.by_func["С"]["Ертегі"]} / ${N.ep_Ертегі}`, "ертегі эпизодында ойын – сынақ", "Хан қызын жаяу бәйге мен күресте жеңгенге береді («Күнікей қыз»)", ERT],
    [`${N.by_func["Қ"]["Жыр"]} / ${N.ep_Жыр}`, "жыр эпизодында ойын батырдың не атының қасиетін көрсетеді", "Тарғын жауға шабуды «Тарлан бәйгесі» дейді («Ер Тарғын»)", ZHYR],
  ];
  fx.forEach(([v, l, ex, col], i) => {
    const y = 1.55 + i * 2.6;
    card(s, 8.0, y, 4.7, 2.35, C.background2, "func-card-" + i);
    txt(s, v, { x: 8.3, y: y + 0.2, w: 1.9, h: 0.85, fontSize: 40, bold: true, color: col, fontFace: THEME.headFontFace, objectName: "func-value-" + i });
    txt(s, l, { x: 10.2, y: y + 0.2, w: 2.35, h: 0.85, fontSize: 14, color: C.text1, valign: "middle", objectName: "func-label-" + i });
    txt(s, ex, { x: 8.3, y: y + 1.15, w: 4.2, h: 1.05, fontSize: 14, italic: true, color: C.accent5, valign: "top", objectName: "func-ex-" + i });
  });
  s.addNotes(`Ойынның оқиғадағы рөлі де әртүрлі болып шықты. Ертегілерде ойын көбіне сынақ: ${N.ep_Ертегі} эпизодтың ${N.by_func["С"]["Ертегі"]}-еуі. Мысалы, «Күн астындағы Күнікей қыз» ертегісінде хан қызын жаяу бәйгеде, кейін күресте жеңгенге береді. Бұл В.Я. Пропп жазған «қаһарманды сынау» функциясына дәл келеді. Ал жырларда ойын көбіне батырдың немесе оның атының қасиетін көрсетеді: ${N.ep_Жыр} эпизодтың ${N.by_func["Қ"]["Жыр"]}-еуі. Той-думан мен шешуші сәт қызметі бірде-бір рет кездеспеді. (40 секунд)`);

  // ================= 8. Асық – оқиғаның басы =================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Нәтижелер" });
  txt(s, "Ең қызық жаңалығым", { x: 0.8, y: 0.6, w: 8, h: 0.45, fontSize: 16, color: C.accent1, objectName: "asyk-kicker" });
  txt(s, "Асық – оқиғаның басы", { x: 0.8, y: 1.05, w: 11.7, h: 0.9, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace, objectName: "asyk-title" });
  txt(s, [{ text: "4 / 4", options: { fontSize: 60, bold: true, color: HEX.accent2, fontFace: THEME.headFontFace, breakLine: true } },
    { text: "асық эпизодының бәрін «сюжетті бастау» деп белгіледім", options: { fontSize: 16, color: HEX.lt2 } }],
    { x: 0.8, y: 2.3, w: 3.4, h: 2.6, valign: "top", objectName: "asyk-stat" });
  const asyk = [
    ["«Ер Төстік»", "Асық ойнап жүріп кемпірдің баласын ұрып қалады – ағаларын іздеп жолға шығады"],
    ["«Алтын сақа»", "Жұртта қалған сақасын алуға қайтып барып, жалмауыз кемпірге кездеседі"],
    ["«Қозы Көрпеш – Баян сұлу»", "Тазшаның сақасын тартып алғанда «жарыңды іздеп тапсайшы» деген сөз естиді"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.3 + i * 1.45;
    card(s, 4.6, y, 8.1, 1.25, C.background1, "asyk-card-" + i);
    await iconCircle(s, [fa.FaRoute, fa.FaRoad, fa.FaHeart][i], 4.85, y + 0.25, 0.75, C.accent2, "asyk-ic-" + i);
    txt(s, [{ text: asyk[i][0], options: { bold: true, color: HEX.dk2, breakLine: true } }, { text: asyk[i][1], options: { color: HEX.dk1 } }],
      { x: 5.85, y, w: 6.65, h: 1.25, fontSize: 15, valign: "middle", objectName: "asyk-text-" + i });
  }
  txt(s, "Асық ойыны баланы үйден алып шығып, үлкен оқиғаның бастамасына айналады", { x: 0.8, y: 6.55, w: 11.9, h: 0.5, fontSize: 16, italic: true, color: C.accent2, objectName: "asyk-note" });
  s.addNotes("Ал ең қызық жаңалығым асыққа қатысты болды. Асық кездескен төрт эпизодтың төртеуін де мен «сюжетті бастау» деп белгіледім. «Ер Төстікте» Төстік асық ойнап жүріп кемпірдің баласын ұрып қалады да, жолға шығады. «Алтын сақада» бала жұртта қалған сақасын алуға қайтып барып, жалмауыз кемпірге кездеседі. «Қозы Көрпеште» Қозы тазша баланың сақасын тартып алғанда, «жарыңды іздеп тапсайшы» деген сөз естиді. Яғни асық баланы үйден алып шығып, үлкен оқиғаның басына айналады. (40 секунд)");

  // ================= 9. Сауалнама =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Сыныптастарым біледі, бірақ ойнамаған", { placeholder: "title" });
  const sg = [...GAMES].sort((a, b) => SV[a].played - SV[b].played);
  s.addChart(pres.charts.BAR, [
    { name: "Біледі", labels: sg, values: sg.map(g => SV[g].know) },
    { name: "Ойнаған не көрген", labels: sg, values: sg.map(g => SV[g].played) },
  ], { x: 0.6, y: 1.4, w: 7.6, h: 5.4, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 45, chartColors: ["B9D3D9", HEX.accent2], ...chartQuiet(),
    catAxisLabelFontSize: 12, valAxisMaxVal: 100, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"", showLegend: true, legendPos: "b",
    showTitle: true, title: `Оқушылар үлесі, % (n = ${N.n_resp})`, titleFontSize: 14, titleColor: HEX.dk1, objectName: "chart-survey" });
  bigStat(s, 8.7, 1.5, 4.0, "10 / 12", "ойынның атын бір оқушы орта есеппен біледі", C.text2, "sv-know");
  bigStat(s, 8.7, 3.15, 4.0, f1(SV["Асық"].played) + "%", "асық ойнаған – ең көп ойналатын әрі ең ұнайтын ойын", C.accent2, "sv-asyk");
  card(s, 8.7, 4.9, 4.0, 1.85, C.background2, "sv-kokpar");
  txt(s, [{ text: "Көкпар: ", options: { bold: true, color: C.text2 } },
    { text: `барлығы (${f1(SV["Көкпар"].know)}%) біледі, бірақ тек 2 оқушы (${f1(SV["Көкпар"].played)}%) көрген`, options: { color: C.text1 } }],
    { x: 8.95, y: 4.9, w: 3.55, h: 1.85, fontSize: 16, valign: "middle", objectName: "sv-kokpar-text" });
  s.addNotes(`Сауалнамаға 22 оқушы жауап берді, бірақ үшеуінің ата-анасы келісім бермегендіктен, олардың жауабын есепке алмадым. Сонымен 5–8-сыныптардың ${N.n_resp} оқушысы қатысты. Сыныптастарым ойындардың атын жақсы біледі: бір оқушы орта есеппен 12 ойынның 10-ын біледі. Бірақ білу мен ойнау – екі түрлі нәрсе. Ең көп ойналатыны – асық, ${f1(SV["Асық"].played)}%. Ал көкпарды барлығы білсе де, оны тек екі оқушы көрген. Ең ұнайтын ойын ретінде де асықты атады. Бір сыныптасым: «Асық ату ұнайды, ол әділдікке, сабырлыққа үйретеді» деп жазды. (45 секунд)`);

  // ================= 10. Кітап пен өмір =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Сауалнама мен мәтінді салыстырдым", { placeholder: "title" });
  const cmp = [
    ["Асық", "4", "эпизод", f1(SV["Асық"].played) + "%", "ойнаған", "Кітапта да, өмірде де бірінші орында", C.accent4],
    ["Бәйге", "4", "эпизод", f1(SV["Бәйге"].played) + "%", "көрген", "Кітапта жиі, бірақ өмірде сирек", C.accent2],
    ["Тоғызқұмалақ", "0", "эпизод", f1(SV["Тоғызқұмалақ"].folk) + "%", "«ертегіден оқыдым» деді", "Мәтінде жоқ, бірақ «оқыдым» дегендер көп", C.accent3],
  ];
  txt(s, "Шығармаларда", { x: 3.55, y: 1.45, w: 2.2, h: 0.4, fontSize: 14, bold: true, color: C.accent5, align: "center", objectName: "cmp-h1" });
  txt(s, "Сауалнамада", { x: 6.05, y: 1.45, w: 2.6, h: 0.4, fontSize: 14, bold: true, color: C.accent5, align: "center", objectName: "cmp-h2" });
  for (let i = 0; i < 3; i++) {
    const y = 1.95 + i * 1.3, [g, a, al, b, bl, note, col] = cmp[i];
    card(s, 0.6, y, 12.1, 1.1, C.background2, "cmp-row-" + i);
    gameChip(s, 0.8, y + 0.17, 2.5, 0.76, g, "cmp-chip-" + i, { size: 17 });
    txt(s, [{ text: a + " ", options: { fontSize: 30, bold: true, color: HEX.dk2, fontFace: THEME.headFontFace } }, { text: al, options: { fontSize: 14, color: HEX.dk1 } }],
      { x: 3.55, y, w: 2.2, h: 1.1, align: "center", valign: "middle", objectName: "cmp-a-" + i });
    txt(s, [{ text: b + " ", options: { fontSize: 30, bold: true, color: col, fontFace: THEME.headFontFace, breakLine: true } }, { text: bl, options: { fontSize: 12, color: HEX.dk1 } }],
      { x: 6.05, y, w: 2.6, h: 1.1, align: "center", valign: "middle", objectName: "cmp-b-" + i });
    txt(s, note, { x: 8.95, y, w: 3.55, h: 1.1, fontSize: 15, italic: true, color: C.text1, valign: "middle", objectName: "cmp-note-" + i });
  }
  card(s, 0.6, 5.95, 12.1, 0.8, C.background1, "cmp-honest", C.accent3);
  txt(s, [{ text: "Сыни ой: ", options: { bold: true, color: C.accent3 } }, { text: "4 оқушы 12 ойынның бәрін «оқыдым» деді – сондықтан бұл сұрақтың жауабына толық сенуге болмайды", options: { color: C.text1 } }],
    { x: 0.9, y: 5.95, w: 11.6, h: 0.8, fontSize: 15, valign: "middle", objectName: "cmp-honest-text" });
  s.addNotes(`Сауалнама мен мәтін талдауын салыстырдым. Асық екеуінде де бірінші орында: шығармаларда ең жиі кездеседі, сыныптастарым оны ең көп ойнайды. Бәйге шығармаларда да 4 рет кездесті, бірақ оны өз көзімен көрген оқушы аз. Ал мені ең таңғалдырғаны – тоғызқұмалақ: мен оқыған 11 шығарманың бірде-бірінде жоқ, ал сыныптастарымның ${f1(SV["Тоғызқұмалақ"].folk)}%-ы ол туралы ертегіден оқыдым деді. Олар мен оқымаған басқа шығармаларды оқыған болуы мүмкін, немесе ойынды жақсы білгендіктен, ертегіден де оқыдым деп ойлаған. Төрт оқушы 12 ойынның бәрін «оқыдым» деп белгіледі, сондықтан бұл сұрақтың жауабына толық сенуге болмайды. (45 секунд)`);

  // ================= 11. Болжам =================
  pres.addSection({ title: "Қорытынды" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Болжамым ішінара расталды", { placeholder: "title" });
  const hyp = [
    ["Ертегіде балалар ойындары жиі кездеседі", "Расталмады: ертегіде күрес жиірек (3 эпизод), асық – 2", false],
    ["Жырда ат ойындары мен күш сынасу ойындары жиі", `Расталды: жырдың ${N.ep_Жыр} эпизодының ${N.by_type["АТ"]["Жыр"] + N.by_type["КҮШ"]["Жыр"]}-еуі`, true],
    ["Жырда ойын көбіне қаһарманды сынайды", "Расталмады: сынақ ертегіде жиірек, жырда ойын батырдың қасиетін көрсетеді", false],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.55 + i * 1.6, [h, r, ok] = hyp[i];
    card(s, 0.6, y, 12.1, 1.35, ok ? "E3F2EA" : "F8E9E3", "hyp-row-" + i);
    await iconCircle(s, ok ? fa.FaCheck : fa.FaTimes, 0.9, y + 0.3, 0.75, ok ? C.accent4 : C.accent3, "hyp-mark-" + i);
    txt(s, h, { x: 1.95, y, w: 4.6, h: 1.35, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "hyp-h-" + i });
    txt(s, r, { x: 6.8, y, w: 5.7, h: 1.35, fontSize: 16, color: C.text1, valign: "middle", objectName: "hyp-r-" + i });
  }
  txt(s, "Болжамның расталмаған бөлігі одан да қызық нәтиже берді: ойынның рөлі жанрға қарай өзгереді",
    { x: 0.6, y: 6.35, w: 12.1, h: 0.45, fontSize: 16, italic: true, color: C.accent5, objectName: "hyp-note" });
  s.addNotes("Болжамымды үш бөлік бойынша тексердім. Бірінші бөлігі расталмады: ертегіде балалар ойындарынан гөрі күрес жиі кездесті. Екінші бөлігі расталды: жырларда ат ойындары басым. Үшінші бөлігі де расталмады: сынақ жырда емес, ертегіде жиірек болды, ал жырда ойын батырдың қасиетін көрсетеді. Сондықтан болжамым ішінара расталды. Бірақ расталмаған бөлігі одан да қызық нәтиже берді. (35 секунд)");

  // ================= 12. Өнім =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Өнім: «Ертегі-ойын» карточкалары", { placeholder: "title" });
  s.addImage({ path: path.join(__dirname, "..", "figs", "kartochka.png"), x: 0.6, y: 1.5, w: 7.0, h: 7.0 * 810 / 1280, objectName: "card-image" });
  const use = [
    [fa.FaChalkboardTeacher, "Әдебиет сабағында", "Шығарманы талдағанда эпизодты оқып, ойынды талқылау"],
    [fa.FaRunning, "Үзілісте және лагерьде", "Ертегідегі ойынды өзің ойнап көру"],
    [fa.FaLayerGroup, "Әр карточкада", "Ойынның атауы мен түрі, ережесі және ертегі эпизоды"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.55 + i * 1.75;
    await iconCircle(s, use[i][0], 8.1, y + 0.1, 0.8, [C.accent1, C.accent2, C.accent3][i], "use-" + i);
    txt(s, use[i][1], { x: 9.1, y, w: 3.6, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "use-title-" + i });
    txt(s, use[i][2], { x: 9.1, y: y + 0.55, w: 3.6, h: 0.9, fontSize: 15, color: C.text1, valign: "top", objectName: "use-text-" + i });
  }
  s.addNotes("Зерттеу нәтижесіне сүйеніп, «Ертегі-ойын» карточкаларын жасадым. Әр карточкада ойынның атауы мен түрі, қысқаша ережесі және сол ойын кездесетін ертегі немесе жыр эпизоды бар. Мысалы, мына карточка – асық. Карточкаларды әдебиет сабағында, сынып сағатында, үзілісте немесе жазғы лагерьде пайдалануға болады: алдымен ертегіден эпизодты оқисың, сосын ойынды өзің ойнап көресің. Осылай ертегі де, ойын да есте жақсы сақталады. (35 секунд)");

  // ================= 13. Қорытынды =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Қорытынды", { placeholder: "title" });
  const concl = [
    `11 шығарманың бәрінде ұлттық ойын бар: ${N.ep_total} эпизод, ең жиісі – бәйге мен асық`,
    "Жырда ат ойындары басым, ертегіде – күрес; 12 ойынның 6-уы мүлде кездеспеді",
    "Ертегіде ойын – сынақ, жырда – батырдың қасиеті, ал асық – оқиғаның бастауы",
    "Сыныптастарым ойындардың атын біледі, бірақ көбін ойнамаған; ең сүйікті ойын – асық",
  ];
  concl.forEach((t, i) => {
    const y = 1.55 + i * 1.02;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "concl-num-bg-" + i });
    txt(s, String(i + 1), { x: 0.6, y: y + 0.08, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "concl-num-" + i });
    txt(s, t, { x: 1.5, y, w: 11.2, h: 0.8, fontSize: 19, color: C.text1, valign: "middle", objectName: "concl-text-" + i });
  });
  card(s, 0.6, 5.75, 5.85, 0.95, C.background2, "limits");
  txt(s, [{ text: "Шектеулер: ", options: { bold: true, color: C.text2 } }, { text: `11 шығарма (4 ертегі), бір мектептің ${N.n_resp} оқушысы`, options: { color: C.text1 } }],
    { x: 0.85, y: 5.75, w: 5.4, h: 0.95, fontSize: 15, valign: "middle", objectName: "limits-text" });
  card(s, 6.85, 5.75, 5.85, 0.95, C.background2, "next");
  txt(s, [{ text: "Келешекте: ", options: { bold: true, color: C.text2 } }, { text: "көбірек ертегі, ата-әжелермен сұхбат", options: { color: C.text1 } }],
    { x: 7.1, y: 5.75, w: 5.4, h: 0.95, fontSize: 15, valign: "middle", objectName: "next-text" });
  s.addNotes(`Қорытындылай келе: бірінші – 11 шығарманың бәрінде ұлттық ойын бар, ең жиісі – бәйге мен асық. Екінші – жырда ат ойындары, ертегіде күрес басым, ал 12 ойынның жартысы мүлде кездеспеді. Үшінші – ертегіде ойын сынақ болады, жырда батырдың қасиетін көрсетеді, ал асық оқиғаны бастайды. Төртінші – сыныптастарым ойындардың атын біледі, бірақ көбін ойнамаған. Жұмыстың шектеулері бар: мен 11 шығарманы ғана оқыдым, сауалнамаға бір мектептің ${N.n_resp} оқушысы қатысты. Келешекте көбірек ертегі оқып, ата-әжелерімнен олардың бала кезгі ойындары туралы сұхбат алғым келеді. (45 секунд)`);

  // ================= 14. Рахмет =================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Қорытынды" });
  s.addText("Ойын – ертегінің жүрегі", { placeholder: "title" });
  txt(s, "Ол батырды сынайды, күшін көрсетеді және жолға шығарады", { x: 0.8, y: 3.0, w: 7.3, h: 1.0, fontSize: 22, color: C.background2, valign: "top", objectName: "end-sub" });
  s.addText("Назарларыңызға рахмет! Сұрақтарыңызға жауап беруге дайынмын", { placeholder: "body" });
  [["Шығарма", "11"], ["Ойын эпизоды", String(N.ep_total)], ["Сауалнама", N.n_resp + " оқушы"]].forEach(([l, vv], i) => {
    const y = 1.6 + i * 1.5;
    card(s, 8.9, y, 3.7, 1.25, C.background1, "end-stat-" + i);
    txt(s, [{ text: vv, options: { fontSize: 30, bold: true, color: HEX.dk2, fontFace: THEME.headFontFace, breakLine: true } }, { text: l, options: { fontSize: 14, color: HEX.accent5 } }],
      { x: 9.2, y, w: 3.2, h: 1.25, valign: "middle", objectName: "end-stat-text-" + i });
  });
  s.addNotes(`Менің басты түйінім: ойын – ертегінің жүрегі. Ол батырды сынайды, күшін көрсетеді және оны жолға шығарады. Назарларыңызға рахмет!

ЫҚТИМАЛ СҰРАҚТАР:
1) Неге осы 11 шығарма? – Мектеп бағдарламасына енген, белгілі және маған қолжетімді шығармаларды алдым. Көбін Ұлттық электрондық кітапханадан (kazneb.kz) оқыдым, басылымдары дереккөздер тізімінде бар.
2) Ойынның қызметін қалай анықтадың? – Алдын ала жазылған бес код бойынша (С, Б, Қ, Т, Ш). Бір эпизод бірнеше қызмет атқаруы мүмкін, мен ең негізгісін таңдадым. Күмәнді жағдайларды жетекшім екеуміз талқылап шештік.
3) Тоғызқұмалақ неге кездеспеді? – Мен оқыған шығармаларда жоқ. Мүмкін ол басқа ертегілерде немесе басқа нұсқаларда бар. Бұл – келесі зерттеуге бағыт.
4) Сауалнамадағы этика қалай сақталды? – Есім мен пошта сұралмады, алғашқы сұрақ ата-ананың келісімі туралы болды. Келісімі жоқ 3 жауапты алып тастадым.
5) Жаяу бәйгені неге қостың? – Ол тізімде жоқ еді, бірақ «Күнікей қызда» нақты кездесті, сондықтан «Басқа» ойын ретінде жаздым.
6) Сауалнамаға қаншалықты сенуге болады? – Білу мен ойнау туралы жауаптар нақты. Ал «ертегіден оқыдың ба?» деген сұраққа оқушылар естеріне сүйеніп жауап берді, сондықтан келесі жолы «қай шығармадан?» деп нақты сұрар едім.
7) ЖИ пайдаландың ба? – Иә, декларацияда ашық көрсеттім: жұмыстың құрылымын жоспарлауға, менің деректерім бойынша мәтінді жазу мен редакциялауға, кесте мен диаграмма жасауға. Шығармаларды өзім оқыдым, эпизодтарды өзім тауып кодтадым, сауалнаманы өзім жүргіздім.`);

  // ================= Қосымша: 16 эпизод =================
  pres.addSection({ title: "Сұрақтарға арналған қосымша" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Сұрақтарға арналған қосымша" });
  s.addText(`Қосымша: табылған ${N.ep_total} эпизод`, { placeholder: "title" });
  const hdr = o => ({ text: o, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } });
  const ep = N.episodes.map(e => [e.work, e.game, { text: e.func, options: { align: "center", bold: true, color: e.genre === "Ертегі" ? ERT : ZHYR } }]);
  const half = Math.ceil(ep.length / 2);
  [ep.slice(0, half), ep.slice(half)].forEach((rows, i) => {
    s.addTable([[hdr("Шығарма"), hdr("Ойын"), hdr("Код")], ...rows], { x: 0.6 + i * 6.2, y: 1.45, w: 5.9, colW: [3.3, 1.85, 0.75], fontSize: 14, fontFace: "+mn-lt",
      color: HEX.dk1, border: { type: "solid", pt: 0.75, color: "C9DDE2" }, valign: "middle", rowH: 0.55, objectName: "table-ep-" + i });
  });
  legendDot(s, 0.6, 6.55, ERT, "Ертегі", "leg-e");
  legendDot(s, 2.4, 6.55, ZHYR, "Жыр", "leg-z");
  txt(s, "С – сынақ, Б – сюжетті бастау, Қ – қасиетін көрсету", { x: 4.2, y: 6.55, w: 8.5, h: 0.38, fontSize: 14, color: C.accent5, valign: "middle", objectName: "ep-codes" });
  s.addNotes("Сұрақ қойылса: шығармалардан табылған барлық эпизод және әрқайсысының қызмет коды. Толық үзінділер жұмыстың А қосымшасында.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });
