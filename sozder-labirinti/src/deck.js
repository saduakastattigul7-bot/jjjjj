// Қорғауға арналған слайдтар (5–7 минут), 16:9
const path = require("path");
const pptxgen = require("pptxgenjs");
const OUT = process.argv[2] || path.join(__dirname, "../Sozder_labirinti_prezentatsiya.pptx");
const FIG = f => path.join(__dirname, "figs", f);

const C = { ink: "1D2846", wall: "2C3A6B", flag: "00809E", flagSoft: "D3EEF4", sun: "F2B233", sp: "B8286F", spSoft: "FBE3EF", sky: "E6F1F4", white: "FFFFFF", soft: "56627E" };
const HF = "Cambria", BF = "Arial";
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 × 7.5
pres.title = "Сөздер лабиринті";

const txt = (s, t, o) => s.addText(t, { isTextBox: true, fontFace: BF, color: C.ink, margin: 0, valign: "top", ...o });
const title = (s, t, o = {}) => txt(s, t, { x: 0.7, y: 0.45, w: 11.9, h: 0.9, fontFace: HF, fontSize: 36, bold: true, color: o.color || C.wall, valign: "middle" });
const card = (s, x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12 });
const shot = (s, f, x, y, h) => s.addImage({ path: FIG(f), x, y, h, w: h * 400 / 760, rounding: false,
  shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 6, offset: 2, angle: 90 } });
const light = () => { const s = pres.addSlide(); s.background = { color: C.white }; return s; };

/* 1. Титул */
{
  const s = pres.addSlide(); s.background = { color: C.wall };
  txt(s, "«Зерде» · Қазақ тілі және әдебиеті · Бастауыш буын", { x: 0.7, y: 0.6, w: 7.5, h: 0.4, fontSize: 16, color: "B9C6EE" });
  txt(s, "СӨЗДЕР\nЛАБИРИНТІ", { x: 0.7, y: 1.3, w: 7.6, h: 2.0, fontFace: HF, fontSize: 54, bold: true, color: C.white });
  txt(s, "Қазақ тілінен интерактивті ойын құрастыру", { x: 0.7, y: 3.5, w: 7.4, h: 0.6, fontSize: 24, color: C.sun });
  txt(s, "Цифрлық жоба · 3-сынып (оқыту орыс тілінде)", { x: 0.7, y: 4.25, w: 7.4, h: 0.4, fontSize: 18, color: "DCE3F7" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 5.6, w: 3.9, h: 1.0, fill: { color: C.wall }, line: { color: "B9C6EE", dashType: "dash", width: 1.25 }, rectRadius: 0.08 });
  txt(s, "Тіркеу коды (шифр)", { x: 0.7, y: 5.6, w: 3.9, h: 1.0, fontSize: 14, color: "B9C6EE", align: "center", valign: "middle" });
  shot(s, "s_maze.png", 9.0, 0.55, 6.4);
  s.addNotes("Сәлеметсіздер ме, құрметті қазылар алқасы! Менің жобамның тақырыбы – «Сөздер лабиринті: қазақ тілінен интерактивті ойын құрастыру».");
}

/* 2. Мәселе */
{
  const s = light(); title(s, "Неге бұл тақырып?");
  txt(s, "Қазақ тілінде орыс тілінде жоқ 9 әріп бар. Біз оларды жиі шатастырамыз.", { x: 0.7, y: 1.4, w: 11.9, h: 0.5, fontSize: 20, color: C.soft });
  const L = [["ә", "а"], ["ғ", "г"], ["қ", "к"], ["ң", "н"], ["ө", "о"], ["ұ", "у"], ["ү", "у"], ["һ", "х"], ["і", "и"]];
  L.forEach(([a, b], i) => {
    const x = 0.7 + i * 1.33;
    card(s, x, 2.3, 1.13, 1.35, C.spSoft);
    txt(s, a, { x, y: 2.3, w: 1.13, h: 1.35, fontFace: HF, fontSize: 54, bold: true, color: C.sp, align: "center", valign: "middle" });
    txt(s, "≠ " + b, { x, y: 3.8, w: 1.13, h: 0.45, fontSize: 20, color: C.soft, align: "center" });
  });
  card(s, 0.7, 4.8, 5.7, 1.8, C.sky);
  txt(s, [{ text: "қоян", options: { bold: true, color: C.sp } }, { text: "  емес  ", options: { color: C.soft } }, { text: "коян", options: { strike: "sngStrike", color: C.soft } }],
    { x: 1.0, y: 5.0, w: 5.2, h: 0.8, fontFace: HF, fontSize: 36 });
  txt(s, "Бір әріп – басқа сөз немесе қате", { x: 1.0, y: 5.9, w: 5.2, h: 0.5, fontSize: 18, color: C.soft });
  card(s, 6.9, 4.8, 5.7, 1.8, C.flagSoft);
  txt(s, "Сөзді жай жаттау – қызықсыз.\nОйын ойнау – қызықты!", { x: 7.2, y: 5.0, w: 5.2, h: 1.4, fontSize: 24, bold: true, color: C.wall, valign: "middle" });
  s.addNotes("Мен орыс тілінде оқитын сыныпта оқимын. Қазақ тілінде 9 ерекше әріп бар. Орыс тілінде олар жоқ. Сондықтан біз жиі қателесеміз: «қоян» орнына «коян» деп жазамыз. Сөздерді жай жаттау қызықсыз, ал ойын ойнау қызықты.");
}

/* 3. Сұрақ, болжам, мақсат */
{
  const s = light(); title(s, "Сұрақ, болжам, мақсат");
  const items = [
    ["❓", "Зерттеу сұрағы", "Менің ойыным сыныптастарыма қазақ сөздерін дұрыс жазуды үйренуге көмектесе ме?", C.flagSoft],
    ["💡", "Болжам", "Ойынды 1–2 апта ойнаса, қорытынды тестте көбірек дұрыс жауап береді.", "FFF1CF"],
    ["🎯", "Мақсат", "Қазақ сөздерін үйренуге арналған лабиринт ойынын жасау және оны сыныптастарыммен тексеру.", C.spSoft],
  ];
  items.forEach(([ic, h, t, f], i) => {
    const x = 0.7 + i * 4.1;
    card(s, x, 1.7, 3.8, 4.9, f);
    txt(s, ic, { x: x + 0.35, y: 2.0, w: 1, h: 0.9, fontSize: 40 });
    txt(s, h, { x: x + 0.35, y: 3.0, w: 3.1, h: 0.6, fontFace: HF, fontSize: 24, bold: true, color: C.wall });
    txt(s, t, { x: x + 0.35, y: 3.75, w: 3.1, h: 2.6, fontSize: 19 });
  });
  s.addNotes("Менің зерттеу сұрағым: мен жасаған ойын сыныптастарыма қазақ сөздерін дұрыс жазуды үйренуге көмектесе ме? Болжамым: ойынды бір-екі апта ойнаса, тестте көбірек дұрыс жауап береді. Мақсатым – осындай ойын жасап, оны тексеру.");
}

/* 4. Жұмыс кезеңдері */
{
  const s = light(); title(s, "Не істедім?");
  const st = [["Сауалнама", "Сыныптастарға не қиын?"], ["70 сөз", "7 тақырып, сөздікпен тексердім"], ["Жоспар", "Дәптерге лабиринт сыздым"],
    ["Claude Code", "Нұсқау бердім, ол код жазды"], ["Сынақ", "3 қате таптым, түзеттік"], ["Тест", "Ойынға дейін және кейін"]];
  st.forEach(([h, t], i) => {
    const x = 0.7 + i * 2.05;
    s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: 1.9, w: 0.8, h: 0.8, fill: { color: i === 3 ? C.sp : C.flag }, line: { color: C.white } });
    txt(s, String(i + 1), { x: x + 0.55, y: 1.9, w: 0.8, h: 0.8, fontSize: 24, bold: true, color: C.white, align: "center", valign: "middle" });
    if (i < st.length - 1) s.addShape(pres.shapes.LINE, { x: x + 1.45, y: 2.3, w: 1.25, h: 0, line: { color: "B7C4CE", width: 2, endArrowType: "triangle" } });
    txt(s, h, { x: x - 0.05, y: 3.0, w: 2.0, h: 0.5, fontFace: HF, fontSize: 18, bold: true, color: C.wall, align: "center" });
    txt(s, t, { x, y: 3.55, w: 1.9, h: 1.2, fontSize: 15, color: C.soft, align: "center" });
  });
  shot(s, "s_themes.png", 0.9, 4.75, 2.45); shot(s, "s_dict.png", 2.4, 4.75, 2.45);
  card(s, 4.3, 5.0, 8.3, 1.75, C.sky);
  txt(s, "Ойынның ережесі: суретке қарап, сөздің әріптерін лабиринттен ретімен жина. Сөз толық жиналғанда ғана шығу ашылады. Қатесіз өтсең – 3 жұлдыз.", { x: 4.6, y: 5.2, w: 7.8, h: 1.4, fontSize: 19, valign: "middle" });
  s.addNotes("Алдымен сыныптастарыма сауалнама жүргіздім. Сосын оқулықтан 7 тақырып бойынша 70 сөз таңдадым. Дәптерге ойынның жоспарын сыздым.");
}

/* 5. Claude Code */
{
  const s = light(); title(s, "Жасанды интеллектпен жұмыс");
  const loop = [["Мен", "ережені сөзбен түсіндірдім", C.flag], ["Claude Code", "HTML, CSS, JavaScript кодын жазды", C.wall], ["Мен", "ойнап көріп, қатені таптым", C.sp]];
  loop.forEach(([h, t, f], i) => {
    const y = 1.6 + i * 1.65;
    card(s, 0.7, y, 5.6, 1.35, f);
    txt(s, h, { x: 1.0, y: y + 0.15, w: 5.0, h: 0.5, fontFace: HF, fontSize: 22, bold: true, color: C.white });
    txt(s, t, { x: 1.0, y: y + 0.65, w: 5.0, h: 0.55, fontSize: 18, color: C.white });
  });
  txt(s, "↻ қайталау", { x: 0.7, y: 6.55, w: 5.6, h: 0.4, fontSize: 16, color: C.soft, align: "center" });
  card(s, 6.9, 1.6, 5.7, 2.4, "FFF1CF");
  txt(s, "«Шешімді адам қабылдайды – жасанды интеллект тек көмектеседі»", { x: 7.2, y: 1.8, w: 5.1, h: 1.4, fontFace: HF, fontSize: 22, italic: true, color: C.wall, valign: "middle" });
  txt(s, "Ережелердің 4.10-тармағы", { x: 7.2, y: 3.35, w: 5.1, h: 0.4, fontSize: 14, color: C.soft });
  txt(s, [
    { text: "Менің шешімім: идея, ереже, кейіпкер, 70 сөз", options: { bullet: true, breakLine: true } },
    { text: "Сынақта 3 қате табылды, бәрі түзетілді", options: { bullet: true, breakLine: true } },
    { text: "ЖИ пайдалану декларацияда ашық жазылған", options: { bullet: true } },
  ], { x: 7.0, y: 4.35, w: 5.6, h: 2.2, fontSize: 19, paraSpaceAfter: 8 });
  s.addNotes("Ойынды Claude Code деген жасанды интеллект көмегімен жасадым. Мен оған ережені сөзбен түсіндірдім, ол код жазды. Мен ойнап көріп, 3 қате таптым, оларды түзеттік. Менің ережем: шешімді адам қабылдайды, жасанды интеллект тек көмектеседі.");
}

/* 6. Ойын */
{
  const s = light(); title(s, "«Сөздер лабиринті» ойыны");
  [["s_home.png", "Басты бет"], ["s_maze.png", "Лабиринт: «қоян»"], ["s_card.png", "Сөз карточкасы"]].forEach(([f, c], i) => {
    const x = 0.8 + i * 3.05; shot(s, f, x, 1.45, 5.0);
    txt(s, c, { x, y: 6.55, w: 2.63, h: 0.4, fontSize: 15, color: C.soft, align: "center" });
  });
  const fx = 10.0;
  [["7", "тақырып"], ["70", "сөз"], ["12", "тест сұрағы"], ["50 КБ", "интернетсіз"]].forEach(([n, l], i) => {
    txt(s, n, { x: fx, y: 1.5 + i * 1.3, w: 2.8, h: 0.75, fontFace: HF, fontSize: 40, bold: true, color: C.flag });
    txt(s, l, { x: fx, y: 2.2 + i * 1.3, w: 2.8, h: 0.4, fontSize: 16, color: C.soft });
  });
  s.addNotes("Қараңыздар: бұл – бота. Суретте қоян тұр. Мен «қ», «о», «я», «н» әріптерін ретімен жинаймын. Мұнда «к» әрпі – бұл қақпан! Оған бассам, ойын «Абай бол!» дейді. Сөз жиналғанда есік ашылады, мен жұлдыз аламын. (Осы жерде ноутбуктан ойынды көрсету, 1 минут.)");
}

/* 7. Сөздік талдауы */
{
  const s = light(); title(s, "70 сөздің 52-сінде қазақ әрпі бар");
  const L = ["қ", "і", "ү", "ғ", "ә", "ұ", "ө", "ң", "һ"], V = [19, 17, 8, 6, 5, 5, 4, 3, 0];
  s.addChart(pres.charts.BAR, [{ name: "Кездесу саны", labels: L, values: V }], {
    x: 0.7, y: 1.5, w: 8.2, h: 5.3, barDir: "col", chartColors: [C.sp], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 16, dataLabelColor: C.ink, dataLabelFontBold: true,
    catAxisLabelFontSize: 24, catAxisLabelColor: C.ink, catAxisLabelFontFace: HF, valAxisLabelColor: C.soft, valAxisLabelFontSize: 12,
    valGridLine: { color: "E3E9EE", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "Әріптердің ойын сөздігінде кездесу саны", titleFontSize: 16, titleColor: C.soft,
  });
  card(s, 9.4, 1.6, 3.2, 2.3, C.spSoft);
  txt(s, "74 %", { x: 9.6, y: 1.8, w: 2.8, h: 1.0, fontFace: HF, fontSize: 54, bold: true, color: C.sp });
  txt(s, "сөзде кемінде бір қазақ әрпі бар", { x: 9.6, y: 2.85, w: 2.8, h: 0.9, fontSize: 16 });
  card(s, 9.4, 4.2, 3.2, 2.5, C.sky);
  txt(s, "«һ» әрпі бар сөз жоқ – келесі нұсқада қосамын (айдаһар, гауһар).", { x: 9.6, y: 4.4, w: 2.8, h: 2.1, fontSize: 17, valign: "middle" });
  s.addNotes("Мен ойындағы 70 сөзді талдадым. Олардың 52-сінде қазақ әрпі бар. «Қ» пен «і» ең жиі кездеседі. «Һ» әрпі бар сөз жоқ – бұл менің ойынымның кемшілігі.");
}

/* 8. Нәтижелер (оқушы толтырады) */
{
  const s = light(); title(s, "Тексеру нәтижелері");
  txt(s, "[__] сыныптасым: бастапқы тест → [__] күн ойын → қорытынды тест", { x: 0.7, y: 1.4, w: 11.9, h: 0.5, fontSize: 20, color: C.soft });
  [["Бастапқы тест", "[__] / 12", C.sky, C.wall], ["Қорытынды тест", "[__] / 12", C.flagSoft, C.flag], ["Өзгеріс", "[+__]", "FFF1CF", "8A5A00"]].forEach(([h, v, f, col], i) => {
    const x = 0.7 + i * 4.1;
    card(s, x, 2.2, 3.8, 2.3, f);
    txt(s, h, { x: x + 0.3, y: 2.4, w: 3.2, h: 0.5, fontSize: 18, color: C.soft });
    txt(s, v, { x: x + 0.3, y: 3.0, w: 3.2, h: 1.1, fontFace: HF, fontSize: 48, bold: true, color: col });
  });
  card(s, 0.7, 4.9, 11.9, 1.8, C.spSoft);
  txt(s, "Ең көп қате: «[__]» әрпі (оның орнына «[__]» басылды). Болжам: [расталды / ішінара расталды / расталмады].", { x: 1.0, y: 5.1, w: 11.3, h: 1.4, fontSize: 22, valign: "middle" });
  s.addNotes("Ойынды [__] сыныптасым ойнады. Бастапқы тестте орташа балл [__] болды, қорытынды тестте – [__]. Ең көп қате «[__]» әрпінде болды. (Нәтиже шыққан соң осы слайдты толтырыңыз; ойынның «Нәтижелер» бөлімінен алыңыз.)");
}

/* 9. Қорытынды */
{
  const s = light(); title(s, "Қорытынды");
  const cols = [
    ["Не шықты", ["Ойын жасалды: 7 тақырып, 70 сөз", "Сабақта да, үйде де ойнауға болады", "Мұғалім қай әріпте қате көп екенін көреді"], C.flagSoft],
    ["Шектеулер", ["Аз оқушы, қысқа уақыт", "«һ» әрпі бар сөз жоқ", "Қазақша дауыс көп құрылғыда жоқ"], C.sky],
    ["Әрі қарай", ["«һ» әрпі бар сөздер", "Мұғалімнің дауысымен жазба", "Сөйлем құрастыру деңгейі"], "FFF1CF"],
  ];
  cols.forEach(([h, items, f], i) => {
    const x = 0.7 + i * 4.1;
    card(s, x, 1.6, 3.8, 4.4, f);
    txt(s, h, { x: x + 0.3, y: 1.85, w: 3.2, h: 0.6, fontFace: HF, fontSize: 24, bold: true, color: C.wall });
    txt(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), { x: x + 0.3, y: 2.7, w: 3.3, h: 3.1, fontSize: 22, paraSpaceAfter: 14 });
  });
  s.addNotes("Ойынды сабақта да, үйде де ойнауға болады. Оған интернет керек емес. Кемшілігі: «һ» әрпі бар сөз жоқ және аз оқушы қатысты. Келесі жолы жаңа сөздер мен дыбыс қосамын.");
}

/* 10. Рефлексия */
{
  const s = pres.addSlide(); s.background = { color: C.wall };
  txt(s, "Мен не үйрендім?", { x: 0.7, y: 0.6, w: 8, h: 0.9, fontFace: HF, fontSize: 40, bold: true, color: C.white });
  txt(s, [
    { text: "Зерттеу сұрағын қоюды және тексеруді", options: { bullet: true, breakLine: true } },
    { text: "Жасанды интеллектке нақты тапсырма беруді", options: { bullet: true, breakLine: true } },
    { text: "Оның жұмысын тексеріп, қатесін табуды", options: { bullet: true } },
  ], { x: 0.7, y: 1.8, w: 7.6, h: 2.6, fontSize: 24, color: "DCE3F7", paraSpaceAfter: 14 });
  txt(s, "Назарларыңызға рахмет!", { x: 0.7, y: 5.3, w: 7.6, h: 0.9, fontFace: HF, fontSize: 36, bold: true, color: C.sun });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.3, y: 1.6, w: 3.3, h: 3.3, fill: { color: C.white }, line: { color: C.white }, rectRadius: 0.1 });
  txt(s, "QR-код:\nойынға сілтеме", { x: 9.3, y: 1.6, w: 3.3, h: 3.3, fontSize: 16, color: C.soft, align: "center", valign: "middle" });
  txt(s, "Ойынды ойнап көріңіз", { x: 9.3, y: 5.05, w: 3.3, h: 0.4, fontSize: 16, color: "DCE3F7", align: "center" });
  s.addNotes("Мен зерттеу жүргізуді, жасанды интеллектке нақты тапсырма беруді және оның жұмысын тексеруді үйрендім. Назарларыңызға рахмет! (QR-кодты stend.js арқылы жасап, осы орынға қойыңыз.)");
}

pres.writeFile({ fileName: OUT }).then(f => console.log("wrote", f));
