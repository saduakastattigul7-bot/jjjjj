const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const path = require("path");

const NAVY = "1F3A5F", GOLD = "C8932E", GREEN = "2F7D5B", BLUE = "2F5D8A";
const TINT = "EEF3F8", GOLD_T = "FBF1DE", GREEN_T = "E3F1EA", INK = "1F2937", MUTED = "5B6472", WHITE = "FFFFFF";
const HEAD = "Times New Roman", BODY = "Arial";

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.title = "Абай мен Пушкин шығармаларындағы жас ұрпақ тәрбиесі";

  const W = 13.33;
  const title = (s, t, sub) => {
    s.addText(t, { x: 0.6, y: 0.4, w: W - 1.2, h: 0.8, fontFace: HEAD, fontSize: 34, bold: true, color: NAVY, margin: 0, isTextBox: true });
    if (sub) s.addText(sub, { x: 0.6, y: 1.15, w: W - 1.2, h: 0.4, fontFace: BODY, fontSize: 15, color: MUTED, italic: true, margin: 0, isTextBox: true });
  };
  const num = (s, n) => s.addText(String(n), { x: W - 1.0, y: 7.0, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 11, color: MUTED, align: "right", margin: 0, isTextBox: true });
  const iconCircle = async (s, x, y, d, name, bg, fg = WHITE) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    const pad = d * 0.25;
    s.addImage({ data: await icon(name, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
  };
  const card = (s, x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill },
  });
  const HL = { highlight: "FFFF00", color: "1F2937" };
  let n = 0;

  // 1. Title
  {
    const s = pres.addSlide(); n++;
    s.background = { color: NAVY };
    s.addText("«Дарын» ғылыми жобалар конкурсы  ·  Секция: Қазақ әдебиеті", { x: 0.8, y: 0.6, w: 11.7, h: 0.4, fontFace: BODY, fontSize: 14, color: "CADCFC", margin: 0, isTextBox: true });
    s.addText("Абай мен Пушкин шығармаларындағы жас ұрпақ тәрбиесі және оның қазіргі қоғамдағы маңызы",
      { x: 0.8, y: 1.5, w: 8.2, h: 2.6, fontFace: HEAD, fontSize: 38, bold: true, color: WHITE, valign: "top", margin: 0, isTextBox: true });
    s.addText("Ұлы классиктердің тәрбиелік ойын жасөспірімнің шешім тіліне аудару", { x: 0.8, y: 4.2, w: 8.2, h: 0.5, fontFace: BODY, fontSize: 17, italic: true, color: GOLD, margin: 0, isTextBox: true });
    s.addText([
      { text: "Орындаған: ", options: { bold: true } }, { text: "[Оқушының аты-жөні], [__] сынып", options: HL }, { text: "", options: { breakLine: true } },
      { text: "Жетекшісі: ", options: { bold: true } }, { text: "[Аты-жөні, лауазымы]", options: HL }, { text: "", options: { breakLine: true } },
      { text: "[Мектеп атауы], [Қала] – 2026", options: HL },
    ], { x: 0.8, y: 5.3, w: 8.2, h: 1.3, fontFace: BODY, fontSize: 15, color: WHITE, margin: 0, paraSpaceAfter: 4, isTextBox: true });
    // TOQTA motif on the right
    const letters = ["Т", "О", "Қ", "Т", "А"], cols = [BLUE, BLUE, GOLD, GREEN, GREEN];
    letters.forEach((l, i) => {
      const y = 1.2 + i * 1.08;
      s.addShape(pres.shapes.OVAL, { x: 10.6, y, w: 0.9, h: 0.9, fill: { color: cols[i] }, line: { color: WHITE, width: 1.5 } });
      s.addText(l, { x: 10.6, y, w: 0.9, h: 0.9, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addNotes("Құрметті әділқазылар алқасы! Менің жобамның тақырыбы – «Абай мен Пушкин шығармаларындағы жас ұрпақ тәрбиесі және оның қазіргі қоғамдағы маңызы». Мен екі ұлы ақынның тәрбиелік ойын бүгінгі жасөспірім күнделікті қолдана алатын нақты құралға айналдыруға тырыстым.");
  }

  // 2. Relevance
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Тақырыптың өзектілігі", "Жасөспірімнің басты қиындығы – ақпараттың аздығы емес, құндылықты әрекетке айналдыра алмау");
    card(s, 0.6, 1.9, 5.2, 4.6, NAVY);
    s.addText("«Адал болу керек екенін білемін...»", { x: 0.9, y: 2.2, w: 4.6, h: 1.4, fontFace: HEAD, fontSize: 26, italic: true, color: WHITE, margin: 0, valign: "top", isTextBox: true });
    s.addText("...бірақ қиын сәтте не істеу керегін бірден таба алмаймын.", { x: 0.9, y: 3.6, w: 4.6, h: 1.0, fontFace: HEAD, fontSize: 20, color: "CADCFC", margin: 0, valign: "top", isTextBox: true });
    s.addText("Психологияда бұл «моральдық білім мен әрекет арасындағы алшақтық» деп аталады (A. Blasi, 1980).", { x: 0.9, y: 5.0, w: 4.6, h: 1.2, fontFace: BODY, fontSize: 13, color: WHITE, margin: 0, valign: "top", isTextBox: true });
    const items = [
      ["FaCopy", "Көшіріп жазу", "Бақылау жұмысында досың дайын жауапты жіберді"],
      ["FaCommentSlash", "Желідегі қорлау", "Чатта сыныптасты келеке етіп жатыр"],
      ["FaUsers", "Топ қысымы", "«Бәрі кетіп жатыр, сен неге қорқасың?»"],
    ];
    for (let i = 0; i < items.length; i++) {
      const y = 1.95 + i * 1.55;
      await iconCircle(s, 6.4, y, 1.0, items[i][0], GOLD);
      s.addText(items[i][1], { x: 7.7, y: y + 0.02, w: 5.0, h: 0.45, fontFace: BODY, fontSize: 19, bold: true, color: INK, margin: 0, isTextBox: true });
      s.addText(items[i][2], { x: 7.7, y: y + 0.5, w: 5.0, h: 0.45, fontFace: BODY, fontSize: 15, color: MUTED, margin: 0, isTextBox: true });
    }
    num(s, n);
    s.addNotes("Бүгінгі оқушы адалдықтың, мейірімнің жақсы қасиет екенін жақсы біледі. Бірақ бақылау жұмысында көшіруге, чаттағы қорлауға немесе достардың қысымына тап болғанда, қандай қадам жасау керегін бірден таба алмайды. Психологияда мұны моральдық білім мен әрекет арасындағы алшақтық деп атайды. Сондықтан классиктердің тәрбиелік ойын бүгінгі шешімдер тіліне аудару маңызды.");
  }

  // 3. Problem / aim / hypothesis
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Зерттеу мәселесі, мақсаты және болжамы");
    const cards = [
      ["FaQuestion", "Мәселе", "Абай мен Пушкиннің тәрбиелік идеяларын қазіргі жасөспірім қолдана алатын ықшам шешім алгоритміне қалай айналдыруға болады?", BLUE, TINT],
      ["FaBullseye", "Мақсат", "Екі ақынның ортақ және дара тәрбиелік ұстанымдарын мәтіндік талдау арқылы анықтап, солардың негізінде практикалық тәрбие моделін ұсыну", GOLD, GOLD_T],
      ["FaLightbulb", "Болжам", "Егер ұстанымдар жалпы үндеу емес, қадамдық алгоритм түрінде берілсе, жасөспірім құндылықты нақты әрекетке жиірек айналдырады", GREEN, GREEN_T],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      card(s, x, 1.6, 3.8, 4.9, cards[i][4]);
      await iconCircle(s, x + 0.35, 1.95, 0.9, cards[i][0], cards[i][3]);
      s.addText(cards[i][1], { x: x + 0.35, y: 3.05, w: 3.1, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: cards[i][3], margin: 0, isTextBox: true });
      s.addText(cards[i][2], { x: x + 0.35, y: 3.75, w: 3.1, h: 2.6, fontFace: BODY, fontSize: 18, color: INK, margin: 0, valign: "top", isTextBox: true });
    }
    num(s, n);
    s.addNotes("Зерттеу мәселесі: екі ақынның тәрбиелік идеяларын жай ұқсастықтарды санамалаумен шектемей, жасөспірім қолдана алатын нақты алгоритмге қалай айналдыруға болады? Мақсатым – екі ақынның ортақ және дара ұстанымдарын анықтап, солардың негізінде тәрбие моделін ұсыну. Болжамым: егер ұстанымдар қадамдық алгоритм түрінде берілсе, оқушы оларды нақты әрекетке жиірек айналдырады.");
  }

  // 4. Tasks
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Зерттеу міндеттері");
    const tasks = [
      "Абайтану, пушкинтану және тәрбие психологиясы бойынша ғылыми әдебиетке шолу жасау",
      "22 мәтін бірлігінен корпус құрып, 7 кодтан тұратын талдау жүйесін әзірлеу",
      "Екі ақынның ортақ және дара тәрбиелік ұстанымдарын салыстыру",
      "«Ар – Ақыл – Жүрек» моделін және «ТОҚТА» алгоритмін құрастыру",
      "Модельді жағдаяттарға қолданып, пилоттық апробация өткізу",
    ];
    tasks.forEach((t, i) => {
      const y = 1.6 + i * 1.0;
      s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.7, h: 0.7, fill: { color: i < 3 ? NAVY : GOLD }, line: { color: i < 3 ? NAVY : GOLD } });
      s.addText(String(i + 1), { x: 0.8, y, w: 0.7, h: 0.7, fontFace: HEAD, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: 1.8, y, w: 10.8, h: 0.7, fontFace: BODY, fontSize: 18, color: INK, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("Әдістер: әдебиетті талдау · салыстырмалы-типологиялық талдау · контент-талдау (кодтау) · герменевтика · модельдеу · сауалнама",
      { x: 0.8, y: 6.65, w: 11.8, h: 0.4, fontFace: BODY, fontSize: 13, italic: true, color: MUTED, margin: 0, isTextBox: true });
    num(s, n);
    s.addNotes("Осы мақсатқа жету үшін бес міндет қойдым. Алдымен ғылыми әдебиетке шолу жасадым. Содан кейін екі ақынның 22 шығармасынан корпус құрып, жеті кодтан тұратын талдау жүйесін әзірледім. Үшіншіден, ақындардың ұстанымдарын салыстырдым. Төртіншіден, модель мен алгоритм құрастырдым. Соңында оны нақты жағдаяттарға қолданып, сыныпта сынап көрдім.");
  }

  // 5. Method
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Зерттеу әдістемесі", "Бес кезең: мәселе → корпус → кодтау → модель → апробация");
    s.addImage({ path: path.join(__dirname, "figs/fig1.png"), x: 0.6, y: 1.75, w: 7.6, h: 2.43 });
    const stats = [["22", "мәтін бірлігі", "Абайдан 11, Пушкиннен 11"], ["7", "тақырыптық код", "білім, ар, мейірім, еңбек, рефлексия, қоғам, жаман әдет"], ["26", "ғылыми дереккөз", "қазақ, орыс, ағылшын тілдерінде"]];
    stats.forEach((st, i) => {
      const y = 1.7 + i * 1.65;
      card(s, 8.7, y, 4.0, 1.45, i === 1 ? GOLD_T : TINT);
      s.addText(st[0], { x: 8.9, y: y + 0.1, w: 1.3, h: 1.25, fontFace: HEAD, fontSize: 48, bold: true, color: i === 1 ? GOLD : NAVY, valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: st[1], options: { bold: true, fontSize: 16, color: INK, breakLine: true } }, { text: st[2], options: { fontSize: 12, color: MUTED } }],
        { x: 10.2, y: y + 0.15, w: 2.4, h: 1.15, fontFace: BODY, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText([
      { text: "Кодтау: ", options: { bold: true } },
      { text: "әр мәтінге құндылық айқын көрінсе «+» қойылды; кодтау нұсқаулығы жасалып, матрицаны жетекші тәуелсіз тексерді.", options: { breakLine: true } },
      { text: "Герменевтика (Х.-Г. Гадамер): ", options: { bold: true } },
      { text: "мәтінді түсіну – оны өз жағдайыңа қолданумен тығыз байланысты.", options: {} },
    ], { x: 0.6, y: 4.5, w: 7.6, h: 1.9, fontFace: BODY, fontSize: 15, color: INK, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true });
    num(s, n);
    s.addNotes("Зерттеу бес кезеңнен тұрды. Абайдан 11, Пушкиннен 11 мәтін бірлігін таңдап алдым. Әр мәтінді жеті код бойынша талдадым: білім, ар, мейірім, еңбек, өзіне есеп беру, қоғам және жаман әдеттен сақтану. Кодтаудың объективтілігі үшін нұсқаулық жасадым, ал матрицаны жетекшім тәуелсіз тексерді. Пушкин шығармаларын түпнұсқада оқып, үзінділерді өзім аудардым.");
  }

  // 6. Abai
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Абай: «Адам бол!»", "Тәрбие тәсілі – тура ақыл-кеңес, үндеу, тізім");
    card(s, 0.6, 1.75, 5.4, 4.9, NAVY);
    s.addImage({ data: await icon("FaQuoteLeft", GOLD), x: 0.9, y: 2.0, w: 0.5, h: 0.5 });
    s.addText("Бес нәрседен қашық бол,\nБес нәрсеге асық бол,\nАдам болам десеңіз...\nӨсек, өтірік, мақтаншақ,\nЕріншек, бекер мал шашпақ –\nБес дұшпаның білсеңіз.\nТалап, еңбек, терең ой,\nҚанағат, рақым, ойлап қой –\nБес асыл іс, көнсеңіз.",
      { x: 0.9, y: 2.6, w: 4.9, h: 3.5, fontFace: HEAD, fontSize: 16, italic: true, color: WHITE, margin: 0, valign: "top", isTextBox: true });
    s.addText("«Ғылым таппай мақтанба»", { x: 0.9, y: 6.1, w: 4.9, h: 0.4, fontFace: BODY, fontSize: 13, color: GOLD, margin: 0, isTextBox: true });
    const pr = [
      ["FaBookOpen", "Адам білімді болып тумайды", "Білім еңбекпен, «жан құмарымен» табылады (19, 7-қара сөз)"],
      ["FaBalanceScale", "Білім мақтан үшін емес", "Білімді оның өзі үшін сүю керек (32-қара сөз)"],
      ["FaHeart", "Толық адам", "Ақыл, қайрат, жүрек бірлігі; билік – жүректе (17-қара сөз)"],
      ["FaSyncAlt", "Өзіңнен өзің есеп ал", "Күнде, аптада, айда бір рет (15-қара сөз)"],
    ];
    for (let i = 0; i < pr.length; i++) {
      const y = 1.8 + i * 1.22;
      await iconCircle(s, 6.5, y, 0.8, pr[i][0], BLUE);
      s.addText(pr[i][1], { x: 7.5, y, w: 5.2, h: 0.4, fontFace: BODY, fontSize: 17, bold: true, color: INK, margin: 0, isTextBox: true });
      s.addText(pr[i][2], { x: 7.5, y: y + 0.42, w: 5.2, h: 0.5, fontFace: BODY, fontSize: 13, color: MUTED, margin: 0, isTextBox: true });
    }
    num(s, n);
    s.addNotes("Абайдың тәрбие туралы ойын төрт ұстанымға топтастырдым. Біріншісі – адам білімді болып тумайды, оны еңбекпен табады. Екіншісі – білім мақтан үшін емес, адамгершілік үшін керек. Үшіншісі – толық адам ақыл, қайрат және жүректің бірлігінен туады. Төртіншісі – адам өзінен өзі есеп алып отыруы керек. Абай ақыл-кеңесті тура айтады: бес дұшпаннан қаш, бес асыл іске ұмтыл.");
  }

  // 7. Pushkin
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Пушкин: «Береги честь смолоду»", "Тәрбие тәсілі – көркем үлгі, кейіпкердің таңдауы");
    const pr = [
      ["FaShieldAlt", "Ар-намысты жастан сақта", "Гринёв ант бұзбайды, Швабрин – ақылды, бірақ арсыз («Капитанская дочка»)"],
      ["FaHandHoldingHeart", "Мейірім қайтады", "Қоян тон эпизоды; «милость к падшим» («Я памятник...»)"],
      ["FaPenFancy", "Үстірт білім мен еріншектік", "«Мы все учились понемногу / Чему-нибудь и как-нибудь» (Онегин)"],
      ["FaHandshake", "Берген сөзге адалдық, достық", "Татьяна; «Друзья мои, прекрасен наш союз!»"],
    ];
    for (let i = 0; i < pr.length; i++) {
      const y = 1.8 + i * 1.22;
      await iconCircle(s, 0.6, y, 0.8, pr[i][0], GOLD);
      s.addText(pr[i][1], { x: 1.6, y, w: 5.3, h: 0.4, fontFace: BODY, fontSize: 17, bold: true, color: INK, margin: 0, isTextBox: true });
      s.addText(pr[i][2], { x: 1.6, y: y + 0.42, w: 5.3, h: 0.5, fontFace: BODY, fontSize: 13, color: MUTED, margin: 0, isTextBox: true });
    }
    card(s, 7.3, 1.75, 5.4, 4.9, GOLD_T);
    s.addImage({ data: await icon("FaQuoteLeft", GOLD), x: 7.6, y: 2.0, w: 0.5, h: 0.5 });
    s.addText("И долго буду тем любезен я народу,\nЧто чувства добрые я лирой пробуждал,\nЧто в мой жестокий век восславил я свободу\nИ милость к падшим призывал.",
      { x: 7.6, y: 2.6, w: 4.9, h: 2.0, fontFace: HEAD, fontSize: 16, italic: true, color: INK, margin: 0, valign: "top", isTextBox: true });
    s.addText("Аудармасы: «Халқыма сол үшін жақын боламын: ізгі сезімдерді ояттым, қатал ғасырымда еркіндікті жырладым және құлағандарға рақым етуге шақырдым».",
      { x: 7.6, y: 4.75, w: 4.9, h: 1.6, fontFace: BODY, fontSize: 13, color: MUTED, margin: 0, valign: "top", isTextBox: true });
    num(s, n);
    s.addNotes("Пушкин тәрбиелік ойын басқаша жеткізеді: ол ақыл айтпай, кейіпкерді таңдау алдына қояды. «Капитан қызында» Гринёв ар-намысын сақтайды, ал Швабрин ақылды болса да, арын сатады. Бұл – білімнің өзі адамгершілікке кепілдік бермейтінін көрсетеді. Пушкин мейірімді, еңбекқорлықты, берген сөзге адалдықты және достықты да кейіпкерлері арқылы көрсетеді.");
  }

  // 8. Chart
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Нәтиже 1: кодтар қалай бөлінді?", "Код кездескен мәтін бірліктерінің саны (әр ақыннан 11 мәтін)");
    const labels = ["Білім, ақыл", "Ар, адалдық", "Мейірім", "Еңбек, қайрат", "Өзіне есеп беру", "Достық, қоғам", "Жаман әдеттен сақтану"];
    s.addChart(pres.charts.BAR, [
      { name: "Абай", labels, values: [9, 2, 2, 6, 3, 3, 3] },
      { name: "Пушкин", labels, values: [1, 6, 5, 3, 4, 4, 6] },
    ], {
      x: 0.5, y: 1.7, w: 8.4, h: 5.0, barDir: "col", barGrouping: "clustered", chartColors: [BLUE, GOLD],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: INK,
      catAxisLabelFontSize: 11, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, valAxisLabelFontSize: 10,
      valAxisMaxVal: 10, valAxisMinVal: 0, valAxisMajorUnit: 2, valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" },
      showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: INK, barGapWidthPct: 60,
    });
    const pts = [
      [NAVY, "Жеті кодтың бәрі екі ақында да бар", "Тәрбиелік мазмұн өрісі ортақ"],
      [BLUE, "Абайда: білім (9) және еңбек (6)", "Тәрбиенің бастауы – білім мен ақыл"],
      [GOLD, "Пушкинде: ар (6), жаман әдет (6), мейірім (5)", "Тәрбиенің бастауы – ар мен ізгі сезім"],
      [GREEN, "Бірге алғанда әр код 7–10 мәтінде", "Екі ақын бір-бірін толықтырады"],
    ];
    pts.forEach((p, i) => {
      const y = 1.8 + i * 1.22;
      s.addShape(pres.shapes.OVAL, { x: 9.3, y: y + 0.08, w: 0.28, h: 0.28, fill: { color: p[0] }, line: { color: p[0] } });
      s.addText(p[1], { x: 9.75, y, w: 3.1, h: 0.5, fontFace: BODY, fontSize: 15, bold: true, color: INK, margin: 0, valign: "top", isTextBox: true });
      s.addText(p[2], { x: 9.75, y: y + 0.55, w: 3.1, h: 0.5, fontFace: BODY, fontSize: 12, color: MUTED, margin: 0, valign: "top", isTextBox: true });
    });
    num(s, n);
    s.addNotes("Кодтау нәтижесі диаграммада көрсетілген. Біріншіден, жеті кодтың барлығы екі ақында да кездеседі, яғни олардың тәрбие туралы түсінігі ортақ. Екіншіден, екпін әр түрлі: Абайда білім мен еңбек басым, ал Пушкинде ар-намыс, мейірім және жаман мінездің салдары жиі көрінеді. Ең маңызды нәтиже: екі ақынды бірге алғанда әр құндылық біркелкі қамтылады, яғни олар бір-бірін толықтырады.");
  }

  // 9. Common vs distinct
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Нәтиже 2: ортақ және дара белгілер");
    const cols = [
      ["АБАЙ", BLUE, TINT, ["«Толық адам»: ақыл – қайрат – жүрек", "Бес дұшпан / бес асыл іс", "Тура үндеу, ақыл-кеңес", "«Не істеу керек» дейді"]],
      ["ОРТАҚ", NAVY, "DCE4EE", ["Адамгершілік білімнен жоғары", "Жастық – мінез қалыптасар кезең", "Адалдық, еңбек, мейірім", "Өзіне есеп беру"]],
      ["ПУШКИН", GOLD, GOLD_T, ["«Береги честь смолоду»", "Кейіпкер таңдауы арқылы үлгі", "«Милость к падшим»", "«Таңдау қалай жасалады» көрсетеді"]],
    ];
    cols.forEach((c, i) => {
      const x = 0.6 + i * 4.1;
      card(s, x, 1.5, 3.8, 4.4, c[2]);
      s.addText(c[0], { x: x + 0.3, y: 1.7, w: 3.2, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: c[1], margin: 0, isTextBox: true });
      s.addText(c[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c[3].length - 1 } })),
        { x: x + 0.3, y: 2.4, w: 3.3, h: 3.3, fontFace: BODY, fontSize: 16, color: INK, margin: 0, valign: "top", paraSpaceAfter: 10, isTextBox: true });
    });
    card(s, 0.6, 6.1, 12.1, 0.9, NAVY);
    s.addText([{ text: "Қорытынды идея:  ", options: { bold: true, color: GOLD } }, { text: "ереже (Абай) + үлгі (Пушкин) = жасөспірімге арналған тәрбие моделінің негізі", options: { color: WHITE } }],
      { x: 0.9, y: 6.1, w: 11.5, h: 0.9, fontFace: BODY, fontSize: 18, valign: "middle", margin: 0, isTextBox: true });
    num(s, n);
    s.addNotes("Салыстыру нәтижесінде екі ақынның ортақ ұстанымы – адамгершілік білімнен жоғары екені анықталды. Абайда ғылым билікті жүрекке береді, ал Пушкинде ақылды, бірақ арсыз Швабрин жеңіліс табады. Айырмашылық тәсілде: Абай не істеу керегін айтады, Пушкин таңдаудың қалай жасалатынын көрсетеді. Осы екеуін біріктіру – менің моделімнің негізі.");
  }

  // 10. Model
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "«Ар – Ақыл – Жүрек» тәрбие моделі", "Абайдың үштігі қазіргі мінез тәрбиесі теориясымен үндеседі (T. Lickona, 1991)");
    const tri = [["FaBrain", "АҚЫЛ", "жақсылықты білу", BLUE], ["FaHeart", "ЖҮРЕК", "жақсылықты қалау", GOLD], ["FaFistRaised", "ҚАЙРАТ", "жақсылықты істеу", GREEN]];
    for (let i = 0; i < 3; i++) {
      const y = 1.8 + i * 1.6;
      await iconCircle(s, 0.7, y, 1.1, tri[i][0], tri[i][3]);
      s.addText(tri[i][1], { x: 2.0, y: y + 0.05, w: 2.6, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true, color: tri[i][3], margin: 0, isTextBox: true });
      s.addText("= " + tri[i][2], { x: 2.0, y: y + 0.55, w: 2.8, h: 0.45, fontFace: BODY, fontSize: 15, color: INK, margin: 0, isTextBox: true });
    }
    const lv = [
      ["1-деңгей. Құндылық өзегі", "Абай: ақыл, қайрат, жүрек; бес асыл іс  ·  Пушкин: ар, мейірім, достық", TINT, BLUE],
      ["2-деңгей. Шешім алгоритмі «ТОҚТА»", "Құндылықты әрекетке айналдыратын «көпір» – бес қадам", GOLD_T, GOLD],
      ["3-деңгей. Нақты жағдаят", "Көшіріп жазу  ·  Желідегі қорлау  ·  Топ қысымы  ·  Жасанды интеллект", GREEN_T, GREEN],
    ];
    lv.forEach((l, i) => {
      const y = 1.8 + i * 1.6;
      card(s, 5.3, y, 7.4, 1.3, l[2]);
      s.addText(l[0], { x: 5.6, y: y + 0.15, w: 6.9, h: 0.45, fontFace: BODY, fontSize: 18, bold: true, color: l[3], margin: 0, isTextBox: true });
      s.addText(l[1], { x: 5.6, y: y + 0.65, w: 6.9, h: 0.5, fontFace: BODY, fontSize: 14, color: INK, margin: 0, isTextBox: true });
      if (i < 2) s.addShape(pres.shapes.DOWN_ARROW, { x: 8.8, y: y + 1.32, w: 0.4, h: 0.26, fill: { color: "9CA3AF" }, line: { color: "9CA3AF" } });
    });
    num(s, n);
    s.addNotes("Зерттеу барысында қызық сәйкестік таптым: американдық ғалым Томас Ликона мінезді жақсылықты білу, қалау және істеу деп сипаттайды. Бұл дәл Абайдың ақыл, жүрек және қайрат үштігі. Осының негізінде үш деңгейлі модель құрдым: бірінші деңгей – құндылықтар, екінші – шешім алгоритмі, үшінші – нақты өмірлік жағдаяттар.");
  }

  // 11. TOQTA
  {
    const s = pres.addSlide(); n++;
    s.background = { color: NAVY };
    s.addText("«ТОҚТА» шешім алгоритмі", { x: 0.6, y: 0.4, w: 12, h: 0.8, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText("Қиын сәтте ең алдымен – тоқта. Бес қадам жадта оңай сақталады.", { x: 0.6, y: 1.15, w: 12, h: 0.4, fontFace: BODY, fontSize: 15, italic: true, color: "CADCFC", margin: 0, isTextBox: true });
    const st = [
      ["Т", "Тоқта", "10 рет дем ал, бірден жауап берме", "«Ақырын жүріп, анық бас» – Абай", BLUE],
      ["О", "Ойлан", "Қай құндылық сынға түсті?", "Бес асыл іс, «честь»", BLUE],
      ["Қ", "Қайтарымын көр", "Салдары: өзіме, өзгеге, ертеңіме", "Гринёв пе, Швабрин бе?", GOLD],
      ["Т", "Таңда да, істе", "«Егер ..., онда мен ...»", "Абайдың «қайраты»", GREEN],
      ["А", "Артынан есеп ал", "Не істедім? Неге? Не өзгертемін?", "«Өзіңнен өзің есеп ал» – 15-қара сөз", GREEN],
    ];
    st.forEach((q, i) => {
      const x = 0.6 + i * 2.48;
      card(s, x, 1.9, 2.25, 4.9, "2A4A73");
      s.addShape(pres.shapes.OVAL, { x: x + 0.62, y: 2.15, w: 1.0, h: 1.0, fill: { color: q[4] }, line: { color: WHITE, width: 1.5 } });
      s.addText(q[0], { x: x + 0.62, y: 2.15, w: 1.0, h: 1.0, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(q[1], { x: x + 0.15, y: 3.35, w: 1.95, h: 0.8, fontFace: BODY, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(q[2], { x: x + 0.15, y: 4.2, w: 1.95, h: 1.3, fontFace: BODY, fontSize: 14, color: WHITE, align: "center", valign: "top", margin: 0, isTextBox: true });
      s.addText(q[3], { x: x + 0.15, y: 5.6, w: 1.95, h: 1.0, fontFace: BODY, fontSize: 12, italic: true, color: GOLD, align: "center", valign: "top", margin: 0, isTextBox: true });
    });
    s.addText(String(n), { x: W - 1.0, y: 7.0, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 11, color: "CADCFC", align: "right", margin: 0, isTextBox: true });
    s.addNotes("Жобаның басты нәтижесі – «ТОҚТА» алгоритмі. Бұл сөз бес қадамның бас әріптерінен құралған және өзі бірінші қадамды еске салады. Т – тоқта, асықпа. О – ойлан: қай құндылық сынға түсіп тұр? Қ – қайтарымын көр: салдары өзіңе, өзгеге, ертеңіңе қандай? Т – таңда да, істе: «егер... онда мен...» деп нақты жоспар құр. А – артынан есеп ал, Абай айтқандай. Әр қадам Абай мен Пушкиннің мәтіндеріне және психология ғылымына негізделген.");
  }

  // 12. Situations
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Алгоритм іс жүзінде", "«Таңда да, істе» қадамындағы нақты шешімдер");
    const sit = [
      ["FaCopy", "Көшіріп жазу", "«Телефонды сөмкеге саламын, түсінбегенімді сабақтан кейін мұғалімнен сұраймын»", "Абай: еріншек / еңбек"],
      ["FaCommentSlash", "Желідегі қорлау", "«Жәбірленушіге қолдау хат жазамын және ересекке айтамын»", "Пушкин: «милость к падшим»"],
      ["FaUsers", "Топ қысымы", "«Бармаймын, бірақ сабақтан кейін бірге барайық» деп балама ұсынамын", "Пушкин: лицей достығы"],
      ["FaRobot", "Жасанды интеллект", "«ЖИ-ді дайын мәтін үшін емес, жоспар құруға ғана қолданамын»", "Абай: талап, еңбек, терең ой"],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 3.08;
      card(s, x, 1.7, 2.85, 5.0, i % 2 ? GOLD_T : TINT);
      await iconCircle(s, x + 0.95, 1.95, 0.95, sit[i][0], i % 2 ? GOLD : BLUE);
      s.addText(sit[i][1], { x: x + 0.2, y: 3.05, w: 2.45, h: 0.5, fontFace: BODY, fontSize: 17, bold: true, color: INK, align: "center", margin: 0, isTextBox: true });
      s.addText(sit[i][2], { x: x + 0.2, y: 3.65, w: 2.45, h: 2.1, fontFace: HEAD, fontSize: 15, italic: true, color: INK, align: "center", valign: "top", margin: 0, isTextBox: true });
      s.addText(sit[i][3], { x: x + 0.2, y: 5.85, w: 2.45, h: 0.6, fontFace: BODY, fontSize: 12, color: MUTED, align: "center", valign: "top", margin: 0, isTextBox: true });
    }
    num(s, n);
    s.addNotes("Алгоритмді төрт жағдаятқа қолдандым. Мысалы, көшіріп жазу кезінде оқушы телефонды сөмкесіне салып, түсінбеген тапсырманы кейін мұғалімнен сұрайды. Чаттағы қорлауда – жәбірленушіні қолдап, ересекке айтады. Топ қысымында – бас тартып, балама ұсынады. Ал жасанды интеллектті дайын мәтін үшін емес, жоспар құру үшін ғана пайдаланады. Әр шешімнің артында Абай мен Пушкиннің нақты ойы тұр.");
  }

  // 13. Pilot
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Пилоттық апробация", "Екі сабақ, дилеммалық жағдаяттар, 0–2 балдық шкала");
    const stats = [["[__]", "оқушы қатысты", BLUE], ["[__] → [__]", "орташа балл (0–2)\nалдын ала → кейін", GOLD], ["[__] %", "нақты әрекет ұсынған\nжауаптар үлесі (кейін)", GREEN]];
    stats.forEach((st, i) => {
      const x = 0.6 + i * 4.1;
      card(s, x, 1.7, 3.8, 2.4, TINT);
      s.addText(st[0], { x: x + 0.2, y: 1.85, w: 3.4, h: 1.1, fontFace: HEAD, fontSize: 40, bold: true, color: st[2], align: "center", valign: "middle", margin: 0, highlight: "FFFF00", isTextBox: true });
      s.addText(st[1], { x: x + 0.2, y: 3.0, w: 3.4, h: 0.9, fontFace: BODY, fontSize: 14, color: MUTED, align: "center", valign: "top", margin: 0, isTextBox: true });
    });
    const sc = [["0", "құндылық байқалмаған / зиянды әрекет"], ["1", "құндылық аталған, әрекет жоқ («адал болу керек»)"], ["2", "құндылық + нақты орындалатын әрекет"]];
    sc.forEach((r, i) => {
      const y = 4.5 + i * 0.7;
      s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: [MUTED, GOLD, GREEN][i] }, line: { color: [MUTED, GOLD, GREEN][i] } });
      s.addText(r[0], { x: 0.8, y, w: 0.5, h: 0.5, fontFace: BODY, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(r[1], { x: 1.5, y, w: 6.0, h: 0.5, fontFace: BODY, fontSize: 15, color: INK, valign: "middle", margin: 0, isTextBox: true });
    });
    card(s, 8.0, 4.4, 4.7, 2.3, GOLD_T);
    s.addText([{ text: "Қорытынды: ", options: { bold: true } }, { text: "[апробация нәтижесі бойынша 1–2 сөйлем жазыңыз]", options: { highlight: "FFFF00" } }],
      { x: 8.25, y: 4.55, w: 4.2, h: 2.0, fontFace: BODY, fontSize: 15, color: INK, margin: 0, valign: "top", isTextBox: true });
    num(s, n);
    s.addNotes("[БҰЛ СЛАЙДТЫ АПРОБАЦИЯДАН КЕЙІН ТОЛТЫРЫҢЫЗ.] Модельді сыныпта екі сабақ көлемінде сынап көрдім. Оқушылар үш жағдаятқа сабаққа дейін және кейін жазбаша жауап берді. Жауаптарды 0-ден 2-ге дейінгі шкаламен бағаладым: 2 балл – құндылықты атап қана қоймай, нақты әрекет ұсынған жауап. Нәтижесінде ... [нақты сандарды айтыңыз].");
  }

  // 14. Conclusions
  {
    const s = pres.addSlide(); n++;
    s.background = { color: WHITE };
    title(s, "Қорытынды және практикалық маңызы");
    const c = [
      "Екі ақынның тәрбиелік мазмұн өрісі ортақ: жеті құндылықтың бәрі екеуінде де бар",
      "Айырмашылық тәсілде: Абай – ереже мен ақыл-кеңес, Пушкин – кейіпкердің таңдауы",
      "«Ар – Ақыл – Жүрек» моделі және бес қадамды «ТОҚТА» алгоритмі ұсынылды",
      "Алгоритм көшіру, кибербуллинг, топ қысымы және ЖИ жағдаяттарында қолданылды",
    ];
    c.forEach((t, i) => {
      const y = 1.6 + i * 1.3;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.05, w: 0.6, h: 0.6, fill: { color: NAVY }, line: { color: NAVY } });
      s.addText(String(i + 1), { x: 0.6, y: y + 0.05, w: 0.6, h: 0.6, fontFace: HEAD, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: 1.45, y, w: 6.4, h: 0.9, fontFace: BODY, fontSize: 17, color: INK, valign: "middle", margin: 0, isTextBox: true });
    });
    card(s, 8.3, 1.5, 4.4, 5.4, GREEN_T);
    s.addText("Қайда қолдануға болады?", { x: 8.6, y: 1.7, w: 3.9, h: 0.5, fontFace: HEAD, fontSize: 20, bold: true, color: GREEN, margin: 0, isTextBox: true });
    const use = [["FaBookOpen", "Әдебиет сабақтары"], ["FaChalkboardTeacher", "Сынып сағаттары"], ["FaUserFriends", "Психолог тренингтері"], ["FaIdCard", "«ТОҚТА» жадынамасы"], ["FaHome", "Отбасылық әңгіме"]];
    for (let i = 0; i < use.length; i++) {
      const y = 2.45 + i * 0.88;
      await iconCircle(s, 8.6, y, 0.6, use[i][0], GREEN);
      s.addText(use[i][1], { x: 9.4, y, w: 3.1, h: 0.6, fontFace: BODY, fontSize: 15, color: INK, valign: "middle", margin: 0, isTextBox: true });
    }
    num(s, n);
    s.addNotes("Қорытындылай келе: екі ақынның тәрбиелік өрісі ортақ, ал айырмашылық тәсілде. Осының негізінде «Ар – Ақыл – Жүрек» моделі мен «ТОҚТА» алгоритмі ұсынылды. Жұмыс нәтижелерін әдебиет сабақтарында, сынып сағаттарында, психолог тренингтерінде және отбасында қолдануға болады. Болашақта зерттеуді бақылау тобы бар экспериментпен жалғастырғым келеді.");
  }

  // 15. Thank you
  {
    const s = pres.addSlide(); n++;
    s.background = { color: NAVY };
    s.addImage({ data: await icon("FaQuoteLeft", GOLD), x: 1.2, y: 1.4, w: 0.8, h: 0.8 });
    s.addText("Адамның адамшылығы істі бастауынан білінеді, қалайша бітіруінен емес.", { x: 1.2, y: 2.4, w: 10.9, h: 1.8, fontFace: HEAD, fontSize: 34, italic: true, color: WHITE, margin: 0, valign: "top", isTextBox: true });
    s.addText("Абай, 37-қара сөз", { x: 1.2, y: 4.2, w: 10.9, h: 0.5, fontFace: BODY, fontSize: 17, color: GOLD, margin: 0, isTextBox: true });
    s.addText("Назарларыңызға рахмет!", { x: 1.2, y: 5.6, w: 10.9, h: 0.8, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addNotes("Абай айтқандай, адамның адамшылығы істі бастауынан білінеді. Менің алгоритмім де дұрыс шешімнің дәл осы алғашқы қадамына арналған. Назарларыңызға рахмет! Сұрақтарыңызға жауап беруге дайынмын.");
  }

  await pres.writeFile({ fileName: path.join(__dirname, "deck.pptx") });
  console.log("written");
})();
