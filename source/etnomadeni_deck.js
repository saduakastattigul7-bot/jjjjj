const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const C = {
  dark: "13334C",
  teal: "1F7A80",
  gold: "D9A441",
  tint: "EEF4F6",
  goldTint: "FBF3E2",
  text: "1E2A33",
  muted: "5B6B75",
  white: "FFFFFF",
};
const FONT = "Arial";

async function svgPng(svg, size = 512) {
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: 256 })
  );
  return svgPng(svg, 256);
}
function shanyrakSvg(color, opacity = 1) {
  const c = "#" + color;
  const lines = [-16, 0, 16]
    .map(
      (d) =>
        `<path d="M -40 ${d} Q 0 ${d - 10} 40 ${d}" />` +
        `<path d="M ${d} -40 Q ${d - 10} 0 ${d} 40" />`
    )
    .join("");
  const rays = Array.from({ length: 24 }, (_, i) => {
    const a = (i * Math.PI * 2) / 24;
    const x1 = Math.cos(a) * 44, y1 = Math.sin(a) * 44;
    const x2 = Math.cos(a) * 49, y2 = Math.sin(a) * 49;
    return `<line x1="${x1.toFixed(2)}" y1="${y1.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y2.toFixed(2)}" />`;
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="-50 -50 100 100">
  <defs><clipPath id="c"><circle r="40"/></clipPath></defs>
  <g fill="none" stroke="${c}" stroke-opacity="${opacity}" stroke-linecap="round">
    <circle r="41" stroke-width="3.5"/>
    <g stroke-width="3" clip-path="url(#c)">${lines}</g>
    <g stroke-width="2.2">${rays}</g>
  </g></svg>`;
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Этномәдени концептілер арқылы оқу сауаттылығын дамыту";

  const shanGold = await svgPng(shanyrakSvg(C.gold));
  const shanGoldFaint = await svgPng(shanyrakSvg(C.gold, 0.35));
  const shanTeal = await svgPng(shanyrakSvg(C.teal, 0.9));

  const I = {};
  const iconList = {
    search: fa.FaSearch, book: fa.FaBookOpen, pen: fa.FaPenFancy,
    key: fa.FaKey, globe: fa.FaGlobeAsia, eye: fa.FaEye, brain: fa.FaLightbulb,
    balance: fa.FaBalanceScale, hand: fa.FaHandsHelping, users: fa.FaUsers,
    chat: fa.FaComments, check: fa.FaCheck, times: fa.FaTimes, flask: fa.FaFlask,
    chart: fa.FaChartLine, lang: fa.FaFont, feather: fa.FaFeatherAlt,
    target: fa.FaBullseye, list: fa.FaListUl, rocket: fa.FaRocket,
    phone: fa.FaMobileAlt, quote: fa.FaQuoteLeft, star: fa.FaStar,
  };
  for (const [k, v] of Object.entries(iconList)) {
    I[k] = await icon(v, C.white);
    I[k + "Teal"] = await icon(v, C.teal);
  }

  let n = 0;
  function contentSlide(title) {
    const s = pres.addSlide();
    n++;
    s.background = { color: C.white };
    s.addText(title, {
      x: 0.5, y: 0.3, w: 8.3, h: 0.75, fontFace: FONT, fontSize: 26, bold: true,
      color: C.dark, margin: 0, valign: "middle", isTextBox: true,
    });
    s.addImage({ data: shanTeal, x: 9.05, y: 0.35, w: 0.55, h: 0.55, transparency: 30 });
    s.addText(String(n), {
      x: 9.1, y: 5.2, w: 0.45, h: 0.3, fontFace: FONT, fontSize: 10, color: C.muted,
      align: "right", margin: 0, isTextBox: true,
    });
    return s;
  }
  function circleIcon(s, img, x, y, d, fill = C.teal) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function card(s, x, y, w, h, fill = C.tint) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12,
    });
  }
  const T = (s, text, o) => s.addText(text, { fontFace: FONT, color: C.text, isTextBox: true, ...o });

  // 1 — Title
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.dark };
    s.addImage({ data: shanGoldFaint, x: 5.9, y: 0.6, w: 4.6, h: 4.6 });
    T(s, "«Қазақ тілі мен әдебиеті: ұлттық код және заманауи білім беру кеңістігі»\nреспубликалық ғылыми-тәжірибелік онлайн конференциясы", {
      x: 0.6, y: 0.45, w: 6.2, h: 0.75, fontSize: 11, color: "B8C7D1", margin: 0, italic: true,
    });
    T(s, "Этномәдени концептілер арқылы қазақ тілі мен әдебиетінде оқу сауаттылығын дамыту", {
      x: 0.6, y: 1.45, w: 6.4, h: 2.0, fontSize: 32, bold: true, color: C.white, margin: 0, valign: "top",
    });
    T(s, "1-секция. Ұлттық құндылықтар мен этномәдени концептілерді оқыту, функционалдық оқу сауаттылығын дамыту", {
      x: 0.6, y: 3.75, w: 6.0, h: 0.6, fontSize: 12, color: C.gold, margin: 0,
    });
    T(s, [
      { text: "Тегі Аты Әкесінің аты", options: { bold: true, breakLine: true } },
      { text: "Жұмыс орнының толық атауы", options: { color: "B8C7D1" } },
    ], { x: 0.6, y: 4.6, w: 6, h: 0.6, fontSize: 13, color: C.white, margin: 0 });
  }

  // 2 — Problem
  {
    const s = contentSlide("Әдістемелік мәселе");
    T(s, "Оқушы дайын тұжырымдарды оңай қайталайды:", {
      x: 0.5, y: 1.25, w: 4.3, h: 0.4, fontSize: 15, bold: true, margin: 0,
    });
    const quotes = ["«Шаңырақ – қасиетті»", "«Үлкенді құрметтеу керек»"];
    quotes.forEach((q, i) => {
      const y = 1.8 + i * 0.85;
      card(s, 0.5, y, 4.3, 0.7, C.goldTint);
      s.addImage({ data: I.quoteTeal, x: 0.7, y: y + 0.2, w: 0.3, h: 0.3 });
      T(s, q, { x: 1.15, y, w: 3.5, h: 0.7, fontSize: 18, italic: true, valign: "middle", margin: 0 });
    });
    T(s, "Бірақ бұл түсінудің дәлелі емес.", {
      x: 0.5, y: 3.6, w: 4.3, h: 0.4, fontSize: 14, color: C.muted, margin: 0,
    });
    // Right: what must be checked
    card(s, 5.2, 1.25, 4.3, 3.85);
    T(s, "Бөлек тексеру керек:", { x: 5.5, y: 1.4, w: 3.8, h: 0.4, fontSize: 15, bold: true, color: C.teal, margin: 0 });
    const checks = [
      [I.book, "сөздің мәтіндегі мағынасын түсінуі"],
      [I.balance, "пікірін дәлелдей алуы"],
      [I.hand, "өмірлік жағдайда саналы шешім қабылдауы"],
    ];
    checks.forEach(([ic, t], i) => {
      const y = 1.95 + i * 0.72;
      circleIcon(s, ic, 5.5, y, 0.5);
      T(s, t, { x: 6.15, y, w: 3.2, h: 0.5, fontSize: 14, valign: "middle", margin: 0 });
    });
    T(s, "Этномәдени мазмұнды оқушының дербес оқу әрекетіне қалай айналдырамыз?", {
      x: 5.5, y: 4.2, w: 3.8, h: 0.75, fontSize: 13, bold: true, italic: true, color: C.dark, margin: 0,
    });
  }

  // 3 — Goal & methods
  {
    const s = contentSlide("Мақсат және зерттеу әдістері");
    card(s, 0.5, 1.25, 9, 1.05, C.dark);
    s.addImage({ data: I.target, x: 0.75, y: 1.52, w: 0.5, h: 0.5 });
    T(s, "Мақсат – этномәдени концептіні талдау жолын және оның орындалуын бағалау өлшемдерін ұсыну", {
      x: 1.5, y: 1.25, w: 7.8, h: 1.05, fontSize: 16, color: C.white, valign: "middle", margin: 0,
    });
    const m = [
      [I.book, "Дереккөздерді салыстыра оқу"],
      [I.search, "Ұғымның контекстік мағынасын ажырату"],
      [I.pen, "Оқу тапсырмасын жобалау"],
    ];
    m.forEach(([ic, t], i) => {
      const x = 0.5 + i * 3.1;
      card(s, x, 2.6, 2.8, 1.45);
      circleIcon(s, ic, x + 0.25, 2.8, 0.55);
      T(s, t, { x: x + 0.25, y: 3.4, w: 2.35, h: 0.6, fontSize: 13.5, bold: true, margin: 0, valign: "top" });
    });
    T(s, [
      { text: "Нысаны: ", options: { bold: true } },
      { text: "ұлттық мазмұндағы мәтінмен жұмыс. " },
      { text: "Сипаты: ", options: { bold: true } },
      { text: "теориялық-әдістемелік; эксперимент нәтижелері берілмейді, тапсырмалар – бейімдеуге арналған үлгілер." },
    ], { x: 0.5, y: 4.3, w: 9, h: 0.7, fontSize: 12.5, color: C.muted, margin: 0 });
  }

  // 4 — Key concepts
  {
    const s = contentSlide("Негізгі ұғымдар: жұмыс анықтамалары");
    const cards = [
      [I.key, "Ұлттық код", "Тілде, көркем мәтінде және күнделікті қарым-қатынаста көрінетін мәдени мағыналар мен құндылық бағдарларының сабақтастығы."],
      [I.globe, "Этномәдени концепт", "Белгілі бір сөздің айналасына жинақталған заттық, бейнелік және құндылықтық мағыналарды талдау бірлігі."],
    ];
    cards.forEach(([ic, h, t], i) => {
      const x = 0.5 + i * 4.6;
      card(s, x, 1.3, 4.4, 2.55);
      circleIcon(s, ic, x + 0.3, 1.55, 0.6);
      T(s, h, { x: x + 1.1, y: 1.55, w: 3.1, h: 0.6, fontSize: 19, bold: true, color: C.dark, valign: "middle", margin: 0 });
      T(s, t, { x: x + 0.3, y: 2.35, w: 3.85, h: 1.35, fontSize: 14, margin: 0, valign: "top" });
    });
    card(s, 0.5, 4.1, 9, 0.9, C.goldTint);
    T(s, "Сөзді жатқа түсіндіру – концептіні меңгерудің көрсеткіші емес: оқушы оның нақты мәтінде қандай қызмет атқарғанын ашуы қажет.", {
      x: 0.75, y: 4.1, w: 8.5, h: 0.9, fontSize: 13.5, italic: true, valign: "middle", margin: 0,
    });
  }

  // 5 — PISA framework
  {
    const s = contentSlide("Теориялық негіз: PISA оқу сауаттылығы");
    T(s, "ЭЫДҰ-ның PISA тұжырымдамасында мәтінді түсіну, пайдалану, бағалау және ол туралы ой қорыту өзара байланысты әрекеттер ретінде қарастырылады [1].", {
      x: 0.5, y: 1.2, w: 9, h: 0.7, fontSize: 14, color: C.muted, margin: 0,
    });
    const steps = [
      [I.search, "Ақпаратты табу", "Мәтіннен қажетті деректі дәл анықтау"],
      [I.brain, "Мағынаны түсіну", "Ұғымды контексте түсіндіру"],
      [I.balance, "Бағалау мен рефлексия", "Дерек сапасын бағалап, пайым жасау"],
    ];
    steps.forEach(([ic, h, t], i) => {
      const x = 0.5 + i * 3.1;
      card(s, x, 2.15, 2.75, 2.0);
      circleIcon(s, ic, x + 1.02, 2.35, 0.7, i === 2 ? C.gold : C.teal);
      T(s, h, { x: x + 0.15, y: 3.15, w: 2.45, h: 0.4, fontSize: 15, bold: true, align: "center", margin: 0 });
      T(s, t, { x: x + 0.15, y: 3.55, w: 2.45, h: 0.5, fontSize: 12, color: C.muted, align: "center", margin: 0 });
      if (i < 2) {
        s.addShape(pres.shapes.CHEVRON, {
          x: x + 2.83, y: 3.0, w: 0.2, h: 0.3, fill: { color: C.gold }, line: { color: C.gold }, rotate: 0,
        });
      }
    });
    T(s, "Бұл жіктеу тапсырма жобалауға негіз етіледі; тапсырмалар ресми PISA тапсырмалары немесе өлшеу шкаласы емес.", {
      x: 0.5, y: 4.45, w: 9, h: 0.55, fontSize: 12, italic: true, color: C.muted, margin: 0,
    });
  }

  // 6 — Three supports
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.dark };
    s.addImage({ data: shanGold, x: 9.05, y: 0.35, w: 0.55, h: 0.55, transparency: 20 });
    T(s, "Әдістемелік ерекшелік: үш тірек", {
      x: 0.5, y: 0.3, w: 8.3, h: 0.75, fontSize: 26, bold: true, color: C.white, margin: 0, valign: "middle",
    });
    const st = [
      ["1", "Мәтіндегі белгі", "Ұғымды танытатын әрекет, сөз, деталь"],
      ["2", "Қорытынды", "Белгіден шығатын мағына, түсіндірме"],
      ["3", "Әрекет", "Бүгінгі ортада қалай жүзеге асатыны"],
    ];
    st.forEach(([num, h, t], i) => {
      const x = 0.5 + i * 3.1;
      s.addShape(pres.shapes.OVAL, { x: x + 0.95, y: 1.35, w: 0.9, h: 0.9, fill: { color: C.gold }, line: { color: C.gold } });
      T(s, num, { x: x + 0.95, y: 1.35, w: 0.9, h: 0.9, fontSize: 30, bold: true, color: C.dark, align: "center", valign: "middle", margin: 0 });
      T(s, h, { x, y: 2.4, w: 2.8, h: 0.45, fontSize: 18, bold: true, color: C.white, align: "center", margin: 0 });
      T(s, t, { x, y: 2.85, w: 2.8, h: 0.6, fontSize: 13, color: "B8C7D1", align: "center", margin: 0 });
      if (i < 2) {
        s.addShape(pres.shapes.LINE, { x: x + 2.05, y: 1.8, w: 1.8, h: 0, line: { color: C.gold, width: 1.5, dashType: "dash" } });
      }
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.8, w: 9, h: 1.2, fill: { color: "1D4863" }, line: { color: "1D4863" }, rectRadius: 0.12 });
    T(s, [
      { text: "Мысалы, «қонақжайлық»: ", options: { bold: true, color: C.gold } },
      { text: "атауды айтумен шектелмей, оны танытатын әрекетті көрсету, мағынасын түсіндіру және бүгін қалай жүзеге асатынын ұсыну. Байланыс үзілген жер – мұғалімге келесі тапсырманың бағыты." },
    ], { x: 0.8, y: 3.8, w: 8.5, h: 1.2, fontSize: 13.5, color: C.white, valign: "middle", margin: 0 });
  }

  // 7 — Shanyrak: materials & story
  {
    const s = contentSlide("«Шаңырақ» концептісі: 6–7-сынып");
    card(s, 0.5, 1.25, 3.3, 3.8);
    s.addImage({ data: shanGold, x: 1.55, y: 1.45, w: 1.2, h: 1.2 });
    T(s, "Материалдар", { x: 0.75, y: 2.8, w: 2.8, h: 0.35, fontSize: 15, bold: true, color: C.dark, margin: 0 });
    T(s, [
      { text: "Қысқа көркем жағдаят", options: { bullet: true, breakLine: true } },
      { text: "Киіз үй туралы танымдық мәтін", options: { bullet: true, breakLine: true } },
      { text: "UNESCO: киіз үй – отбасы мен қонақжайлықтың нышаны [2]", options: { bullet: true } },
    ], { x: 0.75, y: 3.2, w: 2.9, h: 1.7, fontSize: 12.5, margin: 0, paraSpaceAfter: 4, valign: "top" });

    card(s, 4.1, 1.25, 5.4, 3.8, C.goldTint);
    s.addImage({ data: I.quoteTeal, x: 4.35, y: 1.45, w: 0.35, h: 0.35 });
    T(s, "Авторлық шағын мәтін", { x: 4.85, y: 1.43, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: C.teal, margin: 0, valign: "middle" });
    T(s, "Мектептегі мұра көрмесіне Айша атасының шаңырақ үлгісін әкелді. Ол әуелі ағаш бөлшектердің атауын ғана жазды. Атасы: «Мұны бірге көтерген адамдарды да ұмытпа», – деді. Айша сыныптастарына көрмені бірге әзірлеуді ұсынды. Бірі түсіндірме мәтінін жазды, екіншісі келушілерге жол көрсетті. Көрме соңында Айша шаңырақ үлгісінің жанына осы іске қатысқан адамдардың атын қосты.", {
      x: 4.35, y: 1.95, w: 4.95, h: 2.6, fontSize: 12.5, italic: true, margin: 0, valign: "top", lineSpacingMultiple: 1.05,
    });
    T(s, "Дереккөздегі нақты ой мен оқушының жорамалы ажыратылады.", {
      x: 4.35, y: 4.55, w: 4.95, h: 0.4, fontSize: 11, color: C.muted, margin: 0,
    });
  }

  // 8 — Shanyrak tasks
  {
    const s = contentSlide("«Шаңырақ»: тапсырмалар тізбегі");
    const tasks = [
      [I.search, "Табу", "Айшаның бастапқы және кейінгі әрекетін табу."],
      [I.brain, "Түсіндіру", "«Атасының сөзі көрменің мазмұнын қалай өзгертті? Мәтіндегі екі әрекетпен негізде»."],
      [I.balance, "Бағалау", "Танымдық дерек пен көркем жағдаятты салыстыру: «Бұл деректер барлық отбасы туралы қорытындыға жеткілікті ме?»"],
      [I.pen, "Қолдану", "Төменгі сынып оқушысына көрме туралы таныстыру мәтінін жазу: дерек + дәлел + нақты шақыру."],
    ];
    tasks.forEach(([ic, h, t], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.5 + col * 4.6, y = 1.25 + row * 1.6;
      card(s, x, y, 4.4, 1.4);
      circleIcon(s, ic, x + 0.25, y + 0.25, 0.55, i === 3 ? C.gold : C.teal);
      T(s, `${i + 1}. ${h}`, { x: x + 1.0, y: y + 0.2, w: 3.2, h: 0.35, fontSize: 15, bold: true, color: C.dark, margin: 0 });
      T(s, t, { x: x + 1.0, y: y + 0.55, w: 3.25, h: 0.8, fontSize: 11.5, margin: 0, valign: "top" });
    });
    T(s, [
      { text: "Жалпы үндеудің орнына: ", options: { bold: true } },
      { text: "«ұлттық мұраны құрметтейік» емес, «отбасыңдағы бір бұйымның тарихын үлкендерден сұрап, дерек иесін көрсетіп жаз»." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.55, fontSize: 12, color: C.muted, italic: true, margin: 0 });
  }

  // 9 — Uyat: Abai + situation
  {
    const s = contentSlide("«Ұят» концептісі: Абайдың 36-қара сөзі");
    card(s, 0.5, 1.25, 4.2, 3.8);
    circleIcon(s, I.feather, 0.75, 1.45, 0.6);
    T(s, "Абай ажыратады [3]", { x: 1.5, y: 1.45, w: 3.0, h: 0.6, fontSize: 15, bold: true, color: C.dark, valign: "middle", margin: 0 });
    T(s, [
      { text: "Орынсыз қысылу", options: { bullet: true, breakLine: true } },
      { text: "Адамның өз әрекетіне адамгершілік тұрғысынан есеп беруі", options: { bullet: true, breakLine: true } },
      { text: "Қателігін сезінген адамға қайта ауыр сөз айтуды сынайды", options: { bullet: true } },
    ], { x: 0.75, y: 2.25, w: 3.75, h: 1.9, fontSize: 13, margin: 0, paraSpaceAfter: 6, valign: "top" });
    T(s, "Оқушыдан жеке ұялған оқиғасын айту талап етілмейді – талдау ойдан құрастырылған жағдаятқа сүйенеді.", {
      x: 0.75, y: 4.15, w: 3.75, h: 0.8, fontSize: 11, italic: true, color: C.muted, margin: 0, valign: "top",
    });

    card(s, 5.0, 1.25, 4.5, 3.8, C.goldTint);
    circleIcon(s, I.phone, 5.25, 1.45, 0.6, C.gold);
    T(s, "Жағдаят: 8–9-сынып", { x: 6.0, y: 1.45, w: 3.3, h: 0.6, fontSize: 15, bold: true, color: C.dark, valign: "middle", margin: 0 });
    T(s, "Сыныптың ортақ чатында оқушының сәтсіз жауабы түсірілген бейне тарады. Бір оқушы оны күлкілі деп қайта жіберді. Кейін ол сыныптасының ренжігенін білді. Достары хабарламаны өшірсе жеткілікті екенін айтты.", {
      x: 5.25, y: 2.2, w: 4.0, h: 1.7, fontSize: 12.5, italic: true, margin: 0, valign: "top",
    });
    T(s, [
      { text: "? ", options: { bold: true, color: C.teal } },
      { text: "Абай мәтініндегі қай ой бағалауға негіз болады?", options: { breakLine: true } },
      { text: "? ", options: { bold: true, color: C.teal } },
      { text: "Қатені түзету үшін қандай қадам керек?" },
    ], { x: 5.25, y: 3.95, w: 4.0, h: 0.95, fontSize: 12, bold: true, margin: 0, valign: "top" });
  }

  // 10 — Uyat discussion
  {
    const s = contentSlide("«Ұят»: талқылау және қауіпсіз орта");
    T(s, "Салыстырылатын шешімдер", { x: 0.5, y: 1.2, w: 5, h: 0.4, fontSize: 15, bold: true, color: C.teal, margin: 0 });
    const sol = [
      [I.times, "Хабарламаны өшіру"],
      [I.hand, "Кешірім сұрау"],
      [I.users, "Бейнені әрі қарай таратпау туралы ескерту"],
    ];
    sol.forEach(([ic, t], i) => {
      const y = 1.75 + i * 0.75;
      card(s, 0.5, y, 4.6, 0.6);
      circleIcon(s, ic, 0.65, y + 0.1, 0.4);
      T(s, t, { x: 1.2, y, w: 3.8, h: 0.6, fontSize: 13.5, valign: "middle", margin: 0 });
    });
    card(s, 0.5, 4.05, 4.6, 1.0, C.goldTint);
    T(s, [
      { text: "Бағалау өлшемі: ", options: { bold: true } },
      { text: "ұсыныс зардапты қаншалықты азайтады және мәтінмен қалай байланысады." },
    ], { x: 0.7, y: 4.05, w: 4.25, h: 1.0, fontSize: 12.5, valign: "middle", margin: 0 });

    card(s, 5.4, 1.2, 4.1, 3.85, C.dark);
    T(s, "Мұғалім ұстанымы", { x: 5.7, y: 1.4, w: 3.6, h: 0.45, fontSize: 16, bold: true, color: C.gold, margin: 0 });
    T(s, [
      { text: "Баланың жеке басына моральдық баға бермейді", options: { bullet: true, breakLine: true } },
      { text: "Көпшілік алдында ұялтуды тәрбие тәсілі ретінде ұсынбайды", options: { bullet: true, breakLine: true } },
      { text: "Оқушы пікірін еркін негіздеуге мүмкіндік алады", options: { bullet: true } },
    ], { x: 5.7, y: 2.0, w: 3.6, h: 2.8, fontSize: 13.5, color: C.white, margin: 0, paraSpaceAfter: 10, valign: "top" });
  }

  // 11 — Language vs Literature
  {
    const s = contentSlide("Бір мәтін – екі пән");
    const cols = [
      [I.lang, "Қазақ тілі", ["Дәлелді абзац жазу", "Себеп-салдар байланысын жеткізу", "Сөздің контекстік мағынасын ажырату"]],
      [I.book, "Қазақ әдебиеті", ["Автордың пайымдауын талдау", "Мысал таңдауының қызметі", "Оқырманға бағытталған сұрақтың рөлі"]],
    ];
    cols.forEach(([ic, h, items], i) => {
      const x = 0.5 + i * 4.6;
      card(s, x, 1.25, 4.4, 2.75);
      circleIcon(s, ic, x + 0.3, 1.5, 0.65, i ? C.gold : C.teal);
      T(s, h, { x: x + 1.15, y: 1.5, w: 3.0, h: 0.65, fontSize: 19, bold: true, color: C.dark, valign: "middle", margin: 0 });
      T(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), {
        x: x + 0.3, y: 2.4, w: 3.9, h: 1.7, fontSize: 14, margin: 0, paraSpaceAfter: 8, valign: "top",
      });
    });
    T(s, "Тапсырманың көлемі мен күрделілігі мұғалім таңдаған оқу мақсатына сәйкес нақтыланады.", {
      x: 0.5, y: 4.3, w: 9, h: 0.5, fontSize: 13, italic: true, color: C.muted, margin: 0,
    });
  }

  // 12 — Assessment rubric
  {
    const s = contentSlide("Бағалау өлшемдері");
    const crit = [
      "Қажетті ақпаратты дәл табуы",
      "Ұғымды контексте түсіндіруі",
      "Тұжырымын мәтіндік дәлелмен негіздеуі",
      "Жаңа жағдаятқа сәйкес шешім ұсынуы",
    ];
    crit.forEach((t, i) => {
      const y = 1.25 + i * 0.68;
      card(s, 0.5, y, 5.3, 0.55);
      s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.08, w: 0.39, h: 0.39, fill: { color: C.teal }, line: { color: C.teal } });
      T(s, String(i + 1), { x: 0.62, y: y + 0.08, w: 0.39, h: 0.39, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
      T(s, t, { x: 1.2, y, w: 3.8, h: 0.55, fontSize: 13.5, valign: "middle", margin: 0 });
      T(s, "0–2", { x: 4.95, y, w: 0.7, h: 0.55, fontSize: 13, bold: true, color: C.teal, align: "right", valign: "middle", margin: 0 });
    });
    const sc = [["0", "орындалмаған немесе мәтінге қайшы"], ["1", "жартылай, байланысы толық ашылмаған"], ["2", "дұрыс әрі жеткілікті негізделген"]];
    T(s, sc.map(([b, t], i) => [
      { text: b + " – ", options: { bold: true, color: C.dark } },
      { text: t, options: { breakLine: i < 2 } },
    ]).flat(), { x: 0.5, y: 4.05, w: 5.3, h: 0.95, fontSize: 12, color: C.muted, margin: 0, valign: "top" });

    card(s, 6.2, 1.25, 3.3, 3.75, C.dark);
    T(s, "8", { x: 6.2, y: 1.4, w: 3.3, h: 1.5, fontSize: 88, bold: true, color: C.gold, align: "center", valign: "middle", margin: 0 });
    T(s, "балл – ең жоғары нәтиже", { x: 6.4, y: 2.9, w: 2.9, h: 0.4, fontSize: 14, bold: true, color: C.white, align: "center", margin: 0 });
    T(s, "Тек осы қалыптастырушы бағалау құралына қатысты; ұлттық құндылықты меңгерудің әмбебап өлшемі емес.", {
      x: 6.45, y: 3.45, w: 2.8, h: 1.4, fontSize: 11.5, color: "B8C7D1", align: "center", margin: 0, valign: "top",
    });
  }

  // 13 — Answer examples & feedback
  {
    const s = contentSlide("Жауап үлгілері және кері байланыс");
    card(s, 0.5, 1.25, 4.4, 2.2);
    circleIcon(s, I.times, 0.75, 1.45, 0.5, "B5543C");
    T(s, "Қорытынды бар, дәлел жоқ", { x: 1.4, y: 1.45, w: 3.4, h: 0.5, fontSize: 14, bold: true, color: "B5543C", valign: "middle", margin: 0 });
    T(s, "«Айша бірлікті түсінді».", { x: 0.75, y: 2.15, w: 3.95, h: 1.1, fontSize: 15, italic: true, margin: 0, valign: "top" });

    card(s, 5.1, 1.25, 4.4, 2.2);
    circleIcon(s, I.check, 5.35, 1.45, 0.5, C.teal);
    T(s, "Тұжырым + мәтіндік негіз", { x: 6.0, y: 1.45, w: 3.4, h: 0.5, fontSize: 14, bold: true, color: C.teal, valign: "middle", margin: 0 });
    T(s, "«Айша сыныптастарына міндет бөліп, кейін олардың атын жазды; демек, көрмені ортақ еңбектің нәтижесі ретінде танытты».", {
      x: 5.35, y: 2.05, w: 3.95, h: 1.3, fontSize: 12.5, italic: true, margin: 0, valign: "top",
    });

    card(s, 0.5, 3.7, 9, 1.35, C.goldTint);
    circleIcon(s, I.chat, 0.75, 3.95, 0.55, C.gold);
    T(s, [
      { text: "Мұғалімнің нақты кері байланысы: ", options: { bold: true, breakLine: true } },
      { text: "«Екі әрекетті дұрыс таңдадың, енді олардың шаңырақтың ауыспалы мағынасын қалай ашатынын түсіндір».", options: { italic: true } },
    ], { x: 1.5, y: 3.7, w: 7.8, h: 1.35, fontSize: 13, valign: "middle", margin: 0 });
  }

  // 14 — Differentiation
  {
    const s = contentSlide("Саралап оқыту");
    const cols = [
      [I.hand, C.teal, "Қолдауды қажет ететін оқушы", [
        "Мәтіннің қысқартылған нұсқасы",
        "Тірек сөздер",
        "Сөйлем үлгісі: «Менің ойымша… Оған мәтіндегі… дәлел»",
        "Дәлелді бірнеше нұсқадан таңдап, өз сөзімен түсіндіру",
      ]],
      [I.rocket, C.gold, "Дайындығы жоғары оқушы", [
        "Қарсы пайым ұсыну",
        "Оны тексеруге қандай қосымша дерек керегін жазу",
      ]],
    ];
    cols.forEach(([ic, col, h, items], i) => {
      const x = 0.5 + i * 4.6;
      card(s, x, 1.25, 4.4, 2.95);
      circleIcon(s, ic, x + 0.3, 1.45, 0.55, col);
      T(s, h, { x: x + 1.05, y: 1.45, w: 3.2, h: 0.55, fontSize: 15, bold: true, color: C.dark, valign: "middle", margin: 0 });
      T(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), {
        x: x + 0.3, y: 2.2, w: 3.9, h: 1.9, fontSize: 13, margin: 0, paraSpaceAfter: 6, valign: "top",
      });
    });
    card(s, 0.5, 4.4, 9, 0.65, C.dark);
    T(s, "Ортақ талап: пікірдің мәтінге қатысы анық болуы керек.", {
      x: 0.75, y: 4.4, w: 8.5, h: 0.65, fontSize: 14, bold: true, color: C.white, valign: "middle", margin: 0,
    });
  }

  // 15 — Empirical verification
  {
    const s = contentSlide("Тәжірибеде тексеру жолы");
    const steps = [
      [I.book, "Сабақ басында", "Шағын мәтін бойынша жұмыс"],
      [I.feather, "Сабақ соңында", "Мазмұны бөлек, күрделілігі ұқсас мәтін"],
      [I.list, "Бірдей өлшем", "4 критерий × 0–2 балл"],
      [I.chart, "Талдау", "Қай әрекетте ілгерілеу бар?"],
    ];
    steps.forEach(([ic, h, t], i) => {
      const x = 0.5 + i * 2.3;
      circleIcon(s, ic, x + 0.65, 1.3, 0.75, i === 3 ? C.gold : C.teal);
      if (i < 3) s.addShape(pres.shapes.LINE, { x: x + 1.55, y: 1.675, w: 1.25, h: 0, line: { color: C.gold, width: 1.5, dashType: "dash" } });
      T(s, h, { x, y: 2.2, w: 2.05, h: 0.4, fontSize: 14, bold: true, align: "center", margin: 0 });
      T(s, t, { x, y: 2.6, w: 2.05, h: 0.6, fontSize: 12, color: C.muted, align: "center", margin: 0, valign: "top" });
    });
    card(s, 0.5, 3.45, 4.4, 1.6);
    T(s, [
      { text: "Ескерілетін ықпалдар", options: { bold: true, color: C.teal, breakLine: true } },
      { text: "қайталап орындау, мұғалім көмегі, мәтінге таныстық" },
    ], { x: 0.75, y: 3.45, w: 4.0, h: 1.6, fontSize: 13, valign: "middle", margin: 0 });
    card(s, 5.1, 3.45, 4.4, 1.6, C.goldTint);
    T(s, [
      { text: "Жариялауда көрсетіледі", options: { bold: true, color: C.dark, breakLine: true } },
      { text: "қатысушылар саны, тапсырмалар, бағалау тәртібі, зерттеу шектеулері" },
    ], { x: 5.35, y: 3.45, w: 4.0, h: 1.6, fontSize: 13, valign: "middle", margin: 0 });
  }

  // 16 — Conclusion
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.dark };
    s.addImage({ data: shanGoldFaint, x: 6.9, y: 1.1, w: 4.0, h: 4.0 });
    T(s, "Қорытынды", { x: 0.5, y: 0.35, w: 6, h: 0.75, fontSize: 30, bold: true, color: C.white, margin: 0, valign: "middle" });
    const pts = [
      ["«Шаңырақ»", "мағына мен дереккөзді салыстыруға мүмкіндік береді"],
      ["«Ұят»", "әрекет пен жауапкершілікті пайымдауға жетелейді"],
      ["Нәтиже", "құндылық атауын қайталау емес – мәтіндегі белгіні түсіндіріп, қорытындыны негіздеу"],
    ];
    pts.forEach(([h, t], i) => {
      const y = 1.35 + i * 0.95;
      s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.12, w: 0.22, h: 0.22, fill: { color: C.gold }, line: { color: C.gold } });
      T(s, [
        { text: h + " – ", options: { bold: true, color: C.gold } },
        { text: t },
      ], { x: 0.95, y, w: 5.6, h: 0.8, fontSize: 15, color: C.white, margin: 0, valign: "top" });
    });
    T(s, "Ұлттық мұраны таныту мен саналы оқу бір тапсырманың өзара сабақтас міндеттеріне айналады.", {
      x: 0.5, y: 4.3, w: 6.2, h: 0.8, fontSize: 15, italic: true, color: "B8C7D1", margin: 0,
    });
  }

  // 17 — References
  {
    const s = contentSlide("Пайдаланылған әдебиеттер");
    const refs = [
      ["OECD. PISA-2025.", "https://www.oecd.org/kk/publications/pisa-2025-results-volume-i_127aae7d-kk/dafefb09-kk.html"],
      ["UNESCO Multimedia Archives. Traditional Knowledge and Skills in Making Kyrgyz and Kazakh Yurts.", "https://www.unesco.org/archives/multimedia/document-3669"],
      ["Абай Құнанбайұлы. Отыз алтыншы сөз // Абай институты.", "https://www.abai.institute/qara-sozder/otyz-altynsy-soz"],
    ];
    refs.forEach(([t, u], i) => {
      const y = 1.3 + i * 1.15;
      card(s, 0.5, y, 9, 0.95);
      s.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.25, w: 0.45, h: 0.45, fill: { color: C.teal }, line: { color: C.teal } });
      T(s, String(i + 1), { x: 0.7, y: y + 0.25, w: 0.45, h: 0.45, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
      T(s, [
        { text: t, options: { breakLine: true } },
        { text: u, options: { fontSize: 10, color: C.teal, hyperlink: { url: u } } },
      ], { x: 1.4, y, w: 7.9, h: 0.95, fontSize: 13, valign: "middle", margin: 0 });
    });
    T(s, "Қаралған күні: 26.09.2026", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.muted, margin: 0 });
  }

  await pres.writeFile({ fileName: "presentation.pptx" });
  console.log("done");
})();
