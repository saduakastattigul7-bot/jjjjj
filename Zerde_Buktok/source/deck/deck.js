// Қорғауға арналған презентация: PPTX_SKILL=<pptx skill> node deck.js ../../Qorgau_prezentatsiyasy.pptx
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
const N = D.natije;
if (!N || !N.n_resp) throw new Error("Сауалнама деректері жоқ – алдымен build.sh іске қосыңыз");
const SV = N.survey;
const f1 = v => String(Math.round(v * 10) / 10).replace(".", ",");
const pc = v => (v === null || v === undefined) ? "–" : f1(v) + "%";
const big = n => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
const NUMW = n => { n = Math.abs(Math.round(n)); return n === 0 ? "нөл" : n % 10 ? ["", "бір", "екі", "үш", "төрт", "бес", "алты", "жеті", "сегіз", "тоғыз"][n % 10]
  : n % 100 ? ["", "он", "жиырма", "отыз", "қырық", "елу", "алпыс", "жетпіс", "сексен", "тоқсан"][(n % 100) / 10] : n % 1000 ? "жүз" : "мың"; };
const px = n => { const w = NUMW(n), b = /[аоұы][^аоұыәеіөүи]*$/.test(w), v = /[аоұыәеіөүиу]$/.test(w);
  return v ? (b ? "сы" : "сі") : (b ? "ы" : "і"); };
const listJoin = a => a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + " және " + a[a.length - 1];

// Түнгі оқу: қою күрең (кітап), қызғылт (TikTok), көгілдір, сары (шам жарығы)
const THEME = {
  name: "Zerde Buktok",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "231B26", lt1: "FFFFFF", dk2: "3D1A3F", lt2: "F5F0F6",
    accent1: "E0476F", accent2: "1A9E9E", accent3: "E8A33A", accent4: "2E9E5B",
    accent5: "6E6475", accent6: "7B4FA0", hlink: "1A9E9E", folHlink: "7B4FA0",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Буктюб пен буктокерлер: әлеуметтік желі жастардың кітап оқуына әсер ете ме?";
const C = pres.SchemeColor;

// ---------- Макеттер ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.35, w: 7.6, h: 3.4, fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 5.0, w: 7.6, h: 0.9, fontSize: 18, italic: true, color: C.accent3, valign: "top", align: "left", margin: 0 }, text: "" } },
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
// Мотив: түрлі түсті дөңгелектегі иконка
async function iconCircle(s, Comp, x, y, d, fill, name) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + "-circle" });
  const pad = d * 0.24;
  s.addImage({ data: await icon(Comp), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: name + "-icon" });
}
const txt = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, ...o });
function card(s, x, y, w, h, fill, name, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill },
    line: line ? { color: line, width: 1.25 } : { type: "none" }, objectName: name });
}
function bigStat(s, x, y, w, value, label, color, name, size = 54) {
  txt(s, value, { x, y, w, h: 0.95, fontSize: size, bold: true, color, fontFace: THEME.headFontFace, objectName: name + "-value" });
  txt(s, label, { x, y: y + 1.0, w, h: 0.9, fontSize: 15, color: C.text1, valign: "top", objectName: name + "-label" });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };
const chartQuiet = () => ({
  ...chartText, catAxisLabelColor: HEX.dk1, valAxisLabelColor: HEX.accent5, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  dataLabelFontSize: 13, dataLabelColor: HEX.dk1, valGridLine: { color: "E6DFE8", size: 0.75 }, catGridLine: { style: "none" },
  legendFontSize: 13, legendColor: HEX.dk1,
});

// ---------- Деректер ----------
const srcN = N.src_n || {};
const srcSorted = Object.entries(srcN).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
const topSrc = srcSorted.length ? srcSorted[0][0] : null;
const notFin = (SV.FINISH["Жартылай"] || 0) + (SV.FINISH["Жоқ"] || 0);
const soc = N.src["Әлеуметтік желі"] || 0;
const okA = Object.entries(N.src).every(([k, v]) => k === "Әлеуметтік желі" || (v || 0) < soc);
const okB = SV.FINISH.n ? notFin >= 25 : null;
const okC = (SV.KAZ["Иә"] || 0) < 50;
const nOk = [okA, okB, okC].filter(Boolean).length;
const hypShort = nOk === 3 ? "расталды" : nOk === 0 ? "расталмады" : "ішінара расталды";
const BL = N.bloggers || [];
const blLit = BL.filter(x => x.kaz_lit === "Иә").length;
const netTop = Object.entries(SV.NET).filter(([k]) => k !== "n").sort((a, b) => (b[1] || 0) - (a[1] || 0))[0];
const IV = N.interviews || [];
const byWho = w => IV.filter(x => x.who === w);
const shortName = n => n.split(/\s+@|\s+(?=[А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+\s+[А-ЯӘҒҚҢӨҰҮҺІ])/)[0].slice(0, 26);

(async () => {
  let s;
  // ================= 1. Титул =================
  pres.addSection({ title: "Кіріспе" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Кіріспе" });
  txt(s, "«Зерде» республикалық конкурсы  ·  Қазақ тілі және әдебиеті  ·  «Орта буын»", { x: 0.8, y: 0.6, w: 11.5, h: 0.4, fontSize: 14, color: C.accent3, objectName: "kicker" });
  s.addText("Буктюб пен буктокерлер: әлеуметтік желі жастардың кітап оқуына әсер ете ме?", { placeholder: "title" });
  s.addText(`${N.n_resp} оқушыға сауалнама · ${BL.length} қазақ тіліндегі кітап блогері · ${IV.length} сұхбат`, { placeholder: "body" });
  txt(s, "Тіркеу коды (шифр): ______________", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.background2, objectName: "cipher" });
  const tIcons = [[fa.FaTiktok, C.accent1], [fa.FaBookOpen, C.accent3], [fa.FaYoutube, C.accent2], [fa.FaInstagram, C.accent6]];
  for (let i = 0; i < 4; i++)
    await iconCircle(s, tIcons[i][0], 9.0 + (i % 2) * 1.85, 1.7 + Math.floor(i / 2) * 1.85, 1.5, tIcons[i][1], "title-ic-" + i);
  s.addNotes("Құрметті қазылар алқасы! Менің жобамның тақырыбы – «Буктюб пен буктокерлер: әлеуметтік желі жастардың кітап оқуына әсер ете ме?». Мен құрдастарым арасында сауалнама жүргіздім, қазақ тіліндегі кітап блогерлерін талдадым және кітапханашымен, мұғалімдермен, оқушылармен сұхбаттастым. (15 секунд)");

  // ================= 2. Мәселе =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Әлеуметтік желі: кітаптың жауы ма, досы ма?", { placeholder: "title" });
  card(s, 0.6, 1.55, 5.85, 3.3, C.background2, "hook-l");
  await iconCircle(s, fa.FaGlobeEurope, 0.95, 1.85, 0.75, C.accent2, "hook-l-ic");
  txt(s, "Әлемде", { x: 1.9, y: 1.85, w: 4.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "hook-l-t" });
  txt(s, [
    { text: "370 млрд+", options: { bold: true, color: HEX.accent1, fontSize: 30, breakLine: true } },
    { text: "#BookTok бейнелерінің қаралымы (Forbes, 2025)", options: { breakLine: true } },
    { text: "50 млн+", options: { bold: true, color: HEX.accent1, fontSize: 30, breakLine: true } },
    { text: "BookTok ұсынған кітап Еуропаның 6 елінде сатылды (TikTok, 2026)" },
  ], { x: 0.95, y: 2.75, w: 5.2, h: 2.0, fontSize: 15, color: C.text1, valign: "top", objectName: "hook-l-text" });
  card(s, 6.85, 1.55, 5.85, 3.3, C.background2, "hook-r");
  await iconCircle(s, fa.FaUserFriends, 7.2, 1.85, 0.75, C.accent3, "hook-r-ic");
  txt(s, "Мектепте", { x: 8.15, y: 1.85, w: 4.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "hook-r-t" });
  txt(s, "«Әлеуметтік желі баланы кітаптан алыстатады»", { x: 7.2, y: 2.8, w: 5.2, h: 1.2, fontSize: 22, italic: true, color: C.text1, valign: "top", objectName: "hook-r-text" });
  txt(s, "ата-аналар мен мұғалімдердің жиі айтатын пікірі", { x: 7.2, y: 4.15, w: 5.2, h: 0.45, fontSize: 14, color: C.accent5, objectName: "hook-r-src" });
  card(s, 0.6, 5.2, 12.1, 1.35, C.text2, "hook-q");
  txt(s, "Бұл қайшылық біздің мектеп оқушыларына қаншалықты қатысты?",
    { x: 1.0, y: 5.2, w: 11.3, h: 1.35, fontSize: 24, bold: true, color: C.background1, valign: "middle", objectName: "hook-q-text" });
  s.addNotes("Әлемде #BookTok бейнелері 370 миллиардтан астам рет қаралған, ал 2025 жылы Еуропаның алты елінде BookTok ұсынған 50 миллионнан астам кітап сатылған. Ал мектепте біз жиі «әлеуметтік желі баланы кітаптан алыстатады» деген пікірді естиміз. Мен осы қайшылықтың біздің мектеп оқушыларына қаншалықты қатысты екенін тексергім келді. (30 секунд)");

  // ================= 3. Сұрақ, мақсат, болжам =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеу сұрағы мен болжам", { placeholder: "title" });
  card(s, 0.6, 1.55, 12.1, 1.45, C.background2, "q-card");
  await iconCircle(s, fa.FaQuestion, 0.95, 1.85, 0.85, C.accent1, "q");
  txt(s, "Кітап бейнелері 7–9 сынып оқушыларының кітап таңдауы мен оқуына әсер ете ме және қазақ тіліндегі кітап блогерлері оларға қаншалықты таныс?",
    { x: 2.1, y: 1.55, w: 10.3, h: 1.45, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "q-text" });
  const hyp = [
    [fa.FaMobileAlt, "1", "Оқушылар жаңа кітап туралы көбіне әлеуметтік желіден біледі", C.accent1],
    [fa.FaBookmark, "2", "Блогер ұсынған кітапты оқығандардың кемінде төрттен бірі оны аяқтамайды", C.accent3],
    [fa.FaLanguage, "3", "Қазақ тіліндегі кітап блогерін оқушылардың жартысынан азы көреді", C.accent2],
  ];
  txt(s, "Болжам", { x: 0.6, y: 3.3, w: 4, h: 0.5, fontSize: 20, bold: true, color: C.text2, objectName: "hyp-h" });
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 3.9, 3.75, 2.75, C.background1, "hyp-card-" + i, "E0D6E3");
    await iconCircle(s, hyp[i][0], x + 0.3, 4.15, 0.75, hyp[i][3], "hyp-" + i);
    txt(s, hyp[i][1], { x: x + 1.2, y: 4.15, w: 2.3, h: 0.75, fontSize: 28, bold: true, color: hyp[i][3], fontFace: THEME.headFontFace, valign: "middle", objectName: "hyp-n-" + i });
    txt(s, hyp[i][2], { x: x + 0.3, y: 5.05, w: 3.2, h: 1.5, fontSize: 16, color: C.text1, valign: "top", objectName: "hyp-text-" + i });
  }
  s.addNotes("Менің зерттеу сұрағым: кітап бейнелері оқушылардың кітап таңдауы мен оқуына әсер ете ме және қазақ тіліндегі кітап блогерлері оларға қаншалықты таныс? Болжамым үш бөліктен тұрды. Бірінші: оқушылар жаңа кітап туралы көбіне әлеуметтік желіден біледі. Екінші: блогер ұсынған кітапты оқығандардың кемінде төрттен бірі оны аяқтамайды. Үшінші: қазақ тіліндегі кітап блогерін оқушылардың жартысынан азы көреді. (35 секунд)");

  // ================= 4. Ұғымдар =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Кітап туралы бейненің үш алаңы", { placeholder: "title" });
  const kinds = [
    [fa.FaYoutube, "Буктюб", "YouTube", "Толық шолу, талдау, «оқылған кітаптар»", "10–40 минут", C.accent1],
    [fa.FaTiktok, "Буктокер", "TikTok, #BookTok", "Қысқа эмоциялы бейне, тренд, челлендж", "15–90 секунд", C.text2],
    [fa.FaInstagram, "Букстаграм", "Instagram", "Кітап суреттері, карусель, Reels", "фото мен қысқа бейне", C.accent6],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15, [ic, h, net, d, len, col] = kinds[i];
    card(s, x, 1.55, 3.75, 3.3, C.background2, "kind-card-" + i);
    await iconCircle(s, ic, x + 0.3, 1.8, 0.8, col, "kind-" + i);
    txt(s, h, { x: x + 1.25, y: 1.8, w: 2.3, h: 0.5, fontSize: 22, bold: true, color: C.text2, fontFace: THEME.headFontFace, objectName: "kind-h-" + i });
    txt(s, net, { x: x + 1.25, y: 2.3, w: 2.3, h: 0.35, fontSize: 14, color: C.accent5, objectName: "kind-net-" + i });
    txt(s, d, { x: x + 0.3, y: 2.95, w: 3.2, h: 1.0, fontSize: 16, color: C.text1, valign: "top", objectName: "kind-d-" + i });
    txt(s, len, { x: x + 0.3, y: 4.1, w: 3.2, h: 0.45, fontSize: 15, bold: true, color: col, objectName: "kind-len-" + i });
  }
  card(s, 0.6, 5.15, 12.1, 1.45, C.background1, "white-nights", C.accent3);
  await iconCircle(s, fa.FaRedo, 0.9, 5.45, 0.85, C.accent3, "wn-ic");
  txt(s, [{ text: "Ескі кітап қайта «тірілді»: ", options: { bold: true, color: HEX.accent3 } },
    { text: "Ф. Достоевскийдің 1848 жылғы «Ақ түндер» повесі 2024 жылы BookTok арқасында Ұлыбританияда ең көп сатылған кітаптардың біріне айналды", options: { color: HEX.dk1 } }],
    { x: 2.0, y: 5.15, w: 10.5, h: 1.45, fontSize: 17, valign: "middle", objectName: "wn-text" });
  s.addNotes("Кітап туралы бейнелердің үш негізгі алаңы бар. Буктюб – YouTube-тегі ұзақ, толық шолулар. Буктокер – TikTok-тағы #BookTok хэштегімен түсірілетін қысқа, эмоциялы бейнелер. Букстаграм – Instagram-дағы кітап суреттері мен Reels. Бұл бейнелердің күші сонша, Достоевскийдің 1848 жылы жазылған «Ақ түндер» повесі 2024 жылы Ұлыбританияда ең көп сатылған кітаптардың біріне айналды. (35 секунд)");

  // ================= 5. Әдістеме =================
  pres.addSection({ title: "Әдістеме" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Әдістеме" });
  s.addText("Қалай зерттедім: үш әдіс", { placeholder: "title" });
  const nT = byWho("Мұғалім").filter(x => x.idea || x.quote).length, nS = byWho("Оқушы").length, nL = byWho("Кітапханашы").length;
  const steps = [
    [fa.FaPoll, "Сауалнама", `${N.n_resp} оқушы`, `7–9 сынып (7-сынып – ${N.by_class["7"]}, 8-сынып – ${N.by_class["8"]}, 9-сынып – ${N.by_class["9"]}). 13 сұрақ, есім сұралмады`, C.accent1],
    [fa.FaSearch, "Блогерлер", `${BL.length} блогер`, "TikTok, Instagram, YouTube-тен қазақ тіліндегі кітап блогерлерін іздеп, тақырыбын талдадым", C.accent2],
    [fa.FaComments, "Сұхбат", `${IV.length} адам`, `${nL} кітапханашы, ${nT} мұғалім, ${nS} оқушы: негізгі ойы мен дәйексөзі`, C.accent3],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15, [ic, h, n, d, col] = steps[i];
    card(s, x, 1.6, 3.75, 4.3, C.background2, "step-card-" + i);
    await iconCircle(s, ic, x + 0.3, 1.9, 0.9, col, "step-" + i);
    txt(s, h, { x: x + 0.3, y: 3.0, w: 3.2, h: 0.5, fontSize: 21, bold: true, color: C.text2, fontFace: THEME.headFontFace, objectName: "step-h-" + i });
    txt(s, n, { x: x + 0.3, y: 3.5, w: 3.2, h: 0.55, fontSize: 24, bold: true, color: col, objectName: "step-n-" + i });
    txt(s, d, { x: x + 0.3, y: 4.15, w: 3.2, h: 1.6, fontSize: 15, color: C.text1, valign: "top", objectName: "step-d-" + i });
  }
  txt(s, "Ата-аналардан жазбаша келісім алынды · нәтижелер Excel кестесінде пайызбен есептелді", { x: 0.6, y: 6.2, w: 12.1, h: 0.45, fontSize: 15, italic: true, color: C.accent5, objectName: "ethics" });
  s.addNotes(`Зерттеуде үш әдіс қолдандым. Біріншісі – сауалнама: оған 7–9 сыныптың ${N.n_resp} оқушысы қатысты, есімдері сұралмады. Екіншісі – TikTok, Instagram және YouTube-тен қазақ тіліндегі кітап блогерлерін іздеп, олардың тақырыбын талдадым, ${BL.length} блогер таптым. Үшіншісі – сұхбат: кітапханашымен, ${nT} мұғаліммен және ${nS} оқушымен сөйлестім. Ата-аналардан келісім алдым. (40 секунд)`);

  // ================= 6. Сауалнама: жалпы =================
  pres.addSection({ title: "Нәтижелер" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Құрдастарым және әлеуметтік желі", { placeholder: "title" });
  const st = [
    [pc(netTop[1]), `ең жиі қолданатын желі – ${netTop[0]}`, C.accent1, fa.FaTiktok],
    [pc((SV.TIME["1–3 сағат"] || 0) + (SV.TIME["3 сағаттан көп"] || 0)), "күніне 1 сағаттан көп уақыт өткізеді", C.accent3, fa.FaClock],
    [pc(SV.READ["Иә"]), "соңғы айда бағдарламадан тыс кітап оқыған", C.accent4, fa.FaBook],
    [pc(SV.HEARD["Иә"]), "«буктокер» не «буктюбер» сөзін естіген", C.accent2, fa.FaHashtag],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.1;
    card(s, x, 1.6, 2.8, 3.6, C.background2, "st-card-" + i);
    await iconCircle(s, st[i][3], x + 0.3, 1.85, 0.7, st[i][2], "st-" + i);
    bigStat(s, x + 0.3, 2.8, 2.3, st[i][0], st[i][1], st[i][2], "st-" + i, 44);
  }
  txt(s, `Сауалнамаға ${N.n_resp} оқушы қатысты`, { x: 0.6, y: 5.6, w: 12.1, h: 0.5, fontSize: 16, italic: true, color: C.accent5, objectName: "st-note" });
  s.addNotes(`Алдымен құрдастарымның әлеуметтік желіні қалай қолданатынын білдім. Олардың ${pc(netTop[1])}-ы ең жиі ${netTop[0]}-ты қолданады. ${pc((SV.TIME["1–3 сағат"] || 0) + (SV.TIME["3 сағаттан көп"] || 0))}-ы күніне бір сағаттан көп уақыт өткізеді. Соңғы айда мектеп бағдарламасынан тыс кітап оқығандар – ${pc(SV.READ["Иә"])}. Ал «буктокер» сөзін ${pc(SV.HEARD["Иә"])}-ы ғана естіген. (30 секунд)`);

  // ================= 7. Қайдан біледі =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Жаңа кітап туралы кімнен біледі?", { placeholder: "title" });
  const srcAll = Object.entries(srcN).sort((a, b) => a[1] - b[1]);
  s.addChart(pres.charts.BAR, [{ name: "Оқушы саны", labels: srcAll.map(([k]) => k), values: srcAll.map(([, v]) => v) }], {
    x: 0.6, y: 1.5, w: 7.4, h: 5.0, barDir: "bar", ...chartQuiet(), chartColors: [HEX.accent2], showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", valAxisMinVal: 0, valAxisMajorUnit: 1, valAxisHidden: true, valGridLine: { style: "none" },
    showTitle: true, title: `Жауап берген ${N.n_src} оқушы (бірнеше жауап)`, titleFontSize: 14, titleColor: HEX.accent5,
  });
  card(s, 8.4, 1.6, 4.3, 4.9, C.text2, "src-card");
  await iconCircle(s, fa.FaChalkboardTeacher, 8.75, 1.95, 0.9, C.accent3, "src-ic");
  txt(s, topSrc ? `${topSrc} – бірінші орында` : "", { x: 8.75, y: 3.05, w: 3.7, h: 1.0, fontSize: 24, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "top", objectName: "src-h" });
  txt(s, `Әлеуметтік желіні ${srcN["Әлеуметтік желі"]} оқушы ғана атады. Бұл сұраққа ${N.n_resp} оқушының ${N.n_src}-${px(N.n_src)} жауап берді`,
    { x: 8.75, y: 4.15, w: 3.7, h: 2.1, fontSize: 16, color: C.background2, valign: "top", objectName: "src-d" });
  s.addNotes(`Ең күтпеген нәтиже осы болды. «Жаңа кітап туралы қайдан білесің?» деген сұраққа жауап берген ${N.n_src} оқушының ${srcN["Мұғалім"]}-${px(srcN["Мұғалім"])} мұғалімді, ${srcN["Әлеуметтік желі"]}-${px(srcN["Әлеуметтік желі"])} ғана әлеуметтік желіні атады. Яғни біздің мектепте мұғалім әлі де кітап таңдаудың басты көзі. Бірақ бұл сұраққа барлық оқушы жауап бермегенін ескеру керек. (30 секунд)`);

  // ================= 8. Блогер ұсынған кітап =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Кітапты алады, бірақ соңына дейін оқи ма?", { placeholder: "title" });
  card(s, 0.6, 1.6, 4.2, 4.9, C.background2, "tried-card");
  await iconCircle(s, fa.FaHandPointer, 0.9, 1.9, 0.8, C.accent1, "tried-ic");
  bigStat(s, 0.9, 2.95, 3.6, pc(SV.TRIED["Иә"]), "блогер ұсынған кітапты оқып көрген", C.accent1, "tried", 60);
  txt(s, `${SV.FINISH.n} оқушы`, { x: 0.9, y: 5.4, w: 3.6, h: 0.5, fontSize: 16, bold: true, color: C.accent5, objectName: "tried-n" });
  const fin = ["Иә", "Жартылай", "Жоқ"].filter(k => SV.FINISH[k]);
  const finLab = { "Иә": "Соңына дейін оқыды", "Жартылай": "Жартылай оқыды", "Жоқ": "Оқымады" };
  const finCol = { "Иә": HEX.accent4, "Жартылай": HEX.accent3, "Жоқ": HEX.accent1 };
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Блогер ұсынған кітап", labels: fin.map(k => finLab[k]), values: fin.map(k => SV.FINISH[k]) }], {
    x: 5.1, y: 1.5, w: 4.3, h: 5.0, ...chartText, holeSize: 55, chartColors: fin.map(k => finCol[k]),
    showLegend: true, legendPos: "b", legendFontSize: 14, legendColor: HEX.dk1, showPercent: true, showValue: false,
    dataLabelColor: "FFFFFF", dataLabelFontSize: 16, dataLabelFontBold: true,
    showTitle: true, title: `Оқығандар (${SV.FINISH.n} оқушы)`, titleFontSize: 14, titleColor: HEX.accent5,
  });
  card(s, 9.7, 1.6, 3.0, 4.9, C.background1, "fin-card", "E0D6E3");
  txt(s, pc(notFin), { x: 9.95, y: 2.0, w: 2.6, h: 0.95, fontSize: 48, bold: true, color: C.accent3, fontFace: THEME.headFontFace, objectName: "fin-v" });
  txt(s, "оқығандардың кітапты соңына дейін оқымаған бөлігі", { x: 9.95, y: 3.0, w: 2.55, h: 1.2, fontSize: 15, color: C.text1, valign: "top", objectName: "fin-l" });
  txt(s, "Кітапты аяқтау үшін талқылау мен қолдау керек", { x: 9.95, y: 4.4, w: 2.55, h: 1.6, fontSize: 15, bold: true, color: C.text2, valign: "top", objectName: "fin-c" });
  s.addNotes(`Блогер ұсынған кітапты оқушылардың ${pc(SV.TRIED["Иә"])}-ы оқып көрген, бұл – ${SV.FINISH.n} оқушы. Бірақ солардың ${pc(SV.FINISH["Иә"])}-ы ғана кітапты соңына дейін оқыған, қалғаны жартылай ғана оқыған. Яғни бейне кітапты қолға алуға түрткі болады, бірақ оны аяқтау үшін бейненің өзі жеткіліксіз. (30 секунд)`);

  // ================= 9. Қайшылық =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Кедергі дейді, бірақ бейне түсіргісі келеді", { placeholder: "title" });
  card(s, 0.6, 1.6, 5.85, 2.3, C.background2, "c-l");
  await iconCircle(s, fa.FaBan, 0.9, 1.9, 0.75, C.accent1, "c-l-ic");
  txt(s, pc(SV.OPIN["Кедергі жасайды"]), { x: 1.9, y: 1.75, w: 4.3, h: 1.0, fontSize: 50, bold: true, color: C.accent1, fontFace: THEME.headFontFace, objectName: "c-l-v" });
  txt(s, "әлеуметтік желі кітап оқуға кедергі жасайды деп санайды", { x: 0.9, y: 2.85, w: 5.3, h: 0.9, fontSize: 16, color: C.text1, valign: "top", objectName: "c-l-t" });
  card(s, 6.85, 1.6, 5.85, 2.3, C.background2, "c-r");
  await iconCircle(s, fa.FaVideo, 7.15, 1.9, 0.75, C.accent4, "c-r-ic");
  txt(s, pc(SV.VIDEO["Иә"]), { x: 8.15, y: 1.75, w: 4.3, h: 1.0, fontSize: 50, bold: true, color: C.accent4, fontFace: THEME.headFontFace, objectName: "c-r-v" });
  txt(s, "өзі кітап туралы бейне түсіріп көрер еді", { x: 7.15, y: 2.85, w: 5.3, h: 0.9, fontSize: 16, color: C.text1, valign: "top", objectName: "c-r-t" });
  txt(s, "Оқушылар өз сөзімен не деді (ашық сұрақ, 19 жауап)", { x: 0.6, y: 4.15, w: 12.1, h: 0.45, fontSize: 17, bold: true, color: C.text2, objectName: "why-h" });
  const why = [
    ["Кедергі", "зейін әлсірейді, тез жалығады", "4", C.accent1], ["Кедергі", "жеңіл «дофаминге» үйренеді", "3", C.accent1],
    ["Көмек", "блогер қысқа әрі түсінікті ұсынады", "4", C.accent4], ["Көмек", "жарнама, білім, онлайн кітапхана", "4", C.accent4],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.1;
    card(s, x, 4.75, 2.8, 1.85, C.background1, "why-card-" + i, "E0D6E3");
    txt(s, why[i][2], { x: x + 0.25, y: 4.9, w: 0.8, h: 0.75, fontSize: 34, bold: true, color: why[i][3], fontFace: THEME.headFontFace, objectName: "why-n-" + i });
    txt(s, why[i][0], { x: x + 1.05, y: 4.95, w: 1.6, h: 0.6, fontSize: 15, bold: true, color: why[i][3], valign: "middle", objectName: "why-k-" + i });
    txt(s, why[i][1], { x: x + 0.25, y: 5.7, w: 2.35, h: 0.8, fontSize: 14, color: C.text1, valign: "top", objectName: "why-t-" + i });
  }
  s.addNotes(`Менің ойымша, ең қызық нәтиже – осы қайшылық. Оқушылардың ${pc(SV.OPIN["Кедергі жасайды"])}-ы әлеуметтік желі кітап оқуға кедергі жасайды деп санайды. Бірақ ${pc(SV.VIDEO["Иә"])}-ы өзі кітап туралы бейне түсіріп көрер еді. Ашық сұрақта оқушылар желіге үйренген ми ұзақ мәтінге зейін қоя алмайтынын жазды. Ал пайдасы туралы: блогер кітап туралы қысқа әрі түсінікті айтып, нені оқу керегін ұсынады. Демек оқушылар желінің зиянын түсінеді, бірақ оны кітап үшін пайдалануға дайын. (40 секунд)`);

  // ================= 10. Сұхбат =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Сұхбат: кітапты мұқабасы үшін алады", { placeholder: "title" });
  const lib = byWho("Кітапханашы")[0], stu = byWho("Оқушы"), tea = byWho("Мұғалім").filter(x => x.quote);
  const pick = (arr, re) => arr.find(x => re.test(x.quote)) || arr[0];
  const quotes = [
    [fa.FaBookReader, "Кітапханашы", lib && lib.quote, C.accent2],
    [fa.FaChalkboardTeacher, "Мұғалім", pick(tea, /не түсіндің/).quote, C.accent3],
    [fa.FaUserGraduate, "Оқушы", pick(stu, /артта қалмайын/).quote.replace(/^.*?(Мен де артта)/, "$1"), C.accent1],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15, [ic, w, qt, col] = quotes[i];
    card(s, x, 1.6, 3.75, 4.95, C.background2, "qt-card-" + i);
    await iconCircle(s, ic, x + 0.3, 1.85, 0.75, col, "qt-" + i);
    txt(s, w, { x: x + 1.2, y: 1.85, w: 2.3, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle", objectName: "qt-w-" + i });
    txt(s, `«${qt}»`, { x: x + 0.3, y: 2.85, w: 3.2, h: 3.5, fontSize: 17, italic: true, color: C.text1, valign: "top", objectName: "qt-t-" + i });
  }
  s.addNotes("Сұхбат сауалнаманың себебін түсіндірді. Кітапханашы оқушылар кітапты әдемі мұқабасы үшін таласып алатынын, бірақ келесі күні «тілі тым ауыр» деп қайтаратынын айтты. Мұғалім оқушылар бір кітапты бәрі жабылып оқитынын, бірақ мазмұнын айтып бере алмайтынын жеткізді. Ал оқушы кітапты достарынан артта қалмау үшін алғанын, бірақ әлі екі бетін де оқымағанын мойындады. (40 секунд)");

  // ================= 11. Қазақ тіліндегі блогерлер =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Нәтижелер" });
  s.addText("Қазақша блогерлер бар, бірақ оларды аз біледі", { placeholder: "title" });
  const bl = [...BL].filter(x => x.subs != null).sort((a, b) => a.subs - b.subs);
  s.addChart(pres.charts.BAR, [{ name: "Жазылушы", labels: bl.map(x => `${shortName(x.name)} (${x.net})`), values: bl.map(x => x.subs) }], {
    x: 0.6, y: 1.5, w: 7.6, h: 5.1, barDir: "bar", ...chartQuiet(), catAxisLabelFontSize: 12, chartColors: [HEX.accent6], showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "# ##0", valAxisHidden: true, valGridLine: { style: "none" },
    showTitle: true, title: "Жазылушылар саны", titleFontSize: 14, titleColor: HEX.accent5,
  });
  card(s, 8.6, 1.6, 4.1, 2.3, C.text2, "kaz-card");
  txt(s, pc(SV.KAZ["Иә"]), { x: 8.9, y: 1.75, w: 3.6, h: 0.95, fontSize: 50, bold: true, color: C.accent3, fontFace: THEME.headFontFace, objectName: "kaz-v" });
  txt(s, "оқушы ғана қазақ тіліндегі кітап блогерін көреді", { x: 8.9, y: 2.75, w: 3.6, h: 1.0, fontSize: 16, color: C.background1, valign: "top", objectName: "kaz-l" });
  card(s, 8.6, 4.2, 4.1, 2.4, C.background2, "kaz-card2");
  txt(s, [
    { text: `${BL.length} блогердің ${blLit}-${px(blLit)} ұлттық әдебиетке де көңіл бөледі`, options: { breakLine: true } },
    { text: BL.some(x => x.net === "TikTok") ? "" : "Оқушылар көп отыратын TikTok-тан қазақ тіліндегі тұрақты кітап блогерін таппадым", options: { bold: true, color: HEX.accent1 } },
  ], { x: 8.9, y: 4.35, w: 3.6, h: 2.1, fontSize: 15, color: C.text1, valign: "top", paraSpaceAfter: 8, objectName: "kaz-t" });
  s.addNotes(`TikTok, Instagram және YouTube-тен қазақ тілінде кітап туралы жазатын ${BL.length} блогер мен арнаны таптым. Олардың жазылушылары ${big(Math.min(...bl.map(x => x.subs)))}-ден ${big(Math.max(...bl.map(x => x.subs)))}-ге дейін, ${blLit}-${px(blLit)} ұлттық әдебиет туралы да жазады. Бірақ сауалнама бойынша оларды оқушылардың ${pc(SV.KAZ["Иә"])}-ы ғана көреді. Оның үстіне, оқушылар ең көп отыратын TikTok-тан қазақ тіліндегі тұрақты кітап блогерін таппадым. Менің ойымша, бұл – бос тұрған мүмкіндік. (40 секунд)`);

  // ================= 12. Болжамды тексеру =================
  pres.addSection({ title: "Қорытынды" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText(`Болжамым ${hypShort}`, { placeholder: "title" });
  const hr = [
    [okA, "Жаңа кітап туралы көбіне әлеуметтік желіден біледі", okA ? `Иә: әлеуметтік желі бірінші орында` : `Жоқ: ${topSrc ? topSrc.toLowerCase() : "басқа көз"} – ${srcSorted[0] ? srcSorted[0][1] : 0} оқушы, әлеуметтік желі – ${srcN["Әлеуметтік желі"]} оқушы`],
    [okB, "Блогер ұсынған кітапты кемінде төрттен бірі аяқтамайды", `${okB ? "Иә" : "Жоқ"}: ${pc(notFin)} соңына дейін оқымаған`],
    [okC, "Қазақ тіліндегі кітап блогерін жартысынан азы көреді", `${okC ? "Иә" : "Жоқ"}: ${pc(SV.KAZ["Иә"])} ғана көреді`],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.65, [okv, h, d] = hr[i], col = okv ? C.accent4 : C.accent1;
    card(s, 0.6, y, 12.1, 1.4, C.background2, "hr-card-" + i);
    await iconCircle(s, okv ? fa.FaCheck : fa.FaTimes, 0.9, y + 0.27, 0.86, col, "hr-" + i);
    txt(s, h, { x: 2.1, y: y + 0.15, w: 6.0, h: 1.1, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "hr-h-" + i });
    txt(s, d, { x: 8.3, y: y + 0.15, w: 4.2, h: 1.1, fontSize: 16, color: col, bold: true, valign: "middle", objectName: "hr-d-" + i });
  }
  s.addNotes(`Болжамымды үш бөлік бойынша тексердім. Бірінші бөлігі ${okA ? "расталды" : "расталмады: оқушылар жаңа кітап туралы әлеуметтік желіден емес, көбіне мұғалімнен біледі екен"}. Екінші бөлігі ${okB ? "расталды: блогер ұсынған кітапты оқығандардың жартысы оны соңына дейін оқымаған" : "расталмады"}. Үшінші бөлігі ${okC ? `расталды: қазақ тіліндегі кітап блогерін оқушылардың ${pc(SV.KAZ["Иә"])}-ы ғана көреді` : "расталмады"}. Сонымен болжамым ${hypShort}. (35 секунд)`);

  // ================= 13. Оқу шынжыры =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Практикалық өнім: «Оқу шынжыры» моделі", { placeholder: "title" });
  const img = path.join(__dirname, "..", "figs", "shynzhyr.png");
  const ib = fs.readFileSync(img), iw = ib.readUInt32BE(16), ih = ib.readUInt32BE(20), H = 5.1;
  s.addImage({ path: img, x: 0.6, y: 1.45, w: H * iw / ih, h: H, altText: "«Оқу шынжыры» моделі: бейнешолу көру, кітапты алу, оқып шығу, өз пікірін айту, жаңа бейнешолу түсіру", objectName: "model-img" });
  const recs = [
    [fa.FaUserGraduate, "Оқушыларға", "Кітапты «тренд» үшін емес, өзіңе қызық болғаны үшін таңда; оқығаныңа 1 минуттық шолу жаса", C.accent1],
    [fa.FaChalkboardTeacher, "Мұғалімдерге", "Үй тапсырмасы ретінде «60 секундтық шолу»; ұлттық классикаға оқушы бейнелері", C.accent3],
    [fa.FaBookReader, "Кітапханаға", "«BookTok ұсынады» сөресі және оқу клубы: шынжырдың әлсіз буыны – «оқып шығу»", C.accent2],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.7;
    card(s, 7.3, y, 5.4, 1.5, C.background2, "rec-card-" + i);
    await iconCircle(s, recs[i][0], 7.55, y + 0.35, 0.8, recs[i][3], "rec-" + i);
    txt(s, recs[i][1], { x: 8.6, y: y + 0.12, w: 3.9, h: 0.4, fontSize: 17, bold: true, color: C.text2, objectName: "rec-h-" + i });
    txt(s, recs[i][2], { x: 8.6, y: y + 0.5, w: 3.9, h: 0.95, fontSize: 14, color: C.text1, valign: "top", objectName: "rec-t-" + i });
  }
  s.addNotes("Нәтиже бойынша «Оқу шынжыры» моделін ұсындым. Бейнешолу көрген оқушы кітапты алады, оқып шығады, өз пікірін айтады және өзі жаңа бейнешолу түсіреді. Сол бейне келесі оқушыны кітапқа жетелейді. Зерттеу көрсеткендей, шынжырдың әлсіз буыны – кітапты оқып шығу. Сондықтан оқушыларға, мұғалімдерге және кітапханаға ұсыныстар әзірледім: «60 секундтық шолу» тапсырмасы, оқу клубы, «BookTok ұсынады» сөресі. (40 секунд)");

  // ================= 14. Қорытынды =================
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Қорытынды" });
  s.addText("Әлеуметтік желі кітапқа жол ашады, ал оқып шығуға мұғалім мен дос керек", { placeholder: "title" });
  s.addText("Келесі қадам: мектепте қазақ тіліндегі «Оқу шынжыры» парақшасын ашу", { placeholder: "body" });
  const fin2 = [
    [`${pc(SV.TRIED["Иә"])}`, "блогер ұсынған кітапты оқыған"],
    [`${pc(SV.FINISH["Иә"])}`, "солардың соңына дейін оқығаны"],
    [`${pc(SV.KAZ["Иә"])}`, "қазақша кітап блогерін көреді"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.45 + i * 1.6;
    card(s, 8.9, y, 3.8, 1.35, C.background1, "fin-card-" + i);
    txt(s, fin2[i][0], { x: 9.15, y: y + 0.1, w: 3.3, h: 0.7, fontSize: 32, bold: true, color: C.accent1, fontFace: THEME.headFontFace, objectName: "fin-v-" + i });
    txt(s, fin2[i][1], { x: 9.15, y: y + 0.8, w: 3.3, h: 0.45, fontSize: 14, color: C.text1, objectName: "fin-l-" + i });
  }
  txt(s, "Назарларыңызға рахмет!", { x: 0.8, y: 6.3, w: 7, h: 0.5, fontSize: 20, bold: true, color: C.background1, objectName: "thanks" });
  s.addNotes(`Қорытындылай келе: әлеуметтік желі кітапқа жол ашады – оқушылардың ${pc(SV.TRIED["Иә"])}-ы блогер ұсынған кітапты қолына алған. Бірақ оны соңына дейін оқу үшін мұғалімнің кеңесі, достарымен талқылау қажет. Қазақ тіліндегі кітап блогерлері бар, бірақ оларды аз біледі. Шектеулері: сауалнамаға бір мектептің ${N.n_resp} оқушысы ғана қатысты. Келесі қадамым – мектепте қазақ тіліндегі «Оқу шынжыры» парақшасын ашып, ұлттық классикаға оқушы бейнешолуларын жасау. Назарларыңызға рахмет! (40 секунд)`);

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });
