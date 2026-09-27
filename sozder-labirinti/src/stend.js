// Көрме стендінің баспа макеті (Ережелердің 4.16–4.17-т., 9-қосымша):
// сол панель 30×80, орталық 60×80, оң панель 30×80 см. Аты-жөн, мектеп, өңір көрсетілмейді.
// QR-код: STAND_URL=https://... node stend.js  (сілтеме берілмесе, QR орны бос қалады)
const fs = require("fs");
const path = require("path");
const QR = require("qrcode");
const { chromium } = require("/opt/node22/lib/node_modules/playwright");
const OUT = process.argv[2] || path.join(__dirname, "../Stend_maketi.pdf");
const URL_ = process.env.STAND_URL || "";
const img = f => "data:image/png;base64," + fs.readFileSync(path.join(__dirname, "figs", f)).toString("base64");

(async () => {
  const qr = URL_ ? await QR.toDataURL(URL_, { margin: 1, width: 800, color: { dark: "#1D2846", light: "#FFFFFF" } }) : "";
  const html = `<!doctype html><meta charset="utf-8"><style>
  @page side { size: 30cm 80cm; margin: 0 }
  @page mid { size: 60cm 80cm; margin: 0 }
  *{box-sizing:border-box} body{margin:0;font-family:"DejaVu Sans",Arial,sans-serif;color:#1D2846}
  .panel{position:relative;height:80cm;padding:2.2cm 2cm;display:flex;flex-direction:column;gap:1.3cm;overflow:hidden;break-after:page;background:#F4F9FB}
  .side{page:side;width:30cm} .mid{page:mid;width:60cm}
  .band{position:absolute;left:0;right:0;top:0;height:1.2cm;background:
    linear-gradient(135deg,#00809E 25%,transparent 25%) 0 0/2.4cm 2.4cm,
    linear-gradient(225deg,#00809E 25%,transparent 25%) 0 0/2.4cm 2.4cm,#F2B233}
  .band.b{top:auto;bottom:0}
  h1{font-size:88pt;line-height:1.05;margin:0;font-weight:800;color:#2C3A6B}
  h1 span{color:#00809E}
  .sub{font-size:40pt;margin:.3cm 0 0;color:#56627E}
  h2{font-size:34pt;margin:0 0 .4cm;color:#2C3A6B;display:flex;align-items:center;gap:.5cm}
  h2 i{font-style:normal;display:inline-grid;place-items:center;width:1.7cm;height:1.7cm;border-radius:50%;background:#00809E;color:#fff;font-size:26pt}
  p,li{font-size:24pt;line-height:1.4;margin:0}
  ul,ol{margin:0;padding-left:1.1cm;display:flex;flex-direction:column;gap:.35cm}
  .box{background:#fff;border-radius:.8cm;padding:1cm 1.1cm;border:3px solid #CBD8E0}
  .sp{color:#B8286F;font-weight:800}
  .letters{display:grid;grid-template-columns:repeat(3,1fr);gap:.4cm}
  .letters b{display:grid;place-items:center;height:3.2cm;border-radius:.6cm;background:#FBE3EF;color:#B8286F;font-size:66pt}
  .head{display:flex;gap:1.2cm;align-items:center}
  .code{flex:none;width:14cm;height:7cm;border:4px dashed #56627E;border-radius:.3cm;display:grid;place-items:center;font-size:22pt;color:#56627E;text-align:center}
  .shots{display:grid;grid-template-columns:repeat(3,1fr);gap:1.2cm}
  .shots figure{margin:0;text-align:center}
  .shots img{width:100%;border-radius:.6cm;border:3px solid #CBD8E0}
  .shots figcaption{font-size:22pt;color:#56627E;margin-top:.3cm}
  .q{font-size:32pt;line-height:1.35}
  .res{display:grid;grid-template-columns:1fr 1fr 1fr;gap:1cm}
  .res div{background:#fff;border-radius:.8cm;padding:.9cm;text-align:center;border:3px solid #CBD8E0}
  .res b{display:block;font-size:72pt;color:#00809E}
  .res span{font-size:24pt;color:#56627E}
  .fill,.res div.fill{background:#FFF1CF}
  .chart img{width:100%}
  .qr{width:17cm;height:17cm;align-self:center;border-radius:.6cm;background:#fff;border:4px ${qr ? "solid #CBD8E0" : "dashed #56627E"};display:grid;place-items:center;font-size:24pt;color:#56627E;text-align:center}
  .qr img{width:15.5cm;height:15.5cm}
  .refs li{font-size:18pt}
  .grow{flex:1}
  </style>

  <section class="panel side"><div class="band"></div>
    <div class="box"><h2><i>1</i>Өзектілігі</h2>
      <p>Қазақ тілінде орыс тілінде жоқ 9 әріп бар. Орыс тілінде оқитын біз оларды жиі шатастырамыз.</p>
      <div class="letters" style="margin-top:.7cm"><b>ә</b><b>ғ</b><b>қ</b><b>ң</b><b>ө</b><b>ұ</b><b>ү</b><b>һ</b><b>і</b></div>
      <p style="margin-top:.7cm"><span class="sp">қоян</span>, ал «коян» – қате!</p></div>
    <div class="box"><h2><i>2</i>Мақсаты</h2>
      <p>Қазақ сөздерін үйренуге арналған лабиринт ойынын жасау және оны сыныптастарыммен тексеру.</p></div>
    <div class="box"><h2><i>3</i>Міндеттері</h2><ol>
      <li>Сыныптастардан не қиын екенін сұрау</li><li>Оқулықтан 70 сөз таңдау</li><li>Ойын ережесін ойлап табу</li>
      <li>Claude Code көмегімен ойынды жасау және сынау</li><li>Ойынға дейін және кейін тест өткізу</li><li>Нәтижелерді салыстыру</li></ol></div>
    <div class="box grow"><h2><i>4</i>Әдістері</h2><ul>
      <li>сауалнама</li><li>модельдеу (ойын жасау)</li><li>эксперимент: бастапқы және қорытынды тест</li><li>салыстыру, бақылау</li></ul></div>
  <div class="band b"></div></section>

  <section class="panel mid"><div class="band"></div>
    <div class="head"><div class="code">Тіркеу коды<br>(шифр)<br>7 × 14 см</div>
      <div><h1>СӨЗДЕР <span>ЛАБИРИНТІ</span></h1><p class="sub">Қазақ тілінен интерактивті ойын құрастыру</p></div></div>
    <div class="box"><p class="q"><b>Зерттеу сұрағы:</b> менің ойыным сыныптастарыма қазақ сөздерін дұрыс жазуды үйренуге көмектесе ме?</p>
      <p class="q" style="margin-top:.4cm"><b>Болжам:</b> ойынды 1–2 апта ойнаса, қорытынды тестте көбірек дұрыс жауап береді.</p></div>
    <div class="shots">
      <figure><img src="${img("s_home.png")}"><figcaption>Басты бет</figcaption></figure>
      <figure><img src="${img("s_maze.png")}"><figcaption>Лабиринттен «қоян» сөзін жинау</figcaption></figure>
      <figure><img src="${img("s_card.png")}"><figcaption>Сөз карточкасы және жұлдыздар</figcaption></figure></div>
    <div><h2>Негізгі нәтиже</h2>
      <div class="res"><div class="fill"><b>__ / 12</b><span>бастапқы тест (орташа)</span></div>
        <div class="fill"><b>__ / 12</b><span>қорытынды тест (орташа)</span></div>
        <div><b>70</b><span>сөз, 7 тақырып, 52-сінде қазақ әрпі бар</span></div></div>
      <p style="margin-top:.5cm;font-size:20pt;color:#56627E">Сары орындарды тест нәтижесі шыққан соң толтырыңыз.</p></div>
    <div class="box"><h2>Ойын ережесі</h2><ol>
      <li>Суретке қара: бұл қандай сөз?</li>
      <li>Ботаны 🐪 лабиринт бойымен жүргіз, сөздің әріптерін <b>ретімен</b> жина.</li>
      <li>Абай бол: <span class="sp">қ</span> орнына «к», <span class="sp">ұ</span> орнына «у» – қақпан әріптер!</li>
      <li>Сөз толық жиналса, 🏁 шығу ашылады. Қатесіз өтсең – 3 жұлдыз.</li></ol></div>
  <div class="band b"></div></section>

  <section class="panel side"><div class="band"></div>
    <div class="box"><h2><i>5</i>Қорытынды</h2><ul>
      <li>«Сөздер лабиринті» ойыны жасалды: 7 тақырып, 70 сөз, тест</li>
      <li>Сынақта 3 қате табылып, түзетілді</li>
      <li class="fill" style="border-radius:.3cm;padding:.1cm .3cm">Тест нәтижесі: [__] → [__]</li>
      <li>«Шешімді адам қабылдайды – ЖИ тек көмектеседі»</li></ul></div>
    <div class="box"><h2><i>6</i>Практикалық маңызы</h2>
      <p>Ойынды сабақта, үйде және үзілісте ойнауға болады. Интернет қажет емес. Мұғалім қай әріпте қате көп екенін көреді.</p></div>
    <div class="chart box"><h2>Ойын сөздігіндегі қазақ әріптері</h2><img src="${img("letters.png")}"></div>
    <div class="box"><h2>Дереккөздер</h2><ol class="refs">
      <li>Қазақ тілінің түсіндірме сөздігі. – Алматы: Дайк-Пресс, 2008.</li>
      <li>Sozdik.kz – онлайн сөздік.</li>
      <li>Эльконин Д.Б. Психология игры. – М., 1978.</li>
      <li>Prensky M. Digital Game-Based Learning. – 2001.</li>
      <li>Wouters P. et al. // J. of Educational Psychology. – 2013.</li></ol></div>
    <div class="qr">${qr ? `<img src="${qr}">` : "QR-код<br>ойынға сілтеме<br>(STAND_URL арқылы)"}</div>
    <p style="text-align:center;font-size:26pt;font-weight:800">Ойынды ойнап көр!</p>
  <div class="band b"></div></section>`;
  const tmp = path.join(__dirname, "stend.tmp.html");
  fs.writeFileSync(tmp, html);
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto("file://" + tmp);
  await p.pdf({ path: OUT, preferCSSPageSize: true, printBackground: true });
  await b.close();
  fs.unlinkSync(tmp);
  console.log("wrote", OUT, URL_ ? "(QR: " + URL_ + ")" : "(QR орны бос)");
})();
