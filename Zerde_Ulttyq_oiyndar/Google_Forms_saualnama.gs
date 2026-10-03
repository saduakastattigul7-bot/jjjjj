/**
 * «Қазақтың ұлттық ойындары» сауалнамасын Google Forms-та жасайтын скрипт.
 *
 * ҚАЛАЙ ІСКЕ ҚОСУ КЕРЕК
 * 1. https://script.google.com сайтына Google аккаунтыңызбен кіріп, «Жаңа жоба» (New project) басыңыз.
 * 2. Ондағы бар кодты өшіріп, осы файлдың бүкіл мәтінін қойыңыз да, сақтаңыз (Ctrl+S).
 * 3. Жоғарыдағы тізімнен «createSurvey» функциясын таңдап, «Іске қосу» (Run) басыңыз.
 *    Google рұқсат сұрайды – өз аккаунтыңызды таңдап, «Рұқсат беру» (Allow) басыңыз.
 * 4. Төмендегі «Журнал» (Execution log) терезесінде үш сілтеме шығады:
 *    – оқушыларға жіберетін сауалнама сілтемесі;
 *    – форманы өңдеу сілтемесі;
 *    – жауаптар жиналатын Google кесте.
 * 5. Сауалнама біткен соң «kestegeAudaru» функциясын іске қосыңыз. Жауаптар кестесінде
 *    «Кестеге» парағы пайда болады. Оны A2 ұяшығынан бастап көшіріп, Zertteu_kestesi.xlsx
 *    файлындағы «Сауалнама» парағының A2 ұяшығына қойыңыз.
 */

const GAMES = ['Бәйге', 'Көкпар', 'Қыз қуу', 'Теңге алу', 'Аударыспақ', 'Күрес', 'Жамбы ату',
  'Тоғызқұмалақ', 'Жұмбақ айтысу', 'Асық', 'Ақсүйек', 'Алтыбақан'];
const Q_CONSENT = 'Ата-анаң (заңды өкілің) осы сауалнамаға қатысуыңа жазбаша келісім берді ме?';
const Q_CLASS = 'Нешінші сыныпта оқисың?';
const Q_KNOW = 'Бұл ойындарды білесің бе?';
const Q_FOLK = 'Бұл ойындар туралы ертегіден немесе жырдан оқыдың ба?';
const Q_FAV = 'Қай ұлттық ойын саған ең көп ұнайды? Неге? (міндетті емес)';
const KNOW_COLS = ['0 – Білмеймін', '1 – Естігенмін, бірақ ойнамағанмын', '2 – Ойнағанмын немесе көргенмін'];
const FOLK_COLS = ['Иә, оқыдым', 'Жоқ / есімде жоқ'];

function createSurvey() {
  const form = FormApp.create('Қазақтың ұлттық ойындары: сен оларды білесің бе?');
  form.setDescription(
    'Сәлем! Бұл сауалнама «Ертегілер мен жырлардағы ұлттық ойындар» атты зерттеу жобасы үшін жүргізіліп отыр.\n\n' +
    '• Сауалнама 5 минуттай уақыт алады.\n' +
    '• Атыңды, тегіңді, электрондық поштаңды жазудың қажеті жоқ – сауалнама анонимді.\n' +
    '• Қатысу ерікті: қаламасаң, кез келген уақытта тоқтата аласың.\n' +
    '• Жауаптар тек жалпы сан түрінде (мысалы, «20 оқушының 15-і біледі») қолданылады.\n\n' +
    'Дұрыс не бұрыс жауап жоқ – тек шыныңды айт!');
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setShowLinkToRespondAgain(false);
  form.setConfirmationMessage('Рахмет! Жауабың жазылды. Ұлттық ойындарды ойнап көруді ұмытпа!');

  // 1-бөлім: келісім
  const consent = form.addMultipleChoiceItem().setTitle(Q_CONSENT)
    .setHelpText('Келісім берілмесе, «Жоқ» деп белгіле – сауалнама осымен аяқталады.').setRequired(true);

  // 2-бөлім: ойындарды білу
  const page2 = form.addPageBreakItem().setTitle('Ұлттық ойындарды білу')
    .setHelpText('Әр ойын бойынша бір жауап белгіле.');
  form.addListItem().setTitle(Q_CLASS).setChoiceValues(['5', '6', '7', '8']).setRequired(true);
  form.addGridItem().setTitle(Q_KNOW).setRows(GAMES).setColumns(KNOW_COLS).setRequired(true);

  // 3-бөлім: ертегі мен жыр
  form.addPageBreakItem().setTitle('Ертегі мен жыр')
    .setHelpText('Бұл ойынның аты аталатын немесе ол ойналатын ертегіні не жырды оқыдың ба?');
  form.addGridItem().setTitle(Q_FOLK).setRows(GAMES).setColumns(FOLK_COLS).setRequired(true);
  form.addParagraphTextItem().setTitle(Q_FAV);

  consent.setChoices([
    consent.createChoice('Иә', page2),
    consent.createChoice('Жоқ', FormApp.PageNavigationType.SUBMIT),
  ]);

  const ss = SpreadsheetApp.create('Ұлттық ойындар сауалнамасы – жауаптар');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  PropertiesService.getScriptProperties().setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId() });

  Logger.log('Оқушыларға жіберетін сілтеме: ' + form.getPublishedUrl());
  Logger.log('Форманы өңдеу: ' + form.getEditUrl());
  Logger.log('Жауаптар кестесі: ' + ss.getUrl());
}

/** Жауаптарды Zertteu_kestesi.xlsx «Сауалнама» парағының пішініне келтіреді. */
function kestegeAudaru() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const header = ['Код', 'Сынып'];
  GAMES.forEach(g => header.push(g + ': білу (0/1/2)', g + ': ертегіден (1/0)'));
  const rows = [], favs = [];
  let n = 0, skipped = 0;
  form.getResponses().forEach(resp => {
    const ans = {};
    resp.getItemResponses().forEach(ir => { ans[ir.getItem().getTitle()] = ir.getResponse(); });
    if (ans[Q_CONSENT] !== 'Иә') { skipped++; return; }
    n++;
    const code = 'С' + String(n).padStart(2, '0');
    const know = ans[Q_KNOW] || [], folk = ans[Q_FOLK] || [];
    const row = [code, ans[Q_CLASS] || ''];
    GAMES.forEach((g, i) => {
      row.push(know[i] ? Number(String(know[i]).charAt(0)) : '');
      row.push(folk[i] ? (folk[i] === FOLK_COLS[0] ? 1 : 0) : '');
    });
    rows.push(row);
    if (ans[Q_FAV]) favs.push([code, ans[Q_FAV]]);
  });
  const write = (name, head, data) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    if (data.length) sh.getRange(2, 1, data.length, head.length).setValues(data);
  };
  write('Кестеге', header, rows);
  write('Пікірлер', ['Код', Q_FAV], favs);
  Logger.log('Дайын: ' + n + ' жауап «Кестеге» парағына жазылды' + (skipped ? ', келісімсіз ' + skipped + ' жауап алынып тасталды.' : '.'));
  Logger.log(ss.getUrl());
}
