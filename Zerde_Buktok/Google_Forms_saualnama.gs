/**
 * «Буктюб пен буктокерлер» сауалнамасын Google Forms-та жасайтын скрипт.
 *
 * ҚАЛАЙ ІСКЕ ҚОСУ КЕРЕК
 * 1. https://script.google.com сайтына Google аккаунтыңызбен кіріп, «Жаңа жоба» (New project) басыңыз.
 * 2. Ондағы кодты өшіріп, осы файлдың бүкіл мәтінін қойыңыз да, сақтаңыз (Ctrl+S).
 * 3. «createSurvey» функциясын таңдап, «Іске қосу» (Run) басыңыз, рұқсат беріңіз.
 * 4. «Журнал» (Execution log) терезесінде оқушыларға жіберетін сілтеме және жауаптар кестесі шығады.
 * 5. Сауалнама біткен соң жауаптар кестесін Файл → Жүктеп алу → Microsoft Excel (.xlsx) арқылы жүктеп,
 *    жобаға жіберіңіз: жауаптар зерттеу кестесіне өздігінен аударылады.
 * Сұрақтардың мәтінін өзгертпеңіз – кестеге аудару осы мәтін бойынша жүреді.
 */

const Q_CONSENT = 'Ата-анаң (заңды өкілің) осы сауалнамаға қатысуыңа жазбаша келісім берді ме?';

function createSurvey() {
  const form = FormApp.create('Әлеуметтік желі және кітап оқу');
  form.setDescription(
    'Сәлем! Бұл сауалнама «Буктюб пен буктокерлер: әлеуметтік желі жастардың кітап оқуына әсер ете ме?» атты зерттеу жобасы үшін жүргізіліп отыр.\n\n' +
    '• Сауалнама 5–7 минут уақыт алады.\n' +
    '• Атыңды, тегіңді, электрондық поштаңды жазудың қажеті жоқ – сауалнама анонимді.\n' +
    '• Қатысу ерікті: қаламасаң, кез келген уақытта тоқтата аласың.\n' +
    '• Жауаптар тек жалпы сан түрінде қолданылады.\n\n' +
    'Дұрыс не бұрыс жауап жоқ – тек шыныңды айт!');
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setShowLinkToRespondAgain(false);
  form.setConfirmationMessage('Рахмет! Жауабың жазылды. Қызықты кітап оқуды ұмытпа!');

  const consent = form.addMultipleChoiceItem().setTitle(Q_CONSENT)
    .setHelpText('Келісім берілмесе, «Жоқ» деп белгіле – сауалнама осымен аяқталады.').setRequired(true);
  const page2 = form.addPageBreakItem().setTitle('Әлеуметтік желі және кітап');
  form.addListItem().setTitle("Нешінші сыныпта оқисың?").setChoiceValues(["7", "8", "9"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Әлеуметтік желіде күніне қанша уақыт өткізесің?").setChoiceValues(["1 сағаттан аз", "1–3 сағат", "3 сағаттан көп"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Қай желіні жиі қолданасың?").setChoiceValues(["TikTok", "Instagram", "YouTube", "Басқа"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Соңғы айда мектеп бағдарламасынан тыс кітап оқыдың ба?").setChoiceValues(["Иә", "Жоқ"]).setRequired(true);
  form.addCheckboxItem().setTitle("Жаңа кітап туралы қайдан білесің? (бірнеше жауап)").setChoiceValues(["Әлеуметтік желі", "Мұғалім", "Ата-ана", "Дос", "Кітапхана, кітап дүкені", "Басқа"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("«Буктокер» немесе «буктюбер» сөзін естідің бе?").setChoiceValues(["Иә", "Жоқ"]).setRequired(true);

  form.addMultipleChoiceItem().setTitle("Блогер ұсынған кітапты оқып көрдің бе?").setChoiceValues(["Иә", "Жоқ"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Оны соңына дейін оқыдың ба?").setChoiceValues(["Иә", "Жартылай", "Жоқ"]).setHelpText("Блогер ұсынған кітапты оқымаған болсаң, бос қалдыр.").setRequired(false);
  form.addMultipleChoiceItem().setTitle("Қазақ тіліндегі кітап блогерін көресің бе?").setChoiceValues(["Иә", "Жоқ"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Өзің кітап туралы бейне түсіріп көрер ме едің?").setChoiceValues(["Иә", "Білмеймін", "Жоқ"]).setRequired(true);
  form.addMultipleChoiceItem().setTitle("Әлеуметтік желі кітап оқуға көмектесе ме, кедергі жасай ма?").setChoiceValues(["Көмектеседі", "Кедергі жасайды", "Екеуі де", "Білмеймін"]).setRequired(true);
  form.addParagraphTextItem().setTitle("Кітап блогерлерін көресің бе? Есімдерін жаз.");
  form.addParagraphTextItem().setTitle("Әлеуметтік желі кітап оқуға неге көмектеседі не кедергі жасайды? (өз ойың)");

  consent.setChoices([
    consent.createChoice('Иә', page2),
    consent.createChoice('Жоқ', FormApp.PageNavigationType.SUBMIT),
  ]);

  const ss = SpreadsheetApp.create('Әлеуметтік желі және кітап оқу – жауаптар');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  Logger.log('Оқушыларға жіберетін сілтеме: ' + form.getPublishedUrl());
  Logger.log('Форманы өңдеу: ' + form.getEditUrl());
  Logger.log('Жауаптар кестесі: ' + ss.getUrl());
}
