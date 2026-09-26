const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, Footer, PageNumber,
  PageBreak, HeadingLevel, TabStopType, LevelFormat, VerticalAlign,
} = require("docx");

const DIR = __dirname;
const OUT = process.argv[2] || path.join(DIR, "out.docx");
const TOC_PAGES = fs.existsSync(path.join(DIR, "toc_pages.json"))
  ? JSON.parse(fs.readFileSync(path.join(DIR, "toc_pages.json"), "utf8")) : {};

const FONT = "Times New Roman";
const SZ = 28;         // 14 pt
const SZ_T = 24;       // 12 pt in tables
const CM = 567;        // dxa per cm
const PAGE_W = 11906, ML = 30 * 56.7, MR = 15 * 56.7;
const TEXT_W = Math.round(PAGE_W - ML - MR);

// ---------- References ----------
const REFS = {
  zakon: "Қазақстан Республикасының 2007 жылғы 27 шілдедегі № 319-III «Білім туралы» Заңы (өзгерістер мен толықтыруларымен).",
  tokayev: "Тоқаев Қ.-Ж. Абай және ХХІ ғасырдағы Қазақстан // Егемен Қазақстан. – 2020. – Қаңтар.",
  blasi: "Blasi A. Bridging moral cognition and moral action: A critical review of the literature // Psychological Bulletin. – 1980. – Vol. 88, No. 1. – P. 1–45.",
  abai: "Абай (Ибраһим Құнанбайұлы). Шығармаларының екі томдық толық жинағы. – Алматы: Жазушы, 1995. – Т. 1: Өлеңдер мен аудармалар; Т. 2: Поэмалар, аудармалар, қара сөздер.",
  auezov: "Әуезов М. Абай Құнанбайұлы: монографиялық зерттеу. – Алматы: Санат, 1995.",
  silchenko: "Сильченко М.С. Творческая биография Абая. – Алма-Ата: Казгослитиздат, 1957.",
  lotman_bio: "Лотман Ю.М. Александр Сергеевич Пушкин: Биография писателя. – 2-е изд. – Л.: Просвещение, 1983.",
  myrzakhmet: "Мырзахметов М. Абай және Шығыс. – Алматы: Қазақстан, 1994.",
  esim: "Есім Ғ. Хакім Абай. – Алматы: Ғылым, 1994.",
  zhumabaev: "Жұмабаев М. Педагогика. – Алматы: Ана тілі, 1992.",
  lotman_eo: "Лотман Ю.М. Роман А.С. Пушкина «Евгений Онегин»: Комментарий. – Л.: Просвещение, 1980.",
  blagoy: "Благой Д.Д. Творческий путь Пушкина (1826–1830). – М.: Советский писатель, 1967.",
  birtutas: "Оқушыларды тәрбиелеудің «Біртұтас тәрбие» бағдарламасы. – Астана: ҚР Оқу-ағарту министрлігі, 2023.",
  kohlberg: "Kohlberg L. Essays on Moral Development. Vol. 1: The Philosophy of Moral Development. – San Francisco: Harper & Row, 1981.",
  rest: "Rest J.R. Moral Development: Advances in Research and Theory. – New York: Praeger, 1986.",
  lickona: "Lickona T. Educating for Character: How Our Schools Can Teach Respect and Responsibility. – New York: Bantam Books, 1991.",
  bandura: "Bandura A. Social Learning Theory. – Englewood Cliffs, NJ: Prentice Hall, 1977.",
  vygotsky: "Выготский Л.С. Педагогическая психология / под ред. В.В. Давыдова. – М.: Педагогика, 1991.",
  gollwitzer: "Gollwitzer P.M. Implementation intentions: Strong effects of simple plans // American Psychologist. – 1999. – Vol. 54, No. 7. – P. 493–503.",
  steinberg: "Steinberg L., Monahan K.C. Age differences in resistance to peer influence // Developmental Psychology. – 2007. – Vol. 43, No. 6. – P. 1531–1543.",
  mccabe: "McCabe D.L., Treviño L.K., Butterfield K.D. Cheating in academic institutions: A decade of research // Ethics & Behavior. – 2001. – Vol. 11, No. 3. – P. 219–232.",
  kowalski: "Kowalski R.M., Giumetti G.W., Schroeder A.N., Lattanner M.R. Bullying in the digital age: A critical review and meta-analysis of cyberbullying research among youth // Psychological Bulletin. – 2014. – Vol. 140, No. 4. – P. 1073–1137.",
  pushkin: "Пушкин А.С. Полное собрание сочинений: в 10 т. – 4-е изд. – Л.: Наука, 1977–1979.",
  krippendorff: "Krippendorff K. Content Analysis: An Introduction to Its Methodology. – 2nd ed. – Thousand Oaks, CA: Sage, 2004.",
  miller: "Miller G.A. The magical number seven, plus or minus two: Some limits on our capacity for processing information // Psychological Review. – 1956. – Vol. 63, No. 2. – P. 81–97.",
  gadamer: "Гадамер Х.-Г. Истина и метод: Основы философской герменевтики. – М.: Прогресс, 1988.",
};

// ---------- Tables (data) ----------
const TABLES = {
  theories: {
    caption: "Құндылықты білу мен әрекет ету байланысы туралы негізгі тұжырымдар",
    widths: [22, 45, 33],
    header: ["Автор", "Негізгі тұжырым", "Біздің зерттеуге қатысы"],
    rows: [
      ["Л. Колберг [@kohlberg]", "Жасөспірімдер көбіне «конвенционалдық» деңгейде: топтың мақұлдауына бағдарланады", "Топ қысымына осалдықты түсіндіреді"],
      ["А. Блази [@blasi]", "Моральдық ойлау мен әрекеттің байланысы орташа; көпір – моральдық бірегейлік", "«Білемін, бірақ істей алмаймын» мәселесінің негізі"],
      ["Дж. Рест [@rest]", "Моральдық әрекеттің 4 бөлігі: сезімталдық, пайымдау, уәж, мінез", "Алгоритм қадамдарының құрылымы"],
      ["Т. Ликона [@lickona]", "Мінез – жақсылықты білу, қалау және істеу", "Абайдың ақыл–жүрек–қайрат үштігімен сәйкес"],
      ["А. Бандура [@bandura]", "Мінез-құлық үлгіні бақылау арқылы үйреніледі", "Пушкин кейіпкерлері – көркем үлгі"],
      ["Л.С. Выготский [@vygotsky]", "Оқушы өзі тәрбиеленетін орта ұйымдастырылуы керек", "Дайын кеңес емес, құрал ұсыну"],
      ["П. Голвитцер [@gollwitzer]", "«Егер – онда» жоспары мақсатқа жетуді жеңілдетеді", "«Таңда да, істе» қадамы"],
      ["Л. Стейнберг, К. Монахан [@steinberg]", "Құрдастар ықпалына қарсы тұру 14–18 жас аралығында күшейеді", "«Тоқта» қадамының қажеттілігі"],
      ["Д. МакКейб т.б. [@mccabe]", "Көшіріп жазуға құрдастардың мінезі ең күшті ықпал етеді", "1-жағдаят (көшіріп жазу)"],
      ["Р. Ковальски т.б. [@kowalski]", "Кибербуллинг эмпатияның төмендігімен байланысты", "2-жағдаят (желідегі қорлау)"],
    ],
  },
  corpus: {
    caption: "Зерттеу корпусының құрамы",
    widths: [8, 42, 50],
    header: ["№", "Абай", "А.С. Пушкин"],
    rows: [
      ["1", "7-қара сөз", "«Капитанская дочка» (1836)"],
      ["2", "15-қара сөз", "«Евгений Онегин» (1823–1831)"],
      ["3", "17-қара сөз", "«К Чаадаеву» (1818)"],
      ["4", "19-қара сөз", "«19 октября» (1825)"],
      ["5", "25-қара сөз", "«Я памятник себе воздвиг нерукотворный…» (1836)"],
      ["6", "32-қара сөз", "«Сказка о рыбаке и рыбке» (1833)"],
      ["7", "37-қара сөз", "«Моцарт и Сальери» (1830)"],
      ["8", "«Ғылым таппай мақтанба»", "«Няне» (1826)"],
      ["9", "«Интернатта оқып жүр»", "«Сказка о попе и о работнике его Балде» (1830)"],
      ["10", "«Жасымда ғылым бар деп ескермедім»", "«Воспоминание» (1828)"],
      ["11", "«Әсемпаз болма әрнеге»", "«Сказка о мёртвой царевне и о семи богатырях» (1833)"],
    ],
  },
  codes: {
    caption: "Мәтіндерді талдаудың кодтау жүйесі",
    widths: [22, 43, 35],
    header: ["Код", "Анықтамасы", "Мәтіннен мысал"],
    rows: [
      ["Б – Білім, ақыл", "Білімге құштарлық, оқудың, ойлаудың маңызы", "«Білсем екен» деген жан құмары (Абай, 7-қара сөз)"],
      ["А – Ар, адалдық", "Ар-намысты сақтау, берілген сөзге адалдық, өтіріктен аулақ болу", "«Береги честь смолоду» (Пушкин)"],
      ["М – Мейірім, рақым", "Өзгеге жанашырлық, кешірімділік, жақсылық жасау", "«Милость к падшим» (Пушкин)"],
      ["Е – Еңбек, қайрат", "Табандылық, ерік-жігер, істі бастау және аяқтау", "«Талап, еңбек, терең ой» (Абай)"],
      ["Р – Өзіне есеп беру", "Өз әрекетін талдау, қателікті мойындау", "«Өзіңнен өзің есеп ал» (Абай, 15-қара сөз)"],
      ["Қ – Достық, ел, қоғам", "Достыққа адалдық, қоғамға және отанға қызмет", "«Сен де бір кірпіш дүниеге» (Абай)"],
      ["Ж – Жаман әдеттен сақтану", "Ашкөздік, қызғаныш, мақтан, еріншектік, өсектің зиянын көрсету", "«Сказка о рыбаке и рыбке» (Пушкин)"],
    ],
  },
  compare: {
    caption: "Абай мен Пушкиннің тәрбиелік ұстанымдарын салыстыру",
    widths: [19, 27, 27, 27],
    header: ["Өлшем", "Абай", "Пушкин", "Ортақ белгі"],
    rows: [
      ["Тәрбиенің бастауы", "Білім мен ақыл («жан құмары»)", "Ар мен ізгі сезім («честь», «чувства добрые»)", "Жастық – мінез қалыптасатын шешуші кезең"],
      ["Идеал адам", "Толық адам: ақыл, қайрат, жүрек", "Арын сақтаған, мейірімді, адал адам (Гринёв, Татьяна)", "Адамгершілік білімнен жоғары"],
      ["Жамандық түрлері", "Өсек, өтірік, мақтан, еріншек, бекер мал шашпақ", "Ашкөздік, қызғаныш, опасыздық, жалқаулық", "Жаман мінез адамды құлдыратады"],
      ["Тәрбие тәсілі", "Тура үндеу, ақыл-кеңес, тізім", "Көркем үлгі, кейіпкердің таңдауы", "Оқырманның ар-ожданына үндеу"],
      ["Өзіне есеп беру", "«Өзіңнен өзің есеп ал» (15-қара сөз)", "«Строк печальных не смываю» («Воспоминание»)", "Кемелдену қателікті мойындаудан басталады"],
      ["Қоғаммен байланыс", "«Сен де бір кірпіш дүниеге»", "«Отчизне посвятим души прекрасные порывы»", "Жеке тәрбие қоғамға қызметпен ұштасады"],
    ],
  },
  tokta: {
    caption: "«ТОҚТА» алгоритмінің қадамдары және олардың негіздемесі",
    widths: [17, 29, 29, 25],
    header: ["Қадам", "Не істеу керек", "Әдеби негізі", "Ғылыми негізі"],
    rows: [
      ["Т – Тоқта", "10 секунд кідіріп, бірден жауап бермеу", "«Ақырын жүріп, анық бас» (Абай)", "Өзін-өзі реттеу; топ ықпалына осалдық [@steinberg]"],
      ["О – Ойлан", "Қай құндылық сынға түскенін атау", "Бес асыл іс пен бес дұшпан (Абай); «честь» (Пушкин)", "Моральдық сезімталдық [@rest]"],
      ["Қ – Қайтарымын көр", "Салдарды үш деңгейде бағалау: өзіме, өзгеге, ертеңіме", "Гринёв пен Швабриннің тағдыры (Пушкин)", "Моральдық пайымдау [@rest]"],
      ["Т – Таңда да, істе", "«Егер …, онда мен …» түріндегі нақты қадам", "Қайрат (Абай, 17-қара сөз); Гринёвтің таңдауы", "Іске асыру ниеті [@gollwitzer]; моральдық мінез"],
      ["А – Артынан есеп ал", "Кешке 3 сұрақ: не істедім? неге? келесіде не өзгертемін?", "«Өзіңнен өзің есеп ал» (Абай); «Воспоминание» (Пушкин)", "Рефлексия; моральдық бірегейлік [@blasi]"],
    ],
  },
  cases: {
    caption: "«ТОҚТА» алгоритмін үш жағдаятқа қолдану",
    widths: [16, 28, 28, 28],
    header: ["Қадам", "Көшіріп жазу", "Желідегі қорлау", "Топ қысымы"],
    rows: [
      ["Т – Тоқта", "Телефонды бірден ашпау", "Бірден эмодзи қоймау", "Бірден «иә» демей, уақыт алу"],
      ["О – Ойлан", "Адалдық, өз еңбегін құрметтеу", "Мейірім, ар", "Өз шешіміне адалдық, шынайы достық"],
      ["Қ – Қайтарымын көр", "Білім алынбайды, сенім жоғалады", "Адамға жарақат түседі, жазылған сөз өшпейді", "Сабақ пен сенім жоғалады, «бәрі істеді» сылтау емес"],
      ["Т – Таңда да, істе", "«Телефонды сөмкеге салып, түсінбегенімді мұғалімнен сұраймын»", "«Жәбірленушіге қолдау хат жазып, ересекке айтамын»", "«Бармаймын, бірақ сабақтан кейін бірге барайық» деп балама ұсыну"],
      ["А – Артынан есеп ал", "Қиын тақырыпқа алдын ала қалай дайындаламын?", "Үнсіз қалдым ба, әлде көмектестім бе?", "Өз шешімімді айта алдым ба?"],
      ["Классиктер тірегі", "Еріншек / еңбек (Абай); «как-нибудь» (Пушкин)", "Өсек / рақым (Абай); «милость к падшим» (Пушкин)", "Бекер мал шашпақ (Абай); лицей достығы (Пушкин)"],
    ],
  },
  pilot: {
    caption: "Пилоттық апробация нәтижелері (оқушы толтырады)",
    widths: [46, 18, 18, 18],
    header: ["Көрсеткіш", "Алдын ала", "Қайталама", "Өзгеріс"],
    rows: [
      ["Қатысушылар саны (n)", "[___]", "[___]", "–"],
      ["Орташа балл (0–2): 1-жағдаят, көшіріп жазу", "[___]", "[___]", "[___]"],
      ["Орташа балл (0–2): 2-жағдаят, желідегі қорлау", "[___]", "[___]", "[___]"],
      ["Орташа балл (0–2): 3-жағдаят, топ қысымы", "[___]", "[___]", "[___]"],
      ["Орташа балл (0–2): барлық жағдаят", "[___]", "[___]", "[___]"],
      ["2 балл алған жауаптардың үлесі, %", "[___]", "[___]", "[___]"],
      ["Алгоритмнің түсініктілігі (1–5), орташа", "–", "[___]", "–"],
    ],
  },
  enemies: {
    caption: "Абайдың «бес дұшпаны» мен «бес асыл ісінің» бүгінгі көрінісі",
    widths: [20, 42, 38],
    header: ["Абай ұғымы", "Бүгінгі жасөспірім өміріндегі көрінісі", "Ұсынылатын әрекет"],
    rows: [
      ["Өсек", "Чатта біреу туралы тексерілмеген хабар тарату", "Тексерілмеген хабарды бөліспеу"],
      ["Өтірік", "Мұғалімге, ата-анаға сылтау айту; жалған аккаунт ашу", "Қателікті ашық мойындау"],
      ["Мақтаншақ", "«Лайк» пен қаралым үшін өмірін әсірелеп көрсету", "Жетістікті еңбекпен өлшеу"],
      ["Еріншек", "Дайын жауапты көшіру, тапсырманы кейінге қалдыру", "Тапсырманы шағын бөліктерге бөліп орындау"],
      ["Бекер мал шашпақ", "Ақша мен уақытты ойын мен қажетсіз сатып алуға жұмсау", "Уақыт пен қаражатты жоспарлау"],
      ["Талап", "Мақсат қою, жаңа нәрсеге ұмтылу", "Апталық мақсатты жазып қою"],
      ["Еңбек", "Тапсырманы өз бетімен орындау", "Алдымен өзі ойланып, кейін көмек сұрау"],
      ["Терең ой", "Ақпаратты сын тұрғысынан тексеру", "Кемінде екі дереккөзді салыстыру"],
      ["Қанағат", "Бар нәрсеге риза болу, өзгемен жарыспау", "Өзін кешегі өзімен салыстыру"],
      ["Рақым", "Әлсізге, жаңа оқушыға қолдау көрсету", "Жалғыз қалған адамды әңгімеге тарту"],
    ],
  },
  implement: {
    caption: "Модельді мектеп тәжірибесінде пайдалану жолдары",
    widths: [22, 38, 40],
    header: ["Кім пайдаланады", "Қалай пайдаланады", "Нақты мысал"],
    rows: [
      ["Әдебиет пәнінің мұғалімі", "Шығарманы талдауда кейіпкердің таңдауын «ТОҚТА» қадамдары бойынша бағалау", "«Капитан қызы»: Гринёвтің Пугачёвқа ант беруден бас тартуы"],
      ["Сынып жетекшісі", "Сынып сағатында жағдаяттық және рөлдік ойындар өткізу", "«Чаттағы қорлау» жағдаятын рөлмен ойнап, шешім іздеу"],
      ["Мектеп психологы", "Топ қысымына қарсы тұру тренингі", "Әр оқушының «егер – онда» жоспарын жазу жаттығуы"],
      ["Оқушының өзі", "Жадынама карточкасы және кешкі рефлексия", "Күнделікке «бүгінгі есеп» жазу (Абай, 15-қара сөз)"],
      ["Ата-аналар", "Отбасылық әңгіме, шығармаларды бірге оқу", "Абайдың бес асыл ісін отбасылық келісімге айналдыру"],
    ],
  },
  memo: {
    caption: null,
    widths: [12, 30, 58],
    header: ["Әріп", "Қадам", "Өзіңе қоятын сұрақ"],
    rows: [
      ["Т", "ТОҚТА", "Мен асығып тұрған жоқпын ба? 10 рет дем алайын."],
      ["О", "ОЙЛАН", "Бұл жерде қай құндылық сынға түсіп тұр: адалдық па, мейірім бе, ар ма, достық па?"],
      ["Қ", "ҚАЙТАРЫМЫН КӨР", "Бұл әрекеттің салдары өзіме, өзгеге және ертеңіме қандай болады?"],
      ["Т", "ТАҢДА ДА, ІСТЕ", "Мен нақты не істеймін? «Егер …, онда мен …»"],
      ["А", "АРТЫНАН ЕСЕП АЛ", "Не істедім? Неге? Келесіде не өзгертемін?"],
    ],
    big: true,
  },
  lesson: {
    caption: null,
    widths: [30, 12, 58],
    header: ["Кезең", "Уақыты", "Мазмұны"],
    rows: [
      ["1. Ұйымдастыру және алдын ала диагностика", "10 мин", "А қосымшасындағы 1–3 жағдаяттарға жазбаша жауап"],
      ["2. Мәтінмен жұмыс", "12 мин", "«Ғылым таппай мақтанба» өлеңі және «Капитан қызы» повесінен үзінді (әкесінің өсиеті, Гринёвтің таңдауы) бойынша топтық талдау: «Кейіпкер қай құндылықты таңдады? Неге?»"],
      ["3. «ТОҚТА» алгоритмімен танысу", "8 мин", "Ә қосымшасындағы жадынама; мұғалімнің бір жағдаятты үлгі ретінде талдауы"],
      ["4. Жұптық жаттығу", "10 мин", "Жұптар бір жағдаятты алгоритм бойынша талдап, «егер – онда» жоспарын жазады"],
      ["5. Рефлексия", "5 мин", "«Бүгін не үйрендім? Алгоритмнің қай қадамы маған ең қиын?»"],
      ["2-сабақ (бір аптадан кейін)", "15 мин", "Қайталама диагностика: А қосымшасындағы жаңа 1–3 жағдаяттар және алгоритм туралы пікір"],
    ],
  },
};

// Coding matrix
const THEMES = ["Б", "А", "М", "Е", "Р", "Қ", "Ж"];
const MATRIX = {
  abai: [
    ["7-қара сөз", "Б"], ["15-қара сөз", "БР"], ["17-қара сөз", "БМЕ"], ["19-қара сөз", "БР"],
    ["25-қара сөз", "БҚ"], ["32-қара сөз", "БЕЖ"], ["37-қара сөз", "АЕ"],
    ["«Ғылым таппай мақтанба»", "БМЕЖ"], ["«Интернатта оқып жүр»", "БАЕҚ"],
    ["«Жасымда ғылым бар деп ескермедім»", "БР"], ["«Әсемпаз болма әрнеге»", "ЕҚЖ"],
  ],
  push: [
    ["«Капитанская дочка»", "АМЕРҚЖ"], ["«Евгений Онегин»", "БАРЖ"], ["«К Чаадаеву»", "АҚ"],
    ["«19 октября»", "РҚ"], ["«Я памятник себе воздвиг…»", "АМҚ"], ["«Сказка о рыбаке и рыбке»", "МЖ"],
    ["«Моцарт и Сальери»", "АЕЖ"], ["«Няне»", "М"], ["«Сказка о попе и о работнике его Балде»", "ЕЖ"],
    ["«Воспоминание»", "АР"], ["«Сказка о мёртвой царевне…»", "МЖ"],
  ],
};

// ---------- Citation numbering ----------
const content = fs.readFileSync(path.join(DIR, "content.txt"), "utf8").split("\n").filter(l => l.trim() !== "");
const order = [];
function noteCites(s) {
  for (const m of s.matchAll(/\[@([a-z_]+)\]/g)) if (!order.includes(m[1])) order.push(m[1]);
}
for (const line of content) {
  noteCites(line);
  const t = line.match(/^!table (\w+)/);
  if (t && TABLES[t[1]]) TABLES[t[1]].rows.forEach(r => r.forEach(noteCites));
}
for (const k of order) if (!REFS[k]) throw new Error("Missing ref " + k);
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) console.error("Unused refs:", unused);
const cite = s => s.replace(/\[@([a-z_]+)\]/g, (_, k) => `[${order.indexOf(k) + 1}]`);

// ---------- Helpers ----------
const HL_RE = /(\[___\]|\[оқушының[^\]]*\])/;
function runs(text, opts = {}) {
  text = cite(text);
  const parts = text.split(HL_RE).filter(p => p !== "");
  return parts.map(p => new TextRun({
    text: p, font: FONT, size: opts.size || SZ, bold: opts.bold, italics: opts.italics,
    highlight: HL_RE.test(p) || opts.hl ? "yellow" : undefined, color: opts.color,
  }));
}
function para(text, o = {}) {
  return new Paragraph({
    alignment: o.align ?? AlignmentType.JUSTIFIED,
    indent: o.noIndent ? undefined : { firstLine: Math.round(1.25 * CM), ...(o.indent || {}) },
    spacing: { line: 240, before: o.before || 0, after: o.after || 0 },
    keepNext: o.keepNext,
    children: runs(text, o),
  });
}
const tableCaptions = {}; let tNo = 0, fNo = 0;
const figNums = {};

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, w, o = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: o.head ? { fill: "E7EEF5", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 90, right: 90 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { line: 240 },
      children: runs(text, { size: o.size || SZ_T, bold: o.bold || o.head }),
    })],
  });
}
function mkTable(widthsPct, header, rows, o = {}) {
  const widths = widthsPct.map(p => Math.floor(TEXT_W * p / 100));
  widths[widths.length - 1] += TEXT_W - widths.reduce((a, b) => a + b, 0);
  const trs = [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { head: true, center: true, size: o.size })) })];
  rows.forEach(r => trs.push(new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], { center: o.centerCols?.includes(i), bold: o.boldRow?.(r) || (o.boldCol0 && i === 0), size: o.size })),
  })));
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
function tableBlock(name) {
  const out = [];
  if (name === "matrix") {
    tNo++;
    out.push(para(`${tNo}-кесте – Абай мен Пушкин мәтіндерін кодтау матрицасы`, { noIndent: true, keepNext: true, before: 120, after: 60, align: AlignmentType.LEFT }));
    const rows = [];
    const sum = (list) => THEMES.map(t => String(list.filter(([, c]) => c.includes(t)).length));
    const add = (list, start) => list.forEach(([n, c], i) => rows.push([String(start + i), n, ...THEMES.map(t => c.includes(t) ? "+" : "")]));
    rows.push(["", "АБАЙ", "", "", "", "", "", "", ""]);
    add(MATRIX.abai, 1);
    rows.push(["", "Абай бойынша барлығы", ...sum(MATRIX.abai)]);
    rows.push(["", "ПУШКИН", "", "", "", "", "", "", ""]);
    add(MATRIX.push, 12);
    rows.push(["", "Пушкин бойынша барлығы", ...sum(MATRIX.push)]);
    rows.push(["", "Жалпы", ...sum([...MATRIX.abai, ...MATRIX.push])]);
    out.push(mkTable([6, 44, 7, 7, 7, 7, 7, 7, 8], ["№", "Мәтін бірлігі", ...THEMES], rows,
      { centerCols: [0, 2, 3, 4, 5, 6, 7, 8], boldRow: r => /АБАЙ|ПУШКИН|барлығы|Жалпы/.test(r[1]), size: 22 }));
    out.push(para("Ескерту: Б – білім, ақыл; А – ар, адалдық; М – мейірім; Е – еңбек, қайрат; Р – өзіне есеп беру; Қ – достық, ел, қоғам; Ж – жаман әдеттен сақтану.", { noIndent: true, size: 22, before: 60, after: 120 }));
    return out;
  }
  const T = TABLES[name];
  if (T.caption) {
    tNo++;
    out.push(para(`${tNo}-кесте – ${T.caption}`, { noIndent: true, keepNext: true, before: 120, after: 60, align: AlignmentType.LEFT }));
  }
  const o = { boldCol0: true };
  if (T.big) Object.assign(o, { size: 28, centerCols: [0] });
  if (name === "corpus") Object.assign(o, { boldCol0: false, centerCols: [0] });
  if (name === "pilot") Object.assign(o, { boldCol0: false, centerCols: [1, 2, 3] });
  out.push(mkTable(T.widths, T.header, T.rows, o));
  out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  return out;
}
function figBlock(name, caption) {
  fNo++;
  const file = path.join(DIR, "figs", name + ".png");
  const buf = fs.readFileSync(file);
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const maxW = 600, maxH = 560;
  let dw = maxW, dh = Math.round(maxW * h / w);
  if (dh > maxH) { dh = maxH; dw = Math.round(maxH * w / h); }
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: dw, height: dh },
        altText: { title: caption, description: caption, name } })] }),
    para(`${fNo}-сурет – ${caption}`, { noIndent: true, align: AlignmentType.CENTER, after: 160, before: 60 }),
  ];
}
function h1(text, first) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, pageBreakBefore: !first,
    spacing: { after: 240 }, keepNext: true,
    children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: Math.round(1.25 * CM) }, spacing: { before: 240, after: 120 }, keepNext: true,
    children: [new TextRun({ text, font: FONT, size: SZ, bold: true })],
  });
}

// ---------- Front matter ----------
const front = [];
const C = AlignmentType.CENTER;
const line = (t, o = {}) => new Paragraph({ alignment: o.align ?? C, spacing: { line: 240, before: o.before || 0, after: o.after || 0 },
  indent: o.indent, children: runs(t, o) });
front.push(
  line("Қазақстан Республикасы Оқу-ағарту министрлігі"),
  line("«Дарын» республикалық ғылыми-практикалық орталығы", { after: 240 }),
  line("Жалпы білім беретін мектептердің 8–11 сынып оқушыларының", { }),
  line("ғылыми жобалар конкурсы", { after: 1400 }),
  line("Секция: Қазақ әдебиеті", { after: 700 }),
  line("Тақырыбы:", { bold: true }),
  line("АБАЙ МЕН ПУШКИН ШЫҒАРМАЛАРЫНДАҒЫ ЖАС ҰРПАҚ ТӘРБИЕСІ ЖӘНЕ ОНЫҢ ҚАЗІРГІ ҚОҒАМДАҒЫ МАҢЫЗЫ", { bold: true, after: 1400 }),
  line("Орындаған: [Оқушының аты-жөні],", { align: AlignmentType.LEFT, indent: { left: 4800 }, hl: true }),
  line("[__] сынып оқушысы,", { align: AlignmentType.LEFT, indent: { left: 4800 }, hl: true }),
  line("[Мектептің толық атауы]", { align: AlignmentType.LEFT, indent: { left: 4800 }, hl: true, after: 240 }),
  line("Ғылыми жетекшісі: [Аты-жөні],", { align: AlignmentType.LEFT, indent: { left: 4800 }, hl: true }),
  line("[қазақ тілі мен әдебиеті пәнінің мұғалімі]", { align: AlignmentType.LEFT, indent: { left: 4800 }, hl: true, after: 2600 }),
  line("[Қала] – 2026", { hl: true }),
);

const annot = [
  h1("АҢДАТПА", false),
  para("Жұмыста Абай мен А.С. Пушкин шығармаларындағы жас ұрпақ тәрбиесіне қатысты ұстанымдар 22 мәтін бірлігі негізінде жеті тақырыптық код бойынша салыстырылды. Екі ақынның тәрбиелік өрісі ортақ екені, ал айырмашылық екпін мен тәсілде екені анықталды: Абай білім мен еңбекті тура үндеу арқылы, Пушкин ар-намыс пен мейірімді кейіпкердің таңдауы арқылы насихаттайды. Нәтижелер негізінде «Ар – Ақыл – Жүрек» тәрбие моделі және жасөспірімге арналған бес қадамды «ТОҚТА» шешім алгоритмі ұсынылып, көшіріп жазу, желідегі қорлау және топ қысымы жағдаяттарына қолданылды."),
  new Paragraph({ alignment: C, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: "АННОТАЦИЯ", font: FONT, size: SZ, bold: true })] }),
  para("В работе на материале 22 текстовых единиц по семи тематическим кодам сопоставлены воспитательные принципы в произведениях Абая и А.С. Пушкина. Установлено, что содержательное поле воспитательных идей двух поэтов совпадает, а различия проявляются в акцентах и способах воздействия: Абай утверждает ценность знания и труда через прямое назидание, Пушкин – ценность чести и милосердия через выбор героя. На основе результатов предложены воспитательная модель «Честь – Разум – Сердце» и пятишаговый алгоритм принятия решений «ТОҚТА» («Остановись») для подростков, применённый к ситуациям списывания, кибербуллинга и группового давления."),
  new Paragraph({ alignment: C, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: "ABSTRACT", font: FONT, size: SZ, bold: true })] }),
  para("The study compares the educational principles found in the works of Abai and A.S. Pushkin, using 22 text units coded against seven thematic categories. The two poets share the same range of moral values but differ in emphasis and method: Abai promotes knowledge and diligence through direct instruction, while Pushkin promotes honour and mercy through the choices of his characters. Based on these findings, the paper proposes the “Honour – Mind – Heart” model of upbringing and a five-step decision algorithm for teenagers, “TOQTA” (“Stop”), applied to cheating, cyberbullying and peer pressure situations."),
];

// ---------- TOC (manual, page numbers from pass 1) ----------
const tocEntries = [];
for (const l of content) {
  let m;
  if ((m = l.match(/^#\*? (.*)$/))) tocEntries.push({ t: m[1], lvl: 1 });
  else if ((m = l.match(/^## (.*)$/))) tocEntries.push({ t: m[1], lvl: 2 });
}
const toc = [h1("МАЗМҰНЫ", false)];
for (const e of tocEntries) {
  const pg = TOC_PAGES[e.t] ?? "";
  toc.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W, leader: "dot" }],
    indent: e.lvl === 2 ? { left: 400, hanging: 0 } : undefined,
    spacing: { line: 240, after: 40 },
    children: [new TextRun({ text: e.t, font: FONT, size: SZ }), new TextRun({ text: `\t${pg}`, font: FONT, size: SZ })],
  }));
}

// ---------- Body ----------
const body = [];
for (const l of content) {
  let m;
  if ((m = l.match(/^#\*? (.*)$/))) body.push(h1(m[1], false));
  else if ((m = l.match(/^## (.*)$/))) body.push(h2(m[1]));
  else if ((m = l.match(/^> (.*)$/))) body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: Math.round(1.25 * CM), right: 300 },
    spacing: { line: 240, before: 60, after: 60 }, children: runs(m[1], { italics: true }) }));
  else if ((m = l.match(/^-- (.*)$/))) body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, numbering: { reference: "dash", level: 0 },
    spacing: { line: 240 }, children: runs(m[1]) }));
  else if ((m = l.match(/^!fig (\w+) \| (.*)$/))) body.push(...figBlock(m[1], m[2]));
  else if ((m = l.match(/^!table (\w+)/))) body.push(...tableBlock(m[1]));
  else if ((m = l.match(/^!note (.*)$/))) body.push(para("НҰСҚАУ: " + m[1], { bold: true, hl: true, after: 120 }));
  else if (l.startsWith("!references")) order.forEach((k, i) => body.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 }, spacing: { line: 240, after: 60 },
    children: [new TextRun({ text: `${i + 1}\t${REFS[k]}`, font: FONT, size: SZ })],
    tabStops: [{ type: TabStopType.LEFT, position: 567 }] })));
  else body.push(para(l));
}

const doc = new Document({
  creator: "", title: "Абай мен Пушкин шығармаларындағы жас ұрпақ тәрбиесі",
  styles: {
    default: { document: { run: { font: FONT, size: SZ } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "dash", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(1.25 * CM) + 360, hanging: 360 } } } }] }] },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: PAGE_W, height: 16838 },
        margin: { top: 1134, bottom: 1134, left: Math.round(ML), right: Math.round(MR) } },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: C, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }),
      first: new Footer({ children: [] }),
    },
    children: [...front, ...annot, ...toc, ...body],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("written", OUT, "refs:", order.length); });
