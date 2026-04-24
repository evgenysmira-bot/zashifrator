// Ядро маскировки: детекторы, валидаторы и генераторы замен.
// Зависит только от стандартного JS — без Office.js, легко тестировать.

(function (global) {
  'use strict';

  // ---------- ИНН 10 (юрлицо) ----------
  function validateInn10(s) {
    if (!/^\d{10}$/.test(s)) return false;
    const w = [2, 4, 10, 3, 5, 9, 4, 6, 8];
    let sum = 0;
    for (let i = 0; i < 9; i++) sum += (+s[i]) * w[i];
    return ((sum % 11) % 10) === +s[9];
  }

  function generateInn10() {
    let base = '';
    for (let i = 0; i < 9; i++) base += Math.floor(Math.random() * 10);
    const w = [2, 4, 10, 3, 5, 9, 4, 6, 8];
    let sum = 0;
    for (let i = 0; i < 9; i++) sum += (+base[i]) * w[i];
    return base + ((sum % 11) % 10);
  }

  // ---------- ИНН 12 (ИП / физлицо) ----------
  function validateInn12(s) {
    if (!/^\d{12}$/.test(s)) return false;
    const w1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8];
    const w2 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8];
    let s1 = 0, s2 = 0;
    for (let i = 0; i < 10; i++) s1 += (+s[i]) * w1[i];
    for (let i = 0; i < 11; i++) s2 += (+s[i]) * w2[i];
    return ((s1 % 11) % 10) === +s[10] && ((s2 % 11) % 10) === +s[11];
  }

  function generateInn12() {
    let base = '';
    for (let i = 0; i < 10; i++) base += Math.floor(Math.random() * 10);
    const w1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8];
    let s1 = 0;
    for (let i = 0; i < 10; i++) s1 += (+base[i]) * w1[i];
    const d11 = (s1 % 11) % 10;
    const with11 = base + d11;
    const w2 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8];
    let s2 = 0;
    for (let i = 0; i < 11; i++) s2 += (+with11[i]) * w2[i];
    return with11 + ((s2 % 11) % 10);
  }

  // ---------- ОГРН (13) ----------
  function validateOgrn(s) {
    if (!/^\d{13}$/.test(s)) return false;
    const head = s.substring(0, 12);
    return Number((BigInt(head) % 11n) % 10n) === +s[12];
  }

  function generateOgrn() {
    let base = '';
    // первый символ — признак ОГРН (1, 5) — берём 1
    base += '1';
    for (let i = 0; i < 11; i++) base += Math.floor(Math.random() * 10);
    const check = Number((BigInt(base) % 11n) % 10n);
    return base + check;
  }

  // ---------- ОГРНИП (15) ----------
  function validateOgrnip(s) {
    if (!/^\d{15}$/.test(s)) return false;
    const head = s.substring(0, 14);
    return Number((BigInt(head) % 13n) % 10n) === +s[14];
  }

  function generateOgrnip() {
    let base = '3';
    for (let i = 0; i < 13; i++) base += Math.floor(Math.random() * 10);
    const check = Number((BigInt(base) % 13n) % 10n);
    return base + check;
  }

  // ---------- КПП / БИК / р-счёт / к-счёт ----------
  // Без контрольных сумм — просто форматные генераторы, сохраняющие длину и типичный префикс.
  function generateKpp() {
    let out = '';
    for (let i = 0; i < 4; i++) out += Math.floor(Math.random() * 10);   // код налоговой
    out += '01';                                                           // причина постановки (типовая)
    for (let i = 0; i < 3; i++) out += Math.floor(Math.random() * 10);    // порядковый
    return out;
  }
  function generateBik() {
    let out = '04';                                                        // код страны (Россия)
    for (let i = 0; i < 7; i++) out += Math.floor(Math.random() * 10);
    return out;
  }
  function generateRs() {
    let out = '40702810';                                                  // типовой префикс коммерч. счёта в рублях
    for (let i = 0; i < 12; i++) out += Math.floor(Math.random() * 10);
    return out;
  }
  function generateKs() {
    let out = '30101810';                                                  // корр. счёт в рублях
    for (let i = 0; i < 12; i++) out += Math.floor(Math.random() * 10);
    return out;
  }

  // ---------- ФИО ----------
  // Генерируем «И.И.Иванов_N» или «Иванов_N И.И.» — чтобы осталось видно, что это ФИО.
  const FAKE_SURNAMES = ['Иванов', 'Петров', 'Сидоров', 'Кузнецов', 'Смирнов',
                         'Попов', 'Лебедев', 'Козлов', 'Новиков', 'Морозов'];
  const FAKE_INITIALS = ['И.И.', 'П.П.', 'С.С.', 'А.А.', 'В.В.',
                         'Г.Г.', 'Д.Д.', 'Е.Е.', 'К.К.', 'М.М.'];
  // ---------- Адреса ----------
  let _addrCounter = 0;
  function nextAddress() {
    _addrCounter++;
    const n = _addrCounter;
    let zip = '';
    for (let i = 0; i < 6; i++) zip += Math.floor(Math.random() * 10);
    // Сохраняем «форму» адреса, чтобы документ читался.
    return `${zip}, г. Город_${n}, ул. Улица_${n}, д. ${n}, оф. ${n}`;
  }
  function resetAddrCounter(n) { _addrCounter = n | 0; }

  let _fioCounter = 0;
  function nextFio(kind) {
    _fioCounter++;
    const s = FAKE_SURNAMES[(_fioCounter - 1) % FAKE_SURNAMES.length];
    const ini = FAKE_INITIALS[(_fioCounter - 1) % FAKE_INITIALS.length];
    const n = _fioCounter;
    if (kind === 'fio_initials_first') return `${ini}${s}${n}`;
    if (kind === 'fio_initials_last')  return `${s}${n} ${ini}`;
    if (kind === 'fio_full')           return `${s}${n} Имя${n} Отчество${n}`;
    return `${s}${n}`;
  }
  function resetFioCounter(n) { _fioCounter = n | 0; }

  // ---------- Поиск в тексте ----------
  // Возвращает массив находок: { type, value, index, length, meta }
  function findAll(text) {
    const found = [];
    const taken = []; // [start, end) для исключения пересечений

    function overlaps(start, end) {
      return taken.some(([s, e]) => !(end <= s || start >= e));
    }
    function claim(start, end) { taken.push([start, end]); }

    // Порядок важен: сначала более длинные/строгие совпадения.
    const passes = [
      { re: /\b\d{15}\b/g, type: 'ogrnip', validate: validateOgrnip },
      { re: /\b\d{13}\b/g, type: 'ogrn',   validate: validateOgrn   },
      { re: /\b\d{12}\b/g, type: 'inn12',  validate: validateInn12  },
      { re: /\b\d{10}\b/g, type: 'inn10',  validate: validateInn10  },
    ];

    for (const p of passes) {
      let m;
      p.re.lastIndex = 0;
      while ((m = p.re.exec(text)) !== null) {
        if (!p.validate(m[0])) continue;
        if (overlaps(m.index, m.index + m[0].length)) continue;
        claim(m.index, m.index + m[0].length);
        found.push({ type: p.type, value: m[0], index: m.index, length: m[0].length });
      }
    }

    // Названия компаний. Поддерживаем аббревиатуры (ООО, АО, ПАО …) и
    // развёрнутые формы («Публичное акционерное общество» и т.п.).
    // Кавычки: «» (U+00AB/00BB), прямые " ', „ " " (U+201E/201C/201D), '' (U+2018/2019).
    // \b с кириллицей в JS не работает — используем lookbehind с явным списком букв.
    const OPEN = '[\\u00AB\\u0022\\u201C\\u201E\\u2018]';
    const CLOSE = '[\\u00BB\\u0022\\u201D\\u2019]';
    const PREFIXES = [
      'Общество с ограниченной ответственностью',
      'Автономная некоммерческая организация',
      'Непубличное акционерное общество',
      'Публичное акционерное общество',
      'Закрытое акционерное общество',
      'Открытое акционерное общество',
      'Индивидуальный предприниматель',
      'Акционерное общество',
      // Государственные / муниципальные учреждения и образовательные организации
      'Федеральное государственное автономное образовательное учреждение высшего образования',
      'Федеральное государственное бюджетное образовательное учреждение высшего образования',
      'Государственное автономное образовательное учреждение высшего образования',
      'Государственное бюджетное образовательное учреждение высшего образования',
      'Федеральное государственное автономное учреждение',
      'Федеральное государственное бюджетное учреждение',
      'Федеральное государственное унитарное предприятие',
      'Федеральное государственное казённое учреждение',
      'Федеральное государственное казенное учреждение',
      'Государственное автономное образовательное учреждение',
      'Государственное бюджетное образовательное учреждение',
      'Государственное автономное учреждение',
      'Государственное бюджетное учреждение',
      'Государственное унитарное предприятие',
      'Государственное казённое учреждение',
      'Государственное казенное учреждение',
      'Муниципальное бюджетное образовательное учреждение',
      'Муниципальное автономное образовательное учреждение',
      'Муниципальное бюджетное учреждение',
      'Муниципальное казённое учреждение',
      'Муниципальное казенное учреждение',
      'Муниципальное унитарное предприятие',
      'Муниципальное автономное учреждение',
      // Аббревиатуры кириллицей
      'ООО', 'ОАО', 'ПАО', 'ЗАО', 'НАО', 'АНО', 'АО', 'ИП',
      'ФГБУ', 'ФГУП', 'ФГКУ', 'ФГАУ',
      'ФГАОУ ВО', 'ФГБОУ ВО', 'ФГАОУ', 'ФГБОУ',
      'ГБУ', 'ГУП', 'ГКУ', 'ГАУ',
      'ГАОУ ВО', 'ГБОУ ВО', 'ГАОУ', 'ГБОУ',
      'МБУ', 'МКУ', 'МУП', 'МАУ',
      'МБОУ', 'МАОУ',
      'НИУ', 'НИИ',
      // Латинские двойники (часто встречаются в реальных документах)
      'OOO', 'OAO', 'PAO', 'ZAO', 'NAO', 'ANO', 'AO', 'IP',
      // Филиалы/представительства с названием в кавычках
      'Филиал', 'Представительство',
    ].sort((a, b) => b.length - a.length);
    const prefixAlt = PREFIXES.map(p => p.replace(/\s+/g, '\\s+')).join('|');
    // Между префиксом и кавычками может быть описание (напр.,
    // «АНО содействия развитию благотворительной деятельности «ТОЛК»»).
    // Разрешаем 0-10 промежуточных слов — но ТОЛЬКО на той же строке
    // (иначе регулярка перепрыгивает через абзацы и ломает вёрстку).
    // Промежуточные слова между префиксом и «кавычкой» НЕ должны содержать
    // скобки/запятые — иначе регулярка «переедет» через «ПАО Сбербанк),
    // именуемое в дальнейшем «Заказчик»» и схватит последнюю кавычку.
    const companyRe = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])(' + prefixAlt + ')' +
      '(?:[ \\t\\xa0]+[^\\s,()«»"""„"]+){0,10}?[ \\t\\xa0]*' +
      OPEN + '([^\\u00AB\\u00BB\\u0022\\u201C\\u201D\\u201E\\u2018\\u2019]{2,120})' + CLOSE,
      'gi'
    );
    let cm;
    while ((cm = companyRe.exec(text)) !== null) {
      const start = cm.index;
      const end = start + cm[0].length;
      if (overlaps(start, end)) continue;
      if (isPublicEntity(cm[0])) continue;
      claim(start, end);
      found.push({
        type: 'company',
        value: cm[0],
        index: start,
        length: cm[0].length,
        meta: { prefix: cm[1], name: cm[2] }
      });
    }

    // Компания БЕЗ кавычек: «ПАО Сбербанк», «АО Альфа-Банк» — короткие
    // аббревиатуры + одно слово с заглавной кириллицы (допускается дефис).
    // Имя — 1-3 заглавных слова через пробел/дефис.
    // Дополнительные слова разрешены только если за ними НЕ идёт `:` или `№`
    // (иначе это начало нового поля, а не часть имени: «Банк ИНН: 1234»).
    const companyNoQ = /(?<![A-Za-zА-Яа-яЁё0-9])(ПАО|ОАО|АО|ООО|ЗАО|НАО|АНО|PAO|OAO|AO|OOO|ZAO|NAO|ANO|Банк|Банка|Банку)[ \t]+([А-ЯЁ][А-ЯЁа-яё\-]{2,40}(?:[ \t]+[А-ЯЁ][А-ЯЁа-яё\-]{2,40}(?![ \t]*[:№])){0,2})(?![A-Za-zА-Яа-яЁё0-9])/g;
    let cnq;
    while ((cnq = companyNoQ.exec(text)) !== null) {
      const start = cnq.index, end = start + cnq[0].length;
      if (overlaps(start, end)) continue;
      if (isPublicEntity(cnq[0])) continue;
      claim(start, end);
      found.push({
        type: 'company',
        value: cnq[0],
        index: start,
        length: cnq[0].length,
        meta: { prefix: cnq[1], name: cnq[2] }
      });
    }

    // Длинная форма БЕЗ кавычек:
    // «Публичное акционерное общество Сбербанк России (в лице…)»
    // Берём префикс + 1-4 заглавных слова + lookahead на естественный ограничитель.
    const LONG_FORMS = [
      'Общество с ограниченной ответственностью',
      'Автономная некоммерческая организация',
      'Непубличное акционерное общество',
      'Публичное акционерное общество',
      'Закрытое акционерное общество',
      'Открытое акционерное общество',
      'Акционерное общество',
      'Муниципальное бюджетное образовательное учреждение',
      'Муниципальное автономное образовательное учреждение',
      'Муниципальное бюджетное учреждение',
      'Муниципальное казённое учреждение',
      'Муниципальное казенное учреждение',
      'Муниципальное унитарное предприятие',
      'Муниципальное автономное учреждение',
      'Государственное автономное образовательное учреждение',
      'Государственное бюджетное образовательное учреждение',
      'Государственное автономное учреждение',
      'Государственное бюджетное учреждение',
      'Государственное унитарное предприятие',
      'Государственное казённое учреждение',
      'Государственное казенное учреждение',
      'Федеральное государственное автономное образовательное учреждение высшего образования',
      'Федеральное государственное бюджетное образовательное учреждение высшего образования',
      'Федеральное государственное автономное учреждение',
      'Федеральное государственное бюджетное учреждение',
      'Федеральное государственное унитарное предприятие',
    ].map(p => p.replace(/\s+/g, '\\s+')).join('|');
    const companyLongNoQ = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])(' + LONG_FORMS + ')[ \\t]+' +
      '([А-ЯЁ][А-ЯЁа-яё\\-]{2,30}(?:[ \\t]+[А-ЯЁ][А-ЯЁа-яё\\-]{2,30}){0,3})' +
      '(?=\\s*(?:,|\\(|[\\n\\r]|$|[ \\t]+(?:в\\s+лице|в\\s+дальнейшем|именуем[а-яё]+)))',
      'g'
    );
    let clnq;
    while ((clnq = companyLongNoQ.exec(text)) !== null) {
      const start = clnq.index, end = start + clnq[0].length;
      if (overlaps(start, end)) continue;
      if (isPublicEntity(clnq[0])) continue;
      claim(start, end);
      found.push({
        type: 'company',
        value: clnq[0],
        index: start,
        length: clnq[0].length,
        meta: { prefix: clnq[1], name: clnq[2] }
      });
    }

    // ---------- Банковские реквизиты (с меткой) ----------
    // Маскируем ТОЛЬКО значение (группу 1), метку оставляем.
    // Для каждой находки: index — позиция группы, value — только число.
    // Важно: сначала К/с (корр. счёт) — иначе Р/с-регулярка поймает «р.счет» внутри «Кор.счет».
    // Lookbehind на не-букву, чтобы «р» из «Кор» не матчился как начало метки Р/с.
    const NL = '(?<![A-Za-zА-Яа-яЁё0-9])'; // not-letter lookbehind
    const labelledPatterns = [
      // ИНН/ОГРН с явной меткой — на случай, если у числа невалидная контрольная сумма
      // (тестовые заглушки, опечатки и т.п.). Чексам-валидаторы выше уже отработали;
      // overlap-check предотвратит дубль для валидных номеров.
      { type: 'inn10',  re: new RegExp(NL + '(ИНН)\\s*[:№]?\\s*(\\d{10})\\b(?!\\d)', 'gi') },
      { type: 'inn12',  re: new RegExp(NL + '(ИНН)\\s*[:№]?\\s*(\\d{12})\\b(?!\\d)', 'gi') },
      { type: 'ogrn',   re: new RegExp(NL + '(ОГРН)\\s*[:№]?\\s*(\\d{13})\\b(?!\\d)', 'gi') },
      { type: 'ogrnip', re: new RegExp(NL + '(ОГРНИП)\\s*[:№]?\\s*(\\d{15})\\b(?!\\d)', 'gi') },
      { type: 'kpp', re: new RegExp(NL + '(КПП)\\s*[:№]?\\s*(\\d{9})\\b', 'gi') },
      { type: 'bik', re: new RegExp(NL + '(БИК(?:\\s+банка)?)\\s*[:№]?\\s*(\\d{9})\\b', 'gi') },
      { type: 'ks',
        re:  new RegExp(NL + '(К(?:орр?)?(?:\\.?\\s*сч(?:е|ё)т|\\/с|\\.?\\s*с\\.?))\\s*[:№]?\\s*(\\d{20})\\b', 'gi') },
      { type: 'rs',
        re:  new RegExp(NL + '(Р(?:асч(?:е|ё)тный)?(?:[\\.\\/\\s]*с(?:ч(?:е|ё)т)?)?)\\s*[:№]?\\s*(\\d{20})\\b', 'gi') },
      // Альтернативные лейблы счетов: «Номер счета», «Счёт», «л/с», «Лицевой счёт»
      { type: 'rs',
        re:  new RegExp(NL + '(Номер\\s+сч(?:е|ё)та|Сч(?:е|ё)т|[Лл]/с|Лицев(?:ой)?\\s+сч(?:е|ё)т)\\s*[:№]?\\s*(\\d{20})\\b', 'gi') },
      // ОКПО — 8 или 10 цифр
      { type: 'okpo', re: new RegExp(NL + '(ОКПО)\\s*[:№]?\\s*(\\d{8}|\\d{10})\\b', 'gi') },
      // ОКВЭД — NN.NN(.NN)?
      { type: 'okved', re: new RegExp(NL + '(ОКВЭД(?:-?2)?)\\s*[:№]?\\s*(\\d{2}(?:\\.\\d{1,2}){1,2})', 'gi') },
      // ОКТМО — 8 или 11 цифр
      { type: 'oktmo', re: new RegExp(NL + '(ОКТМО)\\s*[:№]?\\s*(\\d{8}|\\d{11})\\b', 'gi') },
      // ОКАТО — 11 цифр
      { type: 'okato', re: new RegExp(NL + '(ОКАТО)\\s*[:№]?\\s*(\\d{11})\\b', 'gi') },
    ];

    // Паспорт РФ: «Паспорт 27 08 №123456», «паспорт: 2708 123456»
    const passportRe = /(?<![A-Za-zА-Яа-яЁё0-9])[Пп]аспорт(?:\s+[А-ЯЁ]{2,6})?\s*[:№]?\s*(\d{2}\s?\d{2})\s*№?\s*(\d{6})(?![\d])/g;
    let prm;
    while ((prm = passportRe.exec(text)) !== null) {
      // Захватываем только номерную часть (серия + номер), а не слово «Паспорт».
      const startNum = prm.index + prm[0].indexOf(prm[1]);
      const endNum = prm.index + prm[0].length;
      if (overlaps(startNum, endNum)) continue;
      claim(startNum, endNum);
      found.push({
        type: 'passport',
        value: text.substring(startNum, endNum),
        index: startNum,
        length: endNum - startNum
      });
    }
    for (const p of labelledPatterns) {
      p.re.lastIndex = 0;
      let mm;
      while ((mm = p.re.exec(text)) !== null) {
        const numStart = mm.index + mm[0].lastIndexOf(mm[2]);
        const numEnd = numStart + mm[2].length;
        if (overlaps(numStart, numEnd)) continue;
        claim(numStart, numEnd);
        found.push({ type: p.type, value: mm[2], index: numStart, length: mm[2].length });
      }
    }

    // ---------- ФИО ----------
    // Важно: между частями ФИО используем только ГОРИЗОНТАЛЬНЫЕ пробелы
    // (пробел, таб, NBSP). Если бы \s разрешал \r\n, регулярка могла бы
    // поймать «А.А.\r\nЭлектронная» как одну ФИО и при замене склеить два абзаца.
    const HS = '[ \\t\\xa0]';
    // Служебные слова — не могут быть фамилиями, даже если похожи.
    const ROLE_WORDS = new Set([
      'Получатель','Получателя','Получателю',
      'Плательщик','Плательщика','Плательщику',
      'Отправитель','Отправителя','Отправителю',
      'Заказчик','Заказчика','Заказчику',
      'Исполнитель','Исполнителя','Исполнителю',
      'Подрядчик','Подрядчика','Подрядчику',
      'Поставщик','Поставщика','Поставщику',
      'Покупатель','Покупателя','Покупателю',
      'Продавец','Продавца','Продавцу',
      'Сотрудник','Сотрудника','Сотруднику',
      'Работник','Работника','Работнику',
      'Представитель','Представителя','Представителю',
      'Гражданин','Гражданина','Гражданину',
      'Сторона','Стороны','Стороне',
      'Клиент','Клиента','Клиенту',
      'Контрагент','Контрагента','Контрагенту',
    ]);
    function isRoleWord(word) { return ROLE_WORDS.has(word); }
    // 1) Инициалы + фамилия: «А.В.Кукса», «А. В. Кукса»
    const fio1 = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])([А-ЯЁ])\\.' + HS + '*([А-ЯЁ])\\.' + HS + '*([А-ЯЁ][а-яё]{1,30})(?![A-Za-zА-Яа-яЁё0-9])',
      'g'
    );
    let fm;
    while ((fm = fio1.exec(text)) !== null) {
      const s = fm.index, e = s + fm[0].length;
      if (overlaps(s, e)) continue;
      if (isRoleWord(fm[3])) continue;   // «И.О. Получателя» — роль, не фамилия
      claim(s, e);
      found.push({ type: 'fio_initials', value: fm[0], index: s, length: fm[0].length,
                   meta: { kind: 'initials_first', surname: fm[3] } });
    }
    // 2) Фамилия + инициалы: «Кукса А.В.»
    const fio2 = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])([А-ЯЁ][а-яё]{1,30})' + HS + '+([А-ЯЁ])\\.' + HS + '*([А-ЯЁ])\\.(?![A-Za-zА-Яа-яЁё0-9])',
      'g'
    );
    while ((fm = fio2.exec(text)) !== null) {
      const s = fm.index, e = s + fm[0].length;
      if (overlaps(s, e)) continue;
      if (isRoleWord(fm[1])) continue;   // «Получателя А.А.» — роль, не фамилия
      claim(s, e);
      found.push({ type: 'fio_initials', value: fm[0], index: s, length: fm[0].length,
                   meta: { kind: 'initials_last', surname: fm[1] } });
    }
    // 3) Полное ФИО — только в контексте (иначе ловим заголовки типа
    //    «Права Обязанности Сторон»). Триггеры:
    //    — ПЕРЕД:   «в лице (ген. директора)», «директор/директора», «представитель(я)», «подписант(а)», «гражданин»
    //    — ПОСЛЕ:   «, действующего/ей …», «(паспорт …»
    // NAME3 — строго case-sensitive ([А-ЯЁ] заглавная, [а-яё] строчная).
    // Флаг i использовать НЕЛЬЗЯ: он превратит [А-ЯЁ] в «любая буква»,
    // и жадный `\S+` в триггере съест половину документа.
    // В NAME3 и триггерах используем HS (горизонтальный пробел) — иначе
    // жадный \s+\S+ может «перепрыгнуть» через перенос строки и испортить вёрстку.
    const NAME3 = '([А-ЯЁ][а-яё]{2,30})' + HS + '+([А-ЯЁ][а-яё]{2,30})' + HS + '+([А-ЯЁ][а-яё]{2,30})(?![A-Za-zА-Яа-яЁё0-9])';
    const fioBefore = new RegExp(
      '(?:' +
        '[Вв]' + HS + '+лице(?:' + HS + '+\\S+){0,6}?' + HS + '+|' +
        '(?:[Гг]енеральн[а-яё]+' + HS + '+)?[Дд]иректор[а-яё]*' + HS + '+|' +
        '[Пп]редставител[а-яё]*' + HS + '+|' +
        '[Пп]одписант[а-яё]*' + HS + '+|' +
        '[Гг]ражданин[а-яё]*' + HS + '+(?:(?:[Рр][Фф]|Российской' + HS + '+Федерации|Республики' + HS + '+[А-ЯЁ][а-яё]+)' + HS + '+)?|' +
        '[Ии]ндивидуальн[а-яё]+' + HS + '+предпринимател[а-яё]+' + HS + '+|' +
        'ИП' + HS + '+|' +
        '[Оо]т' + HS + '+имени' + HS + '+|' +
        '[Нн]а' + HS + '+имя' + HS + '+|' +
        '[Фф]\\.?' + HS + '*[Ии]\\.?' + HS + '*[Оо]\\.?' + HS + '*[: \\t\\xa0]' + HS + '*|' +
        '[Гг]р[\\.\\-]?' + HS + '*[нН][\\.]?' + HS + '*|' +
        '[Уу]чредител[а-яё]*' + HS + '+|' +
        '[Ии]спытуемы[хйм]' + HS + '+' +
      ')' +
      NAME3,
      'g'
    );
    // После ФИО могут идти: «, действующего...», «(паспорт...», а также
    // типичные блоки подписантов: «\nПаспорт...», «\nИНН...», «Банковские реквизиты»,
    // «л/с», «Электронная почта», «e-mail».
    const fioAfter = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])' + NAME3 +
      '(?=\\s*,?\\s*(?:' +
        '[Дд]ействующ[а-яё]+|' +
        '[Пп]одписав[а-яё]+|' +
        '\\(?[Пп]аспорт|' +
        'ИНН\\s*[:№]|' +
        'ОГРНИП\\s*[:№]|' +
        'Банковские\\s+реквизит|' +
        '[Лл]/с\\s*[:№]?|' +
        '[Ээ]лектронная\\s+почт|' +
        'e[-‐]?mail|' +
        'Адрес\\s+регистрации' +
      '))',
      'g'
    );

    // Извлекаем имя как «ФИО-тройку в конце совпадения». HS вместо \s,
    // чтобы имя не пересекало перенос строки (иначе замена склеит абзацы).
    const tailName = new RegExp(
      '([А-ЯЁ][а-яё]{2,30}' + HS + '+[А-ЯЁ][а-яё]{2,30}' + HS + '+[А-ЯЁ][а-яё]{2,30})(?![A-Za-zА-Яа-яЁё0-9])$'
    );
    const collectName3 = (re) => {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(text)) !== null) {
        const inner = tailName.exec(mm[0]);
        if (!inner) continue;
        const name = inner[1];
        const start = mm.index + inner.index;
        const end = start + name.length;
        if (overlaps(start, end)) continue;
        claim(start, end);
        // Фамилия — первое слово из трёх (русский порядок «Фамилия Имя Отчество»).
        const firstWord = (name.match(/[А-ЯЁ][а-яё]+/) || [''])[0];
        found.push({ type: 'fio_full', value: name, index: start, length: name.length,
                     meta: { surname: firstWord } });
      }
    };
    collectName3(fioBefore);
    collectName3(fioAfter);

    // Контактное лицо / ФИО-метка: ловим 2-словные имена («Анна Гунькина»),
    // которые обычно стоят после такой метки. БЕЗ триггера два заглавных слова
    // подряд не ловим — слишком много ложных срабатываний.
    const NAME2 = '([А-ЯЁ][а-яё]{2,30})[ \\t\\xa0]+([А-ЯЁ][а-яё]{2,30})(?![A-Za-zА-Яа-яЁё0-9])';
    const fioContact = new RegExp(
      '(?:[Кк]онтактное[ \\t]+лицо|ФИО|Ф\\.\\s*И\\.\\s*О\\.?)[ \\t]*[:\\-—]?\\s+' + NAME2,
      'g'
    );
    let cp;
    while ((cp = fioContact.exec(text)) !== null) {
      // Берём имя — две последние заглавные группы
      const tail = new RegExp(NAME2);
      const tm = tail.exec(cp[0]);
      if (!tm) continue;
      const nameStart = cp.index + tm.index;
      const name = tm[0];
      const nameEnd = nameStart + name.length;
      if (overlaps(nameStart, nameEnd)) continue;
      claim(nameStart, nameEnd);
      found.push({ type: 'fio_full', value: name, index: nameStart, length: name.length,
                   meta: { surname: tm[2] } }); // во 2-словных именах фамилия — второе слово
    }

    // ---------- Денежные суммы (цена договора + расшифровка прописью) ----------
    // Маскируем только «составной» формат: «1 450 000 (один миллион …) рублей 00 копеек».
    // Просто «1 450 000 рублей» без расшифровки прописью НЕ ловим — слишком много
    // ложных срабатываний (даты, количества, сроки).
    // Не маскируем суммы, рядом с которыми упоминаются пени/штрафы/неустойки.
    const priceRe = /(\d[\d\s]{2,20}?)[ \t]*\(([^)]{4,300})\)[ \t]*(рубл[а-яё]*|руб\.?)(?:[ \t]*\d+[ \t]*(?:копе[а-яё]*|коп\.?))?/gi;
    let amm;
    while ((amm = priceRe.exec(text)) !== null) {
      const start = amm.index;
      const end = start + amm[0].length;
      // Контекст ~120 символов перед суммой — отсекаем санкции.
      const ctx = text.substring(Math.max(0, start - 120), start).toLowerCase();
      if (/пен[ия]|штраф|неустойк|возмещени|убытк/.test(ctx)) continue;
      // Должно начинаться с цифры (не часть длинного числа)
      if (start > 0 && /\d/.test(text[start - 1])) continue;
      if (overlaps(start, end)) continue;
      claim(start, end);
      found.push({ type: 'amount', value: amm[0], index: start, length: amm[0].length });
    }

    // ---------- Адреса ----------
    // Блок от 6-значного индекса до ближайшего «стопа» или конца абзаца.
    // Стопы: email, телефон, ИНН/КПП/ОГРН, «ИП », «почта», «тел.», «e-mail».
    // В блоке обязателен хотя бы один адресный маркер (г./ул./д. 5 и т.п.).
    const addrMarkers = /г\.|(?:^|[ ,])г[,.]|ул\.|(?:^|[ ,])ул[,.]|край|переулок|пер\.|проспект|пр-т|пр-д|пгт|шоссе|(?:^|[ ,])ш[,.]|улица|оф\.|офис|помещ|корпус|корп\.|обл\.|строение|стр\.|р-н|наб\.|площад|пл\.|д\.\s*\d|дом\s*№?\s*\d|литер|этаж\s*\d|комн(?:ата)?\s*\d|ком\.\s*\d|Москва|[Сс]анкт[\-‐\s]?Петербург|Россия|РФ|область|город/i;
    const addrStop =
      '(?=' +
        '[,\\s]*(?:' +
          '[ИиЁё]?ИНН|КПП|ОГРН|БИК|[ИИ]П\\s|' +
          '[Пп]очта|[Тт]ел\\.?|[eE][-‐]?[mM]ail|[Фф]акс|' +
          '[A-Za-z0-9._+\\-]+@|' +                 // email
          '\\+?[78][\\s\\-\\(]?\\d{3}|' +          // телефон
          '\\(далее|на\\s+обработку|на\\s+проведение|с\\s+целью|' + // типовые хвосты в договорах
          '\\)\\s*(?:на|от|с|по|для)\\s|' +
          '[Ии]менуем|[Дд]алее\\s[-–—]|[Дд]ействующ' +
        ')|' +
        '[\\n\\r]|$' +
      ')';
    const addrRe = new RegExp(
      '(?<![A-Za-zА-Яа-яЁё0-9])(\\d{6})[, ]([^\\n\\r]{10,500}?)' + addrStop,
      'g'
    );
    let am;
    while ((am = addrRe.exec(text)) !== null) {
      if (!addrMarkers.test(am[0])) continue;
      // Отрезаем висячую запятую/точку/пробелы с конца.
      let val = am[0].replace(/[,\s\.]+$/, '');
      const s = am.index, e = s + val.length;
      if (overlaps(s, e)) continue;
      claim(s, e);
      found.push({ type: 'address', value: val, index: s, length: val.length });
    }

    // Адрес БЕЗ индекса, но с лейблом «по адресу:», «Адрес:», «Юр. адрес:» и т.п.
    // Останавливаемся на переносе строки, явном лейбле поля или «, [предлог] …»
    // (с/и/в/на/от/под — начало нового причастного оборота в договорной прозе).
    const addrByLabel = new RegExp(
      '(?:[Аа]дрес|по\\s+адресу|[Юю]ридический\\s+адрес|[Фф]актический\\s+адрес|[Мм]есто\\s+нахождения)' +
      '[ \\t]*[:\\-]\\s*' +
      '([^\\n\\r]{10,400}?)' +
      '(?=\\s*(?:[\\n\\r]|$|ИНН|КПП|ОГРН|БИК|[Тт]ел\\.?|[Ээ]лектронная|[eE][-‐]?mail|\\)|' +
      ',\\s+(?:с\\s|и\\s|в\\s|на\\s|от\\s|под\\s|[Ии]менуем|[Дд]ействующ|\\(далее)))',
      'g'
    );
    let alm;
    while ((alm = addrByLabel.exec(text)) !== null) {
      const captured = alm[1];
      if (!addrMarkers.test(captured)) continue;
      const valStart = alm.index + alm[0].indexOf(captured);
      const valEnd = valStart + captured.length;
      const val = captured.replace(/[,\s\.]+$/, '');
      if (overlaps(valStart, valStart + val.length)) continue;
      claim(valStart, valStart + val.length);
      found.push({ type: 'address', value: val, index: valStart, length: val.length });
    }

    // ---------- Email ----------
    const emailRe = /\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g;
    let em;
    while ((em = emailRe.exec(text)) !== null) {
      const s = em.index, e = s + em[0].length;
      if (overlaps(s, e)) continue;
      claim(s, e);
      found.push({ type: 'email', value: em[0], index: s, length: em[0].length });
    }

    // ---------- Телефон (российские форматы) ----------
    // +7 (921) 914-10-02, 8-921-914-1002, +79219141002, 8(921)9141002 — все варианты
    const phoneRe = /(?<![\w\d])(?:\+?7|8)[\s\-\(]*\d{3}[\s\-\)]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}(?![\w\d])/g;
    let pm;
    while ((pm = phoneRe.exec(text)) !== null) {
      const s = pm.index, e = s + pm[0].length;
      if (overlaps(s, e)) continue;
      claim(s, e);
      found.push({ type: 'phone', value: pm[0], index: s, length: pm[0].length });
    }

    // Сортируем по позиции
    found.sort((a, b) => a.index - b.index);

    // Пост-пасс 1: объединяем «A (B)» — длинная форма и аббревиатура одной компании
    // получают одинаковый канонический ключ, чтобы быть помечены одной маской.
    // Также объединяем «Филиал «X» Банка Y» — части одного банка.
    const companyFinds = found.filter(f => f.type === 'company');
    for (let i = 0; i < companyFinds.length - 1; i++) {
      const a = companyFinds[i];
      const b = companyFinds[i + 1];
      const gap = text.substring(a.index + a.length, b.index);
      // «A (B)» — аббревиатура в скобках
      if (/^[ \t\xa0]*\([ \t\xa0]*$/.test(gap)) {
        if (b.meta && a.meta) b.meta.name = a.meta.name;
        continue;
      }
      // «Филиал «X» Банка Y» / «Представительство «X» ПАО Y» — части одного банка.
      // Проверяем: A — филиал/представительство (в любом регистре), B следует через
      // короткий пробел.
      if (a.meta && a.meta.prefix) {
        const ap = a.meta.prefix.toLowerCase();
        if ((ap === 'филиал' || ap === 'представительство') &&
            /^[ \t\xa0]+$/.test(gap) && gap.length <= 5) {
          if (b.meta && a.meta) b.meta.name = a.meta.name;
        }
      }
    }

    // Пост-пасс 2: если адрес переносится на следующий абзац и там
    // продолжение (маркеры ул./д./пом. и т.п.) — расширяем находку.
    const addrFinds = found.filter(f => f.type === 'address');
    for (const f of addrFinds) {
      let end = f.index + f.length;
      // Пропускаем пробелы/табы (не трогаем \n/\r)
      while (end < text.length && (text[end] === ' ' || text[end] === '\t')) end++;
      // Нужен хотя бы один перевод строки
      if (text[end] !== '\n' && text[end] !== '\r') continue;
      while (end < text.length && (text[end] === '\n' || text[end] === '\r')) end++;
      // Читаем следующую строку
      let lineEnd = end;
      while (lineEnd < text.length && text[lineEnd] !== '\n' && text[lineEnd] !== '\r') lineEnd++;
      const nextLine = text.substring(end, lineEnd);
      if (nextLine.length === 0 || nextLine.length > 300) continue;
      // Явно НЕ продолжение — начинается с лейбла «Слово:» или «Слово №»
      if (/^[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё .]*[:№]/.test(nextLine.trim())) continue;
      if (!addrMarkers.test(nextLine)) continue;
      const newEnd = lineEnd;
      if (overlaps(f.index + f.length, newEnd)) continue;
      claim(f.index + f.length, newEnd);
      f.value = text.substring(f.index, newEnd);
      f.length = newEnd - f.index;
    }

    return found;
  }

  // ---------- Публичные организации (не маскируем) ----------
  // Общеизвестные регуляторы/госорганы — их упоминание в договорах не является
  // чувствительной информацией. Если совпадение company-регулярок попало на
  // одну из этих сущностей — пропускаем.
  // Важно: lookbehind `(?<![А-Яа-яЁё])` обязателен, чтобы «Банк/банк» не матчилось
  // внутри «Сбербанк», «Альфа-Банк» и т.п. — только как отдельное слово.
  const PUBLIC_ENTITY_PATTERNS = [
    /(?<![А-Яа-яЁё])Центральн\S*\s+банк\S*\s+Российск\S*\s+Федераци/i,
    /(?<![А-Яа-яЁё])[Бб]анк[ауе]?\s+России(?![А-Яа-яЁё])/,
    /(?<![А-Яа-яЁё])ЦБ\s+РФ(?![А-Яа-яЁё])/,
    /(?<![А-Яа-яЁё])Правительств\S*\s+Российск\S*\s+Федераци/i,
    /(?<![А-Яа-яЁё])Правительств\S*\s+РФ(?![А-Яа-яЁё])/i,
    /(?<![А-Яа-яЁё])[Фф]едеральн\S*\s+налогов\S*\s+служб/i,
    /(?<![А-Яа-яЁё])ФНС(?:\s+России)?(?![А-Яа-яЁё])/,
    /(?<![А-Яа-яЁё])[Пп]енсионн\S*\s+фонд/i,
    /(?<![А-Яа-яЁё])ПФР(?![А-Яа-яЁё])/,
    /(?<![А-Яа-яЁё])[Фф]едеральн\S*\s+служб\S*\s+судебн\S*/i,
    /(?<![А-Яа-яЁё])ФССП(?![А-Яа-яЁё])/,
    /(?<![А-Яа-яЁё])Росреестр/i,
    /(?<![А-Яа-яЁё])[Мм]инистерств\S*\s+[а-яё]+\s+(?:Российск\S*\s+Федераци|РФ)/i,
    /(?<![А-Яа-яЁё])Государственн\S*\s+дум\S*/i,
    /(?<![А-Яа-яЁё])Совет\s+Федераци/i,
  ];
  function isPublicEntity(text) {
    return PUBLIC_ENTITY_PATTERNS.some(re => re.test(text));
  }

  // ---------- Канонический ключ для ФИО ----------
  // «Иванова Ивана Ивановича» (род.п.) и «Иванов И.И.» — один человек.
  // Нормализуем фамилию через примитивный стеммер — срезаем типовые русские
  // суффиксы, получаем общий корень.
  function nameStem(word) {
    const w = (word || '').toLowerCase();
    // Порядок — от длинных к коротким.
    const suffixes = [
      'овна', 'евна', 'ович', 'евич',
      'ового', 'евого', 'ому', 'ему',
      'овым', 'евым', 'иной', 'ыной',
      'ову', 'еву', 'иных', 'ыных',
      'ова', 'ева', 'ина', 'ына',
      'ов', 'ев', 'ин', 'ын',
      'ого', 'ему',
      'ой', 'ая', 'ий', 'ый',
      'их', 'ых',
      'ы', 'а', 'я', 'у', 'ю', 'и', 'е'
    ];
    for (const sfx of suffixes) {
      if (w.endsWith(sfx) && w.length - sfx.length >= 3) {
        return w.substring(0, w.length - sfx.length);
      }
    }
    return w;
  }
  function canonicalFioKey(meta) {
    return meta && meta.surname ? nameStem(meta.surname) : null;
  }

  // ---------- Канонический ключ для компаний ----------
  // Две разные строки — «Публичное акционерное общество «Сбербанк России»»
  // и «ПАО Сбербанк» — это одна и та же компания. Чтобы получить одинаковую
  // маску, группируем по первому слову названия (без учёта регистра и знаков).
  function canonicalCompanyKey(meta) {
    const raw = (meta && meta.name) ? meta.name : '';
    if (!raw) return null;
    const clean = raw.toLowerCase()
      .replace(/[^\p{L}\s\-]/gu, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (!clean) return null;
    return clean.split(' ')[0];
  }

  // ---------- Построение карты замен ----------
  // Все маски имеют формат [ТИП_N]. Одинаковые значения получают одну и ту же маску.
  // Для компаний дополнительно: разные формы одного имени (полная/короткая) получают
  // ОДНУ маску по каноническому ключу, но при демаске восстанавливается исходная форма
  // для каждого вхождения — за счёт хранения массива originals в словаре.
  // Нумерация сквозная в рамках ярлыка (ИНН-10 и ИНН-12 разделяют счётчик [ИНН_N]).
  const PLACEHOLDER_LABELS = {
    inn10:        'ИНН',
    inn12:        'ИНН',
    ogrn:         'ОГРН',
    ogrnip:       'ОГРНИП',
    company:      'КОМПАНИЯ',
    kpp:          'КПП',
    bik:          'БИК',
    rs:           'РАСЧ_СЧЕТ',
    ks:           'КОРР_СЧЕТ',
    fio_initials: 'ФИО',
    fio_full:     'ФИО',
    address:      'АДРЕС',
    email:        'EMAIL',
    phone:        'ТЕЛЕФОН',
    passport:     'ПАСПОРТ',
    okpo:         'ОКПО',
    okved:        'ОКВЭД',
    oktmo:        'ОКТМО',
    okato:        'ОКАТО',
    amount:       'СУММА',
  };

  // dict: { [mask]: { original, type } }
  // pending: Map<original, {mask, type, meta}>
  function buildReplacements(finds, existingDict) {
    const inverse = {}; // original -> mask
    // Карты канонических ключей: по типу
    const companyCanonToMask = {};
    const fioCanonToMask = {};

    for (const [mask, info] of Object.entries(existingDict)) {
      const originals = info.originals || (info.original ? [info.original] : []);
      for (const orig of originals) inverse[orig] = mask;
      if (info.type === 'company' && info.canon) companyCanonToMask[info.canon] = mask;
      if ((info.type === 'fio_initials' || info.type === 'fio_full') && info.canon) {
        fioCanonToMask[info.canon] = mask;
      }
    }

    const counters = {};
    const maskRe = /^\[([А-ЯЁ_]+)_(\d+)\]$/;
    for (const mask of Object.keys(existingDict)) {
      const m = maskRe.exec(mask);
      if (m) {
        const label = m[1], n = +m[2];
        if (!counters[label] || counters[label] < n) counters[label] = n;
      }
    }

    const pending = new Map();

    for (const f of finds) {
      if (inverse[f.value]) continue;
      if (pending.has(f.value)) continue;

      const label = PLACEHOLDER_LABELS[f.type] || 'ДАННЫЕ';
      let mask, canon = null;

      if (f.type === 'company') {
        canon = canonicalCompanyKey(f.meta);
        if (canon && companyCanonToMask[canon]) mask = companyCanonToMask[canon];
      } else if (f.type === 'fio_initials' || f.type === 'fio_full') {
        canon = canonicalFioKey(f.meta);
        if (canon && fioCanonToMask[canon]) mask = fioCanonToMask[canon];
      }

      if (!mask) {
        counters[label] = (counters[label] || 0) + 1;
        mask = `[${label}_${counters[label]}]`;
        if (canon && f.type === 'company') companyCanonToMask[canon] = mask;
        if (canon && (f.type === 'fio_initials' || f.type === 'fio_full')) {
          fioCanonToMask[canon] = mask;
        }
      }

      pending.set(f.value, { mask, type: f.type, meta: f.meta, canon });
      inverse[f.value] = mask;
    }

    return { pending, inverse };
  }

  // Публичный API
  global.Zashifrator = {
    validateInn10, generateInn10,
    validateInn12, generateInn12,
    validateOgrn,  generateOgrn,
    validateOgrnip, generateOgrnip,
    findAll,
    buildReplacements,
    canonicalCompanyKey,
    canonicalFioKey,
    nameStem,
  };
})(typeof window !== 'undefined' ? window : globalThis);
