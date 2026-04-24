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
      'ООО', 'ОАО', 'ПАО', 'ЗАО', 'НАО', 'АНО',
      'АО', 'ИП',
    ].sort((a, b) => b.length - a.length);
    const prefixAlt = PREFIXES.map(p => p.replace(/\s+/g, '\\s+')).join('|');
    const companyRe = new RegExp(
      '(?<![\\wА-Яа-яЁё])(' + prefixAlt + ')\\s*' +
      OPEN + '([^\\u00AB\\u00BB\\u0022\\u201C\\u201D\\u201E\\u2018\\u2019]{2,120})' + CLOSE,
      'gi'
    );
    let cm;
    while ((cm = companyRe.exec(text)) !== null) {
      const start = cm.index;
      const end = start + cm[0].length;
      if (overlaps(start, end)) continue;
      claim(start, end);
      found.push({
        type: 'company',
        value: cm[0],
        index: start,
        length: cm[0].length,
        meta: { prefix: cm[1], name: cm[2] }
      });
    }

    // ---------- Банковские реквизиты (с меткой) ----------
    // Маскируем ТОЛЬКО значение (группу 1), метку оставляем.
    // Для каждой находки: index — позиция группы, value — только число.
    // Важно: сначала К/с (корр. счёт) — иначе Р/с-регулярка поймает «р.счет» внутри «Кор.счет».
    // Lookbehind на не-букву, чтобы «р» из «Кор» не матчился как начало метки Р/с.
    const NL = '(?<![\\wА-Яа-яЁё])'; // not-letter lookbehind
    const labelledPatterns = [
      { type: 'kpp', re: new RegExp(NL + '(КПП)\\s*[:№]?\\s*(\\d{9})\\b', 'gi') },
      { type: 'bik', re: new RegExp(NL + '(БИК(?:\\s+банка)?)\\s*[:№]?\\s*(\\d{9})\\b', 'gi') },
      { type: 'ks',
        re:  new RegExp(NL + '(К(?:орр?)?(?:\\.?\\s*сч(?:е|ё)т|\\/с|\\.?\\s*с\\.?))\\s*[:№]?\\s*(\\d{20})\\b', 'gi') },
      { type: 'rs',
        re:  new RegExp(NL + '(Р(?:асч(?:е|ё)тный)?(?:[\\.\\/\\s]*с(?:ч(?:е|ё)т)?)?)\\s*[:№]?\\s*(\\d{20})\\b', 'gi') },
    ];
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
    // 1) Инициалы + фамилия: «А.В.Кукса», «А. В. Кукса»
    const fio1 = new RegExp(
      '(?<![\\wА-Яа-яЁё])([А-ЯЁ])\\.' + HS + '*([А-ЯЁ])\\.' + HS + '*([А-ЯЁ][а-яё]{1,30})(?![\\wА-Яа-яЁё])',
      'g'
    );
    let fm;
    while ((fm = fio1.exec(text)) !== null) {
      const s = fm.index, e = s + fm[0].length;
      if (overlaps(s, e)) continue;
      claim(s, e);
      found.push({ type: 'fio_initials', value: fm[0], index: s, length: fm[0].length,
                   meta: { kind: 'initials_first' } });
    }
    // 2) Фамилия + инициалы: «Кукса А.В.»
    const fio2 = new RegExp(
      '(?<![\\wА-Яа-яЁё])([А-ЯЁ][а-яё]{1,30})' + HS + '+([А-ЯЁ])\\.' + HS + '*([А-ЯЁ])\\.(?![\\wА-Яа-яЁё])',
      'g'
    );
    while ((fm = fio2.exec(text)) !== null) {
      const s = fm.index, e = s + fm[0].length;
      if (overlaps(s, e)) continue;
      claim(s, e);
      found.push({ type: 'fio_initials', value: fm[0], index: s, length: fm[0].length,
                   meta: { kind: 'initials_last' } });
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
    const NAME3 = '([А-ЯЁ][а-яё]{2,30})' + HS + '+([А-ЯЁ][а-яё]{2,30})' + HS + '+([А-ЯЁ][а-яё]{2,30})(?![\\wА-Яа-яЁё])';
    const fioBefore = new RegExp(
      '(?:' +
        '[Вв]' + HS + '+лице(?:' + HS + '+\\S+){0,6}?' + HS + '+|' +
        '(?:[Гг]енеральн[а-яё]+' + HS + '+)?[Дд]иректор[а-яё]*' + HS + '+|' +
        '[Пп]редставител[а-яё]*' + HS + '+|' +
        '[Пп]одписант[а-яё]*' + HS + '+|' +
        '[Гг]ражданин[а-яё]*' + HS + '+(?:[Рр][Фф]' + HS + '+)?|' +
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
    const fioAfter = new RegExp(
      '(?<![\\wА-Яа-яЁё])' + NAME3 +
      '(?=' + HS + '*,?' + HS + '*(?:[Дд]ействующ[а-яё]+|[Пп]одписав[а-яё]+|\\(?[Пп]аспорт))',
      'g'
    );

    // Извлекаем имя как «ФИО-тройку в конце совпадения». HS вместо \s,
    // чтобы имя не пересекало перенос строки (иначе замена склеит абзацы).
    const tailName = new RegExp(
      '([А-ЯЁ][а-яё]{2,30}' + HS + '+[А-ЯЁ][а-яё]{2,30}' + HS + '+[А-ЯЁ][а-яё]{2,30})(?![\\wА-Яа-яЁё])$'
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
        found.push({ type: 'fio_full', value: name, index: start, length: name.length });
      }
    };
    collectName3(fioBefore);
    collectName3(fioAfter);

    // ---------- Адреса ----------
    // Блок от 6-значного индекса до ближайшего «стопа» или конца абзаца.
    // Стопы: email, телефон, ИНН/КПП/ОГРН, «ИП », «почта», «тел.», «e-mail».
    // В блоке обязателен хотя бы один адресный маркер (г./ул./д. 5 и т.п.).
    const addrMarkers = /г\.|ул\.|край|переулок|пер\.|проспект|пр-т|пр-д|пгт|шоссе|улица|оф\.|офис|помещ|корпус|корп\.|обл\.|строение|стр\.|р-н|наб\.|площад|пл\.|д\.\s*\d|дом\s+\d|литер/i;
    const addrStop =
      '(?=' +
        '[,\\s]*(?:' +
          '[ИиЁё]?ИНН|КПП|ОГРН|БИК|[ИИ]П\\s|' +
          '[Пп]очта|[Тт]ел\\.?|[eE][-‐]?[mM]ail|[Фф]акс|' +
          '[A-Za-z0-9._+\\-]+@|' +                 // email
          '\\+?[78][\\s\\-\\(]?\\d{3}' +           // телефон
        ')|' +
        '[\\n\\r]|$' +
      ')';
    const addrRe = new RegExp(
      '(?<![\\wА-Яа-яЁё\\d])(\\d{6})[, ]([^\\n\\r]{10,500}?)' + addrStop,
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
    return found;
  }

  // ---------- Построение карты замен ----------
  // Все маски имеют формат [ТИП_N]. Одинаковые значения получают одну и ту же маску.
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
  };

  // dict: { [mask]: { original, type } }
  // pending: Map<original, {mask, type, meta}>
  function buildReplacements(finds, existingDict) {
    const inverse = {}; // original -> mask
    for (const [mask, info] of Object.entries(existingDict)) {
      inverse[info.original] = mask;
    }

    // Восстанавливаем счётчики из уже имеющихся масок: берём максимум номера для каждого ярлыка.
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
      counters[label] = (counters[label] || 0) + 1;
      const mask = `[${label}_${counters[label]}]`;

      pending.set(f.value, { mask, type: f.type, meta: f.meta });
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
  };
})(typeof window !== 'undefined' ? window : globalThis);
