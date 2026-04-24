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

    // Сортируем по позиции
    found.sort((a, b) => a.index - b.index);
    return found;
  }

  // ---------- Построение карты замен ----------
  // dict: { [mask]: { original, type } }
  // pending: Map<original, {mask, type, meta}>
  function buildReplacements(finds, existingDict) {
    const inverse = {}; // original -> mask
    for (const [mask, info] of Object.entries(existingDict)) {
      inverse[info.original] = mask;
    }

    let companyN = 0, innN = 0, ogrnN = 0;
    for (const mask of Object.keys(existingDict)) {
      const t = existingDict[mask].type;
      if (t === 'company') companyN++;
      else if (t === 'inn10' || t === 'inn12') innN++;
      else if (t === 'ogrn' || t === 'ogrnip') ogrnN++;
    }

    const pending = new Map();

    for (const f of finds) {
      if (inverse[f.value]) continue; // уже в словаре — используем существующую маску
      if (pending.has(f.value)) continue;

      let mask;
      if (f.type === 'inn10')        mask = generateInn10();
      else if (f.type === 'inn12')   mask = generateInn12();
      else if (f.type === 'ogrn')    mask = generateOgrn();
      else if (f.type === 'ogrnip')  mask = generateOgrnip();
      else if (f.type === 'company') {
        companyN++;
        const prefix = (f.meta && f.meta.prefix) || 'ООО';
        mask = `${prefix} «Организация_${companyN}»`;
      }

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
