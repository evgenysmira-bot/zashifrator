// Связка UI и ядра. Работает с Word через Office.js.
/* global Office, Word, Zashifrator */

const DICT_KEY = 'ZashifratorDict_v1';

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Эта надстройка работает только в Microsoft Word.', 'error');
    return;
  }
  document.getElementById('btnMask').addEventListener('click',   () => onMask().catch(handleError));
  document.getElementById('btnUnmask').addEventListener('click', () => onUnmask().catch(handleError));
});

// ----------- Словарь хранится в настройках документа -----------

function loadDict() {
  return new Promise((resolve) => {
    Office.context.document.settings.refreshAsync(() => {
      const raw = Office.context.document.settings.get(DICT_KEY);
      try { resolve(raw ? JSON.parse(raw) : {}); }
      catch { resolve({}); }
    });
  });
}

function saveDict(dict) {
  Office.context.document.settings.set(DICT_KEY, JSON.stringify(dict));
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

// Приводим старый формат {original, type} к новому {originals[], type, canon}
function upgradeEntry(entry) {
  if (!entry.originals) {
    entry.originals = entry.original ? [entry.original] : [];
    delete entry.original;
  }
  return entry;
}

// ----------- Получить список всех «тел» для поиска -----------
// Main body + все headers/footers во всех секциях (primary, firstPage, evenPages)

async function getAllBodies(context) {
  const doc = context.document;
  const bodies = [doc.body];

  const sections = doc.sections;
  sections.load('items');
  try { await context.sync(); }
  catch (e) { return bodies; } // если sections недоступны — работаем только с body

  if (!sections.items || sections.items.length === 0) return bodies;

  // Берём колонтитулы ТОЛЬКО из первой секции. В большинстве .docx секции
  // линкованы и совместно используют один и тот же физический колонтитул,
  // поэтому достаточно одного прохода. Исключаем мульти-обработку, которая
  // вызывала «рекурсию» масок и GeneralException при батчевом load.
  const firstSection = sections.items[0];
  const kinds = ['Primary', 'FirstPage', 'EvenPages'];
  for (const kind of kinds) {
    try {
      bodies.push(firstSection.getHeader(kind), firstSection.getFooter(kind));
    } catch (e) { /* отсутствующий тип колонтитула — пропускаем */ }
  }
  return bodies;
}

// Получить текст из всех тел (для findAll)
async function getAllText(context) {
  const bodies = await getAllBodies(context);
  for (const b of bodies) b.load('text');
  await context.sync();
  // Разделяем тексты разных тел символом \u2029 — вне регулярок, не затронет матчинг.
  return bodies.map(b => b.text || '').join('\u2029');
}

// Заменить все вхождения `needle` на `replacement` во всех телах.
// Возвращает массив найденных Range-объектов в порядке обхода (для поддержки
// множественных originals на одну маску при демаске).
async function searchAllBodies(context, needle) {
  const bodies = await getAllBodies(context);
  const collections = [];
  for (const b of bodies) {
    const res = b.search(needle, { matchCase: true, matchWholeWord: false });
    res.load('items');
    collections.push(res);
  }
  await context.sync();
  const all = [];
  for (const c of collections) for (const r of c.items) all.push(r);
  return all;
}

async function replaceAllBodies(context, needle, replacement) {
  // 1) Стандартный путь: body.search → insertText (сохраняет форматирование).
  const ranges = await searchAllBodies(context, needle);
  if (ranges.length > 0) {
    for (const r of ranges) r.insertText(replacement, Word.InsertLocation.replace);
    await context.sync();
    return ranges.length;
  }
  // 2) Фолбэк по параграфам — ТОЛЬКО когда body.search не нашёл вообще ничего.
  //    Раньше пробовали запускать его всегда (для добора пропущенных вхождений
  //    «Фамилия И.О.» в параграфах с длинными подчёркиваниями), но это ломало
  //    демаскировку: норма заменяет `\v` (softbreak от Shift+Enter) на пробел,
  //    после чего `paragraph.insertText(replace)` склеивал куски параграфа в
  //    одну строку — в шапке арендного договора арендатор/арендодатель
  //    путались. Возвращаемся к консервативному поведению.
  return await fallbackParagraphReplace(context, needle, replacement);
}

async function fallbackParagraphReplace(context, needle, replacement) {
  let count = 0;
  const norm = s => s.replace(/[\s\u00A0\u00AD\u200B-\u200F\u202A-\u202E\u2060\v]+/g, ' ').trim();
  const normalizedNeedle = norm(needle);
  if (normalizedNeedle.length < 5) return 0;

  const bodies = await getAllBodies(context);
  const allParas = [];
  for (const body of bodies) {
    try {
      const paras = body.paragraphs;
      paras.load('text');
      allParas.push(paras);
    } catch (e) { /* пропускаем тело, если параграфы недоступны */ }
  }
  try { await context.sync(); }
  catch (e) { return count; } // если не смогли загрузить параграфы — выходим

  for (const paras of allParas) {
    let items;
    try { items = paras.items; } catch (e) { continue; }
    if (!items) continue;
    for (const p of items) {
      try {
        const txt = p.text || '';
        if (!txt) continue;
        const normTxt = norm(txt);
        if (!normTxt.includes(normalizedNeedle)) continue;
        const idx = normTxt.indexOf(normalizedNeedle);
        const prefix = normTxt.substring(0, idx).replace(/\s+$/, '');
        const suffix = normTxt.substring(idx + normalizedNeedle.length).replace(/^\s+/, '');
        const sep1 = prefix && !/[\s.,;:]$/.test(prefix) ? ' ' : '';
        const sep2 = suffix && !/^[\s.,;:]/.test(suffix) ? ' ' : '';
        const newText = prefix + sep1 + replacement + sep2 + suffix;
        p.insertText(newText, Word.InsertLocation.replace);
        count++;
      } catch (e) { /* skip this paragraph */ }
    }
  }
  try { await context.sync(); } catch (e) {}
  return count;
}

// ----------- Замаскировать -----------

async function onMask() {
  setStatus('Маскируем…');
  await Word.run(async (context) => {
    const fullText = await getAllText(context);

    const finds = Zashifrator.findAll(fullText);
    if (finds.length === 0) {
      setStatus('Чувствительных данных не найдено.');
      return;
    }

    const dict = await loadDict();
    for (const mask of Object.keys(dict)) upgradeEntry(dict[mask]);

    // Чистим словарь от масок, которых в документе уже нет —
    // иначе счётчик разъезжается (получается [КОМПАНИЯ_3] вместо [КОМПАНИЯ_1]).
    const allMasksInDict = Object.keys(dict);
    for (const mask of allMasksInDict) {
      if (!fullText.includes(mask)) delete dict[mask];
    }

    const { inverse } = Zashifrator.buildReplacements(finds, dict);

    // Собираем уникальные пары original → mask, сохраняя канонический ключ для компаний.
    const pairs = [];
    const seenOrig = new Set();
    for (const f of finds) {
      if (seenOrig.has(f.value)) continue;
      seenOrig.add(f.value);
      const mask = inverse[f.value];
      if (!mask) continue;
      pairs.push({ original: f.value, mask, type: f.type, meta: f.meta });
    }

    // Обновляем словарь: добавляем НОВЫЕ оригиналы в originals[] каждой маски.
    for (const p of pairs) {
      if (!dict[p.mask]) dict[p.mask] = { type: p.type, originals: [] };
      upgradeEntry(dict[p.mask]);
      if (!dict[p.mask].originals.includes(p.original)) {
        dict[p.mask].originals.push(p.original);
      }
      if (p.type === 'company') {
        const canon = (p.meta && Zashifrator.canonicalCompanyKey
          ? Zashifrator.canonicalCompanyKey(p.meta)
          : null);
        if (canon) dict[p.mask].canon = canon;
      } else if (p.type === 'fio_initials' || p.type === 'fio_full') {
        const canon = (p.meta && Zashifrator.canonicalFioKey
          ? Zashifrator.canonicalFioKey(p.meta)
          : null);
        if (canon) dict[p.mask].canon = canon;
      }
    }

    // Длинные оригиналы первыми — чтобы не попасть в подстроку.
    pairs.sort((a, b) => b.original.length - a.original.length);

    let total = 0;
    for (const { original, mask } of pairs) {
      total += await replaceAllBodies(context, original, mask);
    }

    // Пост-пасс: склеиваем подряд идущие одинаковые маски — [X][X] → [X].
    // Такие дубли появляются из-за соседних run-ов в .docx с идентичным текстом
    // (артефакт правок / скрытых полей). Визуально в Word это выглядит как одно
    // значение, но body.search находит его дважды.
    const uniqMasks = [...new Set(pairs.map(p => p.mask))];
    let collapsed = 0;
    for (const mask of uniqMasks) {
      for (let guard = 0; guard < 10; guard++) {
        const n = await replaceAllBodies(context, mask + mask, mask);
        if (n === 0) break;
        collapsed += n;
      }
    }

    await saveDict(dict);
    const suffix = collapsed > 0 ? `, склеено ${collapsed} дубликатов` : '';
    setStatus(`Замаскировано ${total} вхождений (${pairs.length} уникальных)${suffix}.`, 'ok');
  });
}

// ----------- Демаскировать -----------

async function onUnmask() {
  setStatus('Восстанавливаем оригиналы…');
  const dict = await loadDict();
  const entries = Object.entries(dict);
  if (entries.length === 0) {
    setStatus('Словарь пуст — демаскировать нечего.');
    return;
  }

  for (const [, info] of entries) upgradeEntry(info);

  await Word.run(async (context) => {
    // Длинные маски первыми.
    entries.sort((a, b) => b[0].length - a[0].length);

    // Пре-дедуп: склеиваем подряд идущие одинаковые маски ([X][X] → [X]).
    // Фикс для документов, замаскированных старой версией с мульти-секционным багом.
    let cleaned = 0;
    for (const [mask] of entries) {
      for (let guard = 0; guard < 15; guard++) {
        const n = await replaceAllBodies(context, mask + mask, mask);
        if (n === 0) break;
        cleaned += n;
      }
    }

    let total = 0;
    for (const [mask, info] of entries) {
      const originals = info.originals || (info.original ? [info.original] : []);
      if (originals.length === 0) continue;

      const ranges = await searchAllBodies(context, mask);
      for (let i = 0; i < ranges.length; i++) {
        const original = originals[Math.min(i, originals.length - 1)];
        ranges[i].insertText(original, Word.InsertLocation.replace);
      }
      await context.sync();
      total += ranges.length;
    }
    const suffix = cleaned > 0 ? `, схлопнуто ${cleaned} дубликатов масок` : '';
    setStatus(`Восстановлено ${total} вхождений${suffix}.`, 'ok');
  });
}

// ----------- Утилиты -----------

function setStatus(text, kind) {
  const el = document.getElementById('status');
  document.getElementById('statusText').textContent = text;
  el.classList.remove('hidden', 'error', 'ok');
  if (kind) el.classList.add(kind);
}

function handleError(err) {
  console.error(err);
  const msg = (err && err.message) || String(err);
  setStatus('Ошибка: ' + msg, 'error');
}
