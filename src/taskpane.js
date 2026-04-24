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
  const bodies = [];
  const doc = context.document;
  bodies.push(doc.body);

  const sections = doc.sections;
  sections.load('items');
  await context.sync();

  const kinds = ['Primary', 'FirstPage', 'EvenPages'];
  for (const section of sections.items) {
    for (const kind of kinds) {
      try {
        const header = section.getHeader(kind);
        const footer = section.getFooter(kind);
        bodies.push(header, footer);
      } catch (e) {
        // часть секций может не иметь header'а конкретного типа — молча пропускаем
      }
    }
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
  const ranges = await searchAllBodies(context, needle);
  for (const r of ranges) r.insertText(replacement, Word.InsertLocation.replace);
  await context.sync();
  return ranges.length;
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
      }
    }

    // Длинные оригиналы первыми — чтобы не попасть в подстроку.
    pairs.sort((a, b) => b.original.length - a.original.length);

    let total = 0;
    for (const { original, mask } of pairs) {
      total += await replaceAllBodies(context, original, mask);
    }

    await saveDict(dict);
    setStatus(`Замаскировано ${total} вхождений (${pairs.length} уникальных).`, 'ok');
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
    let total = 0;
    for (const [mask, info] of entries) {
      const originals = info.originals || (info.original ? [info.original] : []);
      if (originals.length === 0) continue;

      // Ищем ВСЕ вхождения маски (в т.ч. в колонтитулах).
      // Если originals > 1 — на каждое вхождение подставляем соответствующий
      // вариант по порядку; если вхождений больше — доиспользуем последний.
      const ranges = await searchAllBodies(context, mask);
      for (let i = 0; i < ranges.length; i++) {
        const original = originals[Math.min(i, originals.length - 1)];
        ranges[i].insertText(original, Word.InsertLocation.replace);
      }
      await context.sync();
      total += ranges.length;
    }
    setStatus(`Восстановлено ${total} вхождений.`, 'ok');
  });
}

// ----------- Утилиты -----------

function setStatus(text, kind) {
  const el = document.getElementById('status');
  el.textContent = text;
  el.classList.remove('error', 'ok');
  if (kind) el.classList.add(kind);
}

function handleError(err) {
  console.error(err);
  const msg = (err && err.message) || String(err);
  setStatus('Ошибка: ' + msg, 'error');
}
