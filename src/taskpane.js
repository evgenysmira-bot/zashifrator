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
  const raw = [doc.body];

  const sections = doc.sections;
  sections.load('items');
  await context.sync();

  const kinds = ['Primary', 'FirstPage', 'EvenPages'];
  for (const section of sections.items) {
    for (const kind of kinds) {
      try {
        raw.push(section.getHeader(kind), section.getFooter(kind));
      } catch (e) { /* иногда header нужного типа нет — игнорируем */ }
    }
  }

  // ДЕДУП: если секции ссылаются на один и тот же колонтитул (типичный случай
  // для Word с «Link to Previous»), прокси-объекты разные, но физический
  // контент — один. Без этой дедупликации `body.search` вернёт совпадения
  // N раз (по числу ссылающихся секций), и `insertText` на каждом —
  // вставит маску N раз в одно место. Отсюда «рекурсия» масок.
  for (const b of raw) b.load('text');
  await context.sync();

  const unique = [];
  const seenTexts = new Set();
  // Главное тело всегда первое (его текст может пересекаться с колонтитулом,
  // но это разные физические места).
  unique.push(raw[0]);
  seenTexts.add('__main__::' + (raw[0].text || ''));
  for (let i = 1; i < raw.length; i++) {
    const text = raw[i].text || '';
    if (!text) continue;                    // пустые колонтитулы пропускаем
    if (seenTexts.has(text)) continue;      // уже обработали прокси с этим контентом
    seenTexts.add(text);
    unique.push(raw[i]);
  }
  return unique;
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
  el.textContent = text;
  el.classList.remove('error', 'ok');
  if (kind) el.classList.add(kind);
}

function handleError(err) {
  console.error(err);
  const msg = (err && err.message) || String(err);
  setStatus('Ошибка: ' + msg, 'error');
}
