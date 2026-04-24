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

// ----------- Замаскировать -----------

async function onMask() {
  setStatus('Маскируем…');
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    const fullText = body.text;

    const finds = Zashifrator.findAll(fullText);
    if (finds.length === 0) {
      setStatus('Чувствительных данных не найдено.');
      return;
    }

    const dict = await loadDict();
    const { inverse } = Zashifrator.buildReplacements(finds, dict);

    const pairs = [];
    const seen = new Set();
    for (const f of finds) {
      if (seen.has(f.value)) continue;
      seen.add(f.value);
      const mask = inverse[f.value];
      if (!mask) continue;
      pairs.push({ original: f.value, mask, type: f.type });
    }
    for (const p of pairs) {
      if (!dict[p.mask]) dict[p.mask] = { original: p.original, type: p.type };
    }
    // Сначала длинные — чтобы не попасть в подстроку.
    pairs.sort((a, b) => b.original.length - a.original.length);

    let total = 0;
    for (const { original, mask } of pairs) {
      total += await replaceAll(context, original, mask);
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

  await Word.run(async (context) => {
    entries.sort((a, b) => b[0].length - a[0].length);
    let total = 0;
    for (const [mask, info] of entries) {
      total += await replaceAll(context, mask, info.original);
    }
    setStatus(`Восстановлено ${total} вхождений.`, 'ok');
  });
}

// ----------- Замена всех вхождений -----------

async function replaceAll(context, needle, replacement) {
  const body = context.document.body;
  const results = body.search(needle, { matchCase: true, matchWholeWord: false });
  results.load('items');
  await context.sync();

  const items = results.items;
  for (const r of items) {
    r.insertText(replacement, Word.InsertLocation.replace);
  }
  await context.sync();
  return items.length;
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
