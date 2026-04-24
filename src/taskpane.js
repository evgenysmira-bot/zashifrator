// Связка UI и ядра. Работает с Word через Office.js.
/* global Office, Word, Zashifrator */

const DICT_KEY = 'ZashifratorDict_v1';

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Эта надстройка работает только в Microsoft Word.', true);
    return;
  }
  bind('btnScan',       onScan);
  bind('btnMask',       onMask);
  bind('btnUnmask',     onUnmask);
  bind('btnClearDict',  onClearDict);
  renderDictionary();
});

function bind(id, fn) {
  document.getElementById(id).addEventListener('click', () => fn().catch(handleError));
}

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

// ----------- Получение всего текста документа -----------

async function getDocText(context) {
  const body = context.document.body;
  body.load('text');
  await context.sync();
  return body.text;
}

// ----------- Фильтр по чекбоксам -----------

function filterByOptions(finds) {
  const useInn     = document.getElementById('optInn').checked;
  const useOgrn    = document.getElementById('optOgrn').checked;
  const useCompany = document.getElementById('optCompany').checked;
  const useBank    = document.getElementById('optBank').checked;
  const useFio     = document.getElementById('optFio').checked;
  const useAddress = document.getElementById('optAddress').checked;
  const useContact = document.getElementById('optContact').checked;
  return finds.filter(f => {
    if (f.type === 'inn10' || f.type === 'inn12') return useInn;
    if (f.type === 'ogrn'  || f.type === 'ogrnip') return useOgrn;
    if (f.type === 'company') return useCompany;
    if (f.type === 'kpp' || f.type === 'bik' || f.type === 'rs' || f.type === 'ks') return useBank;
    if (f.type === 'fio_initials' || f.type === 'fio_full') return useFio;
    if (f.type === 'address') return useAddress;
    if (f.type === 'email' || f.type === 'phone') return useContact;
    return false;
  });
}

// ----------- Найти и показать -----------

async function onScan() {
  setStatus('Сканируем документ…');
  await Word.run(async (context) => {
    const text = await getDocText(context);
    const finds = filterByOptions(Zashifrator.findAll(text));
    const uniq = new Set(finds.map(f => f.value));
    if (finds.length === 0) {
      setStatus('Чувствительных данных не найдено.');
    } else {
      const byType = {};
      for (const f of finds) byType[f.type] = (byType[f.type] || 0) + 1;
      const parts = Object.entries(byType).map(([t, n]) => `${typeLabel(t)}: ${n}`);
      setStatus(`Найдено ${finds.length} вхождений (${uniq.size} уникальных). ${parts.join(', ')}.`);
    }
  });
}

// ----------- Замаскировать -----------

async function onMask() {
  setStatus('Маскируем…');
  await Word.run(async (context) => {
    const text = await getDocText(context);
    const finds = filterByOptions(Zashifrator.findAll(text));
    if (finds.length === 0) {
      setStatus('Нечего маскировать — ничего не найдено.');
      return;
    }

    const dict = await loadDict();
    const { pending, inverse } = Zashifrator.buildReplacements(finds, dict);

    // Собираем итоговый список пар (оригинал → маска).
    // Уникализируем по оригиналу.
    const pairs = [];
    const seen = new Set();
    for (const f of finds) {
      if (seen.has(f.value)) continue;
      seen.add(f.value);
      const mask = inverse[f.value];
      if (!mask) continue;
      pairs.push({ original: f.value, mask, type: f.type });
    }

    // Пишем новые пары в словарь.
    for (const p of pairs) {
      if (!dict[p.mask]) dict[p.mask] = { original: p.original, type: p.type };
    }

    // Выполняем замены в документе. Сначала длинные — чтобы не ловить подстроки.
    pairs.sort((a, b) => b.original.length - a.original.length);

    let totalReplaced = 0;
    for (const { original, mask } of pairs) {
      const count = await replaceAll(context, original, mask);
      totalReplaced += count;
    }

    await saveDict(dict);
    renderDictionary(dict);
    setStatus(`Замаскировано ${totalReplaced} вхождений (${pairs.length} уникальных значений).`);
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
    // Длинные маски первыми.
    entries.sort((a, b) => b[0].length - a[0].length);
    let total = 0;
    for (const [mask, info] of entries) {
      total += await replaceAll(context, mask, info.original);
    }
    setStatus(`Восстановлено ${total} вхождений.`);
  });
}

// ----------- Замена всех вхождений строки -----------

async function replaceAll(context, needle, replacement) {
  // Word.search не любит некоторые кавычки/спецсимволы — идём через getRange + search.
  // matchCase=true, matchWholeWord=false.
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

// ----------- Очистка словаря -----------

async function onClearDict() {
  if (!confirm('Удалить словарь замен? Демаскировка после этого будет невозможна.')) return;
  await saveDict({});
  renderDictionary({});
  setStatus('Словарь очищен.');
}

// ----------- Рендер словаря -----------

async function renderDictionary(dict) {
  if (!dict) dict = await loadDict();
  const entries = Object.entries(dict);
  document.getElementById('dictCount').textContent = String(entries.length);
  const host = document.getElementById('dictContainer');
  if (entries.length === 0) {
    host.innerHTML = '<div class="empty">Словарь пуст</div>';
    return;
  }
  const rows = entries.map(([mask, info]) =>
    `<tr>
       <td>${typeLabel(info.type)}</td>
       <td>${escapeHtml(info.original)}</td>
       <td>${escapeHtml(mask)}</td>
     </tr>`
  ).join('');
  host.innerHTML =
    `<table>
       <thead><tr><th>Тип</th><th>Оригинал</th><th>Маска</th></tr></thead>
       <tbody>${rows}</tbody>
     </table>`;
}

// ----------- Утилиты -----------

function typeLabel(t) {
  return ({
    inn10:        'ИНН-10',
    inn12:        'ИНН-12',
    ogrn:         'ОГРН',
    ogrnip:       'ОГРНИП',
    company:      'Компания',
    kpp:          'КПП',
    bik:          'БИК',
    rs:           'Р/с',
    ks:           'К/с',
    fio_initials: 'ФИО (инициалы)',
    fio_full:     'ФИО',
    address:      'Адрес',
    email:        'Email',
    phone:        'Телефон',
  })[t] || t;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => (
    { '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]
  ));
}

function setStatus(text, isError) {
  const el = document.getElementById('status');
  el.textContent = text;
  el.classList.toggle('error', !!isError);
}

function handleError(err) {
  console.error(err);
  const msg = (err && err.message) || String(err);
  setStatus('Ошибка: ' + msg, true);
}
