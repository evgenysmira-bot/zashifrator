// Smoke-тесты ядра: валидаторы + findAll + имитация цикла маскировка→демаскировка.
// Запуск: node tests/smoke.js

const fs = require('fs');
const path = require('path');

// Подгружаем core.js через eval — core.js написан как IIFE на globalThis,
// чтобы одинаково работать в браузере и здесь.
const coreSrc = fs.readFileSync(path.join(__dirname, '..', 'src', 'core.js'), 'utf8');
const sandbox = { globalThis: global };
(new Function('globalThis', coreSrc))(global);
const Z = global.Zashifrator;

let pass = 0, fail = 0;
function check(name, cond, detail) {
  const ok = !!cond;
  console.log(`  ${ok ? 'ok  ' : 'FAIL'}  ${name}${detail ? '  — ' + detail : ''}`);
  if (ok) pass++; else fail++;
}

// ---------- Валидаторы на известных значениях ----------

console.log('\n== Валидаторы ==');
// Реальные ИНН/ОГРН известных организаций (публичные данные из ЕГРЮЛ).
check('ИНН-10 Сбербанк (7707083893)',   Z.validateInn10('7707083893'));
check('ИНН-10 Газпром (7736050003)',    Z.validateInn10('7736050003'));
check('ИНН-10 битый (7707083894)',      !Z.validateInn10('7707083894'));
check('ОГРН Сбербанк (1027700132195)',  Z.validateOgrn('1027700132195'));
check('ОГРН битый (1027700132196)',     !Z.validateOgrn('1027700132196'));

// ---------- Генераторы: то, что сгенерировали, должно пройти валидацию ----------

console.log('\n== Генераторы ==');
for (let i = 0; i < 20; i++) {
  check(`gen ИНН-10 #${i+1}`,   Z.validateInn10(Z.generateInn10()));
  check(`gen ИНН-12 #${i+1}`,   Z.validateInn12(Z.generateInn12()));
  check(`gen ОГРН #${i+1}`,     Z.validateOgrn(Z.generateOgrn()));
  check(`gen ОГРНИП #${i+1}`,   Z.validateOgrnip(Z.generateOgrnip()));
}

// ---------- Поиск в реалистичном тексте ----------

const sampleText = `
ДОГОВОР ПОСТАВКИ № 42/2026

г. Москва                                                        10 марта 2026 г.

Публичное акционерное общество «Сбербанк России», далее — «Поставщик»,
ИНН 7707083893, ОГРН 1027700132195, в лице председателя правления
Иванова И.И., действующего на основании устава, с одной стороны, и

ООО «Ромашка», далее — «Покупатель», ИНН 7736050003,
ОГРН 1027739552642, в лице генерального директора Петрова П.П.,

а также индивидуальный предприниматель Сидоров С.С. (ИНН 500100732259,
ОГРНИП 304500116000157), привлекаемый в качестве Субподрядчика,

заключили настоящий договор о нижеследующем.

Реквизиты: р/с 40702810900000012345 в АО «Альфа-Банк», ИНН 7728168971,
БИК 044525593. Сумма договора: 1 500 000 руб. (один миллион пятьсот тысяч).
`;

console.log('\n== findAll на примере договора ==');
const finds = Z.findAll(sampleText);
const byType = finds.reduce((a, f) => (a[f.type] = (a[f.type] || 0) + 1, a), {});
console.log('  типы:', byType);
for (const f of finds) {
  console.log(`    [${f.type}] "${f.value}"${f.meta ? ' meta=' + JSON.stringify(f.meta) : ''}`);
}
check('Нашли Сбербанк ИНН',     finds.some(f => f.value === '7707083893'));
check('Нашли Сбербанк ОГРН',    finds.some(f => f.value === '1027700132195'));
check('Нашли Ромашка ИНН',      finds.some(f => f.value === '7736050003'));
check('Нашли Альфа ИНН',        finds.some(f => f.value === '7728168971'));
check('Нашли ОГРНИП ИП Сидоров',finds.some(f => f.value === '304500116000157'));
check('Нашли ООО «Ромашка»',    finds.some(f => f.type === 'company' && /Ромашка/.test(f.value)));
check('Нашли АО «Альфа-Банк»',  finds.some(f => f.type === 'company' && /Альфа/.test(f.value)));

// Проверка, что ИНН-10 не поглотил часть ОГРН
check('ИНН-10 не равен куску ОГРН', !finds.some(f => f.type === 'inn10' && sampleText.indexOf('1027700132195') !== -1 && f.value === '1027700132'));

// ---------- Имитация цикла mask → unmask ----------

console.log('\n== Имитация цикла маскировка/демаскировка ==');

function maskText(text) {
  const finds = Z.findAll(text);
  const { inverse } = Z.buildReplacements(finds, {});
  // Составляем словарь: mask -> original
  const dict = {};
  for (const [orig, mask] of Object.entries(inverse)) {
    dict[mask] = { original: orig };
  }
  // Длинные оригиналы заменяем первыми.
  const pairs = Object.entries(inverse).sort((a, b) => b[0].length - a[0].length);
  let masked = text;
  for (const [orig, mask] of pairs) {
    masked = masked.split(orig).join(mask);
  }
  return { masked, dict };
}

function unmaskText(text, dict) {
  // Длинные маски первыми.
  const entries = Object.entries(dict).sort((a, b) => b[0].length - a[0].length);
  let out = text;
  for (const [mask, info] of entries) {
    out = out.split(mask).join(info.original);
  }
  return out;
}

const { masked, dict } = maskText(sampleText);
console.log('  размер словаря:', Object.keys(dict).length);
console.log('  --- первые 400 символов замаскированного ---');
console.log(masked.slice(0, 400).replace(/^/gm, '  | '));

const restored = unmaskText(masked, dict);
check('Round-trip идентичен оригиналу', restored === sampleText,
  restored === sampleText ? '' : 'тексты разошлись');

// Проверяем, что в masked не осталось оригинальных ИНН/ОГРН.
const leaksOriginal =
  masked.includes('7707083893') || masked.includes('1027700132195') ||
  masked.includes('7736050003') || masked.includes('7728168971')   ||
  masked.includes('304500116000157') || masked.includes('Ромашка') ||
  masked.includes('Сбербанк') || masked.includes('Альфа-Банк');
check('В замаскированном тексте нет оригиналов', !leaksOriginal);

// Маски должны сами проходить валидацию.
let maskValidatorOk = true;
for (const [mask, info] of Object.entries(dict)) {
  if (info.original.length === 10 && /^\d{10}$/.test(info.original) && !Z.validateInn10(mask)) maskValidatorOk = false;
  if (info.original.length === 12 && /^\d{12}$/.test(info.original) && !Z.validateInn12(mask)) maskValidatorOk = false;
  if (info.original.length === 13 && /^\d{13}$/.test(info.original) && !Z.validateOgrn(mask))  maskValidatorOk = false;
  if (info.original.length === 15 && /^\d{15}$/.test(info.original) && !Z.validateOgrnip(mask)) maskValidatorOk = false;
}
check('Все числовые маски валидны по контрольной сумме', maskValidatorOk);

// ---------- Итог ----------

console.log(`\nИтог: ${pass} ok, ${fail} fail`);
process.exit(fail === 0 ? 0 : 1);
