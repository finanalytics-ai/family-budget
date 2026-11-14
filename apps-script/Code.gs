/**
 * Семейный бюджет — серверная часть (Google Apps Script, привязан к Google Таблице).
 *
 * Что делает:
 *  - меню «💰 Бюджет» в таблице: первичная настройка, обновление отчётов, подключение бота;
 *  - создаёт листы: Журнал, Остатки, План-факт, ОДДС, Бюджет, Статьи, Кошельки, Контрагенты, Доступ;
 *  - API для Telegram Mini App (doPost): проверяет подпись Telegram и список доступа,
 *    отдаёт данные и записывает операции в Журнал.
 */

// Адрес мини-приложения (GitHub Pages). Можно переопределить свойством скрипта APP_URL.
const DEFAULT_APP_URL = 'https://finanalytics-ai.github.io/family-budget/';

const SH = {
  J: 'Журнал', BAL: 'Остатки', PF: 'План-факт', ODDS: 'ОДДС', B: 'Бюджет',
  ART: 'Статьи', WAL: 'Кошельки', CP: 'Контрагенты', ACC: 'Доступ'
};
const TYPES = ['Расход', 'Доход', 'Перевод'];
const J_HEAD = ['ID', 'Дата', 'Тип', 'Сумма', 'Кошелёк', 'Кошелёк-получатель', 'Группа', 'Статья', 'Контрагент', 'Комментарий', 'Автор', 'Создано'];
const ART_HEAD = ['Группа', 'Статья', 'Тип', 'Иконка', 'Активна'];
const WAL_HEAD = ['Кошелёк', 'Владелец', 'Начальный остаток', 'Иконка', 'Активен'];
const ACC_HEAD = ['Telegram ID', 'Имя', 'Разрешён', 'Группа по умолчанию', 'Кошелёк по умолчанию'];
const MONEY = '#,##0.00';
const MONEY0 = '#,##0';

// [Группа, Статья, Тип, Иконка, План в месяц]
const DEFAULT_ARTICLES = [
  ['Люся', 'Зарплата', 'Доход', '💼', 150000],
  ['Кирилл', 'Зарплата', 'Доход', '💼', 190000],
  ['Общие', 'Прочие доходы', 'Доход', '💰', 0],
  ['Люся', 'Ногти', 'Расход', '💅', 5000],
  ['Люся', 'Брови', 'Расход', '✨', 1000],
  ['Люся', 'Волосы', 'Расход', '💇‍♀️', 3500],
  ['Люся', 'Обеды', 'Расход', '🍱', 10000],
  ['Люся', 'Косметика', 'Расход', '💄', 2000],
  ['Люся', 'Бобик', 'Расход', '🐶', 5500],
  ['Люся', 'Прокладки', 'Расход', '🌸', 1000],
  ['Люся', 'Косметолог', 'Расход', '🧖‍♀️', 7000],
  ['Люся', 'Пилатес', 'Расход', '🧘‍♀️', 4000],
  ['Люся', 'Связь', 'Расход', '📱', 700],
  ['Люся', 'Витамины', 'Расход', '💊', 5000],
  ['Люся', 'Одежда', 'Расход', '👗', 5000],
  ['Кирилл', 'Связь', 'Расход', '📱', 400],
  ['Кирилл', 'Интернет', 'Расход', '🌐', 1600],
  ['Кирилл', 'Одежда', 'Расход', '👕', 5000],
  ['Кирилл', 'ИИ Подписки', 'Расход', '🤖', 4000],
  ['Общие', 'Коммуналка', 'Расход', '🏠', 5000],
  ['Общие', 'Еда', 'Расход', '🛒', 50000],
  ['Общие', 'Развлечения', 'Расход', '🎉', 5000],
  ['Общие', 'Кафе', 'Расход', '☕', 10000],
  ['Общие', 'Подарки', 'Расход', '🎁', 10000]
];
const DEFAULT_WALLETS = [
  ['Карта Кирилла', 'Кирилл', 0, '💳', true],
  ['Карта Люси', 'Люся', 0, '💳', true],
  ['Наличные', 'Общие', 0, '💵', true]
];

/* =========================== МЕНЮ =========================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 Бюджет')
    .addItem('1. Первичная настройка', 'setup')
    .addItem('2. Указать токен бота', 'setBotToken')
    .addItem('3. Подключить приложение к боту', 'connectBot')
    .addSeparator()
    .addItem('Обновить отчёты (ОДДС, План-факт, Остатки)', 'rebuildReports')
    .addItem('Добавить 12 месяцев в бюджет', 'extendBudget')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const months = nextMonths_(firstOfMonth_(new Date()), 12);

  ensureSheet_(ss, SH.ART, ART_HEAD, DEFAULT_ARTICLES.map(a => [a[0], a[1], a[2], a[3], true]), sh => {
    sh.getRange('E2:E').insertCheckboxes(); // сбрасывает значения в FALSE — ниже включаем обратно
    sh.getRange(2, 5, DEFAULT_ARTICLES.length, 1).setValue(true);
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Доход', 'Расход']).build());
    sh.setColumnWidths(1, 2, 140);
  });
  ensureSheet_(ss, SH.WAL, WAL_HEAD, DEFAULT_WALLETS, sh => {
    sh.getRange('E2:E').insertCheckboxes();
    sh.getRange(2, 5, DEFAULT_WALLETS.length, 1).setValue(true);
    sh.getRange('C2:C').setNumberFormat(MONEY);
    sh.setColumnWidth(1, 160);
  });
  ensureSheet_(ss, SH.CP, ['Контрагент'], [], sh => sh.setColumnWidth(1, 220));
  ensureSheet_(ss, SH.ACC, ACC_HEAD, [], sh => {
    sh.getRange('C2:C').insertCheckboxes();
    sh.getRange('A2:A').setNumberFormat('0');
    sh.setColumnWidths(1, 5, 150);
  });
  ensureSheet_(ss, SH.J, J_HEAD, [], sh => {
    sh.getRange('B2:B').setNumberFormat('dd.MM.yyyy');
    sh.getRange('D2:D').setNumberFormat(MONEY);
    sh.getRange('L2:L').setNumberFormat('dd.MM.yyyy HH:mm');
    sh.setColumnWidth(1, 80);
    sh.setColumnWidths(5, 5, 140);
    sh.setColumnWidth(10, 220);
  });
  ensureSheet_(ss, SH.B, ['Группа', 'Статья', 'Тип'].concat(months),
    DEFAULT_ARTICLES.map(a => [a[0], a[1], a[2]].concat(months.map(() => a[4]))), sh => {
      sh.getRange(1, 4, 1, months.length).setNumberFormat('mmm yyyy');
      sh.getRange(2, 4, sh.getMaxRows() - 1, months.length).setNumberFormat(MONEY0);
      sh.setFrozenColumns(3);
      sh.setColumnWidths(1, 2, 130);
    });

  rebuildReports(true);

  // Порядок листов и удаление пустого «Лист1»
  const order = [SH.J, SH.BAL, SH.PF, SH.ODDS, SH.B, SH.ART, SH.WAL, SH.CP, SH.ACC];
  order.forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  ss.getSheets().forEach(s => {
    if (order.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ss.setActiveSheet(ss.getSheetByName(SH.J));

  const msg = 'Готово! Листы созданы. Обновите вкладку таблицы — появится меню «💰 Бюджет». ' +
    'Дальше: начальные остатки на листе «Кошельки», затем меню 💰 Бюджет → «2. Указать токен бота».';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { ss.toast(msg, '💰 Бюджет', 15); Logger.log(msg); }
}

/* ===================== ОТЧЁТНЫЕ ЛИСТЫ ===================== */

function rebuildReports(silent) {
  const ss = SpreadsheetApp.getActive();
  syncBudgetRows_(ss);
  setJournalValidation_(ss);
  buildBalances_(ss);
  buildPlanFact_(ss);
  buildOdds_(ss);
  if (silent !== true) SpreadsheetApp.getActive().toast('Отчёты обновлены', '💰 Бюджет', 4);
}

function extendBudget() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH.B);
  const months = budgetMonths_(ss);
  const last = months.length ? months[months.length - 1] : firstOfMonth_(new Date());
  const add = nextMonths_(addMonths_(last, 1), 12);
  const lastCol = sh.getLastColumn();
  const n = Math.max(sh.getLastRow() - 1, 0);
  sh.getRange(1, lastCol + 1, 1, 12).setValues([add]).setNumberFormat('mmm yyyy');
  if (n) {
    const prev = sh.getRange(2, lastCol, n, 1).getValues();
    sh.getRange(2, lastCol + 1, n, 12).setValues(prev.map(r => add.map(() => r[0]))).setNumberFormat(MONEY0);
  }
  rebuildReports();
}

function setJournalValidation_(ss) {
  const sh = ss.getSheetByName(SH.J);
  const groups = uniq_(readArticles_(ss).map(a => a.group));
  const dv = SpreadsheetApp.newDataValidation;
  sh.getRange('C2:C').setDataValidation(dv().requireValueInList(TYPES).build());
  const wal = dv().requireValueInRange(ss.getSheetByName(SH.WAL).getRange('A2:A')).setAllowInvalid(false).build();
  sh.getRange('E2:F').setDataValidation(wal);
  if (groups.length) sh.getRange('G2:G').setDataValidation(dv().requireValueInList(groups).build());
  sh.getRange('H2:H').setDataValidation(dv().requireValueInRange(ss.getSheetByName(SH.ART).getRange('B2:B')).build());
  sh.getRange('I2:I').setDataValidation(dv().requireValueInRange(ss.getSheetByName(SH.CP).getRange('A2:A')).setAllowInvalid(true).build());
}

/** Добавляет в «Бюджет» строки для новых статей из справочника. */
function syncBudgetRows_(ss) {
  const sh = ss.getSheetByName(SH.B);
  const have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { have[r[0] + '|' + r[1]] = 1; });
  const add = readArticles_(ss).filter(a => !have[a.group + '|' + a.name]);
  if (!add.length) return;
  const nM = Math.max(sh.getLastColumn() - 3, 0);
  const rows = add.map(a => [a.group, a.name, a.type].concat(new Array(nM).fill(0)));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function buildBalances_(ss) {
  const sh = resetSheet_(ss, SH.BAL);
  const J = q_(SH.J), W = q_(SH.WAL);
  const wallets = readWallets_(ss);
  sh.getRange('A1').setValue('Остатки по кошелькам').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Считается автоматически из «Кошельков» (начальный остаток) и «Журнала».').setFontColor('#888');
  const head = ['Кошелёк', 'Начальный остаток', 'Поступления', 'Расходы', 'Переводы (+)', 'Переводы (−)', 'Текущий остаток'];
  const rows = wallets.map((w, i) => {
    const r = i + 5;
    return [w.name,
      `=SUMIFS(${W}$C:$C,${W}$A:$A,$A${r})`,
      `=SUMIFS(${J}$D:$D,${J}$C:$C,"Доход",${J}$E:$E,$A${r})`,
      `=SUMIFS(${J}$D:$D,${J}$C:$C,"Расход",${J}$E:$E,$A${r})`,
      `=SUMIFS(${J}$D:$D,${J}$C:$C,"Перевод",${J}$F:$F,$A${r})`,
      `=SUMIFS(${J}$D:$D,${J}$C:$C,"Перевод",${J}$E:$E,$A${r})`,
      `=B${r}+C${r}-D${r}+E${r}-F${r}`];
  });
  const tr = rows.length + 5;
  const total = ['Итого'].concat(['B', 'C', 'D', 'E', 'F', 'G'].map(c => rows.length ? `=SUM(${c}5:${c}${tr - 1})` : 0));
  const all = [head].concat(rows, [total]);
  sh.getRange(4, 1, all.length, head.length).setValues(all);
  styleHeader_(sh.getRange(4, 1, 1, head.length));
  sh.getRange(5, 2, all.length - 1, 6).setNumberFormat(MONEY);
  sh.getRange(tr, 1, 1, head.length).setFontWeight('bold').setBackground('#eef3fb');
  sh.getRange(5, 7, all.length - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 170); sh.setColumnWidths(2, 6, 130);
  sh.setFrozenRows(4);
}

function buildPlanFact_(ss) {
  const sh = resetSheet_(ss, SH.PF);
  const months = budgetMonths_(ss);
  const arts = readArticles_(ss);
  const J = q_(SH.J), B = q_(SH.B);
  const lastBCol = col_(3 + Math.max(months.length, 1));
  const cur = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM');
  const monthList = months.map(m => Utilities.formatDate(m, ss.getSpreadsheetTimeZone(), 'yyyy-MM'));

  sh.getRange('A1').setValue('План-факт').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Месяц (ГГГГ-ММ):').setFontWeight('bold');
  const sel = sh.getRange('B2');
  sel.setNumberFormat('@').setValue(monthList.indexOf(cur) >= 0 ? cur : (monthList[0] || cur)).setBackground('#fff8d6').setFontWeight('bold');
  if (monthList.length) sel.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(monthList).build());
  sh.getRange('C2').setFormula('=DATE(VALUE(LEFT(B2,4)),VALUE(RIGHT(B2,2)),1)').setNumberFormat('mmmm yyyy').setFontColor('#888');

  const rows = [];   // {a,b,kind,f:{C,D,E,F}}
  const R = () => rows.length + 5; // номер строки листа для следующей записи (заголовок в 4-й)
  const plan = r => `=IFERROR(SUMIFS(INDEX(${B}$D:$${lastBCol},0,MATCH($C$2,${B}$D$1:$${lastBCol}$1,0)),${B}$A:$A,$A${r},${B}$B:$B,$B${r}),0)`;
  const fact = (t, r) => `=SUMIFS(${J}$D:$D,${J}$C:$C,"${t}",${J}$G:$G,$A${r},${J}$H:$H,$B${r},${J}$B:$B,">="&$C$2,${J}$B:$B,"<"&EDATE($C$2,1))`;
  const totals = {};

  ['Доход', 'Расход'].forEach(t => {
    const sign = t === 'Доход' ? 1 : -1;
    rows.push({ a: t === 'Доход' ? 'ДОХОДЫ' : 'РАСХОДЫ', b: '', kind: 'section' });
    const groupRows = [];
    uniq_(arts.filter(a => a.type === t).map(a => a.group)).forEach(g => {
      const gRow = R();
      const items = arts.filter(a => a.type === t && a.group === g);
      const first = gRow + 1, last = gRow + items.length;
      rows.push({ a: g, b: 'Итого ' + g, kind: 'group', f: { C: `=SUM(C${first}:C${last})`, D: `=SUM(D${first}:D${last})` } });
      groupRows.push(gRow);
      items.forEach(a => { const r = R(); rows.push({ a: g, b: a.name, kind: 'item', f: { C: plan(r), D: fact(t, r) } }); });
    });
    const tr = R();
    const sumOf = c => groupRows.length ? '=' + groupRows.map(r => c + r).join('+') : '=0';
    rows.push({ a: t === 'Доход' ? 'Итого доходы' : 'Итого расходы', b: '', kind: 'total', f: { C: sumOf('C'), D: sumOf('D') } });
    totals[t] = tr;
    rows.forEach(x => { if (x.f && x.sign === undefined) x.sign = sign; });
  });
  const sr = R();
  rows.push({ a: 'САЛЬДО (доходы − расходы)', b: '', kind: 'total', sign: 1,
    f: { C: `=C${totals['Доход']}-C${totals['Расход']}`, D: `=D${totals['Доход']}-D${totals['Расход']}` } });

  const head = ['Группа', 'Статья', 'План', 'Факт', 'Отклонение', '% исполнения'];
  const data = rows.map((x, i) => {
    const r = i + 5;
    if (!x.f) return [x.a, x.b, '', '', '', ''];
    const dev = x.sign > 0 ? `=D${r}-C${r}` : `=C${r}-D${r}`;
    return [x.a, x.b, x.f.C, x.f.D, dev, `=IFERROR(D${r}/C${r},"")`];
  });
  sh.getRange(4, 1, 1, 6).setValues([head]);
  styleHeader_(sh.getRange(4, 1, 1, 6));
  sh.getRange(5, 1, data.length, 6).setValues(data);
  sh.getRange(5, 3, data.length, 3).setNumberFormat(MONEY0);
  sh.getRange(5, 6, data.length, 1).setNumberFormat('0%');
  styleRows_(sh, rows, 6, 5);
  sh.getRange(sr, 1, 1, 6).setBackground('#dfe9f7');

  const devRange = sh.getRange(5, 5, data.length, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#d93025').setRanges([devRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#188038').setRanges([devRange]).build()
  ]);
  sh.getRange('E3').setValue('Отклонение: + в вашу пользу (экономия / доход выше плана), − перерасход').setFontColor('#888').setFontSize(9);
  sh.setColumnWidths(1, 2, 160); sh.setColumnWidths(3, 4, 120);
  sh.setFrozenRows(4);
}

function buildOdds_(ss) {
  const sh = resetSheet_(ss, SH.ODDS);
  const months = budgetMonths_(ss);
  const arts = readArticles_(ss);
  const J = q_(SH.J), W = q_(SH.WAL);
  const n = Math.max(months.length, 1);
  const c0 = 3, cEnd = c0 + n - 1, cTot = cEnd + 1;
  const L = i => col_(c0 + i);

  sh.getRange('A1').setValue('ОДДС — отчёт о движении денежных средств').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Переводы между своими кошельками не влияют на итоги. Месяцы — из листа «Бюджет».').setFontColor('#888');

  const rows = []; // {a,b,kind,fn:(L)=>formula, tot:'sum'|'first'|'last'}
  const R = () => rows.length + 4;
  const flow = (t) => (l, r) => `=SUMIFS(${J}$D:$D,${J}$C:$C,"${t}",${J}$G:$G,$A${r},${J}$H:$H,$B${r},${J}$B:$B,">="&${l}$3,${J}$B:$B,"<"&EDATE(${l}$3,1))`;

  const startRow = R();
  rows.push({ a: 'Остаток на начало', b: '', kind: 'total', tot: 'first',
    fn: l => `=SUM(${W}$C$2:$C)+SUMIFS(${J}$D:$D,${J}$C:$C,"Доход",${J}$B:$B,"<"&${l}$3)-SUMIFS(${J}$D:$D,${J}$C:$C,"Расход",${J}$B:$B,"<"&${l}$3)` });

  const sectionTotals = {};
  [['Доход', 'ПОСТУПЛЕНИЯ', 'Итого поступления'], ['Расход', 'ВЫБЫТИЯ', 'Итого выбытия']].forEach(([t, title, totTitle]) => {
    rows.push({ a: title, b: '', kind: 'section' });
    const gRows = [];
    uniq_(arts.filter(a => a.type === t).map(a => a.group)).forEach(g => {
      const items = arts.filter(a => a.type === t && a.group === g);
      const gr = R(); gRows.push(gr);
      rows.push({ a: g, b: 'Итого ' + g, kind: 'group', tot: 'sum', fn: l => `=SUM(${l}${gr + 1}:${l}${gr + items.length})` });
      items.forEach(a => rows.push({ a: g, b: a.name, kind: 'item', tot: 'sum', fn: flow(t) }));
    });
    sectionTotals[t] = R();
    rows.push({ a: totTitle, b: '', kind: 'total', tot: 'sum', fn: l => gRows.length ? '=' + gRows.map(r => l + r).join('+') : '=0' });
  });
  const netRow = R();
  rows.push({ a: 'Чистый денежный поток', b: '', kind: 'total', tot: 'sum', fn: l => `=${l}${sectionTotals['Доход']}-${l}${sectionTotals['Расход']}` });
  rows.push({ a: 'Остаток на конец', b: '', kind: 'total', tot: 'last', fn: l => `=${l}${startRow}+${l}${netRow}` });

  const head = ['Группа', 'Статья'].concat(months.length ? months : [firstOfMonth_(new Date())], ['Итого']);
  const data = rows.map((x, i) => {
    const r = i + 4;
    const line = [x.a, x.b];
    for (let k = 0; k < n; k++) line.push(x.fn ? x.fn(L(k), r) : '');
    const T = x.tot === 'sum' ? `=SUM(${L(0)}${r}:${L(n - 1)}${r})` : x.tot === 'first' ? `=${L(0)}${r}` : x.tot === 'last' ? `=${L(n - 1)}${r}` : '';
    line.push(T);
    return line;
  });
  sh.getRange(3, 1, 1, head.length).setValues([head]);
  sh.getRange(3, c0, 1, n).setNumberFormat('mmm yyyy');
  styleHeader_(sh.getRange(3, 1, 1, head.length));
  sh.getRange(4, 1, data.length, head.length).setValues(data);
  sh.getRange(4, c0, data.length, n + 1).setNumberFormat(MONEY0);
  styleRows_(sh, rows, head.length, 4);
  sh.getRange(4, cTot, data.length, 1).setFontWeight('bold');
  sh.setFrozenRows(3); sh.setFrozenColumns(2);
  sh.setColumnWidths(1, 2, 150); sh.setColumnWidths(c0, n + 1, 105);
}

/* ======================= TELEGRAM / API ======================= */

function setBotToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Токен бота', 'Вставьте токен от @BotFather (вида 123456:ABC...). Он хранится только в настройках этого скрипта.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const token = res.getResponseText().trim();
  const me = tg_(token, 'getMe', {});
  if (!me.ok) { ui.alert('Telegram не принял токен: ' + (me.description || 'ошибка')); return; }
  PropertiesService.getScriptProperties().setProperty('BOT_TOKEN', token);
  ui.alert('Токен сохранён. Бот: @' + me.result.username +
    '\n\nДальше: «Развернуть → Новое развёртывание» (веб-приложение), затем меню 💰 Бюджет → «3. Подключить приложение к боту».');
}

function connectBot() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('BOT_TOKEN');
  if (!token) { ui.alert('Сначала укажите токен бота (пункт 2 меню).'); return; }
  let execUrl = props.getProperty('WEBAPP_URL') || '';
  try { const u = ScriptApp.getService().getUrl(); if (u && /\/exec$/.test(u)) execUrl = u; } catch (e) { }
  if (!/\/exec$/.test(execUrl)) {
    const res = ui.prompt('Адрес веб-приложения', 'Вставьте URL развёртывания (заканчивается на /exec):', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    execUrl = res.getResponseText().trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(execUrl)) { ui.alert('Это не похоже на URL веб-приложения Apps Script.'); return; }
  }
  props.setProperty('WEBAPP_URL', execUrl);
  const appUrl = (props.getProperty('APP_URL') || DEFAULT_APP_URL) + '?api=' + encodeURIComponent(execUrl);
  const r = tg_(token, 'setChatMenuButton', { menu_button: { type: 'web_app', text: 'Бюджет', web_app: { url: appUrl } } });
  if (!r.ok) { ui.alert('Ошибка Telegram: ' + r.description); return; }
  tg_(token, 'setMyCommands', { commands: [{ command: 'start', description: 'Открыть бюджет' }] });
  const me = tg_(token, 'getMe', {});
  ui.alert('Готово! Откройте @' + me.result.username + ' в Telegram и нажмите кнопку «Бюджет» слева от поля ввода.\n\n' +
    'При первом входе приложение покажет ваш Telegram ID и добавит строку на лист «Доступ» — поставьте там галочку «Разрешён».');
}

function tg_(token, method, payload) {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  return JSON.parse(r.getContentText());
}

function doGet() {
  return ContentService.createTextOutput('Семейный бюджет API: OK');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const tgUser = verifyInitData_(req.initData || '');
    const ss = SpreadsheetApp.getActive();
    const user = checkAccess_(ss, tgUser);
    const p = req.payload || {};
    let data;
    switch (req.action) {
      case 'load': data = apiLoad_(ss, user); break;
      case 'add': data = withLock_(() => saveOp_(ss, user, p.op, null)); break;
      case 'update': data = withLock_(() => saveOp_(ss, user, p.op, p.op && p.op.id)); break;
      case 'delete': data = withLock_(() => deleteOp_(ss, p.id)); break;
      case 'saveWallet': withLock_(() => saveWallet_(ss, p.wallet, p.oldName)); data = apiLoad_(ss, user); break;
      default: throw new Error('Неизвестное действие');
    }
    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err), code: err && err.code, id: err && err.tgId };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Проверка подписи initData по алгоритму Telegram Mini Apps. */
function verifyInitData_(initData) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) throw err_('Сервер не настроен: не указан токен бота', 'config');
  if (!initData) throw err_('Откройте приложение через Telegram', 'no_tg');
  const params = {};
  let hash = '';
  initData.split('&').forEach(part => {
    const i = part.indexOf('=');
    if (i < 0) return;
    const k = decodeURIComponent(part.slice(0, i));
    const v = decodeURIComponent(part.slice(i + 1).replace(/\+/g, '%20'));
    if (k === 'hash') hash = v; else params[k] = v;
  });
  const dcs = Object.keys(params).sort().map(k => k + '=' + params[k]).join('\n');
  const secret = Utilities.computeHmacSha256Signature(bytes_(token), bytes_('WebAppData'));
  const sig = Utilities.computeHmacSha256Signature(bytes_(dcs), secret);
  const hex = sig.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  if (!hash || hex !== hash) throw err_('Подпись Telegram не прошла проверку', 'bad_sign');
  const age = Date.now() / 1000 - Number(params.auth_date || 0);
  if (age > 7 * 86400) throw err_('Сессия устарела, переоткройте приложение', 'expired');
  return JSON.parse(params.user || '{}');
}

function checkAccess_(ss, u) {
  const sh = ss.getSheetByName(SH.ACC);
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  const row = vals.find(r => String(r[0]).trim() === String(u.id));
  const name = [u.first_name, u.last_name].filter(Boolean).join(' ') || u.username || String(u.id);
  if (!row) {
    withLock_(() => {
      sh.appendRow([u.id, name, false, '', '']);
      sh.getRange(sh.getLastRow(), 3).insertCheckboxes();
    });
  }
  if (!row || row[2] !== true) {
    const e = err_('Нет доступа', 'forbidden'); e.tgId = u.id; throw e;
  }
  return { id: u.id, name: row[1] || name, group: row[3] || '', wallet: row[4] || '' };
}

function apiLoad_(ss, user) {
  const tz = ss.getSpreadsheetTimeZone();
  const budget = {};
  const bs = ss.getSheetByName(SH.B);
  if (bs.getLastRow() > 1 && bs.getLastColumn() > 3) {
    const v = bs.getRange(1, 1, bs.getLastRow(), bs.getLastColumn()).getValues();
    const heads = v[0].slice(3).map(d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM') : '');
    v.slice(1).forEach(r => {
      heads.forEach((m, i) => {
        if (!m) return;
        const x = Number(r[3 + i]) || 0;
        if (!x) return;
        (budget[m] = budget[m] || {})[r[0] + '|' + r[1]] = x;
      });
    });
  }
  return {
    user: user,
    articles: readArticles_(ss),
    wallets: readWallets_(ss),
    counterparties: readCounterparties_(ss),
    budget: budget,
    ops: readOps_(ss)
  };
}

function readOps_(ss) {
  const sh = ss.getSheetByName(SH.J);
  const tz = ss.getSpreadsheetTimeZone();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const v = sh.getRange(2, 1, n, J_HEAD.length).getValues();
  const out = [];
  let fixed = false;
  v.forEach((r, i) => {
    if (!(r[1] instanceof Date) || !r[2] || !r[3]) return;
    if (!r[0]) { r[0] = newId_(); sh.getRange(i + 2, 1).setValue(r[0]); fixed = true; }
    out.push({
      id: String(r[0]), date: Utilities.formatDate(r[1], tz, 'yyyy-MM-dd'), type: r[2], amount: Number(r[3]) || 0,
      wallet: r[4], wallet2: r[5], group: r[6], article: r[7], cp: r[8], comment: r[9], author: r[10]
    });
  });
  return out;
}

function saveOp_(ss, user, op, id) {
  if (!op) throw new Error('Нет данных операции');
  const arts = readArticles_(ss), wals = readWallets_(ss).map(w => w.name);
  const type = String(op.type);
  const amount = Math.round(Number(op.amount) * 100) / 100;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(op.date || '');
  if (TYPES.indexOf(type) < 0) throw new Error('Неверный тип операции');
  if (!(amount > 0)) throw new Error('Сумма должна быть больше нуля');
  if (!m) throw new Error('Неверная дата');
  if (wals.indexOf(op.wallet) < 0) throw new Error('Кошелёк не найден: ' + op.wallet);
  let group = '', article = '', wallet2 = '';
  if (type === 'Перевод') {
    if (wals.indexOf(op.wallet2) < 0 || op.wallet2 === op.wallet) throw new Error('Выберите другой кошелёк-получатель');
    wallet2 = op.wallet2;
  } else {
    const a = arts.find(x => x.group === op.group && x.name === op.article && x.type === type);
    if (!a) throw new Error('Статья не найдена: ' + op.group + ' / ' + op.article);
    group = a.group; article = a.name;
  }
  const cp = String(op.cp || '').trim().slice(0, 100);
  const comment = String(op.comment || '').trim().slice(0, 500);
  const date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  const sh = ss.getSheetByName(SH.J);
  let rowIdx;
  if (id) {
    rowIdx = findRow_(sh, id);
    if (!rowIdx) throw new Error('Операция не найдена (возможно, удалена)');
  } else {
    id = newId_();
    rowIdx = sh.getLastRow() + 1;
  }
  const author = id && rowIdx <= sh.getLastRow() ? (sh.getRange(rowIdx, 11).getValue() || user.name) : user.name;
  sh.getRange(rowIdx, 1, 1, J_HEAD.length).setValues([[id, date, type, amount, op.wallet, wallet2, group, article, cp, comment, author, new Date()]]);
  if (cp) addCounterparty_(ss, cp);
  const tz = ss.getSpreadsheetTimeZone();
  return { id: id, date: Utilities.formatDate(date, tz, 'yyyy-MM-dd'), type: type, amount: amount, wallet: op.wallet,
    wallet2: wallet2, group: group, article: article, cp: cp, comment: comment, author: author };
}

function deleteOp_(ss, id) {
  const sh = ss.getSheetByName(SH.J);
  const r = findRow_(sh, id);
  if (!r) throw new Error('Операция не найдена');
  sh.deleteRow(r);
  return { id: id };
}

/** Создание / изменение кошелька. При переименовании обновляет Журнал и Доступ. */
function saveWallet_(ss, w, oldName) {
  if (!w) throw new Error('Нет данных кошелька');
  const name = String(w.name || '').trim().slice(0, 60);
  if (!name) throw new Error('Введите название кошелька');
  const initial = Math.round((Number(w.initial) || 0) * 100) / 100;
  const sh = ss.getSheetByName(SH.WAL);
  const n = Math.max(sh.getLastRow() - 1, 0);
  const names = n ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()) : [];
  const oldIdx = oldName ? names.indexOf(String(oldName)) : -1;
  if (oldName && oldIdx < 0) throw new Error('Кошелёк не найден: ' + oldName);
  const dup = names.findIndex(x => x.toLowerCase() === name.toLowerCase());
  if (dup >= 0 && dup !== oldIdx) throw new Error('Кошелёк «' + name + '» уже есть');
  const row = [name, String(w.owner || ''), initial, String(w.icon || '💳'), w.active !== false];
  const r = oldIdx >= 0 ? oldIdx + 2 : sh.getLastRow() + 1;
  if (oldIdx < 0) sh.getRange(r, 5).insertCheckboxes();
  sh.getRange(r, 1, 1, 5).setValues([row]);
  sh.getRange(r, 3).setNumberFormat(MONEY);
  if (oldIdx >= 0 && oldName !== name) {
    renameInColumns_(ss.getSheetByName(SH.J), [5, 6], oldName, name);
    renameInColumns_(ss.getSheetByName(SH.ACC), [5], oldName, name);
  }
  buildBalances_(ss);
}

function renameInColumns_(sh, cols, from, to) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  cols.forEach(c => {
    const rg = sh.getRange(2, c, n, 1);
    const v = rg.getValues();
    let changed = false;
    v.forEach(r => { if (String(r[0]) === from) { r[0] = to; changed = true; } });
    if (changed) rg.setValues(v);
  });
}

/* ========================= СПРАВОЧНИКИ ========================= */

function readArticles_(ss) {
  const sh = ss.getSheetByName(SH.ART);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => r[0] && r[1] && (r[2] === 'Доход' || r[2] === 'Расход'))
    .map(r => ({ group: String(r[0]).trim(), name: String(r[1]).trim(), type: r[2], icon: r[3] || '', active: r[4] !== false }));
}

function readWallets_(ss) {
  const sh = ss.getSheetByName(SH.WAL);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), owner: r[1] || '', initial: Number(r[2]) || 0, icon: r[3] || '', active: r[4] !== false }));
}

function readCounterparties_(ss) {
  const sh = ss.getSheetByName(SH.CP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function addCounterparty_(ss, name) {
  const list = readCounterparties_(ss).map(s => s.toLowerCase());
  if (list.indexOf(name.toLowerCase()) < 0) ss.getSheetByName(SH.CP).appendRow([name]);
}

/* =========================== УТИЛИТЫ =========================== */

function ensureSheet_(ss, name, head, rows, format) {
  let sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() > 0) return sh;
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  styleHeader_(sh.getRange(1, 1, 1, head.length));
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  if (format) format(sh);
  return sh;
}

function resetSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.setFrozenRows(0); sh.setFrozenColumns(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.clear(); sh.clearConditionalFormatRules();
  return sh;
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#2b5797').setFontColor('#ffffff').setVerticalAlignment('middle');
}

function styleRows_(sh, rows, width, startRow) {
  rows.forEach((x, i) => {
    const rg = sh.getRange(startRow + i, 1, 1, width);
    if (x.kind === 'section') rg.setFontWeight('bold').setBackground('#e8eaed');
    else if (x.kind === 'group') rg.setFontWeight('bold').setBackground('#f5f7fa');
    else if (x.kind === 'total') rg.setFontWeight('bold').setBackground('#eef3fb');
    else if (x.kind === 'item') sh.getRange(startRow + i, 1).setFontColor('#999999');
  });
}

function budgetMonths_(ss) {
  const sh = ss.getSheetByName(SH.B);
  if (!sh || sh.getLastColumn() < 4) return [];
  return sh.getRange(1, 4, 1, sh.getLastColumn() - 3).getValues()[0].filter(d => d instanceof Date);
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function firstOfMonth_(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function addMonths_(d, k) { return new Date(d.getFullYear(), d.getMonth() + k, 1); }
function nextMonths_(start, k) { const a = []; for (let i = 0; i < k; i++) a.push(addMonths_(start, i)); return a; }
function uniq_(a) { return a.filter((x, i) => a.indexOf(x) === i); }
function q_(name) { return "'" + name + "'!"; }
function bytes_(s) { return Utilities.newBlob(s).getBytes(); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function err_(msg, code) { const e = new Error(msg); e.code = code; return e; }
function col_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
