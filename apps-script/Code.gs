/**
 * Анкета гостей юбилея: Надира Мырзагулова, 50 лет.
 *
 * Сайт отправляет сюда ответы гостей. Скрипт:
 *  - записывает ответ строкой на лист «Ответы»;
 *  - обновляет итог на листе «Итог» (сколько человек придёт);
 *  - отправляет письмо владельцу таблицы и всем, у кого есть доступ
 *    к таблице на редактирование. Новому человеку достаточно дать
 *    доступ «Редактор», и письма начнут приходить ему тоже.
 */

const SHEET_ANSWERS = 'Ответы';
const SHEET_TOTAL = 'Итог';

/** Запустите один раз вручную: создаёт листы, заголовки и формулы. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let answers = ss.getSheetByName(SHEET_ANSWERS);
  if (!answers) answers = ss.insertSheet(SHEET_ANSWERS);
  answers.getRange(1, 1, 1, 6)
    .setValues([['Когда ответил', 'Имя', 'Придёт', 'Сколько человек', 'Пожелание', 'Язык']])
    .setFontWeight('bold').setBackground('#4A0F1F').setFontColor('#F6ECE2');
  answers.setFrozenRows(1);
  answers.setColumnWidths(1, 1, 150);
  answers.setColumnWidth(2, 220);
  answers.setColumnWidths(3, 2, 130);
  answers.setColumnWidth(5, 320);

  let total = ss.getSheetByName(SHEET_TOTAL);
  if (!total) total = ss.insertSheet(SHEET_TOTAL, 0);
  // Формулы без разделителей аргументов: работают в любой локали таблицы.
  const A = `'${SHEET_ANSWERS}'`;
  total.getRange('A1:B4').setValues([
    ['Всего придёт человек', `=SUMPRODUCT((${A}!C2:C2000="Да")*${A}!D2:D2000)`],
    ['Ответили «приду»', `=SUMPRODUCT(--(${A}!C2:C2000="Да"))`],
    ['Ответили «не смогу»', `=SUMPRODUCT(--(${A}!C2:C2000="Нет"))`],
    ['Всего ответов', `=COUNTA(${A}!B2:B2000)`],
  ]);
  total.getRange('A1:A4').setFontWeight('bold');
  total.getRange('B1').setFontSize(20).setFontWeight('bold').setFontColor('#7A1C33');
  total.setColumnWidth(1, 220);

  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(def);

  // Проверка, что почта работает: придёт тестовое письмо владельцу.
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    'Анкета юбилея подключена',
    'Таблица готова. Ответы гостей будут приходить на эту почту.');
}

/** Приём ответа с сайта. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    const name = clean_(p.name, 80);
    if (!name) return json_({ ok: false, error: 'no name' });

    const yes = p.attend === 'yes';
    const guests = yes ? Math.min(Math.max(parseInt(p.guests, 10) || 1, 1), 50) : 0;
    const comment = clean_(p.comment, 500);
    const lang = p.lang === 'ru' ? 'RU' : 'KG';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ANSWERS);
    sheet.appendRow([new Date(), name, yes ? 'Да' : 'Нет', guests, comment, lang]);
    SpreadsheetApp.flush();

    notify_(ss, { name, yes, guests, comment });
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, info: 'Анкета юбилея работает' });
}

function notify_(ss, a) {
  const total = ss.getSheetByName(SHEET_TOTAL);
  const people = total ? total.getRange('B1').getValue() : '';
  const answers = total ? total.getRange('B4').getValue() : '';

  const emails = new Set();
  const owner = ss.getOwner();
  if (owner) emails.add(owner.getEmail());
  ss.getEditors().forEach(u => { if (u.getEmail()) emails.add(u.getEmail()); });
  if (!emails.size) emails.add(Session.getEffectiveUser().getEmail());

  const status = a.yes ? `придёт, ${a.guests} чел.` : 'не сможет прийти';
  const subject = `Юбилей: ${a.name} — ${status}`;
  const html =
    `<div style="font-family:Georgia,serif;font-size:16px;color:#3A0A16">` +
    `<p style="font-size:20px"><b>${esc_(a.name)}</b> — ${status}</p>` +
    (a.comment ? `<p>Пожелание: <i>${esc_(a.comment)}</i></p>` : '') +
    `<hr style="border:0;border-top:1px solid #C9A35E">` +
    `<p>Сейчас всего придёт: <b>${people}</b> чел. (ответов: ${answers})</p>` +
    `<p><a href="${ss.getUrl()}" style="color:#7A1C33">Открыть таблицу с ответами</a></p></div>`;

  MailApp.sendEmail({ to: [...emails].join(','), subject, htmlBody: html, name: 'Юбилей — анкета' });
}

function clean_(v, max) {
  return String(v || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max)
    .replace(/^[=+\-@]/, "'$&"); // не даём ответу стать формулой
}
function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
