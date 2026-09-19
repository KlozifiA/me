/**
 * Приём заявок с сайта Klozifi → Google Таблица + письмо на почту.
 *
 * Как подключить (один раз, ~5 минут):
 * 1. Создай Google Таблицу (sheets.new), назови например «Заявки Klozifi».
 * 2. В таблице: Расширения → Apps Script. Удали всё в редакторе и вставь этот файл.
 * 3. Сохрани (Ctrl+S). Нажми «Начать развертывание» → «Новое развертывание».
 *    Тип: «Веб-приложение».
 *    Выполнять как: «От моего имени».
 *    У кого есть доступ: «Все».
 * 4. Нажми «Развернуть», разреши доступ к своему аккаунту (Google покажет
 *    предупреждение «приложение не проверено» — это твой собственный скрипт,
 *    жми «Дополнительно» → «Перейти к проекту»).
 * 5. Скопируй «URL веб-приложения» (заканчивается на /exec) и вставь его
 *    в index.html в строку  const LEAD_ENDPOINT = '';
 *
 * Если потом поменяешь код скрипта — делай «Управление развертываниями» →
 * карандаш → Версия: «Новая версия», иначе изменения не применятся.
 */

// Куда слать письма. Пусто = на почту владельца скрипта (твою).
const NOTIFY_EMAIL = '';
const SHEET_NAME = 'Заявки';
const HEADERS = ['Дата', 'Имя', 'Контакт', 'Услуга', 'Сроки', 'Сообщение', 'Страница'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Ловушка для спам-ботов: реальный человек это поле не видит и не заполняет
  if (p.website) return json({ ok: true });
  if (!p.name || !p.contact) return json({ ok: false, error: 'missing fields' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const row = [
      new Date(),
      clean(p.name, 80),
      clean(p.contact, 100),
      clean(p.service, 100),
      clean(p.deadline, 50),
      clean(p.message, 2000),
      clean(p.page, 200),
    ];
    sheet.appendRow(row);
    notify(row);
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Открой ссылку /exec в браузере — если видишь {"ok":true,...}, скрипт работает
function doGet() {
  return json({ ok: true, info: 'Klozifi lead endpoint' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function notify(row) {
  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  const [date, name, contact, service, deadline, message, page] = row;
  const body = [
    'Новая заявка с сайта',
    '',
    'Имя: ' + name,
    'Контакт: ' + contact,
    'Услуга: ' + service,
    'Сроки: ' + deadline,
    '',
    'Сообщение:',
    message || '—',
    '',
    'Страница: ' + page,
    'Время: ' + Utilities.formatDate(date, 'Europe/Moscow', 'dd.MM.yyyy HH:mm'),
    '',
    'Все заявки: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  ].join('\n');
  MailApp.sendEmail({ to: to, subject: 'Новая заявка: ' + name + ' — ' + service, body: body });
}

// Обрезает длину и не даёт вставить формулу в таблицу (=, +, -, @ в начале)
function clean(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
