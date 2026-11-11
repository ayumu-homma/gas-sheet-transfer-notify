/**
 * gas-sheet-transfer-notify
 *
 * 公開用のサンプルであり、受注成果物ではありません。
 * 「受付」シートで状態が「要対応」の行を「処理済み」へ転記し、
 * テスト用アドレスへメールを1通送り、状態を「転記済」にします。
 *
 * 重要:
 *   メール送信があるため、単純な onEdit トリガーでは動きません。
 *   インストール可能なトリガーから handleFormSubmit / handleEdit を呼んでください。
 *   関数名を onEdit にしていないのは、単純トリガーとの二重起動と権限不足を避けるためです。
 */

var CONFIG = Object.freeze({
  // スプレッドシートにバインドして使う場合は空のままでよい。
  // clasp のスタンドアロンで使う場合だけ、スプレッドシート ID を入れる。
  // その場合は appsscript.json のスコープを spreadsheets に広げる必要がある。
  SPREADSHEET_ID: '',
  INTAKE_SHEET: '受付',
  DONE_SHEET: '処理済み',
  STATUS_NEED: '要対応',
  STATUS_DONE: '転記済',
  HEADERS: ['タイムスタンプ', '名前（ダミー）', '内容', '状態', 'メール（ダミー）'],
  EMAIL_SUBJECT: '【サンプル】新しい要対応行を転記しました'
});

/**
 * インストール可能なトリガー「フォーム送信時」から呼び出す。
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function handleFormSubmit(e) {
  try {
    var row = e && e.range ? e.range.getRow() : null;
    if (!row || row < 2) {
      throw new Error('フォーム送信イベントから行番号を取得できませんでした。');
    }
    processIntakeRow_(row, 'formSubmit');
  } catch (err) {
    logError_('handleFormSubmit', err);
    throw err;
  }
}

/**
 * インストール可能なトリガー「編集時」から呼び出す。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleEdit(e) {
  try {
    if (!e || !e.range) {
      return;
    }
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.INTAKE_SHEET) {
      return;
    }
    var row = e.range.getRow();
    if (row < 2) {
      return;
    }
    processIntakeRow_(row, 'edit');
  } catch (err) {
    logError_('handleEdit', err);
    throw err;
  }
}

/**
 * 受付シートの先頭の「要対応」行を1件処理する（トリガーなしの動作確認用）。
 */
function testProcessFirstNeed() {
  var ss = getSpreadsheet_();
  var intake = getSheetOrThrow_(ss, CONFIG.INTAKE_SHEET);
  var map = getHeaderMap_(intake);
  var last = intake.getLastRow();
  if (last < 2) {
    throw new Error('受付シートにデータ行がありません。先に setupOnce を実行してください。');
  }
  var statuses = intake.getRange(2, map['状態'], last - 1, 1).getDisplayValues();
  var i;
  for (i = 0; i < statuses.length; i++) {
    if (String(statuses[i][0]).trim() === CONFIG.STATUS_NEED) {
      processIntakeRow_(i + 2, 'manualTest');
      Logger.log('row ' + (i + 2) + ' を処理しました。メールと処理済みシートを確認してください。');
      return;
    }
  }
  throw new Error('状態が「要対応」の行がありません。受付シートの状態列を「要対応」に変更してから再実行してください。');
}

/**
 * 初期セットアップ。このスクリプトをスプレッドシートにバインドした状態で実行する。
 * 既存の同名シートがある場合は、中身を上書きしない。
 */
function setupOnce() {
  var ss = getSpreadsheet_();
  var intake = ensureSheet_(ss, CONFIG.INTAKE_SHEET, CONFIG.HEADERS);
  ensureSheet_(ss, CONFIG.DONE_SHEET, CONFIG.HEADERS);

  if (intake.getLastRow() < 2) {
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    intake.getRange(2, 1, 2, CONFIG.HEADERS.length).setValues([
      [now, 'サンプルA', '週次の在庫メモを転記してほしい（ダミー）', CONFIG.STATUS_NEED, 'sample-user@example.com'],
      [now, 'サンプルB', '対応不要のテスト行（ダミー）', '対応不要', 'sample-user2@example.com']
    ]);
  }

  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_TO');
  if (!to) {
    Logger.log('次に、プロジェクトの設定 > スクリプト プロパティへ NOTIFY_TO（自分のテスト用メールアドレス）を追加してください。');
  }
  Logger.log('setupOnce が完了しました。NOTIFY_TO を設定したあと、testProcessFirstNeed で動作確認できます。');
}

function processIntakeRow_(row, source) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('同時実行のロックを取得できませんでした。しばらくしてから同じ行を再実行してください。');
  }

  var statusCell;
  var previousStatus;
  var appendedDone = false;
  var doneSheet;
  var doneAppendRow;

  try {
    var ss = getSpreadsheet_();
    var intake = getSheetOrThrow_(ss, CONFIG.INTAKE_SHEET);
    doneSheet = getSheetOrThrow_(ss, CONFIG.DONE_SHEET);
    var map = getHeaderMap_(intake);
    var doneMap = getHeaderMap_(doneSheet);

    statusCell = intake.getRange(row, map['状態']);
    var status = String(statusCell.getDisplayValue()).trim();

    if (status === CONFIG.STATUS_DONE) {
      console.log('skip: row ' + row + ' is already ' + CONFIG.STATUS_DONE + ' (source=' + source + ')');
      return;
    }

    var treatAsNeed = (status === CONFIG.STATUS_NEED) ||
      (source === 'formSubmit' && status === '');
    if (!treatAsNeed) {
      console.log('skip: row ' + row + ' status="' + status + '" (source=' + source + ')');
      return;
    }

    var rec = readRowByHeaders_(intake, row, map);
    if (!String(rec['内容'] || '').trim() && !String(rec['名前（ダミー）'] || '').trim()) {
      console.log('skip: row ' + row + ' is empty');
      return;
    }

    var notifyTo = getNotifyTo_();
    var timestamp = rec['タイムスタンプ'] ||
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    // 先に転記済にして二重処理を防ぐ。失敗時はこの関数の末尾で戻す。
    previousStatus = status;
    statusCell.setValue(CONFIG.STATUS_DONE);

    var doneLastCol = doneSheet.getLastColumn();
    var doneHeaders = doneSheet.getRange(1, 1, 1, doneLastCol).getDisplayValues()[0];
    var out = [];
    var i;
    for (i = 0; i < doneLastCol; i++) {
      out[i] = '';
    }
    for (i = 0; i < doneHeaders.length; i++) {
      var key = String(doneHeaders[i]).trim();
      if (key && Object.prototype.hasOwnProperty.call(rec, key)) {
        out[i] = rec[key];
      }
    }
    if (doneMap['状態']) {
      out[doneMap['状態'] - 1] = CONFIG.STATUS_DONE;
    }
    if (doneMap['タイムスタンプ'] && !String(out[doneMap['タイムスタンプ'] - 1] || '').trim()) {
      out[doneMap['タイムスタンプ'] - 1] = timestamp;
    }

    doneAppendRow = doneSheet.getLastRow() + 1;
    doneSheet.getRange(doneAppendRow, 1, 1, doneLastCol).setValues([out]);
    appendedDone = true;

    var body = [
      'これは公開サンプルの通知です。受注案件の通知ではありません。',
      '',
      '発生元: ' + source,
      'タイムスタンプ: ' + timestamp,
      '名前（ダミー）: ' + rec['名前（ダミー）'],
      '内容: ' + rec['内容'],
      'メール（ダミー・参考）: ' + rec['メール（ダミー）'],
      '',
      'このメールはスクリプト プロパティ NOTIFY_TO のテスト用アドレスへ送っています。'
    ].join('\n');

    MailApp.sendEmail({
      to: notifyTo,
      subject: CONFIG.EMAIL_SUBJECT,
      body: body
    });

    console.log('ok: row ' + row + ' transferred and notified (source=' + source + ')');
  } catch (err) {
    if (statusCell && previousStatus !== undefined) {
      try {
        statusCell.setValue(previousStatus);
      } catch (rollbackErr) {
        logError_('rollbackStatus', rollbackErr);
      }
    }
    if (appendedDone && doneSheet && doneAppendRow) {
      try {
        doneSheet.deleteRow(doneAppendRow);
      } catch (rollbackErr) {
        logError_('rollbackDoneRow', rollbackErr);
      }
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートを特定できません。スプレッドシートから開くか、CONFIG.SPREADSHEET_ID を設定してください。');
  }
  return ss;
}

function getNotifyTo_() {
  var to = String(PropertiesService.getScriptProperties().getProperty('NOTIFY_TO') || '').trim();
  if (!to) {
    throw new Error('スクリプト プロパティ NOTIFY_TO が未設定です。自分のテスト用アドレスを設定してください。');
  }
  return to;
}

function getSheetOrThrow_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。名前の誤記がないか確認するか、setupOnce を実行してください。');
  }
  return sheet;
}

function getHeaderMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};
  var i;
  for (i = 0; i < headers.length; i++) {
    var key = String(headers[i]).trim();
    if (key) {
      map[key] = i + 1;
    }
  }
  var missing = [];
  for (i = 0; i < CONFIG.HEADERS.length; i++) {
    if (!map[CONFIG.HEADERS[i]]) {
      missing.push(CONFIG.HEADERS[i]);
    }
  }
  if (missing.length) {
    throw new Error('「' + sheet.getName() + '」の1行目に必要な列がありません: ' + missing.join(', '));
  }
  return map;
}

function readRowByHeaders_(sheet, row, headerMap) {
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  var rec = {};
  var keys = Object.keys(headerMap);
  var i;
  for (i = 0; i < keys.length; i++) {
    rec[keys[i]] = values[headerMap[keys[i]] - 1];
  }
  return rec;
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function logError_(where, err) {
  var msg = (err && err.stack) ? String(err.stack) : String(err);
  console.error('[' + where + '] ' + msg);
  Logger.log('[' + where + '] ' + msg);
}
