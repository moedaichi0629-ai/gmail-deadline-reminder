/**
 * Google Sheets × Gmail 納期リマインド自動送信ツール
 *
 * 「案件管理」シートの納期を毎日確認し、設定したタイミング
 * （例：3日前・前日・当日）に担当者へGmailでリマインドメールを送信する。
 * 送信結果は「送信履歴」シートに記録し、同じ案件・同じタイミングの
 * 二重送信を防止する。
 *
 * このスクリプトは対象スプレッドシートに紐づくコンテナバインド型
 * Apps Scriptとして利用する前提。
 */

/* =====================================================================
 * 設定（運用に合わせてここだけ変更する）
 * ===================================================================== */

// 通知タイミング（納期の何日前に通知するか）。0 は納期当日。
// 例：[3, 1, 0] → 3日前・前日・当日の3回通知する
var REMINDER_DAYS = [3, 1, 0];

// 「完了」として扱うステータス。ここに含まれないステータスはすべて通知対象。
var COMPLETED_STATUSES = ['完了'];

// メール件名・本文のテンプレート。
// 使用できる差し込み項目：
//   {案件名} {担当者} {納期} {残り日数} {期限}
//   {期限} は「本日」「明日」「3日後」のような表記になる
var EMAIL_SUBJECT_TEMPLATE = '【納期リマインド】{案件名}の納期は{期限}（{納期}）です';
var EMAIL_BODY_TEMPLATE = [
  '{担当者}さん',
  '',
  'お疲れさまです。納期が近づいている案件のお知らせです。',
  '',
  '■ 案件名：{案件名}',
  '■ 納期　：{納期}（{期限}）',
  '',
  '進捗をご確認のうえ、ご対応をお願いいたします。',
  '',
  '※このメールは「案件管理」シートの情報をもとに自動送信しています。'
].join('\n');

// 納期の表示形式（メール本文・送信履歴に使用）
var DEADLINE_DISPLAY_FORMAT = 'yyyy/MM/dd';

// シート名
var SHEET_NAME = '案件管理';
var HISTORY_SHEET_NAME = '送信履歴';

/* =====================================================================
 * 案件管理シートの読み込み
 * ===================================================================== */

// 必須ヘッダー一覧
var REQUIRED_HEADERS = ['案件名', '担当者', 'メールアドレス', '納期', 'ステータス', '最終通知日'];

/**
 * 「案件管理」シートを取得する。
 * シートが存在しない場合はエラーとする。
 */
function getCaseSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('シート「' + SHEET_NAME + '」が見つかりません。');
  }
  return sheet;
}

/**
 * ヘッダー行（1行目）を取得し、ヘッダー名 → 列インデックス(0始まり)の
 * マップを作成する。必須ヘッダーが不足している場合はエラーとする。
 */
function getHeaderMap_(sheet) {
  var lastColumn = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  var headerMap = {};
  for (var i = 0; i < headerRow.length; i++) {
    var headerName = String(headerRow[i]).trim();
    if (headerName !== '') {
      headerMap[headerName] = i;
    }
  }

  var missingHeaders = [];
  for (var j = 0; j < REQUIRED_HEADERS.length; j++) {
    if (!Object.prototype.hasOwnProperty.call(headerMap, REQUIRED_HEADERS[j])) {
      missingHeaders.push(REQUIRED_HEADERS[j]);
    }
  }

  if (missingHeaders.length > 0) {
    throw new Error('必須ヘッダーが不足しています: ' + missingHeaders.join(', '));
  }

  return headerMap;
}

/**
 * 「案件管理」シートから2行目以降の案件データを取得し、
 * 後続処理で扱いやすい形（オブジェクトの配列）に変換する。
 * データが0件の場合は空配列を返す。
 */
function getCaseList() {
  var sheet = getCaseSheet_();
  var headerMap = getHeaderMap_(sheet);

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  // ヘッダーのみ、またはデータ0件の場合
  if (lastRow < 2) {
    return [];
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  var dataRows = dataRange.getValues();

  var caseList = [];
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];
    caseList.push({
      caseName: row[headerMap['案件名']],
      assignee: row[headerMap['担当者']],
      email: row[headerMap['メールアドレス']],
      deadline: row[headerMap['納期']],
      status: row[headerMap['ステータス']],
      lastNotifiedDate: row[headerMap['最終通知日']],
      rowNumber: r + 2 // 元のシート行番号（1行目はヘッダーのため+2）
    });
  }

  return caseList;
}

/* =====================================================================
 * 日付の扱い
 * ===================================================================== */

/**
 * 値がDateオブジェクトかどうかを判定する。
 *
 * Apps Scriptの実行環境によっては、スプレッドシートの日付セルから
 * 取得した値が実際にはDateであっても「value instanceof Date」が
 * falseになるケースがある（実機デバッグで確認済み）。
 * そのため、より確実な判定方法として
 * Object.prototype.toString.call(value) === '[object Date]' を用いる。
 */
function isDateObject_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

/**
 * 案件データの日付項目（納期など）を「年月日」のみを表す
 * 文字列（yyyy-MM-dd）に変換する。時刻部分の違いを無視して
 * 日付だけで比較できるようにするための関数。
 *
 * 対応する値の形式：
 * 1. Dateオブジェクト（スプレッドシートの日付セルから取得される通常の形）
 *    → スプレッドシートのタイムゾーンで年月日を解釈する
 * 2. "yyyy/MM/dd" 形式の文字列（例: "2026/09/20"）
 * 3. "yyyy-MM-dd" 形式の文字列（例: "2026-09-20"）
 *    → 文字列は年月日がそのまま記載されているとみなし、
 *      タイムゾーン変換は行わない（曖昧さがないため）
 *
 * new Date(文字列) のような環境依存・曖昧な自動パースは行わず、
 * 上記の形式にのみ明示的にマッチさせる。
 * 存在しない日付（例: "2026/13/40"）や空欄・不正な文字列（例: "abc"）は
 * すべてnullを返す。
 */
function formatDateKey_(value, timeZone) {
  if (isDateObject_(value)) {
    if (isNaN(value.getTime())) {
      return null;
    }
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }

  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (trimmed === '') {
      return null;
    }

    var m = trimmed.match(/^(\d{4})\/(\d{2})\/(\d{2})$/) ||
      trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) {
      return null;
    }

    return buildDateKeyIfValid_(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
  }

  return null;
}

/**
 * 年・月・日（数値）が実在するカレンダー上の日付かどうかを検証し、
 * 有効な場合のみ "yyyy-MM-dd" 形式の文字列を返す。
 * 無効な場合（例: 2026年13月40日）はnullを返す。
 *
 * UTC基準のDateで一度組み立て、年月日が入力どおりに保持されているかを
 * 確認することで、存在しない日付を確実に検出する
 * （JavaScriptのDateは繰り上げ計算をしてしまうため、単純に new Date() の
 * 結果を信用しない）。
 */
function buildDateKeyIfValid_(year, month, day) {
  var utcDate = new Date(Date.UTC(year, month - 1, day));
  if (
    utcDate.getUTCFullYear() !== year ||
    utcDate.getUTCMonth() !== month - 1 ||
    utcDate.getUTCDate() !== day
  ) {
    return null;
  }
  return year + '-' + pad2_(month) + '-' + pad2_(day);
}

/**
 * 数値を2桁の文字列に変換する（例: 3 → "03"）。
 */
function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

/**
 * 「yyyy-MM-dd」形式の日付キーをUTC基準のミリ秒に変換する。
 */
function dateKeyToUtcMs_(dateKey) {
  var parts = dateKey.split('-');
  return Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

/**
 * 実行日（todayKey）から納期（deadlineKey）までの残り日数を返す。
 * 納期が今日なら0、明日なら1、過ぎていれば負の値になる。
 *
 * 年月日をUTC基準のDateへ変換して差を取るため（UTCにはDSTが存在しない）、
 * 月末・年末・うるう年をまたぐ場合でもカレンダー上の日数として正しく求められる。
 */
function daysUntil_(todayKey, deadlineKey) {
  return Math.round((dateKeyToUtcMs_(deadlineKey) - dateKeyToUtcMs_(todayKey)) / (24 * 60 * 60 * 1000));
}

/**
 * 「yyyy-MM-dd」形式の日付キーを表示用（yyyy/MM/dd）に変換する。
 */
function formatDateKeyForDisplay_(dateKey) {
  return dateKey.replace(/-/g, '/');
}

/* =====================================================================
 * 通知判定
 * ===================================================================== */

/**
 * ステータスが「完了」扱い（COMPLETED_STATUSES に含まれる）かどうかを判定する。
 * 前後の空白は無視して比較する。
 */
function isCompletedStatus_(status) {
  var normalized = String(status === null || status === undefined ? '' : status).trim();
  return COMPLETED_STATUSES.indexOf(normalized) !== -1;
}

/**
 * 残り日数を通知タイミングの表示名に変換する（例: 3 → "3日前"、1 → "前日"、0 → "当日"）。
 */
function getTimingLabel_(daysLeft) {
  if (daysLeft === 0) {
    return '当日';
  }
  if (daysLeft === 1) {
    return '前日';
  }
  return daysLeft + '日前';
}

/**
 * 残り日数をメール用の期限表記に変換する（例: 0 → "本日"、1 → "明日"、3 → "3日後"）。
 */
function getDueLabel_(daysLeft) {
  if (daysLeft === 0) {
    return '本日';
  }
  if (daysLeft === 1) {
    return '明日';
  }
  return daysLeft + '日後';
}

/**
 * 同じ案件・同じ納期・同じ通知タイミングを一意に表す「通知キー」を作る。
 * 送信履歴シートにこのキーが「送信成功」で記録されていれば、同じ通知は二度と送らない。
 *
 * 行番号は行の並べ替え・挿入でずれるため使用せず、
 * 案件名・送信先・納期・タイミングの組み合わせで判定する。
 * 納期が変更された場合は別のキーになるため、新しい納期に対して改めて通知される。
 */
function buildNotificationKey_(c, deadlineKey, daysLeft) {
  return [
    String(c.caseName).trim(),
    String(c.email).trim(),
    deadlineKey,
    daysLeft
  ].join('|');
}

/**
 * 対象案件が今日の通知条件をすべて満たすか判定する。
 *
 * 条件：
 * 1. 納期が日付として正しい
 * 2. ステータスが「完了」扱いではない
 * 3. 納期までの残り日数が REMINDER_DAYS のいずれかと一致する
 * 4. 同じ通知キー（案件・納期・タイミング）がまだ送信成功していない
 *
 * 時刻部分は無視し、年月日単位で比較する。
 * 判定に通った場合は、メール作成・履歴記録に必要な情報も返す。
 */
function checkEligibility_(c, todayKey, sentKeys, timeZone) {
  var deadlineKey = formatDateKey_(c.deadline, timeZone);
  if (!deadlineKey) {
    return { eligible: false, reason: '納期が空欄または日付として不正です' };
  }

  if (isCompletedStatus_(c.status)) {
    return { eligible: false, reason: 'ステータスが完了扱いです（ステータス=' + c.status + '）' };
  }

  var daysLeft = daysUntil_(todayKey, deadlineKey);
  if (REMINDER_DAYS.indexOf(daysLeft) === -1) {
    return { eligible: false, reason: '本日は通知タイミングではありません（納期=' + deadlineKey + '、残り' + daysLeft + '日）' };
  }

  var key = buildNotificationKey_(c, deadlineKey, daysLeft);
  if (sentKeys[key]) {
    return { eligible: false, alreadySent: true, reason: getTimingLabel_(daysLeft) + '通知は送信済みです' };
  }

  return {
    eligible: true,
    reason: '',
    deadlineKey: deadlineKey,
    daysLeft: daysLeft,
    timingLabel: getTimingLabel_(daysLeft),
    key: key
  };
}

/**
 * 通知対象となった案件について、安全にメール送信できるか
 * 最低限の必須項目（案件名・担当者・メールアドレス・納期）を確認する。
 * 不足がある場合はスキップ対象として理由を返す。
 */
function validateCaseForSending_(c, timeZone) {
  if (!c.caseName || String(c.caseName).trim() === '') {
    return { valid: false, reason: '案件名が空欄です' };
  }
  if (!c.assignee || String(c.assignee).trim() === '') {
    return { valid: false, reason: '担当者が空欄です' };
  }
  if (!c.email || String(c.email).trim() === '') {
    return { valid: false, reason: 'メールアドレスが空欄です' };
  }
  if (!formatDateKey_(c.deadline, timeZone)) {
    return { valid: false, reason: '納期が空欄または不正です' };
  }
  return { valid: true, reason: '' };
}

/* =====================================================================
 * メール
 * ===================================================================== */

/**
 * テンプレート中の {項目名} を values の値で置き換える。
 * values に存在しない項目名はそのまま残す（設定ミスに気付けるようにするため）。
 */
function renderTemplate_(template, values) {
  return template.replace(/\{([^{}]+)\}/g, function (match, name) {
    return Object.prototype.hasOwnProperty.call(values, name) ? String(values[name]) : match;
  });
}

/**
 * 案件と判定結果から、メールの件名・本文を組み立てる。
 */
function buildReminderEmail_(c, eligibility) {
  var values = {
    '案件名': String(c.caseName).trim(),
    '担当者': String(c.assignee).trim(),
    '納期': formatDateKeyForDisplay_(eligibility.deadlineKey),
    '残り日数': eligibility.daysLeft,
    '期限': getDueLabel_(eligibility.daysLeft)
  };
  return {
    subject: renderTemplate_(EMAIL_SUBJECT_TEMPLATE, values),
    body: renderTemplate_(EMAIL_BODY_TEMPLATE, values)
  };
}

/**
 * 指定案件へリマインドメールを送信する。
 */
function sendReminderEmail_(c, mail) {
  GmailApp.sendEmail(String(c.email).trim(), mail.subject, mail.body);
}

/* =====================================================================
 * 送信履歴
 * ===================================================================== */

// 送信履歴シートの列構成（最後の「通知キー」は二重送信防止の判定に使用する）
var HISTORY_HEADERS = ['送信日時', '案件名', '担当者', '送信先', '納期', '通知タイミング', '件名', '結果', 'エラー内容', '通知キー'];
var HISTORY_RESULT_SUCCESS = '送信成功';
var HISTORY_RESULT_FAILURE = '送信失敗';

/**
 * 「送信履歴」シートを取得する。存在しない場合はヘッダー付きで作成する。
 */
function getOrCreateHistorySheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(HISTORY_SHEET_NAME);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 送信履歴シートから「送信成功」の通知キーを読み込み、
 * { 通知キー: true } 形式のオブジェクトで返す。
 */
function loadSentKeys_(historySheet) {
  var sentKeys = {};
  var lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    return sentKeys;
  }

  var resultIndex = HISTORY_HEADERS.indexOf('結果');
  var keyIndex = HISTORY_HEADERS.indexOf('通知キー');
  var rows = historySheet.getRange(2, 1, lastRow - 1, HISTORY_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][resultIndex] === HISTORY_RESULT_SUCCESS && rows[i][keyIndex]) {
      sentKeys[String(rows[i][keyIndex])] = true;
    }
  }
  return sentKeys;
}

/**
 * 送信履歴シートへ1行追記する。
 */
function appendHistory_(historySheet, sentAt, c, eligibility, subject, result, errorMessage) {
  historySheet.appendRow([
    sentAt,
    String(c.caseName).trim(),
    String(c.assignee).trim(),
    String(c.email).trim(),
    formatDateKeyForDisplay_(eligibility.deadlineKey),
    eligibility.timingLabel,
    subject,
    result,
    errorMessage || '',
    eligibility.key
  ]);
}

/* =====================================================================
 * メイン処理
 * ===================================================================== */

// LockServiceでロック取得を待機する最大時間（ミリ秒）
var LOCK_WAIT_MS = 30 * 1000;

/**
 * メイン処理（毎日の時間主導トリガーから実行する）。
 *
 * 「案件管理」シートの各案件について、
 * 1. 通知条件（納期が正しい／完了扱いでない／今日が通知タイミング／未送信）を判定
 * 2. 満たす場合、必須項目の不足がないか確認（不足時はスキップ）
 * 3. 問題なければGmailでリマインドメールを送信
 * 4. 送信結果（成功・失敗）を「送信履歴」シートへ記録
 * 5. 送信に成功した場合は「最終通知日」を本日の日付に更新
 *
 * 1件のエラー・不備が発生しても、他の案件の処理は継続する。
 *
 * LockService（スクリプトロック）により、複数の実行が
 * ほぼ同時に走った場合の二重送信を防止する。ロックを取得できない場合は
 * 案件取得・通知判定・送信処理を一切開始せずに終了する。
 * ロックは正常終了・異常終了のいずれでも必ず解放する（try/finally）。
 *
 * 戻り値として処理件数のサマリーを返す。
 */
function sendDeadlineReminders() {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(LOCK_WAIT_MS);
  } catch (lockErr) {
    lockAcquired = false;
  }

  if (!lockAcquired) {
    Logger.log('[ロック取得失敗] 他の処理が実行中のため、今回はリマインド送信を開始せずに終了しました。');
    return {
      total: 0,
      targeted: 0,
      sent: 0,
      failed: 0,
      skipped: 0,
      alreadySent: 0,
      excluded: 0
    };
  }

  try {
    var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var now = new Date();
    var todayKey = formatDateKey_(now, timeZone);

    var sheet = getCaseSheet_();
    var headerMap = getHeaderMap_(sheet);
    var lastNotifiedColumn = headerMap['最終通知日'] + 1; // 0始まり→1始まりの列番号へ変換

    var caseList = getCaseList();
    var historySheet = getOrCreateHistorySheet_();
    var sentKeys = loadSentKeys_(historySheet);

    var summary = {
      total: caseList.length,
      targeted: 0,
      sent: 0,
      failed: 0,
      skipped: 0,
      alreadySent: 0,
      excluded: 0
    };

    Logger.log('=== リマインド処理開始（本日=' + todayKey + ', 通知タイミング=' + REMINDER_DAYS.join(',') + '日前） ===');
    Logger.log('処理対象案件数: ' + caseList.length);

    for (var i = 0; i < caseList.length; i++) {
      var c = caseList[i];

      try {
        var eligibility = checkEligibility_(c, todayKey, sentKeys, timeZone);
        if (!eligibility.eligible) {
          if (eligibility.alreadySent) {
            summary.alreadySent++;
          } else {
            summary.excluded++;
          }
          Logger.log('[対象外] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' 理由: ' + eligibility.reason);
          continue;
        }

        summary.targeted++;
        Logger.log('[通知対象] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' タイミング=' + eligibility.timingLabel);

        var validation = validateCaseForSending_(c, timeZone);
        if (!validation.valid) {
          summary.skipped++;
          Logger.log('[スキップ] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' 理由: ' + validation.reason);
          continue;
        }

        var mail = buildReminderEmail_(c, eligibility);

        try {
          sendReminderEmail_(c, mail);
        } catch (sendErr) {
          summary.failed++;
          appendHistory_(historySheet, new Date(), c, eligibility, mail.subject, HISTORY_RESULT_FAILURE, sendErr.message);
          Logger.log('[送信失敗] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' 原因: ' + sendErr.message);
          continue;
        }

        // 送信成功を記録する（この記録が二重送信防止の判定に使われる）
        sentKeys[eligibility.key] = true;
        appendHistory_(historySheet, new Date(), c, eligibility, mail.subject, HISTORY_RESULT_SUCCESS, '');
        sheet.getRange(c.rowNumber, lastNotifiedColumn).setValue(now);

        summary.sent++;
        Logger.log('[送信成功] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' タイミング=' + eligibility.timingLabel + ' 宛先=' + c.email);

      } catch (err) {
        summary.failed++;
        Logger.log('[処理失敗] 行' + c.rowNumber + ' 案件名=' + c.caseName + ' 原因: ' + err.message);
      }
    }

    Logger.log('=== リマインド処理終了 ===');
    Logger.log(
      '処理件数=' + summary.total +
      ', 通知対象=' + summary.targeted +
      ', 送信成功=' + summary.sent +
      ', 送信失敗=' + summary.failed +
      ', スキップ=' + summary.skipped +
      ', 送信済み=' + summary.alreadySent +
      ', 対象外=' + summary.excluded
    );

    return summary;

  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 * 時間主導トリガー
 * ===================================================================== */

// sendDeadlineReminders を実行対象とする時間主導トリガーの関数名
var REMINDER_TRIGGER_FUNCTION = 'sendDeadlineReminders';

// トリガーの基準時（時単位）。
// Apps Scriptの時間主導トリガーは指定した時刻ちょうどの実行を
// 保証しないため、あくまで「9時頃」の基準としてのみ使用する。
var REMINDER_TRIGGER_HOUR = 9;

/**
 * 指定したハンドラー関数名を対象とする時間主導トリガーの件数を数える。
 */
function countTriggersByHandler_(handlerFunctionName) {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerFunctionName) {
      count++;
    }
  }
  return count;
}

/**
 * sendDeadlineReminders() を毎日9時頃に実行する時間主導トリガーを作成する。
 *
 * 「9時頃」とは、Apps Scriptの時間主導トリガーの仕様上、
 * 指定した時刻ちょうどの実行が保証されないことを踏まえた表現であり、
 * 9:00ちょうどに実行されることを意味しない。
 *
 * 曜日を限定せず毎日実行される設定のため、土日・祝日も通常どおり
 * 実行される。
 *
 * すでに sendDeadlineReminders 用の時間主導トリガーが存在する場合は、
 * 重複登録を避けるため新規作成を行わない。
 */
function setupDailyTrigger() {
  var existingCount = countTriggersByHandler_(REMINDER_TRIGGER_FUNCTION);

  if (existingCount > 0) {
    Logger.log(
      '[トリガー作成スキップ] 「' + REMINDER_TRIGGER_FUNCTION + '」用の時間主導トリガーは' +
      'すでに' + existingCount + '件登録されているため、新規作成しませんでした。'
    );
    return;
  }

  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  ScriptApp.newTrigger(REMINDER_TRIGGER_FUNCTION)
    .timeBased()
    .inTimezone(timeZone)
    .everyDays(1)
    .atHour(REMINDER_TRIGGER_HOUR)
    .create();

  Logger.log(
    '[トリガー作成] 「' + REMINDER_TRIGGER_FUNCTION + '」を毎日' + REMINDER_TRIGGER_HOUR +
    '時頃に実行する時間主導トリガーを作成しました' +
    '（' + REMINDER_TRIGGER_HOUR + ':00ちょうどの実行は保証されません）。'
  );
}

/**
 * sendDeadlineReminders() を対象とする時間主導トリガーのみを削除する。
 * ハンドラー関数名が一致するものだけを削除するため、
 * 本案件以外の他の関数に紐づくトリガーは削除しない。
 */
function removeDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removedCount = 0;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === REMINDER_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
      removedCount++;
    }
  }

  Logger.log('[トリガー削除] 「' + REMINDER_TRIGGER_FUNCTION + '」用の時間主導トリガーを' + removedCount + '件削除しました。');
}
