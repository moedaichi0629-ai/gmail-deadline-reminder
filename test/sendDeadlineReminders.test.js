'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGasLogic } = require('../test-support/loadGasLogic');

// 実行日時：日本時間 2026-10-05 09:00
const NOW = '2026-10-05T09:00:00+09:00';

const HEADER = ['案件名', '担当者', 'メールアドレス', '納期', 'ステータス', '最終通知日'];

// サンプルデータはすべて架空の情報（example.com）
function sampleRows() {
  return [
    HEADER,
    ['コーポレートサイト制作', '山田', 'demo-yamada@example.com', '2026/10/08', '対応中', ''], // 3日前
    ['LP制作', '佐藤', 'demo-sato@example.com', '2026/10/06', '未対応', ''], // 前日
    ['バナー制作', '鈴木', 'demo-suzuki@example.com', '2026/10/05', '対応中', ''], // 当日
    ['ECサイト改修', '田中', 'demo-tanaka@example.com', '2026/10/07', '対応中', ''], // 2日前（対象外）
    ['採用サイト制作', '高橋', 'demo-takahashi@example.com', '2026/10/06', '完了', ''], // 完了（対象外）
  ];
}

function historyResults(sheets) {
  return sheets['送信履歴'].rows.slice(1).map((r) => ({
    caseName: r[1],
    to: r[3],
    deadline: r[4],
    timing: r[5],
    result: r[7],
    error: r[8],
  }));
}

test('通知タイミングの案件だけに送信し、送信履歴と最終通知日を記録する', () => {
  const { gas, sheets, sentEmails } = loadGasLogic({ now: NOW, caseRows: sampleRows() });

  const summary = gas.sendDeadlineReminders();

  assert.deepEqual(sentEmails.map((m) => m.to), [
    'demo-yamada@example.com',
    'demo-sato@example.com',
    'demo-suzuki@example.com',
  ]);
  assert.equal(summary.sent, 3);
  assert.equal(summary.excluded, 2);
  assert.equal(summary.failed, 0);

  // 送信履歴シートが自動作成され、ヘッダーと結果が記録される
  const history = sheets['送信履歴'];
  assert.deepEqual(history.rows[0], ['送信日時', '案件名', '担当者', '送信先', '納期', '通知タイミング', '件名', '結果', 'エラー内容', '通知キー']);
  assert.deepEqual(
    historyResults(sheets).map((h) => [h.caseName, h.timing, h.result]),
    [
      ['コーポレートサイト制作', '3日前', '送信成功'],
      ['LP制作', '前日', '送信成功'],
      ['バナー制作', '当日', '送信成功'],
    ]
  );
  assert.ok(history.rows[1][0] instanceof Date || Object.prototype.toString.call(history.rows[1][0]) === '[object Date]');

  // 送信した行だけ最終通知日が更新される
  const caseRows = sheets['案件管理'].rows;
  assert.notEqual(caseRows[1][5], '');
  assert.notEqual(caseRows[2][5], '');
  assert.notEqual(caseRows[3][5], '');
  assert.equal(caseRows[4][5], '');
  assert.equal(caseRows[5][5], '');
});

test('二重送信防止：同じ日に2回実行しても2回目は送信しない', () => {
  const env = loadGasLogic({ now: NOW, caseRows: sampleRows() });
  env.gas.sendDeadlineReminders();
  const second = env.gas.sendDeadlineReminders();

  assert.equal(env.sentEmails.length, 3);
  assert.equal(second.sent, 0);
  assert.equal(second.alreadySent, 3);
  assert.equal(historyResults(env.sheets).length, 3);
});

test('複数通知：同じ案件に3日前・前日・当日の3回だけ送信される', () => {
  const rows = [HEADER, ['コーポレートサイト制作', '山田', 'demo-yamada@example.com', '2026/10/08', '対応中', '']];
  let caseRows = rows;
  let historyRows;
  const sentPerDay = {};

  // 10/5〜10/9 の毎朝9時に実行した場合をシミュレーション（同じ日に2回実行する日も含める）
  for (const day of ['05', '05', '06', '07', '07', '08', '09']) {
    const env = loadGasLogic({ now: '2026-10-' + day + 'T09:00:00+09:00', caseRows, historyRows });
    const summary = env.gas.sendDeadlineReminders();
    sentPerDay[day] = (sentPerDay[day] || 0) + summary.sent;
    caseRows = env.sheets['案件管理'].rows;
    historyRows = env.sheets['送信履歴'].rows;
  }

  assert.deepEqual(sentPerDay, { '05': 1, '06': 0, '07': 1, '08': 1, '09': 0 });
  assert.deepEqual(
    historyRows.slice(1).map((r) => r[5]),
    ['3日前', '前日', '当日']
  );
});

test('送信済みでも別タイミングなら通知できる（既存の送信履歴を引き継いで判定する）', () => {
  // 前回実行（10/5）で3日前通知を送信済みの状態から、10/7（前日）に実行する
  const first = loadGasLogic({ now: NOW, caseRows: [HEADER, ['コーポレートサイト制作', '山田', 'demo-yamada@example.com', '2026/10/08', '対応中', '']] });
  first.gas.sendDeadlineReminders();

  const env = loadGasLogic({
    now: '2026-10-07T09:00:00+09:00',
    caseRows: first.sheets['案件管理'].rows,
    historyRows: first.sheets['送信履歴'].rows,
  });
  const summary = env.gas.sendDeadlineReminders();

  assert.equal(summary.sent, 1);
  assert.match(env.sentEmails[0].subject, /納期は明日/);
});

test('送信失敗：失敗を送信履歴に記録し、他の案件の送信は続ける', () => {
  const { gas, sheets, sentEmails } = loadGasLogic({
    now: NOW,
    caseRows: sampleRows(),
    failEmails: ['demo-sato@example.com'],
  });

  const summary = gas.sendDeadlineReminders();

  assert.equal(summary.sent, 2);
  assert.equal(summary.failed, 1);
  assert.deepEqual(sentEmails.map((m) => m.to), ['demo-yamada@example.com', 'demo-suzuki@example.com']);

  const failed = historyResults(sheets).find((h) => h.caseName === 'LP制作');
  assert.equal(failed.result, '送信失敗');
  assert.match(failed.error, /送信テスト用のエラー/);
  // 失敗した行の最終通知日は更新しない
  assert.equal(sheets['案件管理'].rows[2][5], '');
});

test('送信失敗した通知は、同じ日の再実行で再送できる', () => {
  const first = loadGasLogic({ now: NOW, caseRows: sampleRows(), failEmails: ['demo-sato@example.com'] });
  first.gas.sendDeadlineReminders();

  const retry = loadGasLogic({
    now: '2026-10-05T10:00:00+09:00',
    caseRows: first.sheets['案件管理'].rows,
    historyRows: first.sheets['送信履歴'].rows,
  });
  const summary = retry.gas.sendDeadlineReminders();

  assert.equal(summary.sent, 1);
  assert.equal(retry.sentEmails[0].to, 'demo-sato@example.com');
});

test('同じ案件の行が重複していても、同じ実行内で二重送信しない', () => {
  const row = ['コーポレートサイト制作', '山田', 'demo-yamada@example.com', '2026/10/08', '対応中', ''];
  const { gas, sentEmails } = loadGasLogic({ now: NOW, caseRows: [HEADER, row, row.slice()] });

  const summary = gas.sendDeadlineReminders();

  assert.equal(sentEmails.length, 1);
  assert.equal(summary.alreadySent, 1);
});

test('メールアドレスが空欄の案件はスキップし、送信履歴には残さない', () => {
  const { gas, sheets, sentEmails } = loadGasLogic({
    now: NOW,
    caseRows: [HEADER, ['コーポレートサイト制作', '山田', '', '2026/10/08', '対応中', '']],
  });

  const summary = gas.sendDeadlineReminders();

  assert.equal(summary.skipped, 1);
  assert.equal(sentEmails.length, 0);
  assert.equal(sheets['送信履歴'].rows.length, 1); // ヘッダーのみ
});

test('二重実行防止：ロックを取得できない場合は何も送信しない', () => {
  const { gas, sentEmails, sheets } = loadGasLogic({ now: NOW, caseRows: sampleRows(), lockAvailable: false });

  const summary = gas.sendDeadlineReminders();

  assert.equal(summary.sent, 0);
  assert.equal(summary.total, 0);
  assert.equal(sentEmails.length, 0);
  assert.equal(sheets['送信履歴'], undefined);
});

test('必須ヘッダーが不足している場合はエラーにし、ロックは必ず解放する', () => {
  const { gas, lockState } = loadGasLogic({
    now: NOW,
    caseRows: [['案件名', '担当者', '納期', 'ステータス'], ['LP制作', '佐藤', '2026/10/06', '対応中']],
  });

  assert.throws(() => gas.sendDeadlineReminders(), /必須ヘッダーが不足しています: メールアドレス, 最終通知日/);
  assert.equal(lockState.acquired, 1);
  assert.equal(lockState.released, 1);
});

test('案件データが0件でもエラーにならない', () => {
  const { gas, sentEmails } = loadGasLogic({ now: NOW, caseRows: [HEADER] });
  const summary = gas.sendDeadlineReminders();
  assert.equal(summary.total, 0);
  assert.equal(sentEmails.length, 0);
});
