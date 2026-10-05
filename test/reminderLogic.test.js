'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGasLogic, TIME_ZONE } = require('../test-support/loadGasLogic');

const TODAY = '2026-10-05';

function makeCase(overrides = {}) {
  return {
    caseName: 'コーポレートサイト制作',
    assignee: '山田',
    email: 'demo-yamada@example.com',
    deadline: '2026/10/08',
    status: '対応中',
    lastNotifiedDate: '',
    rowNumber: 2,
    ...overrides,
  };
}

function check(gas, c, sentKeys = {}, todayKey = TODAY) {
  return gas.checkEligibility_(c, todayKey, sentKeys, TIME_ZONE);
}

// ---------------------------------------------------------------------
// 通知タイミングの判定
// ---------------------------------------------------------------------

test('3日前：納期の3日前は「3日前」通知の対象になる', () => {
  const { gas } = loadGasLogic();
  const result = check(gas, makeCase({ deadline: '2026/10/08' }));
  assert.equal(result.eligible, true);
  assert.equal(result.daysLeft, 3);
  assert.equal(result.timingLabel, '3日前');
});

test('前日：納期の1日前は「前日」通知の対象になる', () => {
  const { gas } = loadGasLogic();
  const result = check(gas, makeCase({ deadline: '2026/10/06' }));
  assert.equal(result.eligible, true);
  assert.equal(result.daysLeft, 1);
  assert.equal(result.timingLabel, '前日');
});

test('当日：納期当日は「当日」通知の対象になる', () => {
  const { gas } = loadGasLogic();
  const result = check(gas, makeCase({ deadline: '2026/10/05' }));
  assert.equal(result.eligible, true);
  assert.equal(result.daysLeft, 0);
  assert.equal(result.timingLabel, '当日');
});

test('通知対象外の日：2日前・4日前・納期超過は通知しない', () => {
  const { gas } = loadGasLogic();
  for (const deadline of ['2026/10/07', '2026/10/09', '2026/10/04', '2026/09/01', '2026/12/31']) {
    const result = check(gas, makeCase({ deadline }));
    assert.equal(result.eligible, false, deadline + ' は対象外のはず');
    assert.match(result.reason, /通知タイミングではありません/);
  }
});

test('納期が空欄・不正な値・存在しない日付の場合は通知しない', () => {
  const { gas } = loadGasLogic();
  for (const deadline of ['', 'abc', '2026/13/40', '2026/02/30', null, undefined, 12345]) {
    const result = check(gas, makeCase({ deadline }));
    assert.equal(result.eligible, false);
    assert.match(result.reason, /納期が空欄または日付として不正/);
  }
});

test('納期がスプレッドシートの日付（Dateオブジェクト）でも判定できる', () => {
  const { gas } = loadGasLogic();
  // 日本時間 2026-10-08 00:00（UTCでは前日 15:00）
  const deadline = new Date('2026-10-07T15:00:00Z');
  const result = check(gas, makeCase({ deadline }));
  assert.equal(result.eligible, true);
  assert.equal(result.deadlineKey, '2026-10-08');
  assert.equal(result.timingLabel, '3日前');
});

// ---------------------------------------------------------------------
// 境界値（月末・年末・うるう年）
// ---------------------------------------------------------------------

test('境界値：月末をまたいでも残り日数を正しく数える', () => {
  const { gas } = loadGasLogic();
  const result = check(gas, makeCase({ deadline: '2026/11/01' }), {}, '2026-10-29');
  assert.equal(result.eligible, true);
  assert.equal(result.daysLeft, 3);
});

test('境界値：年末をまたいでも前日を正しく判定する', () => {
  const { gas } = loadGasLogic();
  const result = check(gas, makeCase({ deadline: '2027/01/01' }), {}, '2026-12-31');
  assert.equal(result.eligible, true);
  assert.equal(result.timingLabel, '前日');
});

test('境界値：うるう年の2月29日を含めて日数を数える', () => {
  const { gas } = loadGasLogic();
  assert.equal(gas.daysUntil_('2028-02-28', '2028-03-01'), 2);
  assert.equal(gas.daysUntil_('2028-02-29', '2028-03-01'), 1);
  assert.equal(gas.daysUntil_('2027-02-28', '2027-03-01'), 1);
  const result = check(gas, makeCase({ deadline: '2028/03/03' }), {}, '2028-02-29');
  assert.equal(result.timingLabel, '3日前');
});

// ---------------------------------------------------------------------
// ステータス
// ---------------------------------------------------------------------

test('完了案件を除外する（前後の空白があっても完了扱い）', () => {
  const { gas } = loadGasLogic();
  for (const status of ['完了', ' 完了 ']) {
    const result = check(gas, makeCase({ status }));
    assert.equal(result.eligible, false);
    assert.match(result.reason, /完了扱い/);
  }
});

test('完了以外のステータス（対応中・未対応・未完了・空欄）はすべて通知対象', () => {
  const { gas } = loadGasLogic();
  for (const status of ['対応中', '未対応', '未完了', '', null]) {
    const result = check(gas, makeCase({ status }));
    assert.equal(result.eligible, true, String(status) + ' は通知対象のはず');
  }
});

test('完了扱いのステータスは設定で追加できる', () => {
  const { gas } = loadGasLogic();
  gas.COMPLETED_STATUSES = ['完了', 'キャンセル'];
  assert.equal(check(gas, makeCase({ status: 'キャンセル' })).eligible, false);
  assert.equal(check(gas, makeCase({ status: '対応中' })).eligible, true);
});

// ---------------------------------------------------------------------
// 複数通知・二重送信防止
// ---------------------------------------------------------------------

test('複数通知：同じ案件が3日前・前日・当日のそれぞれで通知対象になる', () => {
  const { gas } = loadGasLogic();
  const c = makeCase({ deadline: '2026/10/08' });
  const labels = ['2026-10-05', '2026-10-07', '2026-10-08'].map((today) => check(gas, c, {}, today).timingLabel);
  assert.deepEqual(labels, ['3日前', '前日', '当日']);
});

test('二重送信防止：同じタイミングの通知が送信済みなら対象外になる', () => {
  const { gas } = loadGasLogic();
  const c = makeCase({ deadline: '2026/10/08' });
  const first = check(gas, c);
  const sentKeys = { [first.key]: true };
  const second = check(gas, c, sentKeys);
  assert.equal(second.eligible, false);
  assert.equal(second.alreadySent, true);
  assert.match(second.reason, /3日前通知は送信済み/);
});

test('送信済みでも別タイミングなら通知できる（3日前送信済み → 前日は送る）', () => {
  const { gas } = loadGasLogic();
  const c = makeCase({ deadline: '2026/10/08' });
  const threeDays = check(gas, c, {}, '2026-10-05');
  const sentKeys = { [threeDays.key]: true };
  const dayBefore = check(gas, c, sentKeys, '2026-10-07');
  assert.equal(dayBefore.eligible, true);
  assert.equal(dayBefore.timingLabel, '前日');
  assert.notEqual(dayBefore.key, threeDays.key);
});

test('納期が変更された場合は新しい納期に対して改めて通知できる', () => {
  const { gas } = loadGasLogic();
  const original = check(gas, makeCase({ deadline: '2026/10/08' }));
  const sentKeys = { [original.key]: true };
  const changed = check(gas, makeCase({ deadline: '2026/10/06' }), sentKeys);
  assert.equal(changed.eligible, true);
  assert.equal(changed.timingLabel, '前日');
});

// ---------------------------------------------------------------------
// 設定（通知タイミング）
// ---------------------------------------------------------------------

test('通知タイミングは REMINDER_DAYS の変更だけで切り替えられる', () => {
  const { gas } = loadGasLogic();
  gas.REMINDER_DAYS = [7];
  assert.equal(check(gas, makeCase({ deadline: '2026/10/12' })).timingLabel, '7日前');
  assert.equal(check(gas, makeCase({ deadline: '2026/10/08' })).eligible, false);
});

// ---------------------------------------------------------------------
// メールテンプレート
// ---------------------------------------------------------------------

test('テンプレートの差し込み：未定義の項目はそのまま残す', () => {
  const { gas } = loadGasLogic();
  const text = gas.renderTemplate_('{案件名}は{納期}まで（{不明}）', { '案件名': 'LP制作', '納期': '2026/10/08' });
  assert.equal(text, 'LP制作は2026/10/08まで（{不明}）');
});

test('件名・本文に案件名・担当者・納期・期限が差し込まれる', () => {
  const { gas } = loadGasLogic();
  const c = makeCase({ deadline: '2026/10/08' });
  const mail = gas.buildReminderEmail_(c, check(gas, c));
  assert.equal(mail.subject, '【納期リマインド】コーポレートサイト制作の納期は3日後（2026/10/08）です');
  assert.match(mail.body, /^山田さん/);
  assert.match(mail.body, /■ 案件名：コーポレートサイト制作/);
  assert.match(mail.body, /■ 納期　：2026\/10\/08（3日後）/);
});

test('期限の表記：当日は「本日」、前日は「明日」になる', () => {
  const { gas } = loadGasLogic();
  const today = makeCase({ deadline: '2026/10/05' });
  const tomorrow = makeCase({ deadline: '2026/10/06' });
  assert.match(gas.buildReminderEmail_(today, check(gas, today)).subject, /納期は本日（2026\/10\/05）/);
  assert.match(gas.buildReminderEmail_(tomorrow, check(gas, tomorrow)).subject, /納期は明日（2026\/10\/06）/);
});

test('件名テンプレートは設定で変更できる（{残り日数}も使える）', () => {
  const { gas } = loadGasLogic();
  gas.EMAIL_SUBJECT_TEMPLATE = '[{担当者}] {案件名} 残り{残り日数}日';
  const c = makeCase({ deadline: '2026/10/08' });
  assert.equal(gas.buildReminderEmail_(c, check(gas, c)).subject, '[山田] コーポレートサイト制作 残り3日');
});

// ---------------------------------------------------------------------
// 送信前の必須項目チェック
// ---------------------------------------------------------------------

test('案件名・担当者・メールアドレスが空欄の場合は送信前にスキップする', () => {
  const { gas } = loadGasLogic();
  assert.match(gas.validateCaseForSending_(makeCase({ caseName: ' ' }), TIME_ZONE).reason, /案件名が空欄/);
  assert.match(gas.validateCaseForSending_(makeCase({ assignee: '' }), TIME_ZONE).reason, /担当者が空欄/);
  assert.match(gas.validateCaseForSending_(makeCase({ email: '' }), TIME_ZONE).reason, /メールアドレスが空欄/);
  assert.equal(gas.validateCaseForSending_(makeCase(), TIME_ZONE).valid, true);
});
