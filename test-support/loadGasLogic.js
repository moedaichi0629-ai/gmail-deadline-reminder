'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const GAS_FILE_PATH = path.join(__dirname, '..', 'Code.gs');
const TIME_ZONE = 'Asia/Tokyo';

// Code.gs はGoogle Apps Script用のファイル（module.exports等は持たない）。
// このファイル自体は一切変更せず、Node上の仮想環境（vm）でそのまま読み込み、
// SpreadsheetApp・GmailApp・LockService などのGAS固有オブジェクトを
// テスト用の簡易な偽物（フェイク）に差し替えて実行する。
// これにより、本番コードと全く同じ実装をテストできる。

/**
 * Utilities.formatDate の簡易版。
 * yyyy / MM / dd / HH / mm / ss のみ対応する（Code.gs で使用する範囲）。
 */
function formatDate(date, timeZone, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  })
    .formatToParts(date)
    .forEach((p) => {
      parts[p.type] = p.value;
    });
  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('HH', parts.hour)
    .replace('mm', parts.minute)
    .replace('ss', parts.second);
}

/**
 * 2次元配列で値を保持するシートのフェイク。
 */
class FakeSheet {
  constructor(name, rows) {
    this.name = name;
    this.rows = rows ? rows.map((r) => r.slice()) : [];
  }

  getLastRow() {
    return this.rows.length;
  }

  getLastColumn() {
    return this.rows.reduce((max, r) => Math.max(max, r.length), 0);
  }

  getRange(row, column, numRows = 1, numColumns = 1) {
    const sheet = this;
    return {
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const src = sheet.rows[row - 1 + r] || [];
          const line = [];
          for (let c = 0; c < numColumns; c++) {
            const v = src[column - 1 + c];
            line.push(v === undefined ? '' : v);
          }
          out.push(line);
        }
        return out;
      },
      setValues(values) {
        values.forEach((line, r) => {
          line.forEach((v, c) => sheet.setCell(row + r, column + c, v));
        });
        return this;
      },
      setValue(value) {
        sheet.setCell(row, column, value);
        return this;
      },
    };
  }

  setCell(row, column, value) {
    while (this.rows.length < row) {
      this.rows.push([]);
    }
    const line = this.rows[row - 1];
    while (line.length < column) {
      line.push('');
    }
    line[column - 1] = value;
  }

  appendRow(values) {
    this.rows.push(values.slice());
    return this;
  }

  setFrozenRows() {
    return this;
  }
}

/**
 * GAS環境のフェイクを作り、Code.gs を読み込んだコンテキストを返す。
 *
 * options:
 *   now           : 実行日時（「今日」を固定するため。new Date() の引数なし呼び出しがこの日時になる）
 *   caseRows      : 「案件管理」シートの内容（1行目はヘッダー）。省略時はシートなし
 *   historyRows   : 「送信履歴」シートの内容（省略時はシートなし＝自動作成される）
 *   failEmails    : 送信に失敗させるメールアドレスの配列
 *   lockAvailable : false にするとロック取得に失敗する
 */
function loadGasLogic(options = {}) {
  const fixedNow = options.now ? new Date(options.now).getTime() : Date.now();

  class FakeDate extends Date {
    constructor(...args) {
      if (args.length === 0) {
        super(fixedNow);
      } else {
        super(...args);
      }
    }

    static now() {
      return fixedNow;
    }
  }

  const sheets = {};
  if (options.caseRows) {
    sheets['案件管理'] = new FakeSheet('案件管理', options.caseRows);
  }
  if (options.historyRows) {
    sheets['送信履歴'] = new FakeSheet('送信履歴', options.historyRows);
  }

  const spreadsheet = {
    getSheetByName: (name) => sheets[name] || null,
    insertSheet: (name) => {
      sheets[name] = new FakeSheet(name);
      return sheets[name];
    },
    getSpreadsheetTimeZone: () => TIME_ZONE,
  };

  const sentEmails = [];
  const failEmails = options.failEmails || [];
  const logs = [];
  const lockState = { acquired: 0, released: 0 };

  const sandbox = {
    Date: FakeDate,
    Math,
    String,
    Object,
    Error,
    isNaN,
    parseInt,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
    },
    GmailApp: {
      sendEmail(to, subject, body) {
        if (failEmails.indexOf(to) !== -1) {
          throw new Error('送信テスト用のエラー');
        }
        sentEmails.push({ to, subject, body });
      },
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => {
          if (options.lockAvailable === false) {
            return false;
          }
          lockState.acquired++;
          return true;
        },
        releaseLock: () => {
          lockState.released++;
        },
      }),
    },
    Logger: {
      log: (message) => logs.push(String(message)),
    },
    Utilities: {
      formatDate,
    },
  };

  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(GAS_FILE_PATH, 'utf8'), context, { filename: 'Code.gs' });

  return { gas: context, sheets, sentEmails, logs, lockState, TIME_ZONE, FakeDate };
}

module.exports = { loadGasLogic, FakeSheet, formatDate, TIME_ZONE };
