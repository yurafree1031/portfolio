// 請求書のデモの手元テスト用の共通部品。GAS の部品（スプレッドシート・ドライブ・Gmail など）を、メモリ上の偽物に差し替えて動かす
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DEMO_DIR = path.join(__dirname, '..');
const SOURCES = ['Errors.gs', 'Sheet.gs', 'Alert.gs', 'Button.gs', 'Code.gs', 'Invoice.gs', 'Payment.gs', 'Monthly.gs'];
// class・const はファイルの外から見えないので、最後に必要なものを globalThis に出す
const EXPORTS = ['UserError', 'AdminError', 'CONFIG', 'PROP', 'SHEET', 'TABLES', 'STATUS', 'RESULT', 'SETTINGS',
  'ERROR_ROW_COLOR', 'WARN_ROW_COLOR', 'INPUT_COLOR', 'DONE_COLOR', 'RUN_CELLS'];

const { FakeSpreadsheet, FakeSheet, formatDate, gasContext, newEnv } = require('../../../lib/gas/tests/fake-gas');

// ---- このデモで使う、共通の偽物にない部品 ----
FakeSheet.prototype.setRowHeight = function () {};
FakeSheet.prototype.getImages = function () { return this.images || []; };
FakeSheet.prototype.insertImage = function (blob, col, row) {
  const img = { blob, col, row, fn: null };
  const api = {
    setWidth: () => api, setHeight: () => api,
    assignScript: (fn) => { img.fn = fn; return api; },
    getAnchorCell: () => ({ getRow: () => row }),
  };
  (this.images = this.images || []).push(api);
  this.buttons = (this.buttons || []).concat([img]);
  return api;
};
FakeSheet.prototype.getActiveRange = function () { return this.getRange(this.ss.activeRow || 1, 1); };

/** 偽のファイル（中身は文字で持つ。Shift_JIS のCSVは sjis: true で、UTF-8 で読むと文字化けする） */
let nextFileId = 1;
function fakeBlob(data, type, name, sjis = false) {
  const b = {
    data, type, name, sjis,
    getAs: (t) => fakeBlob(`PDF:${data}`, t, name),
    setName: (n) => { b.name = n; return b; },
    getName: () => b.name,
    getContentType: () => b.type,
    setContentType: (t) => { b.type = t; return b; },
    setDataFromString: (s, charset) => { b.data = s; b.sjis = charset === 'Shift_JIS'; return b; },
    getDataAsString: (charset) => (b.sjis && charset !== 'Shift_JIS' ? b.data.replace(/[^\x00-\x7f]/g, '�') : b.data),
  };
  return b;
}
function fakeFile(blob, folder) {
  const id = `FILE${nextFileId++}`;
  const f = {
    id, blob, folder, name: blob.name, updated: Date.now(),
    getId: () => id,
    getName: () => f.name,
    setName: (n) => { if (f.renameError) throw new Error('no permission'); f.name = n; return f; },
    getUrl: () => `https://drive.google.com/file/d/${id}/view`,
    getBlob: () => f.blob,
    getLastUpdated: () => new Date(f.updated),
  };
  return f;
}
function fakeFolder(id) {
  const folder = {
    id, files: [], createError: null,
    getId: () => id,
    getName: () => `folder-${id}`,
    getUrl: () => `https://drive.google.com/drive/folders/${id}`,
    createFile: (blob) => {
      if (folder.createError) throw folder.createError;
      const f = fakeFile(blob, folder);
      folder.files.push(f);
      return f;
    },
    getFiles: () => {
      let i = 0;
      const list = folder.files.slice();
      return { hasNext: () => i < list.length, next: () => list[i++] };
    },
  };
  return folder;
}

/**
 * コードを偽の GAS 部品と一緒に読み込む
 * 戻り値: { ctx（読み込んだ関数）, ss, env（メール・下書き・トリガー・プロパティなどの記録）, folders: { pdf, csv } }
 */
function load({ me = 'keiri@example.com' } = {}) {
  const ss = new FakeSpreadsheet({ firstSheet: 'シート1' });
  const env = newEnv(me);
  env.drafts = [];
  env.deletedDrafts = [];
  env.promptAnswers = [];
  env.holidays = [];
  const folders = { pdf: fakeFolder('PDFFOLDER0001'), csv: fakeFolder('CSVFOLDER0001') };
  const allFolders = () => Object.values(folders);
  const ctx = gasContext(ss, env);
  const ui = ctx.SpreadsheetApp.getUi();
  ui.ButtonSet.OK_CANCEL = 'OK_CANCEL';
  ui.Button.OK = 'OK';
  ui.prompt = (title, message) => {
    env.alerts.push({ title, message, prompt: true });
    const a = env.promptAnswers.length ? env.promptAnswers.shift() : { button: 'OK', text: '金額の誤り' };
    return { getSelectedButton: () => a.button, getResponseText: () => a.text };
  };
  const props = ctx.PropertiesService.getScriptProperties();
  props.getKeys = () => Object.keys(env.props);
  ctx.PropertiesService = { getScriptProperties: () => props };
  ctx.Utilities = {
    formatDate,
    newBlob: (data, type, name) => fakeBlob(data, type, name),
    base64Decode: (s) => [...Buffer.from(s, 'base64')],
  };
  ctx.DriveApp = {
    getFolderById: (id) => {
      const f = allFolders().find((x) => x.id === id);
      if (!f) throw new Error('No item with the given ID could be found');
      return f;
    },
    getFileById: (id) => {
      for (const f of allFolders()) { const file = f.files.find((x) => x.id === id); if (file) return file; }
      throw new Error('not found');
    },
  };
  ctx.GmailApp = {
    createDraft: (to, subject, body, opts) => {
      if (env.draftError) throw env.draftError;
      const d = { id: `DRAFT${env.drafts.length + 1}`, to, subject, body, opts };
      env.drafts.push(d);
      return { getId: () => d.id };
    },
    getDraft: (id) => {
      const d = env.drafts.find((x) => x.id === id && !x.sent);
      if (!d) throw new Error('Invalid draft');
      return { deleteDraft: () => { env.deletedDrafts.push(id); d.deleted = true; } };
    },
  };
  ctx.CalendarApp = {
    getCalendarById: () => ({ getEvents: () => env.holidays.map((s) => ({ getStartTime: () => ctx.toDate_(s) })) }),
  };
  vm.createContext(ctx);
  const code = SOURCES.map((f) => fs.readFileSync(path.join(DEMO_DIR, f), 'utf8')).join('\n');
  vm.runInContext(`${code}\n${EXPORTS.map((n) => `globalThis.${n} = ${n};`).join('\n')}\nglobalThis.VMDate = Date;`, ctx, { filename: 'invoice' });
  return { ctx, ss, env, folders };
}

/** 日付（GAS 側の Date で作る。コードの「instanceof Date」が効くように） */
function d(t, s) { return t.ctx.toDate_(s); }

/** 「設定」に書く */
function setSettings(t, values) {
  t.ctx.ensureAllSheets_();
  Object.entries(values).forEach(([k, v]) => t.ctx.writeSetting_(k, v));
}

/** 動かせる状態にする（設定・取引先を入れる） */
function setup(t, { clients = CLIENTS } = {}) {
  setSettings(t, {
    company: '株式会社テスト制作',
    regno: 'T1123456789012',
    bank: 'テスト銀行 本店 普通 1234567',
    pdfFolder: t.folders.pdf.getUrl(),
    csvFolder: t.folders.csv.getUrl(),
    admin: 'admin@example.com',
  });
  addRows(t, '取引先', clients);
  t.ctx.ensureAllSheets_();
}

const CLIENTS = [
  { 取引先名: '株式会社サンプル商事', 先方の担当者: '山田 太郎', メールアドレス: 'yamada@example.com', 締め日: '末日', 支払月: '翌月', 支払日: '末日', '振込名義（カナ）': 'カ)サンプルシヨウジ' },
  { 取引先名: '有限会社テスト工務店', 先方の担当者: '', メールアドレス: 'office@example.net', 締め日: '末日', 支払月: '翌月', 支払日: '末日', '振込名義（カナ）': 'ユ)テストコウムテン' },
  { 取引先名: '合同会社デモ食品', 先方の担当者: '鈴木 花子', メールアドレス: 'suzuki@example.org', 締め日: '20日', 支払月: '翌月', 支払日: '10日', '振込名義（カナ）': 'ド)デモシヨクヒン' },
  { 取引先名: '株式会社見本不動産', 先方の担当者: '', メールアドレス: '', 締め日: '末日', 支払月: '翌々月', 支払日: '10日', '振込名義（カナ）': 'カ)ミホンフドウサン' },
];

/** 表のシートの下に行を足す（rows は { 見出し: 値 }。日付は 'yyyy/mm/dd' の文字なら日付にする） */
function addRows(t, sheetName, rows) {
  const tb = t.ctx.openTable_(sheetName);
  let r = tb.lastRow() + 1;
  rows.forEach((row) => {
    Object.entries(row).forEach(([h, v]) => {
      // 'yyyy/mm/dd' は日付にする（2月31日のような、ない日は、本物と同じく文字のまま）
      const value = typeof v === 'string' && /^\d{4}\/\d{2}\/\d{2}$/.test(v) ? d(t, v) || v : v;
      tb.sheet.set(r, tb.map[h], value);
    });
    r++;
  });
}

/** 表を { 見出し: 値 } の配列で読む */
function rowsOf(t, sheetName) {
  return t.ctx.openTable_(sheetName).rows();
}

/** 1行の背景（「確認事項」の列の色で見る） */
function colorOf(t, sheetName, row) {
  const tb = t.ctx.openTable_(sheetName);
  return tb.sheet.bgOf(row, tb.map['確認事項']);
}

/** 今の時刻を固定する（Date.now と new Date() を差し替える） */
function setNow(t, iso) {
  const ms = Date.parse(iso);
  const Real = t.ctx.VMDate;
  vm.runInContext(`
    (function () {
      const Real = Date;
      const fixed = ${ms};
      function FakeDate(...a) { return a.length ? new Real(...a) : new Real(fixed); }
      FakeDate.prototype = Real.prototype;
      FakeDate.now = () => fixed;
      FakeDate.UTC = Real.UTC;
      FakeDate.parse = Real.parse;
      globalThis.Date = FakeDate;
    })();`, t.ctx);
  return Real;
}

/** CSVを入金明細のフォルダに置く */
function putCsv(t, name, lines, { sjis = false } = {}) {
  const text = lines.map((l) => l.map((c) => `"${c}"`).join(',')).join('\r\n');
  const blob = fakeBlob(text, 'text/csv', name, sjis);
  return t.folders.csv.createFile(blob);
}

module.exports = { load, setup, setSettings, addRows, rowsOf, colorOf, setNow, putCsv, d, CLIENTS };
