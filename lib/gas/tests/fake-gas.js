// GAS の部品（スプレッドシート・メール・トリガーなど）の偽物。各デモの tests/helper.js から使う（どのデモでも共通）
// デモに合わせた偽物（フォームなど）は、各デモの helper.js に書く
'use strict';

const colToNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

// ---- 偽のスプレッドシート ----
class FakeRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    if (row < 1 || col < 1) throw new Error(`範囲がシートの外です: ${row},${col}`);
    if (row + numRows - 1 > sheet.maxRows || col + numCols - 1 > sheet.maxCols) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  each(fn) { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) fn(this.row + r, this.col + c, r, c); }
  getValue() { return this.sheet.get(this.row, this.col); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) { out.push([]); for (let c = 0; c < this.numCols; c++) out[r].push(this.sheet.get(this.row + r, this.col + c)); }
    return out;
  }
  setValue(v) { this.sheet.set(this.row, this.col, v); return this; }
  setValues(rows) {
    if (rows.length !== this.numRows || rows.some((r) => r.length !== this.numCols)) throw new Error('setValues の大きさが範囲と違います');
    this.each((r, c, i, j) => this.sheet.set(r, c, rows[i][j]));
    return this;
  }
  clearContent() { this.each((r, c) => this.sheet.cells.delete(`${r},${c}`)); return this; }
  setBackground(color) {
    this.each((r, c) => (color == null ? this.sheet.bg.delete(`${r},${c}`) : this.sheet.bg.set(`${r},${c}`, color)));
    return this;
  }
  getBackgrounds() { const out = []; for (let r = 0; r < this.numRows; r++) { out.push([]); for (let c = 0; c < this.numCols; c++) out[r].push(this.sheet.bg.get(`${this.row + r},${this.col + c}`) || null); } return out; }
  setBackgrounds(rows) { this.each((r, c, i, j) => (rows[i][j] == null ? this.sheet.bg.delete(`${r},${c}`) : this.sheet.bg.set(`${r},${c}`, rows[i][j]))); return this; }
  setDataValidation(rule) { this.sheet.validations.push({ row: this.row, col: this.col, numRows: this.numRows, rule }); return this; }
  clearDataValidations() { this.sheet.validations = this.sheet.validations.filter((v) => !(v.col === this.col && v.row === this.row)); return this; }
  setWrap() { return this; }
  setVerticalAlignment() { return this; }
  clearFormat() { this.setBackground(null); (this.sheet.cleared = this.sheet.cleared || []).push(this.row); return this; }
  setFontWeight() { return this; }
  setFontColor() { return this; }
  setFontSize() { return this; }
  setNumberFormat() { return this; }
  setBorder() { return this; }
  setHorizontalAlignment(a) { this.sheet.aligns.push({ row: this.row, col: this.col, numRows: this.numRows, a }); return this; }
  protect() {
    const p = { description: '', warningOnly: false, range: this };
    p.setDescription = (d) => { p.description = d; return p; };
    p.setWarningOnly = (w) => { p.warningOnly = w; return p; };
    p.getDescription = () => p.description;
    this.sheet.protections.push(p);
    return p;
  }
  getA1Notation() { let n = this.col; let c = ''; while (n > 0) { c = String.fromCharCode(65 + ((n - 1) % 26)) + c; n = Math.floor((n - 1) / 26); } return `${c}${this.row}`; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getLastColumn() { return this.col + this.numCols - 1; }
  getSheet() { return this.sheet; }
}

class FakeSheet {
  constructor(ss, name) {
    Object.assign(this, { ss, name, cells: new Map(), bg: new Map(), validations: [], protections: [], aligns: [], rules: [], hidden: [], widths: {} });
    this.maxRows = 1000;
    this.maxCols = 26;
    this.id = ss.nextId++;
  }
  get(r, c) { const k = `${r},${c}`; return this.cells.has(k) ? this.cells.get(k) : ''; }
  set(r, c, v) {
    // 本物と同じく、先頭の ' は「文字として入れる」印で、値には残らない
    if (typeof v === 'string' && v.startsWith("'")) v = v.slice(1);
    if (v === '' || v == null) this.cells.delete(`${r},${c}`); else this.cells.set(`${r},${c}`, v);
  }
  getName() { return this.name; }
  getFormUrl() { return this.formUrl || null; } // フォームの回答のシートなら、テストで formUrl を入れる
  getSheetId() { return this.id; }
  getIndex() { return this.ss.sheets.indexOf(this) + 1; }
  activate() { this.ss.active = this; return this; }
  getLastRow() { let m = 0; for (const k of this.cells.keys()) m = Math.max(m, Number(k.split(',')[0])); return m; }
  getLastColumn() { let m = 0; for (const k of this.cells.keys()) m = Math.max(m, Number(k.split(',')[1])); return m; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  getRange(a, b, c, d) {
    if (typeof a === 'number') return new FakeRange(this, a, b, c || 1, d || 1);
    // 「B2」のような1つのセルと、「B3:C4」のような範囲
    const m = String(a).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    const row = Number(m[2]);
    const col = colToNum(m[1]);
    if (!m[3]) return new FakeRange(this, row, col);
    return new FakeRange(this, row, col, Number(m[4]) - row + 1, colToNum(m[3]) - col + 1);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  // 行・列を足すと、それより下（右）の値と色がずれる
  shift(map, axis, at, n) {
    const out = new Map();
    for (const [k, v] of map) {
      const [r, c] = k.split(',').map(Number);
      const nr = axis === 'row' && r >= at ? r + n : r;
      const nc = axis === 'col' && c >= at ? c + n : c;
      out.set(`${nr},${nc}`, v);
    }
    return out;
  }
  insertRowsBefore(row, n) { this.cells = this.shift(this.cells, 'row', row, n); this.bg = this.shift(this.bg, 'row', row, n); this.maxRows += n; }
  insertRowBefore(row) { this.insertRowsBefore(row, 1); }
  insertRowsAfter(row, n) { this.insertRowsBefore(row + 1, n); }
  deleteRow(row) {
    const out = new Map();
    for (const [k, v] of this.cells) {
      const [r, c] = k.split(',').map(Number);
      if (r === row) continue;
      out.set(`${r > row ? r - 1 : r},${c}`, v);
    }
    this.cells = out;
    this.maxRows--;
  }
  insertColumnBefore(col) { this.cells = this.shift(this.cells, 'col', col, 1); this.bg = this.shift(this.bg, 'col', col, 1); this.maxCols++; }
  insertColumnAfter(col) { this.insertColumnBefore(col + 1); }
  hideColumns(col) { if (!this.hidden.includes(col)) this.hidden.push(col); }
  setFrozenRows(n) { this.frozen = n; }
  setColumnWidth(col, w) { this.widths[col] = w; }
  getColumnWidth(col) { return this.widths[col] || 100; }
  setHiddenGridlines() {}
  getConditionalFormatRules() { return this.rules; }
  setConditionalFormatRules(rules) { this.rules = rules; }
  getProtections() { return this.protections; }
  // グラフ（作ったときの設定をそのまま持つ）
  newChart() {
    const chart = { type: null, ranges: [], options: {}, headers: 0, position: null };
    const b = {
      setChartType(t) { chart.type = t; return b; },
      addRange(r) { chart.ranges.push(r); return b; },
      setNumHeaders(n) { chart.headers = n; return b; },
      setPosition(row, col) { chart.position = [row, col]; return b; },
      setOption(k, v) { chart.options[k] = v; return b; },
      build() { return chart; },
    };
    return b;
  }
  insertChart(c) { (this.charts = this.charts || []).push(c); }
  getCharts() { return (this.charts || []).slice(); }
  removeChart(c) { this.charts = (this.charts || []).filter((x) => x !== c); }
  // ---- テスト用 ----
  header(row) { const out = []; for (let c = 1; c <= this.getLastColumn(); c++) out.push(this.get(row, c)); return out; }
  col(name, headerRow) { return this.header(headerRow).indexOf(name) + 1; }
  rowObj(row, headerRow) { const h = this.header(headerRow); const o = {}; h.forEach((n, i) => { if (n) o[n] = this.get(row, i + 1); }); return o; }
  bgOf(row, col) { return this.bg.get(`${row},${col}`) || null; }
}

class FakeSpreadsheet {
  constructor({ firstSheet = 'シート1', formUrl = 'https://docs.google.com/forms/d/FORM/edit' } = {}) {
    this.nextId = 100;
    this.sheets = [];
    this.toasts = [];
    this.alerts = [];
    this.answers = [];
    this.menus = [];
    this.formUrl = formUrl;
    this.tz = 'Asia/Tokyo';
    this.active = this.add(firstSheet);
  }
  add(name) { const s = new FakeSheet(this, name); this.sheets.push(s); return s; }
  getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
  getSheets() { return this.sheets; }
  getActiveSheet() { return this.active; }
  insertSheet(name, index) {
    const s = new FakeSheet(this, name);
    this.sheets.splice(index == null ? this.sheets.length : index, 0, s);
    this.active = s; // 本物と同じく、作ったシートが開かれる
    return s;
  }
  getFormUrl() { return this.formUrl; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/SS/edit'; }
  getSpreadsheetTimeZone() { return this.tz; }
  toast(message, title, seconds) { this.toasts.push({ message, title, seconds }); }
  get lastToast() { return this.toasts[this.toasts.length - 1]; }
}


// ---- ほかの GAS の部品 ----
function pad(n, w = 2) { return String(n).padStart(w, '0'); }
function formatDate(date, tz, fmt) {
  // テストでは日本時間（UTC+9）だけを扱う
  const d = new Date(date.getTime() + 9 * 3600 * 1000);
  return fmt.replace('yyyy', d.getUTCFullYear()).replace('MM', pad(d.getUTCMonth() + 1)).replace('dd', pad(d.getUTCDate()))
    .replace('HH', pad(d.getUTCHours())).replace('mm', pad(d.getUTCMinutes()))
    .replace(/\bM\b/, d.getUTCMonth() + 1).replace(/\bd\b/, d.getUTCDate());
}

function validationBuilder() {
  const rule = {};
  const b = {
    requireValueInList(list, show) { rule.list = list; rule.show = show; return b; },
    requireTextIsEmail() { rule.email = true; return b; },
    requireNumberBetween(a, z) { rule.between = [a, z]; return b; },
    requireCheckbox() { rule.checkbox = true; return b; },
    requireDate() { rule.date = true; return b; },
    requireFormulaSatisfied(f) { rule.formula = f; return b; },
    setAllowInvalid(v) { rule.allowInvalid = v; return b; },
    setHelpText(t) { rule.help = t; return b; },
    build() { return rule; },
  };
  return b;
}
function conditionalBuilder() {
  const rule = {};
  const b = {
    whenTextEqualTo(t) { rule.text = t; return b; },
    setBackground(c) { rule.bg = c; return b; },
    setRanges(r) { rule.ranges = r; return b; },
    build() {
      return { ...rule, getBooleanCondition: () => ({ getCriteriaType: () => 'TEXT_EQUAL_TO', getCriteriaValues: () => [rule.text] }) };
    },
  };
  return b;
}

/**
 * Code.gs を偽の GAS 部品と一緒に読み込む
 * 戻り値: { ctx（読み込んだ関数）, ss, form, env（メール・トリガー・プロパティなどの記録） }

/**
 * どのデモでも使う GAS の部品の偽物を作る。env にメール・トリガー・プロパティ・確認画面などを記録する
 * env.uiAnswers に答え（'YES' など）を入れておくと、確認画面でその順に答える（空なら 'YES'）
 * 戻り値: vm に渡す部品（デモの helper.js で、FormApp などを足して使う）
 */
function gasContext(ss, env) {
  const ui = {
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
    Button: { YES: 'YES', NO: 'NO', CANCEL: 'CANCEL', CLOSE: 'CLOSE' },
    createMenu(title) {
      const menu = { title, items: [] };
      const chain = {
        addItem(label, fn) { menu.items.push([label, fn]); return chain; },
        addSeparator() { menu.items.push(['---']); return chain; },
        addToUi() { ss.menus.push(menu); return chain; },
      };
      return chain;
    },
    alert(title, message) { env.alerts.push({ title, message }); return env.uiAnswers.length ? env.uiAnswers.shift() : 'YES'; },
  };
  const lock = (key) => ({ tryLock: () => !env[key], waitLock: () => { if (env[key]) throw new Error('Lock timeout'); }, releaseLock() { env.lockReleases = (env.lockReleases || 0) + 1; } });
  const triggerBuilder = (handler) => {
    const t = { handler, kind: null };
    const b = {
      forSpreadsheet() { t.kind = 'formSubmit'; return b; },
      onFormSubmit() { return b; },
      timeBased() { t.kind = 'time'; return b; },
      everyDays(n) { t.every = n; return b; },
      everyHours(n) { t.everyHours = n; return b; },
      onWeekDay(d) { t.weekDay = d; return b; },
      atHour(h) { t.hour = h; return b; },
      inTimezone() { return b; },
      after(ms) { t.after = ms; return b; },
      create() { t.owner = env.me; t.getHandlerFunction = () => handler; env.triggers.push(t); return t; },
    };
    return b;
  };
  return {
    console,
    Logger: { log: (m) => env.logs.push(String(m)) },
    SpreadsheetApp: {
      getActive: () => ss,
      getUi: () => ui,
      flush() {},
      newDataValidation: validationBuilder,
      newConditionalFormatRule: conditionalBuilder,
      BooleanCriteria: { TEXT_EQUAL_TO: 'TEXT_EQUAL_TO' },
      ProtectionType: { RANGE: 'RANGE' },
    },
    MailApp: {
      getRemainingDailyQuota: () => env.quota,
      sendEmail: (m) => {
        if (env.mailError) throw env.mailError;
        env.quota -= String(m.to).split(',').length;
        env.mails.push(m);
      },
    },
    LockService: { getScriptLock: () => lock('scriptLockBusy'), getDocumentLock: () => lock('docLockBusy') },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in env.props ? env.props[k] : null),
        setProperty: (k, v) => { env.props[k] = String(v); },
        setProperties: (o) => { Object.entries(o).forEach(([k, v]) => { env.props[k] = String(v); }); },
        deleteProperty: (k) => { delete env.props[k]; },
      }),
    },
    ScriptApp: {
      // 本物と同じく、見えるのは今のアカウントのトリガーだけ
      getProjectTriggers: () => env.triggers.filter((t) => t.owner === env.me),
      newTrigger: triggerBuilder,
      deleteTrigger: (t) => { env.triggers = env.triggers.filter((x) => x !== t); },
      WeekDay: { SUNDAY: 'SUNDAY', MONDAY: 'MONDAY', TUESDAY: 'TUESDAY', WEDNESDAY: 'WEDNESDAY', THURSDAY: 'THURSDAY', FRIDAY: 'FRIDAY', SATURDAY: 'SATURDAY' },
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => env.me }) },
    Utilities: { formatDate },
    Charts: { ChartType: { LINE: 'LINE', COLUMN: 'COLUMN', PIE: 'PIE' } },
  };
}

/** gasContext に渡す env の初期値 */
function newEnv(me) {
  return { me, props: {}, mails: [], triggers: [], logs: [], quota: 100, mailError: null, scriptLockBusy: false, docLockBusy: false, uiAnswers: [], alerts: [] };
}

module.exports = { FakeRange, FakeSheet, FakeSpreadsheet, formatDate, validationBuilder, conditionalBuilder, gasContext, newEnv };
