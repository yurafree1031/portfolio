/** デモ用（開発用。デモのスプレッドシートを整えたら消す。お客様には渡さない） */

/** 2. 請求書を作ったあとに、一覧の金額から架空のネットバンクのCSV（Shift_JIS）を作って、入金明細のフォルダに置く */
function makeDemoBankCsv() {
  const s = readSettings_();
  const folder = openFolder_(s.csvFolder, '入金明細のCSVを置くフォルダ', 'E22');
  const inv = openTable_(SHEET.INVOICES).rows().filter((r) => r['状態'] === STATUS.UNPAID);
  const total = (name) => inv.filter((r) => r['取引先'] === name).map((r) => toNumber_(r['請求額（円）']));
  const sample = total('株式会社サンプル商事'); // 2枚（明細の分と、毎月の保守費）
  const test = total('有限会社テスト工務店');
  const demo = total('合同会社デモ食品');
  const rows = [
    ['日付', '内容', '出金金額(円)', '入金金額(円)', '残高(円)', 'メモ'],
    ['2026/11/05', '振込＊ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', String(demo[0] - 440), '', ''], // 振込手数料を引いた入金
    ['2026/11/10', 'カード利用 サンプル文具', '3,300', '', '', ''], // 出金（読まない）
    ['2026/11/28', '振込＊ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', String(sample.reduce((a, b) => a + b, 0)), '', ''], // 2枚分まとめた入金
    ['2026/11/30', '振込＊ﾔﾏﾀﾞ ﾀﾛｳ', '', String(test[0]), '', ''], // 名義が違う（社長の個人名で振込）
    ['2026/11/30', '利息', '', '12', '', ''], // 請求書の支払いではない
  ];
  const text = rows.map((r) => r.map((c) => `"${c}"`).join(',')).join('\r\n');
  folder.createFile(Utilities.newBlob('').setDataFromString(text, 'Shift_JIS').setName('nyukin_202611.csv').setContentType('text/csv'));
}
