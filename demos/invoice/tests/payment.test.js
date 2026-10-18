// 「入金を取り込む」ボタンと、「入金」の「確認」のチェック
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, setup, setSettings, addRows, rowsOf, colorOf, putCsv } = require('./helper');

const HEAD = ['日付', '内容', '出金金額(円)', '入金金額(円)', '残高(円)'];

/** 請求書を3枚作った状態（2026-10-001 サンプル商事 110,000・002 サンプル商事 22,000・003 デモ食品 55,000） */
function ready() {
  const t = load();
  setup(t);
  addRows(t, '明細', [
    { 請求日: '2026/10/31', 取引先: '株式会社サンプル商事', 品目: '制作費', 数量: 1, '単価（円）': 100000, 税率: '10%' },
    { 請求日: '2026/10/15', 取引先: '株式会社サンプル商事', 品目: '撮影', 数量: 1, '単価（円）': 20000, 税率: '10%' },
    { 請求日: '2026/10/20', 取引先: '合同会社デモ食品', 品目: '更新', 数量: 1, '単価（円）': 50000, 税率: '10%' },
  ]);
  t.ctx.createInvoices();
  t.env.alerts.length = 0;
  return t;
}
const status = (t) => Object.fromEntries(rowsOf(t, '請求書一覧').map((r) => [r['取引先'] + r['請求額（円）'], r['状態']]));

test('名義と金額が合う入金は自動で入金済みに。手数料を引いた入金・名義の違う入金・利息は黄色で確認待ち', () => {
  const t = ready();
  putCsv(t, 'nyukin.csv', [HEAD,
    ['2026/11/28', '振込＊ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '110,000', ''],
    ['2026/11/05', '振込＊ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', ''],
    ['2026/11/10', 'カード利用', '3300', '', ''],
    ['2026/11/30', 'ﾔﾏﾀﾞ ﾀﾛｳ', '', '22000', ''],
    ['2026/11/30', '利息', '', '12', ''],
  ], { sjis: true });
  t.ctx.importPayments();
  const result = t.ss.getSheetByName('操作').getRange('B18').getValue();
  assert.match(result, /入金 4件 を取り込みました。入金済みにした請求書 1枚　確認待ち 3件/);

  const p = rowsOf(t, '入金');
  assert.equal(p.length, 4);
  assert.equal(t.ctx.ymd_(p[0]['入金日']), '2026-11-05'); // 日付の順
  const by = (amount) => p.find((r) => r['入金額（円）'] === amount);
  assert.equal(by(110000)['結果'], '入金済みにしました');
  assert.equal(by(110000)['請求書番号'], '2026-10-001');
  assert.equal(by(54560)['結果'], '確認待ち');
  assert.equal(by(54560)['請求書番号'], '2026-10-003');
  assert.match(by(54560)['確認事項'], /440円 少ない入金/);
  assert.equal(colorOf(t, '入金', by(54560)._row), t.ctx.WARN_ROW_COLOR);
  assert.equal(by(22000)['請求書番号'], '2026-10-002');
  assert.match(by(22000)['確認事項'], /振込名義が「取引先」の振込名義（カナ）と合いません/);
  assert.equal(by(12)['請求書番号'], '');

  const inv = rowsOf(t, '請求書一覧').find((r) => r['請求書番号'] === '2026-10-001');
  assert.equal(inv['状態'], '入金済み');
  assert.equal(t.ctx.ymd_(inv['入金日']), '2026-11-28');
  assert.equal(inv['入金額（円）'], 110000);
  assert.equal(status(t)['合同会社デモ食品55000'], '未入金'); // 確認待ちは、まだ変えない
});

test('「確認」にチェックすると反映。手数料の差は最後の請求書から引き、理由を残す。番号が空なら「請求書の支払いではない」', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/05', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', ''], ['2026/11/30', '利息', '', '12', '']]);
  t.ctx.importPayments();
  const tb = t.ctx.openTable_('入金');
  const check = (amount) => {
    const r = rowsOf(t, '入金').find((x) => x['入金額（円）'] === amount);
    tb.write(r._row, '確認', true);
    t.ctx.onEdit({ range: tb.sheet.getRange(r._row, tb.map['確認']) });
    return rowsOf(t, '入金').find((x) => x['入金額（円）'] === amount);
  };
  const fee = check(54560);
  assert.equal(fee['結果'], '入金済みにしました');
  assert.match(fee['確認事項'], /振込手数料 440円 を引いた入金として反映しました/);
  assert.equal(colorOf(t, '入金', fee._row), null);
  const inv = rowsOf(t, '請求書一覧').find((r) => r['請求書番号'] === '2026-10-003');
  assert.equal(inv['状態'], '入金済み');
  assert.equal(inv['入金額（円）'], 54560);

  const interest = check(12);
  assert.equal(interest['結果'], '請求書の支払いではない');
  assert.equal(colorOf(t, '入金', interest._row), null);
});

test('「確認」のとき、番号が違えばチェックを外して理由を書く。差の大きい入金は黄色で差を残して反映し、チェックを外すと取り消せる。2枚分はカンマ区切り', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/30', 'ﾔﾏﾀﾞ ﾀﾛｳ', '', '131560', '']]);
  t.ctx.importPayments();
  const tb = t.ctx.openTable_('入金');
  const row = rowsOf(t, '入金')[0]._row;
  const check = (value) => {
    tb.write(row, '確認', value);
    t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['確認']) });
    return rowsOf(t, '入金')[0];
  };
  const confirm = (nos) => { tb.write(row, '請求書番号', nos); return check(true); };
  const inv = (no) => rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === no);
  let r = confirm('2026-10-999');
  assert.equal(r['確認'], false);
  assert.match(r['確認事項'], /2026-10-999 という請求書がありません/);

  // 差の大きい入金：反映して、差を黄色で残す
  r = confirm('2026-10-002');
  assert.equal(r['結果'], '入金済みにしました');
  assert.match(r['確認事項'], /請求額の合計（22,000円）より 109,560円 多い入金のまま、入金済みにしました/);
  assert.equal(colorOf(t, '入金', r._row), t.ctx.WARN_ROW_COLOR);
  assert.equal(inv('2026-10-002')['入金額（円）'], 131560);

  // 番号を間違えた → チェックを外すと、請求書が未入金に戻る
  r = check(false);
  assert.equal(r['結果'], '確認待ち');
  assert.match(r['確認事項'], /反映を取り消しました/);
  assert.equal(inv('2026-10-002')['状態'], '未入金');
  assert.equal(inv('2026-10-002')['入金日'], '');
  assert.equal(inv('2026-10-002')['入金額（円）'], '');

  r = confirm('2026-10-001、2026-10-002');
  assert.equal(r['結果'], '入金済みにしました');
  assert.equal(inv('2026-10-001')['入金額（円）'], 110000);
  assert.equal(inv('2026-10-002')['入金額（円）'], 21560); // 手数料440円は最後の1枚から
  assert.match(r['確認事項'], /振込手数料 440円/);
});

test('「請求書一覧」で未入金に戻すと、入金日・入金額を消し、その請求書で反映した「入金」の行を確認待ちに戻す', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/28', 'ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '110000', '']]);
  t.ctx.importPayments();
  const list = t.ctx.openTable_('請求書一覧');
  const row = rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === '2026-10-001')._row;
  list.write(row, '状態', '未入金');
  t.ctx.onEdit({ range: list.sheet.getRange(row, list.map['状態']) });
  const inv = rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === '2026-10-001');
  assert.equal(inv['入金日'], '');
  assert.equal(inv['入金額（円）'], '');
  const p = rowsOf(t, '入金')[0];
  assert.equal(p['結果'], '確認待ち');
  assert.equal(p['確認'], false);
  assert.match(p['確認事項'], /未入金に戻されたため/);
});

test('ほかの処理の途中に「確認」にチェックしたら、チェックを外して理由を書く（黙って何もしない、にしない）', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/05', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', '']]);
  t.ctx.importPayments();
  const tb = t.ctx.openTable_('入金');
  const row = rowsOf(t, '入金')[0]._row;
  t.env.scriptLockBusy = true;
  tb.write(row, '確認', true);
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['確認']) });
  const r = rowsOf(t, '入金')[0];
  assert.equal(r['確認'], false);
  assert.match(r['確認事項'], /ほかの処理の途中だったため、反映できませんでした/);
  assert.equal(rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === '2026-10-003')['状態'], '未入金');
});

test('取り込みが途中で止まっても、「入金」の行は残り、次に押すと続きを反映する（請求書だけが入金済みにならない）', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/28', 'ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '110000', ''], ['2026/11/29', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '55000', '']]);
  const orig = t.ctx.applyPayment_;
  let n = 0;
  t.ctx.applyPayment_ = (...a) => { if (++n === 2) throw new Error('timeout'); return orig(...a); };
  t.ctx.importPayments();
  assert.match(t.env.alerts.pop().title, /管理者に連絡/);
  let p = rowsOf(t, '入金');
  assert.equal(p.length, 2);
  assert.equal(p[0]['結果'], '入金済みにしました');
  assert.equal(p[1]['結果'], '確認待ち');
  assert.match(p[1]['確認事項'], /反映の途中で止まりました/);
  assert.equal(rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === '2026-10-003')['状態'], '未入金');

  t.ctx.applyPayment_ = orig;
  t.ctx.importPayments();
  p = rowsOf(t, '入金');
  assert.equal(p.length, 2); // 同じ入金を足さない
  assert.equal(p[1]['結果'], '入金済みにしました');
  assert.equal(rowsOf(t, '請求書一覧').find((x) => x['請求書番号'] === '2026-10-003')['状態'], '入金済み');
});

test('時刻つき・和暦の日付も読む。読めない行は、ほかの入金を止めずに知らせ、そのCSVは次に押したときにもう一度読む', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD,
    ['2026/11/28 10:15', 'ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '+110,000', ''],
    ['R8.11.29', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '55000', ''],
    ['11月30日', 'ﾘｿｸ', '', '12', ''],
  ]);
  t.ctx.importPayments();
  const a = t.env.alerts.pop();
  assert.equal(a.title, '読めなかったものがあります');
  assert.match(a.message, /入金 2件 を取り込みました/);
  assert.match(a.message, /「a\.csv」：入金の 1行 の日付か金額を読めませんでした（例：4行目「11月30日」）/);
  const p = rowsOf(t, '入金');
  assert.equal(p.length, 2);
  assert.equal(t.ctx.ymd_(p[1]['入金日']), '2026-11-29');
  t.ctx.importPayments();
  assert.equal(t.env.alerts.pop().title, '読めなかったものがあります'); // 消えずに、また知らせる
  assert.equal(rowsOf(t, '入金').length, 2);
});

test('列の名前が合わないCSVがあっても、ほかのCSVは取り込み、合わないCSVの名前と直し方を出す。CSV以外のファイルは読まない', () => {
  const t = ready();
  putCsv(t, 'other.csv', [['取引日', '摘要', '金額'], ['2026/11/05', 'x', '1']]);
  putCsv(t, 'good.csv', [HEAD, ['2026/11/05', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', '']]);
  t.folders.csv.createFile({ name: 'memo.txt', getDataAsString: () => '', getName: () => 'memo.txt' });
  t.ctx.importPayments();
  const a = t.env.alerts.pop();
  assert.equal(a.title, '読めなかったものがあります');
  assert.match(a.message, /「other\.csv」：「日付」「内容」「入金金額\(円\)」の列が見つかりません/);
  assert.equal(rowsOf(t, '入金').length, 1);
  setSettings(t, { csvDate: '取引日', csvName: '摘要', csvAmount: '金額' });
  t.ctx.importPayments();
  assert.equal(rowsOf(t, '入金').length, 2);
});

test('取消・入金済みの請求書の番号で「確認」すると、その状態を伝える', () => {
  const t = ready();
  putCsv(t, 'a.csv', [HEAD, ['2026/11/30', 'ﾔﾏﾀﾞ ﾀﾛｳ', '', '1', '']]);
  t.ctx.importPayments();
  const inv = t.ctx.openTable_('請求書一覧');
  inv.write(rowsOf(t, '請求書一覧')[0]._row, '状態', '入金済み');
  const tb = t.ctx.openTable_('入金');
  const row = rowsOf(t, '入金')[0]._row;
  tb.write(row, '請求書番号', '2026-10-001');
  tb.write(row, '確認', true);
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['確認']) });
  assert.match(rowsOf(t, '入金')[0]['確認事項'], /2026-10-001 は「入金済み」です/);
});

test('同じCSVをもう一度押しても、期間が重なるCSVを置いても、同じ入金を2回足さない。同じ日・同じ名義・同じ金額の2件は2件のまま', () => {
  const t = ready();
  putCsv(t, 'oct.csv', [HEAD, ['2026/11/05', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', ''], ['2026/11/06', 'ﾘｿｸ', '', '5', ''], ['2026/11/06', 'ﾘｿｸ', '', '5', '']]);
  t.ctx.importPayments();
  assert.equal(rowsOf(t, '入金').length, 3);
  t.ctx.importPayments();
  assert.match(t.ss.getSheetByName('操作').getRange('B18').getValue(), /新しい入金はありませんでした/);
  putCsv(t, 'nov.csv', [HEAD, ['2026/11/06', 'ﾘｿｸ', '', '5', ''], ['2026/11/06', 'ﾘｿｸ', '', '5', ''], ['2026/11/28', 'ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '22000', '']]);
  t.ctx.importPayments();
  assert.equal(rowsOf(t, '入金').length, 4);
});

test('「確認」の列の空の行に「チェックなし」が入っていても、入金は表のすぐ下に書く。チェックボックスは入金のある行だけ', () => {
  const t = ready();
  const tb = t.ctx.openTable_('入金');
  for (let r = tb.headerRow + 1; r <= 1000; r++) tb.sheet.set(r, tb.map['確認'], false); // 本物でチェックボックスを全部の行に付けたときの形
  putCsv(t, 'a.csv', [HEAD, ['2026/11/05', 'ﾄﾞ)ﾃﾞﾓｼﾖｸﾋﾝ', '', '54560', '']]);
  t.ctx.importPayments();
  const p = rowsOf(t, '入金');
  assert.equal(p.length, 1);
  assert.equal(p[0]._row, tb.headerRow + 1);
  t.ctx.ensureAllSheets_(); // 見た目をそろえると、空の行の「チェックなし」は消える
  assert.equal(tb.sheet.get(tb.headerRow + 2, tb.map['確認']), '');
  const boxes = tb.sheet.validations.filter((v) => v.col === tb.map['確認'] && v.rule.checkbox);
  assert.ok(boxes.every((v) => v.row === tb.headerRow + 1 && v.numRows === 1));
});

test('入金明細のフォルダが空欄なら、設定を案内して止まる', () => {
  const t = ready();
  setSettings(t, { csvFolder: '' });
  t.ctx.importPayments();
  assert.match(t.env.alerts.pop().message, /「入金明細のCSVを置くフォルダ」を書いてください/);
});

test('取消した請求書とは照らし合わせない', () => {
  const t = ready();
  t.ctx.cancelInvoice_('2026-10-001', 'テスト', true);
  putCsv(t, 'a.csv', [HEAD, ['2026/11/28', 'ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', '', '110000', '']]);
  t.ctx.importPayments();
  const p = rowsOf(t, '入金')[0];
  assert.equal(p['結果'], '確認待ち');
  assert.match(p['確認事項'], /金額の合う未入金の請求書がありません/);
});
