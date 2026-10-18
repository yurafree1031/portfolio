// 計算・照らし合わせ・日付のロジック（シートを使わない部分）
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, d } = require('./helper');

const t = load();
const c = t.ctx;
const ymd = (x) => c.ymd_(x);
const client = (o = {}) => ({ name: 'A社', closing: '末日', payMonth: '翌月', payDay: '末日', ...o });

test('税率：「10%」「10％」「0.1」「10」は10、8%も同じ、それ以外は0', () => {
  for (const v of ['10%', '10％', 0.1, 10, '１０%']) assert.equal(c.taxRateOf_(v), 10, String(v));
  for (const v of ['8%', 0.08, 8]) assert.equal(c.taxRateOf_(v), 8, String(v));
  for (const v of ['5%', 'ten', 0.05, '']) assert.equal(c.taxRateOf_(v), 0, String(v));
});

test('端数：切り捨て・四捨五入・切り上げ。マイナスは0に向けて。小数の誤差で1円ずれない', () => {
  assert.equal(c.roundYen_(1234.5, '切り捨て'), 1234);
  assert.equal(c.roundYen_(1234.5, '四捨五入'), 1235);
  assert.equal(c.roundYen_(1234.1, '切り上げ'), 1235);
  assert.equal(c.roundYen_(1234, '切り上げ'), 1234);
  assert.equal(c.roundYen_(-1234.5, '切り捨て'), -1234);
  assert.equal(c.roundYen_(0.1 * 3 * 10, '切り捨て'), 3); // 2.9999… を 2 にしない
  assert.equal(c.roundYen_(3.3 * 3, '切り上げ'), 10);
});

test('金額：税率ごとに合計してから消費税を1回だけ計算する（明細ごとに計算しない）', () => {
  // 明細ごとに切り捨てると 99+99+99=297 になるが、合計の 2997*0.1=299.7 → 299 が正しい
  const lines = [{ rate: 10, amount: 999 }, { rate: 10, amount: 999 }, { rate: 10, amount: 999 }, { rate: 8, amount: 5880 }];
  const r = c.invoiceTotals_(lines, '切り捨て');
  assert.deepEqual(JSON.parse(JSON.stringify(r.byRate)), [{ rate: 10, base: 2997, tax: 299 }, { rate: 8, base: 5880, tax: 470 }]);
  assert.equal(r.base, 8877);
  assert.equal(r.tax, 769);
  assert.equal(r.total, 9646);
  assert.equal(c.invoiceTotals_([{ rate: 10, amount: 2997 }], '四捨五入').tax, 300);
});

test('金額：使っていない税率の行は出さない', () => {
  const r = c.invoiceTotals_([{ rate: 10, amount: 1000 }], '切り捨て');
  assert.equal(r.byRate.length, 1);
});

test('支払期限：末日締め翌月末払い・20日締め翌月10日払い・翌々月・月末の日数の違い', () => {
  assert.equal(ymd(c.dueDate_(d(t, '2026/10/31'), client())), '2026-11-30');
  assert.equal(ymd(c.dueDate_(d(t, '2026/10/15'), client())), '2026-11-30');
  assert.equal(ymd(c.dueDate_(d(t, '2026/01/31'), client())), '2026-02-28'); // 2月は28日まで
  const twenty = client({ closing: '20日', payDay: '10日' });
  assert.equal(ymd(c.dueDate_(d(t, '2026/10/20'), twenty)), '2026-11-10');
  assert.equal(ymd(c.dueDate_(d(t, '2026/10/21'), twenty)), '2026-12-10'); // 締め日を過ぎたら次の締め
  assert.equal(ymd(c.dueDate_(d(t, '2026/12/31'), client({ payMonth: '翌々月', payDay: '10日' }))), '2027-02-10'); // 年をまたぐ
  // 当月払いで支払日が請求日より前 → 次の月の支払日
  assert.equal(ymd(c.dueDate_(d(t, '2026/10/25'), client({ payMonth: '当月', payDay: '10日' }))), '2026-11-10');
});

test('日付の読み取り：いろいろな書き方。ない日（2月31日）はじく', () => {
  for (const s of ['2026/10/31', '2026-10-31', '20261031', '2026年10月31日', '２０２６/１０/３１']) assert.equal(ymd(c.toDate_(s)), '2026-10-31', s);
  // 銀行のCSVに多い書き方：時刻つき・和暦
  for (const s of ['2026/10/31 10:15', '2026-10-31T09:00:00', 'R8.10.31', 'R08/10/31', '令和8年10月31日']) assert.equal(ymd(c.toDate_(s)), '2026-10-31', s);
  assert.equal(ymd(c.toDate_('令和元年5月1日')), '2019-05-01');
  assert.equal(ymd(c.toDate_('H31.4.30')), '2019-04-30');
  assert.equal(c.toDate_('10/31'), null); // 年のない日付は、決めつけずに読まない
  assert.equal(c.toDate_('2026/02/31'), null);
  assert.equal(c.toDate_('来週'), null);
  assert.equal(c.toDate_(''), null);
});

test('数字の読み取り：カンマ・円・全角。数でなければ NaN', () => {
  assert.equal(c.toNumber_('1,280円'), 1280);
  assert.equal(c.toNumber_('１２８０'), 1280);
  assert.equal(c.toNumber_('-500'), -500);
  assert.equal(c.toNumber_('¥12,000'), 12000);
  assert.equal(c.toNumber_('+33,000'), 33000);
  assert.ok(Number.isNaN(c.toNumber_('千円')));
  assert.ok(Number.isNaN(c.toNumber_('')));
});

test('登録番号の表記ゆれをそろえる', () => {
  assert.equal(c.normalizeRegNo_('ｔ1123-4567-8901-2'), 'T1123456789012');
  assert.equal(c.normalizeRegNo_('Ｔ１１２３４５６７８９０１２'), 'T1123456789012');
});

test('振込名義：半角カナ・小さい字・会社の種類の略・先頭の「振込」の違いをなくす', () => {
  assert.equal(c.kanaKey_('振込＊ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ'), c.kanaKey_('カ)サンプルシヨウジ'));
  assert.equal(c.kanaKey_('ｻﾝﾌﾟﾙｼｮｳｼﾞ(ｶ'), c.kanaKey_('サンプルショウジ'));
  assert.equal(c.kanaKey_('ユ）テスト　コウムテン'), 'テストコウムテン');
  assert.equal(c.kanaKey_('さんぷる'), 'サンプル');
  assert.ok(c.nameMatches_('ﾌﾘｺﾐ ｶ)ｻﾝﾌﾟﾙｼﾖｳｼﾞ', 'カ)サンプルシヨウジ'));
  assert.ok(c.nameMatches_('ｶ)ﾐﾎﾝﾌﾄﾞｳｻﾝｶﾌﾞｼｷｶ', 'カ)ミホンフドウサン')); // 後ろに余計な字
  assert.ok(c.nameMatches_('ｶ)ﾐﾎﾝﾌﾄﾞｳ', 'カ)ミホンフドウサンカイハツ')); // 銀行が途中で切った
  assert.ok(!c.nameMatches_('ｶ)ﾐﾎﾝ', 'カ)ミホンフドウサン')); // 短すぎるものは合わせない
  assert.ok(!c.nameMatches_('ﾔﾏﾀﾞ ﾀﾛｳ', 'カ)サンプルシヨウジ'));
});

test('CSVの読み取り：「""」で囲んだカンマ・改行、CRLF', () => {
  const rows = c.parseCsv_('"日付","内容"\r\n"2026/11/05","振込,カンマ"\r\n"2026/11/06","改行\nあり"\r\n');
  assert.deepEqual(JSON.parse(JSON.stringify(rows)), [['日付', '内容'], ['2026/11/05', '振込,カンマ'], ['2026/11/06', '改行\nあり']]);
});

test('CSVから入金を取り出す：見出しの前の行を飛ばし、出金・金額が空の行は読まない', () => {
  const s = { csvDate: '日付', csvName: '内容', csvAmount: '入金金額(円)' };
  const text = ['口座：普通 1234567', '日付,内容,出金金額(円),入金金額(円)', '2026/11/05,振込＊ｶ)ｻﾝﾌﾟﾙ,,"110,000"', '2026/11/06,カード,3300,', '2026/11/07,利息,,12'].join('\n');
  const r = c.depositsFromCsv_(text, s);
  assert.equal(r.deposits.length, 2);
  assert.equal(r.deposits[0].amount, 110000);
  assert.equal(r.deposits[1].name, '利息');
  const bad = c.depositsFromCsv_('Date,Desc,Amount\n2026/11/05,x,1', s);
  assert.match(bad.error, /列が見つかりません/);
});

// ---- 入金の照らし合わせ ----
const clients = [
  { name: 'A社', kana: 'カ)エーシヤ' },
  { name: 'B社', kana: 'ユ)ビーシヤ' },
  { name: 'C社', kana: '' },
];
const inv = (no, client, total, due = '2026/11/30') => ({ no, client, total, due: d(t, due) });
const match = (dep, unpaid, reserved = new Set()) => JSON.parse(JSON.stringify(c.matchDeposit_({ name: dep[0], amount: dep[1] }, unpaid, clients, 880, reserved)));

test('照らし合わせ：名義と金額が合う1枚は自動', () => {
  const r = match(['ｶ)ｴｰｼﾔ', 110000], [inv('1', 'A社', 110000), inv('2', 'B社', 110000)]);
  assert.deepEqual([r.auto, r.nos], [true, ['1']]);
});

test('照らし合わせ：2枚分まとめた入金は、組み合わせが1つなら自動', () => {
  const r = match(['ｶ)ｴｰｼﾔ', 165000], [inv('1', 'A社', 110000), inv('2', 'A社', 55000), inv('3', 'A社', 33000)]);
  assert.deepEqual([r.auto, r.nos.sort()], [true, ['1', '2']]);
});

test('照らし合わせ：組み合わせが2通り以上なら、確認待ち', () => {
  const r = match(['ｶ)ｴｰｼﾔ', 30000], [inv('1', 'A社', 10000), inv('2', 'A社', 20000), inv('3', 'A社', 15000), inv('4', 'A社', 15000)]);
  assert.equal(r.auto, false);
  assert.match(r.note, /組み合わせが2通り/);
});

test('照らし合わせ：振込手数料ぶん少ない入金は、候補を入れて確認待ち', () => {
  const r = match(['ｶ)ｴｰｼﾔ', 109560], [inv('1', 'A社', 110000)]);
  assert.deepEqual([r.auto, r.nos], [false, ['1']]);
  assert.match(r.note, /440円 少ない/);
  const far = match(['ｶ)ｴｰｼﾔ', 100000], [inv('1', 'A社', 110000)]);
  assert.deepEqual([far.auto, far.nos], [false, []]); // 差が大きすぎる
  assert.match(far.note, /金額の合う未入金の請求書がありません/);
});

test('照らし合わせ：同じ取引先・同じ金額が2枚なら、期限の早いほうを候補に確認待ち。2件目の入金は残りの1枚を候補に', () => {
  const unpaid = [inv('1', 'A社', 33000, '2026/10/31'), inv('2', 'A社', 33000, '2026/11/30')];
  const reserved = new Set();
  const r1 = match(['ｶ)ｴｰｼﾔ', 33000], unpaid, reserved);
  assert.deepEqual([r1.auto, r1.nos], [false, ['1']]);
  reserved.add('1');
  const r2 = match(['ｶ)ｴｰｼﾔ', 33000], unpaid, reserved);
  assert.deepEqual([r2.auto, r2.nos], [false, ['2']]);
});

test('照らし合わせ：名義が合わなくても、金額が合う請求書が1枚なら候補にして確認待ち', () => {
  const r = match(['ﾔﾏﾀﾞ ﾀﾛｳ', 22000], [inv('1', 'B社', 22000), inv('2', 'A社', 110000)]);
  assert.deepEqual([r.auto, r.nos], [false, ['1']]);
  assert.match(r.note, /振込名義が「取引先」の振込名義（カナ）と合いません/);
});

test('照らし合わせ：振込名義（カナ）が空の取引先は、名義では合わせない', () => {
  const r = match(['ｼｰｼﾔ', 5000], [inv('1', 'C社', 5000)]);
  assert.equal(r.auto, false);
});

test('照らし合わせ：合うものがなければ、番号なしで確認待ち（利息など）', () => {
  const r = match(['ﾘｿｸ', 12], [inv('1', 'A社', 110000)]);
  assert.deepEqual([r.auto, r.nos], [false, []]);
  assert.match(r.note, /合う請求書が見つかりません/);
});

// ---- 次に来る締め日（請求日の自動入力） ----

test('次の締め日：末日締めは今月末、締め日を過ぎたら来月の締め日、今日が締め日なら今日', () => {
  assert.equal(ymd(c.nextClosing_(d(t, '2026/10/18'), client())), '2026-10-31');
  assert.equal(ymd(c.nextClosing_(d(t, '2026/10/31'), client())), '2026-10-31');
  const twenty = client({ closing: '20日' });
  assert.equal(ymd(c.nextClosing_(d(t, '2026/10/20'), twenty)), '2026-10-20');
  assert.equal(ymd(c.nextClosing_(d(t, '2026/10/21'), twenty)), '2026-11-20');
  assert.equal(ymd(c.nextClosing_(d(t, '2026/12/25'), twenty)), '2027-01-20'); // 年をまたぐ
  assert.equal(ymd(c.nextClosing_(d(t, '2027/02/01'), client())), '2027-02-28');
});

test('日付の印：請求日と期限の印を、ほかの印を残して書き換える', () => {
  let m = c.withDateMark_('毎月', '請求日', '2026-10-31');
  m = c.withDateMark_(m, '期限', '2026-11-30');
  assert.equal(m, '毎月 請求日:2026-10-31 期限:2026-11-30');
  assert.equal(c.dateMark_(m, '期限'), '2026-11-30');
  assert.equal(c.withDateMark_(m, '請求日', '2026-11-30'), '毎月 期限:2026-11-30 請求日:2026-11-30');
  assert.equal(c.withDateMark_(m, '期限', ''), '毎月 請求日:2026-10-31');
});
