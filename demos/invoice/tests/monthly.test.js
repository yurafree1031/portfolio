// 自動の処理（始める・止める・毎朝の処理）
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, setup, setSettings, addRows, rowsOf, setNow } = require('./helper');

/** 「明細」の上の帯の文 */
const band = (t) => t.ss.getSheetByName('明細').get(t.ctx.TABLES['明細'].alertRow, 1);

function ready(now = '2026-10-18T09:00:00+09:00') {
  const t = load();
  setup(t);
  setNow(t, now);
  addRows(t, '毎月の請求', [
    { 取引先: '株式会社サンプル商事', 品目: '保守費（{月}分）', 数量: 1, '単価（円）': 30000, 税率: '10%' },
    { 取引先: '合同会社デモ食品', 品目: 'サーバー管理費', 数量: '', '単価（円）': 5000, 税率: '' },
  ]);
  return t;
}

test('始める：今月の毎月の明細を入れるかを聞き、「OK」で今月の締め日を請求日にして入れる。トリガーを入れる', () => {
  const t = ready();
  t.ctx.startAuto();
  const ask = t.env.alerts.find((a) => a.title === '10月の毎月の明細を入れますか？');
  assert.match(ask.message, /「毎月の請求」の 2行/);
  const lines = rowsOf(t, '明細');
  assert.equal(lines.length, 2);
  assert.equal(t.ctx.ymd_(lines[0]['請求日']), '2026-10-31');
  assert.equal(lines[0]['品目'], '保守費（10月分）');
  assert.equal(t.ctx.ymd_(lines[1]['請求日']), '2026-10-20'); // 20日締め
  assert.equal(t.ctx.ymd_(lines[0]['支払期限']), '2026-11-30');
  assert.equal(t.ctx.ymd_(lines[1]['支払期限']), '2026-11-10');
  assert.equal(lines[1]['数量'], 1);
  assert.equal(lines[1]['税率'], '10%');
  assert.equal(t.env.triggers.length, 1);
  assert.equal(t.env.triggers[0].handler, 'dailyRun');
  assert.equal(t.env.triggers[0].hour, 8);
  assert.match(t.env.alerts.pop().title, /自動の処理を始めました/);
  assert.match(t.ctx.readSettings_().state, /動いています/);
});

test('始める：「いいえ」なら今月の分は入れず、来月から。毎朝の処理でも今月は入れない', () => {
  const t = ready();
  t.env.uiAnswers.push('NO');
  t.ctx.startAuto();
  assert.equal(rowsOf(t, '明細').length, 0);
  t.ctx.dailyRun();
  assert.equal(rowsOf(t, '明細').length, 0);
  setNow(t, '2026-11-01T08:10:00+09:00');
  t.ctx.dailyRun();
  const lines = rowsOf(t, '明細');
  assert.equal(lines.length, 2);
  assert.equal(lines[0]['品目'], '保守費（11月分）');
  assert.equal(t.ctx.ymd_(lines[0]['請求日']), '2026-11-30');
  t.ctx.dailyRun(); // 同じ月に2回動いても、2回入れない
  assert.equal(rowsOf(t, '明細').length, 2);
});

test('始める：管理者のメールアドレスがないと始めない', () => {
  const t = ready();
  setSettings(t, { admin: '' });
  t.ctx.startAuto();
  assert.match(t.env.alerts.pop().message, /「管理者のメールアドレス」を書いてください/);
  assert.equal(t.env.triggers.length, 0);
});

test('始める：ほかの人が動かしているときに押すと、送信元を替えるかを確かめる。前の人のトリガーは何もせず片づく', () => {
  const t = ready();
  t.ctx.startAuto();
  t.env.me = 'other@example.com';
  t.env.uiAnswers.push('NO');
  t.ctx.startAuto();
  assert.equal(t.env.alerts.pop().title, '送信元を替えますか？');
  assert.equal(t.env.props.AUTO_SENDER, 'keiri@example.com');
  t.ctx.startAuto();
  assert.equal(t.env.props.AUTO_SENDER, 'other@example.com');
  t.env.me = 'keiri@example.com';
  t.ctx.dailyRun(); // 前の送信元のトリガーが動いた
  assert.equal(t.env.triggers.filter((x) => x.owner === 'keiri@example.com').length, 0);
});

test('止める：トリガーを外し、止めている間は毎朝の処理が動いても何もしない', () => {
  const t = ready();
  t.env.uiAnswers.push('NO');
  t.ctx.startAuto();
  t.ctx.stopAuto();
  assert.equal(t.env.triggers.length, 0);
  assert.equal(t.ctx.readSettings_().state, '止まっています');
  setNow(t, '2026-11-01T08:10:00+09:00');
  t.ctx.dailyRun();
  assert.equal(rowsOf(t, '明細').length, 0);
  // 再開すると、11月分を入れるかを聞く
  t.ctx.startAuto();
  assert.ok(t.env.alerts.some((a) => a.title === '11月の毎月の明細を入れますか？'));
  assert.equal(rowsOf(t, '明細').length, 2);
});

test('毎朝：毎月の請求に入れられない行（取引先がない）は赤くし、管理者にメールと帯で知らせる。ほかの行は入れる', () => {
  const t = ready();
  addRows(t, '毎月の請求', [{ 取引先: '消えた会社', 品目: '保守', 数量: 1, '単価（円）': 1000, 税率: '10%' }]);
  t.env.uiAnswers.push('NO');
  t.ctx.startAuto();
  setNow(t, '2026-11-01T08:10:00+09:00');
  t.ctx.dailyRun();
  assert.equal(rowsOf(t, '明細').length, 2);
  const bad = rowsOf(t, '毎月の請求').find((r) => r['取引先'] === '消えた会社');
  assert.match(bad['確認事項'], /「取引先」シートにない取引先です（2026年11月の明細に入れられませんでした）/);
  const alert = t.env.mails.find((x) => x.to === 'admin@example.com' && /エラー/.test(x.subject));
  assert.match(alert.body, /\[E27\]/);
  assert.match(band(t), /\[E27\]/);

  // 取引先名を直すと、次の朝に11月の明細として入り、帯が消える
  const mt = t.ctx.openTable_('毎月の請求');
  mt.write(bad._row, '取引先', '株式会社見本不動産');
  setNow(t, '2026-11-02T08:10:00+09:00');
  t.ctx.dailyRun();
  const added = rowsOf(t, '明細').find((r) => r['取引先'] === '株式会社見本不動産');
  assert.equal(t.ctx.ymd_(added['請求日']), '2026-11-30');
  assert.equal(rowsOf(t, '毎月の請求').find((r) => r['取引先'] === '株式会社見本不動産')['確認事項'], '');
  assert.equal(band(t), '');
  t.ctx.dailyRun(); // 2回入れない
  assert.equal(rowsOf(t, '明細').filter((r) => r['取引先'] === '株式会社見本不動産').length, 1);
});

test('毎朝：設定が正しくないときは、管理者にメールで知らせて帯を出す。直ってうまく動いたら帯を消す', () => {
  const t = ready();
  t.env.uiAnswers.push('NO');
  t.ctx.startAuto();
  setSettings(t, { fee: 'たくさん' });
  t.ctx.dailyRun();
  assert.match(band(t), /\[E27\].*振込手数料として見る差/);
  assert.ok(t.env.mails.some((m) => m.to === 'admin@example.com'));
  setSettings(t, { fee: 880 });
  t.ctx.dailyRun();
  assert.equal(band(t), '');
});

