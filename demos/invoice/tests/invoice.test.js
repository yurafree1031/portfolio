// 「請求書を作る」ボタンと、「選んだ請求書を取消す」の流れ
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, setup, setSettings, addRows, rowsOf, colorOf, setNow } = require('./helper');

function ready() {
  const t = load();
  setup(t);
  return t;
}
const line = (o) => ({ 請求日: '2026/10/31', 取引先: '株式会社サンプル商事', 品目: '制作費', 数量: 1, '単価（円）': 100000, 税率: '10%', ...o });

test('開いたとき：シートが決まった並びでそろい、ボタン2つとメニューができる', () => {
  const t = load();
  t.ctx.onOpen();
  assert.deepEqual(t.ss.getSheets().map((s) => s.getName()).filter((n) => n !== 'シート1'),
    ['操作', '明細', '請求書一覧', '入金', '取引先', '毎月の請求', '設定']);
  assert.deepEqual(t.ss.getSheetByName('操作').buttons.map((b) => b.fn), ['createInvoices', 'importPayments']);
  assert.equal(t.ss.menus[0].title, '請求書');
  t.ctx.onOpen(); // 2回開いても、ボタンは増えない
  assert.equal(t.ss.getSheetByName('操作').buttons.length, 2);
  assert.equal(t.ss.getActiveSheet().getName(), 'シート1'); // 画面を勝手に動かさない
});

test('設定が空のまま押すと、どれが空かを出して止まる（設定シートを開く）', () => {
  const t = load();
  t.ctx.onOpen();
  t.ctx.createInvoices();
  const a = t.env.alerts.pop();
  assert.equal(a.title, '確認してください');
  assert.match(a.message, /「会社名」・「登録番号」・「振込先」・「請求書を保存するフォルダ」を書いてください/);
  assert.equal(t.ss.getActiveSheet().getName(), '設定');
});

test('スプレッドシートのタイムゾーンが日本時間でないと、日付がずれるので、直し方を出して止める', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ss.tz = 'America/Los_Angeles';
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /タイムゾーンが「America\/Los_Angeles」.*「\(GMT\+09:00\) 東京」にしてから/);
  t.ctx.importPayments();
  assert.match(t.env.alerts.pop().message, /タイムゾーン/);
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
});

test('登録番号の形が違うと止まる。全角・ハイフンはそろえて通す', () => {
  const t = ready();
  t.ctx.writeSetting_('regno', 'T123');
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /登録番号」は T と13けた/);
  t.ctx.writeSetting_('regno', 'Ｔ1123-4567-8901-2');
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
});

test('フォルダのURLの貼り間違い・開けないフォルダは、それぞれの直し方を出す', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.writeSetting_('pdfFolder', 'フォルダ');
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /フォルダを開いたときのアドレス（URL）を貼って/);
  t.ctx.writeSetting_('pdfFolder', 'https://drive.google.com/drive/folders/NOTSHARED00001');
  t.ctx.createInvoices();
  const a = t.env.alerts.pop();
  assert.equal(a.title, '管理者に連絡してください');
  assert.match(a.message, /\[E21\]/);
});

test('取引先と請求日が同じ明細を1枚にまとめ、PDF・一覧・明細の番号・下書きができる', () => {
  const t = ready();
  addRows(t, '明細', [
    line({ 品目: '採用ページ制作', '単価（円）': 180000 }),
    line({ 品目: '撮影時のお弁当', 数量: 6, '単価（円）': 980, 税率: '8%' }),
    line({ 取引先: '合同会社デモ食品', 請求日: '2026/10/20', 品目: '商品ページ更新', 数量: 12, '単価（円）': 3500 }),
  ]);
  t.ctx.createInvoices();
  const confirm = t.env.alerts.find((a) => a.title === '請求書を作りますか？');
  assert.match(confirm.message, /請求書を 2枚（合計 /);
  assert.match(confirm.message, /メールは送りません/);

  const list = rowsOf(t, '請求書一覧');
  assert.equal(list.length, 2);
  const a = list.find((r) => r['取引先'] === '株式会社サンプル商事');
  assert.equal(a['請求書番号'], '2026-10-001');
  assert.equal(a['税抜（円）'], 185880);
  assert.equal(a['消費税（円）'], 18000 + 470); // 10%：180000→18000、8%：5880→470
  assert.equal(a['請求額（円）'], 185880 + 18470);
  assert.equal(t.ctx.ymd_(a['支払期限']), '2026-11-30');
  assert.equal(a['状態'], '未入金');
  assert.match(a['PDF'], /^=HYPERLINK\(".+","請求書_2026-10-001_株式会社サンプル商事\.pdf"\)$/);
  assert.equal(list.find((r) => r['取引先'] === '合同会社デモ食品')['請求書番号'], '2026-10-002');
  assert.equal(t.ctx.ymd_(list.find((r) => r['取引先'] === '合同会社デモ食品')['支払期限']), '2026-11-10');

  assert.deepEqual(rowsOf(t, '明細').map((r) => r['請求書番号']), ['2026-10-001', '2026-10-001', '2026-10-002']);
  assert.equal(t.folders.pdf.files.length, 2);
  assert.equal(t.env.drafts.length, 2);
  const draft = t.env.drafts[0];
  assert.equal(draft.to, 'yamada@example.com');
  assert.equal(draft.subject, '請求書のご送付（2026年10月分）');
  assert.match(draft.body, /^株式会社サンプル商事\n山田 太郎 様/);
  assert.match(draft.body, /ご請求金額：204,350円（税込）/);
  assert.match(draft.body, /お支払期限：2026\/11\/30/);
  assert.equal(draft.opts.attachments[0].name, '請求書_2026-10-001_株式会社サンプル商事.pdf');
  assert.equal(t.env.mails.length, 0); // メールは送らない
  assert.match(t.ss.getSheetByName('操作').getRange('B9').getValue(), /請求書 2枚（合計 .+円）を作りました。下書き 2通/);
});

test('請求書のPDFの中身：インボイスの項目（登録番号・税率ごとの合計と消費税・軽減税率の印）と、宛名の「御中」', () => {
  const t = ready();
  addRows(t, '明細', [
    line({ 取引先: '有限会社テスト工務店', 品目: '保守費 <10月分>', '単価（円）': 15000 }),
    line({ 取引先: '有限会社テスト工務店', 品目: 'お茶', 数量: 2, '単価（円）': 150, 税率: '8%' }),
  ]);
  t.ctx.createInvoices();
  const html = t.folders.pdf.files[0].blob.data;
  assert.match(html, /有限会社テスト工務店 御中/);
  assert.match(html, /登録番号 T1123456789012/);
  assert.match(html, /10%対象（税抜）<\/td><td class="r">15,000円/);
  assert.match(html, /消費税（10%）<\/td><td class="r">1,500円/);
  assert.match(html, /8%対象（税抜）<\/td><td class="r">300円/);
  assert.match(html, /消費税（8%）<\/td><td class="r">24円/);
  assert.match(html, /お茶 ※/);
  assert.match(html, /※は軽減税率（8%）の対象です/);
  assert.match(html, /保守費 &lt;10月分&gt;/); // 品目の記号はそのまま文字として出す
  assert.match(html, /ご請求金額（税込）<\/td><td class="v">16,824 円/);
  assert.match(html, /テスト銀行 本店 普通 1234567/);
});

test('空の数量は1・空の税率は10%として入れ、「1,280円」のような単価は数字に直す', () => {
  const t = ready();
  addRows(t, '明細', [line({ 数量: '', 税率: '', '単価（円）': '1,280円' })]);
  t.ctx.createInvoices();
  const r = rowsOf(t, '明細')[0];
  assert.equal(r['数量'], 1);
  assert.equal(r['税率'], '10%');
  assert.equal(r['単価（円）'], 1280);
  assert.equal(rowsOf(t, '請求書一覧')[0]['請求額（円）'], 1408);
});

test('直すところのある明細は赤くして飛ばし、同じ請求書のほかの明細は黄色で待たせる。ほかの請求書は作る', () => {
  const t = ready();
  addRows(t, '明細', [
    line({ 品目: '制作費' }),
    line({ 品目: '追加作業', '単価（円）': '千円' }),
    line({ 取引先: '株式会社ない会社' }),
    line({ 取引先: '合同会社デモ食品', 請求日: '2026/10/20' }),
    line({ 請求日: '2026/02/31', 取引先: '有限会社テスト工務店' }),
  ]);
  t.ctx.createInvoices();
  const rows = rowsOf(t, '明細');
  assert.match(rows[1]['確認事項'], /単価は数字にしてください/);
  assert.equal(colorOf(t, '明細', rows[1]._row), t.ctx.ERROR_ROW_COLOR);
  assert.match(rows[0]['確認事項'], /同じ請求書.*直すところがあるため/);
  assert.equal(colorOf(t, '明細', rows[0]._row), t.ctx.WARN_ROW_COLOR);
  assert.match(rows[2]['確認事項'], /「取引先」シートにない取引先/);
  assert.match(rows[4]['確認事項'], /請求日を 2026\/10\/31 の形/);
  assert.equal(rows[3]['請求書番号'], '2026-10-001');
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
  assert.match(t.env.alerts.find((a) => a.title === '請求書を作りますか？').message, /赤い行（3行）/);

  // 直してもう一度押すと、色が戻って作られる
  const tb = t.ctx.openTable_('明細');
  tb.write(rows[1]._row, '単価（円）', 1000);
  t.ctx.createInvoices();
  const after = rowsOf(t, '明細');
  assert.equal(after[0]['請求書番号'], '2026-10-002');
  assert.equal(after[0]['確認事項'], '');
  assert.equal(colorOf(t, '明細', after[0]._row), null);
});

test('取引先の支払条件が空なら、その取引先の明細を赤くする', () => {
  const t = load();
  setup(t, { clients: [{ 取引先名: 'X社', メールアドレス: 'x@example.com', 締め日: '末日', 支払月: '', 支払日: '末日' }] });
  addRows(t, '明細', [line({ 取引先: 'X社' })]);
  t.ctx.createInvoices();
  assert.match(rowsOf(t, '明細')[0]['確認事項'], /「X社」の支払月を選んでください/);
});

test('合計が0円以下の請求書は作らない（値引きが大きすぎる）', () => {
  const t = ready();
  addRows(t, '明細', [line({ '単価（円）': 10000 }), line({ 品目: '値引き', '単価（円）': -20000 })]);
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
  assert.match(rowsOf(t, '明細')[0]['確認事項'], /合計が0円以下/);
});

test('作る明細がないときは、作らずにそう伝える。「いいえ」を押したら何も作らない', () => {
  const t = ready();
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /作る請求書がありません/);
  addRows(t, '明細', [line()]);
  t.env.uiAnswers.push('NO');
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
  assert.equal(rowsOf(t, '明細')[0]['請求書番号'], '');
});

test('2回押しても、作成済みの明細でもう1枚作らない。番号は続きから', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
  addRows(t, '明細', [line({ 請求日: '2026/10/15' }), line({ 請求日: '2026/11/30' })]);
  t.ctx.createInvoices();
  assert.deepEqual(rowsOf(t, '請求書一覧').map((r) => r['請求書番号']).sort(), ['2026-10-001', '2026-10-002', '2026-11-001']);
});

test('一覧の行を消しても、番号を使い回さない', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  const tb = t.ctx.openTable_('請求書一覧');
  tb.sheet.deleteRow(tb.headerRow + 1);
  addRows(t, '明細', [line({ 請求日: '2026/10/15' })]);
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧')[0]['請求書番号'], '2026-10-002');
});

test('実行中にもう一度押されたら、待ってもらう（同じ明細で2枚作らない）', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.env.scriptLockBusy = true;
  t.ctx.createInvoices();
  assert.match(t.ss.lastToast.message, /終わってから、もう一度押してください/);
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
});

test('メールアドレスがない取引先は、下書きを作らず一覧を黄色に。下書きが作れなくても請求書はできる', () => {
  const t = ready();
  addRows(t, '明細', [line({ 取引先: '株式会社見本不動産' }), line({ 取引先: '有限会社テスト工務店' })]);
  t.env.draftError = new Error('Gmail error');
  t.ctx.createInvoices();
  const list = rowsOf(t, '請求書一覧');
  assert.equal(list.length, 2);
  assert.match(list.find((r) => r['取引先'] === '株式会社見本不動産')['確認事項'], /メールアドレスがないため/);
  assert.match(list.find((r) => r['取引先'] === '有限会社テスト工務店')['確認事項'], /下書きを作れませんでした/);
  assert.equal(colorOf(t, '請求書一覧', list[0]._row), t.ctx.WARN_ROW_COLOR);
});

test('PDFを保存できないときは、そこで止めて、それまでに作った枚数を出す（PDFのない請求書を一覧に残さない）', () => {
  const t = ready();
  addRows(t, '明細', [line(), line({ 取引先: '有限会社テスト工務店' })]);
  let n = 0;
  const orig = t.folders.pdf.createFile;
  t.folders.pdf.createFile = (b) => { if (++n === 2) throw new Error('Access denied'); return orig(b); };
  t.ctx.createInvoices();
  const a = t.env.alerts.pop();
  assert.match(a.message, /\[E23\]/);
  assert.match(a.message, /ここまでに 1枚 を作りました/);
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
  assert.equal(rowsOf(t, '明細').filter((r) => r['請求書番号']).length, 1);
});

test('時間が足りないときは途中で止め、「残り ○枚」を出す。もう一度押すと続きを作る', () => {
  const t = ready();
  addRows(t, '明細', [line(), line({ 取引先: '有限会社テスト工務店' }), line({ 取引先: '合同会社デモ食品' })]);
  const orig = t.folders.pdf.createFile;
  t.folders.pdf.createFile = (b) => { t.ctx.CONFIG.TIME_BUDGET_MS = -1; return orig(b); }; // 1枚作ったら時間切れ
  t.ctx.createInvoices();
  t.folders.pdf.createFile = orig;
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
  assert.match(t.ss.getSheetByName('操作').getRange('B9').getValue(), /残り 2枚 は、もう一度押してください/);
  t.ctx.CONFIG.TIME_BUDGET_MS = 270000;
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 3);
});

test('PDFを作っている間に明細が並べ替えられても、番号は正しい行に書く', () => {
  const t = ready();
  addRows(t, '明細', [line({ 品目: 'A' }), line({ 取引先: '有限会社テスト工務店', 品目: 'B' })]);
  const tb = t.ctx.openTable_('明細');
  const orig = t.folders.pdf.createFile;
  let swapped = false;
  t.folders.pdf.createFile = (b) => {
    if (!swapped) { // 1枚目のPDFを作っている間に、2行を入れ替える
      swapped = true;
      const r1 = tb.headerRow + 1;
      const a = tb.read(r1);
      const b2 = tb.read(r1 + 1);
      tb.spec.header.forEach((h) => { tb.write(r1, h, b2[h]); tb.write(r1 + 1, h, a[h]); });
    }
    return orig(b);
  };
  t.ctx.createInvoices();
  const rows = rowsOf(t, '明細');
  const no = (item) => rows.find((r) => r['品目'] === item)['請求書番号'];
  assert.equal(no('A'), '2026-10-001');
  assert.equal(no('B'), '2026-10-002');
});

test('取消：一覧は「取消」と理由、明細は番号を消して黄色、PDFは名前に「取消」、下書きは消す。作り直すと訂正版になる', () => {
  const t = ready();
  addRows(t, '明細', [line(), line({ 品目: '追加' })]);
  t.ctx.createInvoices();
  const sheet = t.ss.getSheetByName('請求書一覧');
  sheet.activate();
  t.ss.activeRow = rowsOf(t, '請求書一覧')[0]._row;
  t.env.promptAnswers.push({ button: 'OK', text: '金額の誤り' });
  t.ctx.cancelSelectedInvoice();
  const done = t.env.alerts.pop();
  assert.equal(done.title, '2026-10-001 を取消しました');
  assert.match(done.message, /明細 2行 を作り直せる状態に戻しました/);
  assert.match(done.message, /Gmailの下書きを消しました/);
  const inv = rowsOf(t, '請求書一覧')[0];
  assert.equal(inv['状態'], '取消');
  assert.match(inv['取消・訂正'], /^取消：金額の誤り（\d{4}\/\d{2}\/\d{2}）$/);
  assert.match(inv['PDF'], /【取消】請求書_2026-10-001/);
  assert.match(t.folders.pdf.files[0].name, /^【取消】/);
  assert.deepEqual(t.env.deletedDrafts, ['DRAFT1']);
  const lines = rowsOf(t, '明細');
  assert.ok(lines.every((r) => r['請求書番号'] === '' && /取消した請求書 2026-10-001 の明細です/.test(r['確認事項'])));

  // 直して作り直す
  t.ctx.openTable_('明細').write(lines[0]._row, '単価（円）', 90000);
  t.ctx.createInvoices();
  const list = rowsOf(t, '請求書一覧');
  assert.equal(list[1]['請求書番号'], '2026-10-002');
  assert.equal(list[1]['取消・訂正'], '2026-10-001 の訂正版');
  assert.match(list[0]['取消・訂正'], /→ 2026-10-002 で作り直し$/);
  assert.match(t.folders.pdf.files[1].blob.data, /2026-10-001 の訂正版です/);
  assert.ok(rowsOf(t, '明細').every((r) => r['確認事項'] === ''));
});

test('取消：入金済み・取消済み・一覧以外のシートでは取消さない。理由が空なら取消さない。送ったあとの下書きは見つからないと伝える', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  t.ss.getSheetByName('明細').activate();
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /「請求書一覧」で、取消す請求書の行/);

  const tb = t.ctx.openTable_('請求書一覧');
  tb.sheet.activate();
  const row = rowsOf(t, '請求書一覧')[0]._row;
  t.ss.activeRow = row;
  tb.write(row, '状態', '入金済み');
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /入金済みです/);
  tb.write(row, '状態', '未入金');

  t.env.promptAnswers.push({ button: 'OK', text: '  ' });
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /取消す理由を書いてください/);
  assert.equal(rowsOf(t, '請求書一覧')[0]['状態'], '未入金');

  t.env.drafts[0].sent = true; // お客様に送ったあと
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /下書きは見つかりませんでした.*取引先に取消の連絡/);
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /もう取消してあります/);
});

test('取消：黄色い行の文を消しても、作り直すと訂正版としてつながる', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  t.ctx.cancelInvoice_('2026-10-001', '金額の誤り', true);
  const tb = t.ctx.openTable_('明細');
  const row = rowsOf(t, '明細')[0]._row;
  tb.write(row, '確認事項', ''); // 「見かた」のとおりに、確かめて文を消した
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['確認事項']) });
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧')[1]['取消・訂正'], '2026-10-001 の訂正版');
  assert.equal(rowsOf(t, '明細')[0]['管理用'], '');
});

test('取消：作り直さないを選ぶと、明細に「（取消）」を付けて作る対象から外す', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  const sheet = t.ss.getSheetByName('請求書一覧');
  sheet.activate();
  t.ss.activeRow = rowsOf(t, '請求書一覧')[0]._row;
  t.env.promptAnswers.push({ button: 'OK', text: '取引がなくなった' });
  t.env.uiAnswers.push('NO'); // 作り直さない
  t.ctx.cancelSelectedInvoice();
  assert.match(t.env.alerts.pop().message, /明細 1行 に「取消」の印を付けました/);
  const l = rowsOf(t, '明細')[0];
  assert.equal(l['請求書番号'], '2026-10-001（取消）');
  assert.equal(colorOf(t, '明細', l._row), null);
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /作る請求書がありません/);
  assert.equal(rowsOf(t, '請求書一覧').length, 1);
});

test('取消：明細を作り直すかの画面で「キャンセル」なら、何も変えない', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  t.ss.getSheetByName('請求書一覧').activate();
  t.ss.activeRow = rowsOf(t, '請求書一覧')[0]._row;
  t.env.uiAnswers.push('CANCEL');
  t.ctx.cancelSelectedInvoice();
  assert.equal(rowsOf(t, '請求書一覧')[0]['状態'], '未入金');
});

test('PDF：請求日が締め日なら対象期間（前の締め日の翌日〜締め日）、それ以外の日は取引日を載せる', () => {
  const t = ready();
  addRows(t, '明細', [line(), line({ 請求日: '2026/10/15' }), line({ 取引先: '合同会社デモ食品', 請求日: '2026/10/20' })]);
  t.ctx.createInvoices();
  const html = (i) => t.folders.pdf.files[i].blob.data;
  assert.match(html(0), /対象期間<\/td><td>2026\/10\/01〜2026\/10\/31/);
  assert.match(html(1), /取引日<\/td><td>2026\/10\/15/);
  assert.match(html(2), /対象期間<\/td><td>2026\/09\/21〜2026\/10\/20/); // 20日締め
});

test('「状態」のセルの色：未入金は薄い赤、入金済みは薄い緑（行全体の黄色と見分けられるように）', () => {
  const t = ready();
  const rules = t.ss.getSheetByName('請求書一覧').getConditionalFormatRules();
  assert.equal(rules.find((r) => r.text === '未入金').bg, t.ctx.ERROR_ROW_COLOR);
  assert.equal(rules.find((r) => r.text === '入金済み').bg, t.ctx.DONE_COLOR);
});

test('CC：取引先のCCを下書きに入れる。形の違うアドレスは外して、一覧を黄色に', () => {
  const t = ready();
  const ct = t.ctx.openTable_('取引先');
  const r = rowsOf(t, '取引先')[0]._row;
  ct.write(r, 'CCのメールアドレス', 'keiri@example.com, 部長');
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  assert.equal(t.env.drafts[0].opts.cc, 'keiri@example.com');
  assert.match(rowsOf(t, '請求書一覧')[0]['確認事項'], /CCの「部長」がメールアドレスの形でないため/);
});

test('確認画面を長く開いたままにしたら、1枚も作らずに「もう一度押して」と出す（6分の上限で途中で止まらないように）', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.CONFIG.TIME_BUDGET_MS = -1;
  t.ctx.createInvoices();
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
  assert.match(t.ss.getSheetByName('操作').getRange('B9').getValue(), /残り 1枚 は、もう一度押してください/);
});

// ---- 支払期限（取引先の条件で自動で入り、契約で違えば書き換えられる） ----
const editLine = (t, row, name, value) => {
  const tb = t.ctx.openTable_('明細');
  tb.write(row, name, value);
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map[name]) });
  return rowsOf(t, '明細').find((r) => r._row === row);
};

test('取引先を入れると、請求日（次の締め日）と支払期限が入る。取引先を変えると両方入れ直す', () => {
  const t = ready();
  setNow(t, '2026-10-18T10:00:00+09:00');
  const row = t.ctx.openTable_('明細').headerRow + 1;
  let r = editLine(t, row, '取引先', '株式会社サンプル商事');
  assert.equal(t.ctx.ymd_(r['請求日']), '2026-10-31');
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-11-30');
  r = editLine(t, row, '取引先', '合同会社デモ食品'); // 20日締め・翌月10日払い
  assert.equal(t.ctx.ymd_(r['請求日']), '2026-10-20');
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-11-10');
});

test('請求日を手で書き換えると、支払期限はその日から入れ直し、請求日は取引先を変えても変えない', () => {
  const t = ready();
  setNow(t, '2026-10-18T10:00:00+09:00');
  const row = t.ctx.openTable_('明細').headerRow + 1;
  editLine(t, row, '取引先', '株式会社サンプル商事');
  let r = editLine(t, row, '請求日', t.ctx.toDate_('2026/11/15'));
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-12-31');
  r = editLine(t, row, '取引先', '有限会社テスト工務店');
  assert.equal(t.ctx.ymd_(r['請求日']), '2026-11-15');
  r = editLine(t, row, '請求日', ''); // 空にすると、次の締め日に戻る
  assert.equal(t.ctx.ymd_(r['請求日']), '2026-10-31');
});

test('請求日が空の明細は、請求書を作るときに次の締め日を入れて作る', () => {
  const t = ready();
  setNow(t, '2026-10-18T10:00:00+09:00');
  addRows(t, '明細', [line({ 請求日: '' })]);
  t.ctx.createInvoices();
  const inv = rowsOf(t, '請求書一覧')[0];
  assert.equal(t.ctx.ymd_(inv['請求日']), '2026-10-31');
  assert.equal(t.ctx.ymd_(rowsOf(t, '明細')[0]['請求日']), '2026-10-31');
});
test('支払期限：手で書き換えた日付は、請求日を変えても上書きしない。空にすると自動に戻る', () => {
  const t = ready();
  const row = t.ctx.openTable_('明細').headerRow + 1;
  editLine(t, row, '請求日', t.ctx.toDate_('2026/10/31'));
  editLine(t, row, '取引先', '株式会社サンプル商事');
  let r = editLine(t, row, '支払期限', t.ctx.toDate_('2026/11/10')); // 契約で違う期限
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-11-10');
  r = editLine(t, row, '請求日', t.ctx.toDate_('2026/10/30'));
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-11-10');
  r = editLine(t, row, '支払期限', '');
  assert.equal(t.ctx.ymd_(r['支払期限']), '2026-11-30');
});

test('支払期限：書き換えた期限で請求書を作る。同じ取引先・請求日でも、期限が違えば別の請求書', () => {
  const t = ready();
  addRows(t, '明細', [
    line({ 品目: '保守費' }),
    line({ 品目: '制作費（着手金）', 支払期限: '2026/11/10' }),
  ]);
  t.ctx.createInvoices();
  const list = rowsOf(t, '請求書一覧');
  assert.equal(list.length, 2);
  assert.deepEqual(list.map((r) => t.ctx.ymd_(r['支払期限'])).sort(), ['2026-11-10', '2026-11-30']);
  assert.match(t.folders.pdf.files.find((f) => /2026\/11\/10/.test(f.blob.data)).blob.data, /制作費（着手金）/);
  const lines = rowsOf(t, '明細');
  assert.equal(t.ctx.ymd_(lines[0]['支払期限']), '2026-11-30'); // 空だった行には、自動の期限が入る
});

test('支払期限：請求日より前・日付でない期限は、赤い行で止める', () => {
  const t = ready();
  addRows(t, '明細', [line({ 支払期限: '2026/10/01' }), line({ 取引先: '合同会社デモ食品', 請求日: '2026/10/20', 支払期限: '来月' })]);
  t.ctx.createInvoices();
  const rows = rowsOf(t, '明細');
  assert.match(rows[0]['確認事項'], /支払期限が請求日より前です/);
  assert.match(rows[1]['確認事項'], /支払期限を 2026\/11\/30 の形で/);
  assert.equal(rowsOf(t, '請求書一覧').length, 0);
});

test('支払期限：取引先の条件を変えたら、自動で入った期限は作るときに入れ直す（手で書いた期限はそのまま）', () => {
  const t = ready();
  const row = t.ctx.openTable_('明細').headerRow + 1;
  editLine(t, row, '請求日', t.ctx.toDate_('2026/10/31'));
  editLine(t, row, '取引先', '株式会社サンプル商事');
  editLine(t, row, '品目', '制作費');
  editLine(t, row, '単価（円）', 1000);
  addRows(t, '明細', [line({ 取引先: '株式会社サンプル商事', 品目: '別契約', 支払期限: '2026/12/15' })]);
  const ct = t.ctx.openTable_('取引先');
  ct.write(rowsOf(t, '取引先')[0]._row, '支払月', '翌々月');
  t.ctx.createInvoices();
  const dues = rowsOf(t, '請求書一覧').map((r) => t.ctx.ymd_(r['支払期限'])).sort();
  assert.deepEqual(dues, ['2026-12-15', '2026-12-31']);
});

test('「請求書一覧」で「入金済み」にすると、空の入金日・入金額を埋める', () => {
  const t = ready();
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  const tb = t.ctx.openTable_('請求書一覧');
  const row = rowsOf(t, '請求書一覧')[0]._row;
  tb.write(row, '状態', '入金済み');
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['状態']) });
  const r = rowsOf(t, '請求書一覧')[0];
  assert.ok(r['入金日'] instanceof t.ctx.VMDate);
  assert.equal(r['入金額（円）'], 110000);
});

test('「確認事項」を消すと、行の色が戻る（入力する列は薄い青）', () => {
  const t = ready();
  addRows(t, '明細', [line({ 取引先: '株式会社ない会社' })]);
  t.ctx.createInvoices();
  const tb = t.ctx.openTable_('明細');
  const row = tb.headerRow + 1;
  tb.write(row, '確認事項', '');
  t.ctx.onEdit({ range: tb.sheet.getRange(row, tb.map['確認事項']) });
  assert.equal(tb.sheet.bgOf(row, tb.map['取引先']), t.ctx.INPUT_COLOR);
  assert.equal(tb.sheet.bgOf(row, tb.map['確認事項']), null);
});

test('取引先を足すと、明細の取引先のプルダウンが入れ直される', () => {
  const t = ready();
  addRows(t, '取引先', [{ 取引先名: '新しい会社', 締め日: '末日', 支払月: '翌月', 支払日: '末日' }]);
  const tb = t.ctx.openTable_('取引先');
  t.ctx.onEdit({ range: tb.sheet.getRange(tb.lastRow(), 1) });
  const lines = t.ss.getSheetByName('明細');
  const rule = lines.validations.filter((v) => v.rule.list && v.rule.list.includes('株式会社サンプル商事')).pop();
  assert.ok(rule.rule.list.includes('新しい会社'));
});

test('見出しの「品目」を消されたら、推測で書かずに E24 で止める', () => {
  const t = ready();
  const tb = t.ctx.openTable_('明細');
  tb.sheet.set(tb.headerRow, tb.map['品目'], '');
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /\[E24\] 「明細」シートの見出しの行が見つかりません/);
});

test('見出しの列（確認事項など）を消されたら、右に足し直す', () => {
  const t = ready();
  const tb = t.ctx.openTable_('明細');
  tb.sheet.set(tb.headerRow, tb.map['確認事項'], '');
  const again = t.ctx.openTable_('明細');
  assert.ok(again.map['確認事項'] > tb.map['確認事項'] - 1);
  assert.equal(tb.sheet.get(again.headerRow, again.map['確認事項']), '確認事項');
});

test('設定の消された項目は、下に足し直す。消された設定以外は初期値に戻さない', () => {
  const t = ready();
  const { sheet, rows } = t.ctx.ensureSettingsSheet_();
  sheet.set(rows['振込手数料として見る差（円）'], 1, '');
  sheet.set(rows['振込手数料として見る差（円）'], 2, '');
  t.ctx.ensureSettingsSheet_();
  const s = t.ctx.readSettings_();
  assert.equal(s.fee, 880);
  assert.equal(s.company, '株式会社テスト制作');
});

test('設定の数字が正しくないと止まる', () => {
  const t = ready();
  setSettings(t, { fee: 'たくさん' });
  addRows(t, '明細', [line()]);
  t.ctx.createInvoices();
  assert.match(t.env.alerts.pop().message, /振込手数料として見る差（円）」は0以上の数字/);
});
