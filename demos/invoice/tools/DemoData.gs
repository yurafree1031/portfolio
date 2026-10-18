/**
 * デモ用の架空のデータを入れる（開発用。デモのスプレッドシートを整えたら、このファイルは消す。お客様には渡さない）
 * 会社・取引先・口座はすべて架空。登録番号は、検査数字が合わない番号（実在しえない）にしている
 */

/** 1. 設定・取引先・毎月の請求・明細を入れる。フォルダは「請求書デモ」の下に2つ作る */
function loadDemoData() {
  // 何度でも入れ直せるよう、ツールのシートと記録を作り直す（フォルダは、あれば使い回す）
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TIME_ZONE); // clasp で作ったシートは日本時間でないことがある（お客様のシートは、ツールが止めて案内する）
  if (!ss.getSheetByName('シート1')) ss.insertSheet('シート1');
  Object.values(SHEET).forEach((n) => { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
  PropertiesService.getScriptProperties().deleteAllProperties();
  ensureAllSheets_();
  const sub = (parent, name) => { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); };
  const root = sub(DriveApp.getRootFolder(), '請求書デモ');
  const pdf = sub(root, '請求書PDF');
  const csv = sub(root, '入金明細CSV');
  [pdf, csv].forEach((f) => { const it = f.getFiles(); while (it.hasNext()) it.next().setTrashed(true); }); // 前のテストのファイルはゴミ箱へ（戻せる）
  const values = {
    company: '株式会社サンプルWeb制作',
    address: '〒100-0000\n東京都千代田区サンプル1-2-3 見本ビル5F',
    tel: '03-0000-0000',
    regno: 'T1123456789012',
    bank: 'サンプル銀行 本店営業部 普通 1234567\nカ)サンプルウエブセイサク',
    remarks: '恐れ入りますが、振込手数料は貴社にてご負担ください。',
    pdfFolder: pdf.getUrl(),
    csvFolder: csv.getUrl(),
    admin: 'admin@example.com',
  };
  Object.entries(values).forEach(([k, v]) => writeSetting_(k, v));

  const clients = openTable_(SHEET.CLIENTS);
  const c = [
    ['株式会社サンプル商事', '山田 太郎', 'yamada@example.com', 'keiri@example.com', '末日', '翌月', '末日', 'カ)サンプルシヨウジ'],
    ['有限会社テスト工務店', '', 'office@example.net', '', '末日', '翌月', '末日', 'ユ)テストコウムテン'],
    ['合同会社デモ食品', '鈴木 花子', 'suzuki@example.org', '', '20日', '翌月', '10日', 'ド)デモシヨクヒン'],
    ['株式会社見本不動産', '', '', '', '末日', '翌々月', '10日', 'カ)ミホンフドウサン'],
  ];
  writeRows_(clients, c);

  const monthly = openTable_(SHEET.MONTHLY);
  writeRows_(monthly, [
    ['株式会社サンプル商事', 'Webサイト保守費（{月}分）', 1, 30000, '10%'],
    ['有限会社テスト工務店', 'Webサイト保守費（{月}分）', 1, 15000, '10%'],
    ['有限会社テスト工務店', 'サーバー管理費（{月}分）', 1, 5000, '10%'],
  ]);

  const lines = openTable_(SHEET.LINES);
  const d = (s) => toDate_(s);
  writeRows_(lines, [
    [d('2026/10/31'), '株式会社サンプル商事', '', '採用ページの追加制作', 1, 180000, '10%'],
    [d('2026/10/15'), '株式会社サンプル商事', '', '写真撮影（半日）', 1, 45000, '10%'],
    [d('2026/10/15'), '株式会社サンプル商事', '', '撮影時のお弁当', 6, 980, '8%'],
    [d('2026/10/20'), '合同会社デモ食品', '', '商品ページの更新', 12, 3500, '10%'],
    [d('2026/10/20'), '合同会社デモ食品', '', 'バナー制作', 3, 8000, '10%'],
    [d('2026/10/31'), '株式会社見本不動産', '', '物件検索ページの改修', 1, 350000, '10%'],
    [d('2026/10/31'), '株式会社見本不動産', '', '値引き', 1, -20000, '10%'],
  ]);
  prepareMonthly_(2026, 10);
  PropertiesService.getScriptProperties().setProperty(PROP.MONTHLY_DONE, '2026-10');
  ensureAllSheets_();
  Logger.log(`フォルダ: ${root.getUrl()}`);
}

function writeRows_(t, rows) {
  const start = t.lastRow() + 1;
  rows.forEach((r, i) => r.forEach((v, j) => t.write(start + i, t.spec.header[j], v)));
  for (let i = 0; i < rows.length; i++) t.paintRow(start + i, null);
}
