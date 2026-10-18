/**
 * 請求書を作り、入金を照らし合わせる
 *
 * 流れ:
 *   - 「明細」に請求する内容を入れて「請求書を作る」ボタン（Invoice.gs）
 *     → 取引先と請求日が同じ明細を1枚にまとめ、インボイスの項目がそろった請求書のPDFを、決めたフォルダに保存
 *     → 「請求書一覧」に1行残す → PDFを付けたGmailの下書きを作る（送るのは人）
 *   - 「入金を取り込む」ボタン（Payment.gs）
 *     → 決めたフォルダのネットバンクのCSVを読み、「入金」に並べる → 振込名義と金額で請求書を探し、合えば「入金済み」
 *   - 毎朝の自動の処理（Monthly.gs。管理者が「自動の処理を始める」で入れる）
 *     → 月の初めに「毎月の請求」の明細を「明細」に入れる
 *
 * シートは開いたときにそろう（なければ作る）。フォルダは会社が今使っているものを「設定」で指定する（作らない）
 * エラーの種類（UserError・AdminError）と errorText_ は Errors.gs、シートの部品は Sheet.gs、管理者への知らせは Alert.gs（lib/gas/ のコピー）
 */

// このデモの AdminError の番号（README の「エラー番号の一覧」と合わせる。E01〜E05 はAIの部品で使う番号なので避ける）
//   E21 請求書を保存するフォルダを開けない（消された・共有されていない）
//   E22 入金明細のCSVを置くフォルダを開けない
//   E23 請求書のPDFを保存できない（フォルダの編集権限がない、ドライブの容量がいっぱい）
//   E24 シートの見出しの行が見つからない
//   E27 自動の処理で読んだ「設定」「取引先」の値が正しくない
//   E28 スプレッドシートのタイムゾーンが日本時間でない（自動の処理で見つけた。ボタンのときは UserError で止める）

const CONFIG = {
  MENU_NAME: '請求書',
  MAIL_SENDER_NAME: '請求書ツール',
  // GASは1回6分で強制終了する。この時間を過ぎたら次の請求書に進まずに止める（残りは次に押したとき）
  // 1枚にかかるのはPDFの作成と保存・下書きで数秒〜十数秒。最後の1枚が長引いても6分に収まるよう、90秒残す
  TIME_BUDGET_MS: 270 * 1000,
  LOCK_WAIT_MS: 1000,
  DAILY_HOUR: 8, // 毎朝の自動の処理の時刻（この時刻から1時間のうちに動く）
  CONFIRM_LIST_MAX: 10, // 「請求書を作りますか？」の画面に並べる枚数（多いときは「ほか○枚」）
  MAX_COMBINE: 12, // まとめた入金を探すときに見る、1社の未入金の請求書の数（組み合わせが多くなりすぎないように）
};

const TIME_ZONE = 'Asia/Tokyo'; // 日付はすべて日本時間で扱う（プロジェクトの設定に左右されないように）
const JST_OFFSET_MS = 9 * 60 * 60 * 1000; // 日本時間は夏時間がないので、9時間ずらすだけで日付を決められる
const DAY_MS = 24 * 60 * 60 * 1000;

// 色は「スプレッドシートの見た目ルール」に合わせる。1つの色に意味は1つ
const ERROR_ROW_COLOR = '#f4cccc'; // 直す必要がある行（薄い赤）
const WARN_ROW_COLOR = '#fff2cc'; // 確かめてほしい行（薄い黄）
const DONE_COLOR = '#d9ead3'; // 「状態」のセルだけ：済んだ（入金済み）
const HEADER_COLOR = '#f3f3f3'; // 表の見出し（薄いグレー）
const INPUT_COLOR = '#e8f0fe'; // 使う人が入力・選択するセル（薄い青）
const NOTE_COLOR = '#666666'; // 補足の文字
const PROTECT_NOTE = '請求書ツールの見出し'; // 見出しの行に付ける保護（消そうとすると警告が出る）の名前

const SHEET = {
  RUN: '操作',
  LINES: '明細',
  INVOICES: '請求書一覧',
  PAYMENTS: '入金',
  CLIENTS: '取引先',
  MONTHLY: '毎月の請求',
  SETTINGS: '設定',
};

const TAX_RATES = ['10%', '8%'];
const CLOSING_DAYS = ['5日', '10日', '15日', '20日', '25日', '末日'];
const PAY_MONTHS = ['当月', '翌月', '翌々月'];
const ROUNDINGS = ['切り捨て', '四捨五入', '切り上げ'];

const STATUS = { UNPAID: '未入金', PAID: '入金済み', CANCELED: '取消' };
const RESULT = { DONE: '入金済みにしました', WAIT: '確認待ち', NONE: '請求書の支払いではない' };

/**
 * 表のシートの決まり
 *   header: 見出し（この順で作る。使う人が並べ替えても、名前で探す）
 *   key: 見出しの行を探すときに使う名前（このシートの中で、ほかのセルと重ならないもの）
 *   headerRow: 作るときの見出しの行（上に説明・帯を書く）
 *   help: 上の行に書く説明（毎回書き直す）
 *   alertRow: 自動の処理のエラーを書く帯の行（ないシートは 0）
 *   inputs: 使う人が入れる列（薄い青）。inputRow(行の値) が false の行は塗らない
 *   widths・wrap・numberFormats・right: 見た目
 *   hidden: ツールが使う隠し列
 */
const TABLES = {
  [SHEET.LINES]: {
    header: ['請求日', '取引先', '支払期限', '品目', '数量', '単価（円）', '税率', '請求書番号', '確認事項', '管理用'],
    key: '品目',
    headerRow: 6,
    alertRow: 5,
    help: [
      '請求する内容を1行に1品目ずつ入れます。取引先・請求日・支払期限が同じ行が、1枚の請求書になります。',
      '支払期限は、取引先と請求日を入れると、取引先の条件から自動で入ります。契約で違うときは、日付を書き換えてください。',
      '単価は税抜です。数量が空なら1、税率が空なら10%にします。値引きは単価をマイナスにします。',
      '「請求書番号」はツールが書きます（番号のある行は、請求書を作成済み）。',
    ],
    inputs: ['請求日', '取引先', '支払期限', '品目', '数量', '単価（円）', '税率'],
    inputRow: (row) => !row['請求書番号'],
    widths: { 請求日: 100, 取引先: 220, 支払期限: 100, 品目: 260, 数量: 70, '単価（円）': 110, 税率: 70, 請求書番号: 110, 確認事項: 380 },
    wrap: ['品目', '確認事項'],
    numberFormats: { 請求日: 'yyyy/mm/dd', 支払期限: 'yyyy/mm/dd', 数量: '#,##0.##', '単価（円）': '#,##0', 税率: '@' }, // 税率は文字のまま（「10%」が数の0.1に変わらないように）
    right: ['数量', '単価（円）', '税率'],
    hidden: ['管理用'], // 取消した請求書の番号（作り直したときに「訂正版」とつなぐ）・毎月の請求から入れた印・自動で入れた支払期限
  },
  [SHEET.INVOICES]: {
    header: ['請求書番号', '請求日', '取引先', '税抜（円）', '消費税（円）', '請求額（円）', '支払期限', '状態', '入金日', '入金額（円）', 'PDF', '取消・訂正', 'メモ', '確認事項', '管理用'],
    key: '請求書番号',
    headerRow: 3,
    alertRow: 0,
    help: [
      '「請求書を作る」で1枚ごとに1行増えます。入金されたら「状態」を「入金済み」にします（「入金を取り込む」でも変わります）。',
      '直すときは、その行を選んでメニュー「請求書 → 選んだ請求書を取消す」を押します。',
    ],
    inputs: ['状態', '入金日', '入金額（円）', 'メモ'],
    inputRow: (row) => row['状態'] !== STATUS.CANCELED,
    restInput: false,
    lastRowKey: '請求書番号',
    widths: { 請求書番号: 110, 請求日: 100, 取引先: 220, '税抜（円）': 100, '消費税（円）': 100, '請求額（円）': 110, 支払期限: 100, 状態: 90, 入金日: 100, '入金額（円）': 110, PDF: 280, '取消・訂正': 260, メモ: 200, 確認事項: 320 },
    wrap: ['取消・訂正', 'メモ', '確認事項'],
    numberFormats: { 請求日: 'yyyy/mm/dd', 支払期限: 'yyyy/mm/dd', 入金日: 'yyyy/mm/dd', '税抜（円）': '#,##0', '消費税（円）': '#,##0', '請求額（円）': '#,##0', '入金額（円）': '#,##0' },
    right: ['税抜（円）', '消費税（円）', '請求額（円）', '入金額（円）'],
    hidden: ['管理用'],
  },
  [SHEET.PAYMENTS]: {
    header: ['入金日', '振込名義', '入金額（円）', '請求書番号', '結果', '確認', '確認事項', '管理用'],
    key: '振込名義',
    headerRow: 5,
    alertRow: 0,
    help: [
      '「入金を取り込む」で、ネットバンクのCSVの入金がここに並びます。黄色い行は「確認事項」を読んでください。',
      '「請求書番号」が合っていれば（違えば直して）「確認」にチェックを入れると、請求書が「入金済み」になります。',
      '2枚分なら番号をカンマで区切ります。請求書の支払いでない入金（利息など）は、番号を空にしてチェックします。',
      '間違えたときは、チェックを外すと取り消せます。',
    ],
    inputs: ['請求書番号', '確認'],
    inputRow: (row) => row['結果'] === RESULT.WAIT,
    restInput: false,
    lastRowKey: '入金額（円）',
    widths: { 入金日: 100, 振込名義: 220, '入金額（円）': 110, 請求書番号: 180, 結果: 150, 確認: 60, 確認事項: 440 },
    wrap: ['確認事項'],
    numberFormats: { 入金日: 'yyyy/mm/dd', '入金額（円）': '#,##0' },
    right: ['入金額（円）'],
    hidden: ['管理用'],
  },
  [SHEET.CLIENTS]: {
    header: ['取引先名', '先方の担当者', 'メールアドレス', 'CCのメールアドレス', '締め日', '支払月', '支払日', '振込名義（カナ）', '確認事項'],
    key: '取引先名',
    headerRow: 4,
    alertRow: 0,
    help: [
      '取引先を1行に1社。「締め日」「支払月」「支払日」から支払期限を決めます（月末締め翌月末払い＝末日・翌月・末日）。',
      '「先方の担当者」が空なら、宛名は「御中」にします。「CC」は下書きにCCで入れる人（2人以上はカンマで区切る）。',
      '「振込名義（カナ）」は通帳に出る名前（例：カ)サンプルシヨウジ）。入金の照らし合わせに使います。',
    ],
    inputs: ['取引先名', '先方の担当者', 'メールアドレス', 'CCのメールアドレス', '締め日', '支払月', '支払日', '振込名義（カナ）'],
    inputRow: () => true,
    widths: { 取引先名: 220, 先方の担当者: 140, メールアドレス: 220, CCのメールアドレス: 240, 締め日: 70, 支払月: 70, 支払日: 70, '振込名義（カナ）': 200, 確認事項: 300 },
    wrap: ['CCのメールアドレス', '確認事項'],
    numberFormats: {},
    right: [],
  },
  [SHEET.MONTHLY]: {
    header: ['取引先', '品目', '数量', '単価（円）', '税率', '確認事項'],
    key: '品目',
    headerRow: 3,
    alertRow: 0,
    help: [
      '毎月同じ内容で請求する明細（保守費など）。1社に何行でも書けます。品目の {月} には「10月」のように月が入ります。',
      '毎月1日の朝に、その月の締め日を請求日にして「明細」に入ります。請求書を作るのは「請求書を作る」ボタンです。',
    ],
    inputs: ['取引先', '品目', '数量', '単価（円）', '税率'],
    inputRow: () => true,
    widths: { 取引先: 220, 品目: 260, 数量: 70, '単価（円）': 110, 税率: 70, 確認事項: 320 },
    wrap: ['品目', '確認事項'],
    numberFormats: { 数量: '#,##0.##', '単価（円）': '#,##0', 税率: '@' },
    right: ['数量', '単価（円）', '税率'],
  },
};

const DEFAULT_BODY = [
  '{宛名}',
  '',
  'いつもお世話になっております。{会社名}です。',
  '{請求月}分の請求書をお送りします。',
  '',
  '請求書番号：{請求書番号}',
  'ご請求金額：{請求額}円（税込）',
  'お支払期限：{支払期限}',
  '',
  'お手数ですが、期限までに下記の口座へお振り込みをお願いいたします。',
  '{振込先}',
  '',
  'ご不明な点がございましたら、お気軽にお問い合わせください。',
].join('\n');

// 「設定」シートの行（A列の名前で探す。消されたら、下に足し直す）
const SETTINGS = [
  { key: 'company', label: '会社名', note: '請求書に載せる自社の名前', required: true },
  { key: 'address', label: '住所', note: '請求書に載せる。郵便番号から改行して書ける（Alt＋Enter で改行）' },
  { key: 'tel', label: '電話番号', note: '請求書に載せる' },
  { key: 'regno', label: '登録番号', note: '適格請求書発行事業者の登録番号（T＋13けた）', required: true },
  { key: 'bank', label: '振込先', note: '請求書とメールに載せる。例：〇〇銀行 △△支店 普通 1234567 カ)〇〇', required: true },
  { key: 'remarks', label: '請求書の備考', note: '請求書の下に載せる文（なければ空）' },
  { key: 'pdfFolder', label: '請求書を保存するフォルダ', note: '今使っているGoogleドライブのフォルダを開き、上のアドレス（URL）を貼る', required: true },
  { key: 'csvFolder', label: '入金明細のCSVを置くフォルダ', note: 'ネットバンクからダウンロードしたCSVを入れるフォルダのURL。請求書のフォルダと別にする' },
  { key: 'csvDate', label: 'CSVの入金日の列', note: 'CSVの1行目（見出し）にある、日付の列の名前', initial: '日付' },
  { key: 'csvName', label: 'CSVの振込名義の列', note: '振込んだ人の名前が入る列の名前（「内容」「摘要」など）', initial: '内容' },
  { key: 'csvAmount', label: 'CSVの入金額の列', note: '入金の金額が入る列の名前', initial: '入金金額(円)' },
  { key: 'fee', label: '振込手数料として見る差（円）', note: '入金が請求額よりこの金額まで少ないときは、手数料を引いた入金の候補として黄色で出す', initial: 880 },
  { key: 'rounding', label: '消費税の端数', note: '税率ごとの合計にかけた消費税の、1円未満の扱い', initial: '切り捨て', choices: ROUNDINGS },
  { key: 'admin', label: '管理者のメールアドレス', note: '自動の処理（毎月の明細）のエラーがここへ届く' },
  { key: 'subject', label: '下書きの件名', note: '{請求月} {取引先} {請求書番号} が使える', initial: '請求書のご送付（{請求月}分）' },
  { key: 'body', label: '下書きの本文', note: '{宛名} {会社名} {請求月} {請求書番号} {請求額} {支払期限} {振込先} が使える', initial: DEFAULT_BODY },
  { key: 'state', label: '自動の処理', note: 'ツールが書く。始める・止めるは、メニュー「請求書」から（管理者）', tool: true },
];
const SETTINGS_HEADER = ['項目', '内容', '説明'];
const SETTINGS_HEADER_ROW = 3;

// 「操作」シートのセル
const RUN_CELLS = { createResult: 'B9', paymentResult: 'B18' };
const RUN_BUTTONS = [
  { row: 6, fn: 'createInvoices', png: () => CREATE_BUTTON_PNG_BASE64 },
  { row: 15, fn: 'importPayments', png: () => PAYMENT_BUTTON_PNG_BASE64 },
];

const PROP = {
  SEQ_PREFIX: 'INVOICE_SEQ_', // 月ごとの、最後に使った請求書番号の通し番号（INVOICE_SEQ_2026-10）
  READ_FILES: 'PAYMENT_READ_FILES', // 読み終えたCSV（{ファイルID: 更新日時}）。同じファイルを何度も読まないように
  SENDER: 'AUTO_SENDER', // 自動の処理を動かすアカウント（最後に「自動の処理を始める」を押した人）
  STOPPED: 'AUTO_STOPPED', // 「自動の処理を止める」で '1'
  MONTHLY_DONE: 'MONTHLY_DONE', // 毎月の明細を入れ終えた月（2026-10）
};
const TRIGGER_HANDLERS = ['dailyRun'];

// ---- 開いたとき・編集したとき ----

/** スプレッドシートを開いたときに、メニューとシートを用意する */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem('選んだ請求書を取消す', 'cancelSelectedInvoice')
    .addSeparator()
    .addItem('自動の処理を始める（管理者）', 'startAuto')
    .addItem('自動の処理を止める（管理者）', 'stopAuto')
    .addToUi();
  try {
    ensureAllSheets_();
  } catch (e) {
    Logger.log(e.stack || e); // 作れなかったシートは、ボタンを押したときに作り直す
  }
}

/**
 * セルが編集されたときに呼ばれる
 * - 「入金」の「確認」にチェック → 請求書を「入金済み」にする。チェックを外す → 反映を取り消す（Payment.gs）
 * - 「請求書一覧」の「状態」を「入金済み」にした → 入金日・入金額が空なら、今日・請求額を入れる
 *   「未入金」に戻した → 入金日・入金額を消し、その請求書で反映した「入金」の行を確認待ちに戻す
 * - 「明細」の取引先を入れた → 請求日（次の締め日）と支払期限を取引先の条件で入れる（手で書き換えた日付は変えない。Invoice.gs）
 * - 「取引先」を書き換えた → 「明細」「毎月の請求」の取引先のプルダウンを入れ直す
 * - 「確認事項」を消した → その行の色を戻す
 */
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const name = sheet.getName();
    if (name === SHEET.PAYMENTS) confirmPaymentRows_(sheet, e.range);
    if (name === SHEET.INVOICES) syncPaidRows_(sheet, e.range);
    if (name === SHEET.LINES) fillDates_(e.range);
    if (name === SHEET.CLIENTS) refreshClientChoices_();
    if (TABLES[name]) restoreClearedRows_(sheet, e.range);
  } catch (err) {
    Logger.log(err.stack || err); // 編集のたびの処理なので、失敗しても画面には出さない（ボタンを押したときに整う）
  }
}

/**
 * 「請求書一覧」で「状態」を変えた行に合わせる
 * 入金済みにした → 空の入金日・入金額を埋める。未入金に戻した → 入金日・入金額を消し、その請求書で反映した「入金」の行を確認待ちに戻す
 */
function syncPaidRows_(sheet, range) {
  const t = openTable_(SHEET.INVOICES);
  const statusCol = t.map['状態'];
  if (range.getColumn() > statusCol || range.getLastColumn() < statusCol) return;
  const first = Math.max(range.getRow(), t.headerRow + 1);
  const last = range.getRow() + range.getNumRows() - 1;
  for (let r = first; r <= last; r++) {
    const row = t.read(r);
    if (row['状態'] === STATUS.PAID) {
      if (row['入金日'] === '') t.write(r, '入金日', today_());
      if (row['入金額（円）'] === '') t.write(r, '入金額（円）', row['請求額（円）']);
    }
    if (row['状態'] === STATUS.UNPAID && (row['入金日'] !== '' || row['入金額（円）'] !== '')) {
      t.write(r, '入金日', '');
      t.write(r, '入金額（円）', '');
      reopenPaymentsOf_(String(row['請求書番号']).trim(), '「請求書一覧」で未入金に戻されたため、反映を取り消しました。合う請求書を確かめて、もう一度「確認」にチェックを入れてください');
    }
  }
}

/** 「確認事項」を空にした行の色を戻す（入力する列は薄い青に） */
function restoreClearedRows_(sheet, range) {
  const t = openTable_(sheet.getName());
  const noteCol = t.map['確認事項'];
  if (!noteCol || range.getColumn() > noteCol || range.getLastColumn() < noteCol) return;
  const first = Math.max(range.getRow(), t.headerRow + 1);
  const last = range.getRow() + range.getNumRows() - 1;
  for (let r = first; r <= last; r++) if (t.read(r)['確認事項'] === '') t.paintRow(r, null);
}

// ---- シートを用意する ----

/** すべてのシートを、決まった並びでそろえる（あるものはそのまま。説明の文とボタンは毎回書き直す） */
function ensureAllSheets_() {
  ensureRunSheet_();
  [SHEET.LINES, SHEET.INVOICES, SHEET.PAYMENTS, SHEET.CLIENTS, SHEET.MONTHLY].forEach((name) => openTable_(name, true));
  ensureSettingsSheet_(true);
  refreshClientChoices_();
}

/**
 * 表のシートを開く（なければ作る）。見出しの列が消されていたら、右に足し直す
 * style が true なら、見た目と入力の決まりもそろえる（開いたとき・ボタンを押したとき。編集のたびには重いのでしない）
 * 戻り値: { sheet, headerRow, map（見出し → 列）, spec, read(row), write(row, name, value), rows(), paintRow(row, color), lastRow() }
 */
function openTable_(name, style) {
  const spec = TABLES[name];
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  let created = false;
  if (!sheet) {
    const active = ss.getActiveSheet();
    sheet = ss.insertSheet(name, sheetIndex_(name));
    sheet.getRange(spec.headerRow, 1, 1, spec.header.length).setValues([spec.header]);
    created = true;
    if (active) active.activate(); // シートを作っても、元の画面に戻す
  }
  const headerRow = findHeaderRow_(sheet, spec.key);
  if (!headerRow) {
    throw new AdminError('E24', `「${name}」シートの見出しの行が見つかりません`, `見出しの「${spec.key}」を消すか書き換えると出ます。「編集 → 元に戻す」か「ファイル → 変更履歴」で戻してください`);
  }
  const map = headerMap_(sheet, headerRow);
  const missing = spec.header.filter((h) => !map[h]);
  missing.forEach((h) => {
    const col = sheet.getLastColumn() + 1;
    if (col > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());
    sheet.getRange(headerRow, col).setValue(h);
    map[h] = col;
  });
  const t = tableApi_(sheet, headerRow, map, spec);
  if (created || missing.length) layoutTable_(t, created ? spec.header : missing);
  if (style || created || missing.length) styleTable_(t);
  return t;
}

/** シートを作る位置（決まった並びの中で、今あるシートの後ろ） */
function sheetIndex_(name) {
  const order = Object.values(SHEET);
  const ss = SpreadsheetApp.getActive();
  const before = order.slice(0, order.indexOf(name)).map((n) => ss.getSheetByName(n)).filter(Boolean);
  return before.length ? Math.max(...before.map((s) => s.getIndex())) : 0;
}

/** 表を読み書きする道具 */
function tableApi_(sheet, headerRow, map, spec) {
  const width = () => Math.max(...Object.values(map));
  const t = {
    sheet, headerRow, map, spec,
    /**
     * データの最後の行。lastRowKey のあるシートは、その列で数える
     * （チェックボックスの列は、空の行にも「チェックなし」が入っていることがあり、シートの最後の行だけでは決められないため）
     */
    lastRow() {
      const last = Math.max(sheet.getLastRow(), headerRow);
      if (!spec.lastRowKey || last === headerRow) return last;
      const values = sheet.getRange(headerRow + 1, map[spec.lastRowKey], last - headerRow, 1).getValues();
      for (let i = values.length - 1; i >= 0; i--) if (values[i][0] !== '') return headerRow + 1 + i;
      return headerRow;
    },
    /** 1行を { 見出し: 値 } で読む */
    read(row) {
      const values = sheet.getRange(row, 1, 1, width()).getValues()[0];
      const o = { _row: row };
      Object.entries(map).forEach(([h, c]) => { o[h] = values[c - 1]; });
      return o;
    },
    /** データの行を全部読む（空の行は飛ばす） */
    rows() {
      const n = t.lastRow() - headerRow;
      if (n <= 0) return [];
      const values = sheet.getRange(headerRow + 1, 1, n, width()).getValues();
      return values.map((v, i) => {
        const o = { _row: headerRow + 1 + i };
        Object.entries(map).forEach(([h, c]) => { o[h] = v[c - 1]; });
        return o;
      }).filter((o) => spec.header.some((h) => h !== '確認事項' && h !== '管理用' && o[h] !== '' && o[h] !== false));
    },
    write(row, name, value) { sheet.getRange(row, map[name]).setValue(value); },
    /** 1行に、色と「確認事項」を書く（color が null なら、確認事項を消して色を戻す） */
    paintRow(row, color, note) {
      if (map['確認事項']) {
        if (note !== undefined) sheet.getRange(row, map['確認事項']).setValue(note);
        else if (color === null) sheet.getRange(row, map['確認事項']).clearContent();
      }
      sheet.getRange(row, 1, 1, width()).setBackground(color);
      if (color === null) {
        const values = t.read(row);
        if (!spec.inputRow || spec.inputRow(values)) spec.inputs.forEach((h) => sheet.getRange(row, map[h]).setBackground(INPUT_COLOR));
      }
    },
  };
  return t;
}

/** 作ったとき・列を足し直したときだけ：列幅と隠し列（使う人が自分で変えた幅は戻さない） */
function layoutTable_(t, headers) {
  headers.forEach((h) => {
    const col = t.map[h];
    if (t.spec.widths[h]) t.sheet.setColumnWidth(col, t.spec.widths[h]);
    if ((t.spec.hidden || []).includes(h)) t.sheet.hideColumns(col);
  });
}

/**
 * 毎回：説明の文・見出し・入力の青・書式をそろえる（色の付いた行＝確認事項のある行は、色を変えない）
 * 一度に読み書きして、行が多くても遅くならないようにする
 */
function styleTable_(t) {
  const { sheet, headerRow, map, spec } = t;
  const width = Math.max(...Object.values(map));
  spec.help.forEach((text, i) => sheet.getRange(1 + i, 1).setValue(text).setFontColor(NOTE_COLOR));
  sheet.setFrozenRows(headerRow);
  const headerRange = sheet.getRange(headerRow, 1, 1, width);
  styleHeader_(headerRange);
  protectHeader_(sheet, headerRange);
  const dataRows = sheet.getMaxRows() - headerRow;
  if (dataRows <= 0) return;
  Object.entries(spec.numberFormats).forEach(([h, f]) => sheet.getRange(headerRow + 1, map[h], dataRows, 1).setNumberFormat(f));
  spec.right.forEach((h) => sheet.getRange(headerRow, map[h], dataRows + 1, 1).setHorizontalAlignment('right'));
  spec.wrap.forEach((h) => sheet.getRange(headerRow + 1, map[h], dataRows, 1).setWrap(true));

  // 入力する列の青。データの行は1行ずつ決め、その下の空の行は列ごとにまとめて塗る
  const last = t.lastRow();
  if (last > headerRow) {
    const n = last - headerRow;
    const range = sheet.getRange(headerRow + 1, 1, n, width);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const noteCol = map['確認事項'];
    const inputCols = spec.inputs.map((h) => map[h]);
    const next = values.map((v, i) => {
      if (noteCol && v[noteCol - 1] !== '') return backgrounds[i]; // 赤・黄の行はそのまま
      const row = {};
      Object.entries(map).forEach(([h, c]) => { row[h] = v[c - 1]; });
      const input = !spec.inputRow || spec.inputRow(row);
      return backgrounds[i].map((_, c) => (input && inputCols.includes(c + 1) ? INPUT_COLOR : null));
    });
    range.setBackgrounds(next);
  }
  const rest = sheet.getMaxRows() - last;
  if (rest > 0) {
    sheet.getRange(last + 1, 1, rest, width).setBackground(null);
    // ツールが行を足すシート（請求書一覧・入金）は、空の行に入力の青を塗らない（使う人が行を足す場所ではないので）
    if (spec.restInput !== false) spec.inputs.forEach((h) => sheet.getRange(last + 1, map[h], rest, 1).setBackground(INPUT_COLOR));
  }
  applyValidations_(t, dataRows);
  if (t.sheet.getName() === SHEET.INVOICES) applyStatusColors_(t, dataRows);
}

/** 入力の決まり（プルダウン・日付・数字・メールアドレス） */
function applyValidations_(t, dataRows) {
  const { sheet, headerRow, map } = t;
  const at = (h) => sheet.getRange(headerRow + 1, map[h], dataRows, 1);
  const rule = () => SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  const list = (values, help) => rule().requireValueInList(values, true).setHelpText(help).build();
  const date = rule().requireDate().setHelpText('日付を 2026/10/31 の形で入れてください').build();
  const name = t.sheet.getName();
  if (name === SHEET.LINES || name === SHEET.MONTHLY) {
    at('数量').setDataValidation(rule().requireNumberBetween(0.01, 1000000).setHelpText('数量は0より大きい数字で入れてください').build());
    at('単価（円）').setDataValidation(rule().requireNumberBetween(-100000000, 100000000).setHelpText('単価は数字で入れてください（税抜。値引きはマイナス）').build());
    at('税率').setDataValidation(list(TAX_RATES, '10% か 8% を選んでください'));
  }
  if (name === SHEET.LINES) {
    at('請求日').setDataValidation(date);
    at('支払期限').setDataValidation(date);
  }
  if (name === SHEET.INVOICES) {
    at('入金日').setDataValidation(date);
    at('入金額（円）').setDataValidation(rule().requireNumberBetween(0, 10000000000).setHelpText('入金額は数字で入れてください').build());
    // 「状態」のプルダウン。取消した行は、ツールが「取消」と書く（プルダウンでは選べない）
    const statusCol = map['状態'];
    const values = sheet.getRange(headerRow + 1, statusCol, dataRows, 1).getValues();
    values.forEach(([v], i) => {
      if (v === STATUS.CANCELED) sheet.getRange(headerRow + 1 + i, statusCol).clearDataValidations();
    });
    const statusRule = list([STATUS.UNPAID, STATUS.PAID], '未入金 か 入金済み を選んでください（取消は、メニュー「請求書 → 選んだ請求書を取消す」から）');
    let start = 0;
    for (let i = 0; i <= values.length; i++) {
      if (i === values.length || values[i][0] === STATUS.CANCELED) {
        if (i > start) sheet.getRange(headerRow + 1 + start, statusCol, i - start, 1).setDataValidation(statusRule);
        start = i + 1;
      }
    }
  }
  if (name === SHEET.PAYMENTS) {
    // チェックボックスは入金のある行だけに付ける（空の行に付けると「チェックなし」の値が入り、最後の行が分からなくなる）
    const last = t.lastRow();
    if (last > headerRow) setCheckboxes_(t, headerRow + 1, last - headerRow);
    const rest = sheet.getMaxRows() - last;
    if (rest > 0) sheet.getRange(last + 1, map['確認'], rest, 1).clearDataValidations().clearContent();
  }
  if (name === SHEET.CLIENTS) {
    at('締め日').setDataValidation(list(CLOSING_DAYS, '締め日を選んでください'));
    at('支払月').setDataValidation(list(PAY_MONTHS, '支払月を選んでください'));
    at('支払日').setDataValidation(list(CLOSING_DAYS, '支払日を選んでください'));
    const email = rule().requireTextIsEmail().setHelpText('メールアドレスの形で入れてください').build();
    at('メールアドレス').setDataValidation(email);
  }
}

/** 「入金」の「確認」の列の、row から n 行にチェックボックスを付ける */
function setCheckboxes_(t, row, n) {
  t.sheet.getRange(row, t.map['確認'], n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

/**
 * 「状態」のセルの色（見た目ルールの例外：まだ済んでいない＝薄い赤、済んだ＝薄い緑。取消は色なし）
 * 行全体の黄色（確かめてほしい行）と見分けられるよう、色は「状態」のセルだけに付ける。セルの文字で状態が分かるので、色だけに頼らない
 */
function applyStatusColors_(t, dataRows) {
  const range = t.sheet.getRange(t.headerRow + 1, t.map['状態'], dataRows, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.UNPAID).setBackground(ERROR_ROW_COLOR).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.PAID).setBackground(DONE_COLOR).setRanges([range]).build(),
  ];
  t.sheet.setConditionalFormatRules(rules);
}

/** 「明細」「毎月の請求」の取引先を、「取引先」シートの名前から選べるようにする */
function refreshClientChoices_() {
  const names = readClients_().list.map((c) => c.name);
  if (!names.length) return; // 取引先がまだないときは、プルダウンを付けない（入れた文字は、作るときに確かめる）
  [SHEET.LINES, SHEET.MONTHLY].forEach((sheetName) => {
    const t = openTable_(sheetName);
    const rows = t.sheet.getMaxRows() - t.headerRow;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false)
      .setHelpText('「取引先」シートにある取引先から選んでください（ないときは、先に「取引先」シートに足す）').build();
    t.sheet.getRange(t.headerRow + 1, t.map['取引先'], rows, 1).setDataValidation(rule);
  });
}

/** 「操作」シート（ボタンを置く画面）。いちばん左に作る。文は毎回書き直す */
function ensureRunSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET.RUN);
  if (!sheet) {
    const active = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET.RUN, sheetIndex_(SHEET.RUN));
    sheet.setHiddenGridlines(true);
    sheet.setColumnWidth(1, 24);
    sheet.setColumnWidth(2, 720);
    RUN_BUTTONS.forEach((b) => sheet.setRowHeight(b.row, 70));
    sheet.getRange(RUN_CELLS.createResult).setValue('まだ作っていません');
    sheet.getRange(RUN_CELLS.paymentResult).setValue('まだ取り込んでいません');
    if (active) active.activate();
  }
  sheet.getRange('B2').setValue('請求書の作成と入金の確認').setFontSize(16).setFontWeight('bold');
  sheet.getRange('B3:B4').setValues([
    ['1. 「明細」に請求する内容を入れる（「毎月の請求」の分は、月の初めに自動で入ります）'],
    ['2. 下のボタンを押す → 請求書のPDFと、PDFを付けたGmailの下書きができます。中身を確かめてから、Gmailの「下書き」から送ってください'],
  ]);
  sheet.getRange('B8').setValue('前回の結果').setFontWeight('bold');
  sheet.getRange('B11').setValue('入金の確認').setFontSize(14).setFontWeight('bold');
  sheet.getRange('B12:B13').setValues([
    ['1. ネットバンクからダウンロードした入金明細のCSVを、「設定」の「入金明細のCSVを置くフォルダ」に入れる'],
    ['2. 下のボタンを押す → 合う請求書は自動で「入金済み」になります。迷うものは「入金」シートに黄色で残ります'],
  ]);
  sheet.getRange('B17').setValue('前回の結果').setFontWeight('bold');
  sheet.getRange('B20').setValue('見かた').setFontWeight('bold');
  sheet.getRange('B21:B25').setValues([
    ['・赤い行 ＝ 直す必要がある行。右の「確認事項」のとおりに直して、もう一度ボタンを押してください'],
    ['・黄色い行 ＝ 確かめてほしい行。「確認事項」を読んで、確かめたら「確認事項」の文を消すと色が戻ります（「入金」シートは「確認」にチェック）'],
    ['・結果に「残り ○枚」と出たら、もう一度ボタンを押してください'],
    ['・ボタンはパソコンのブラウザで使ってください（スマートフォンのアプリでは動きません）'],
    ['・メールは自動では送りません。下書きを見て、送ってよいものだけを送ってください'],
  ]).setFontColor(NOTE_COLOR);
  sheet.getRange('B3:B25').setWrap(true);
  const images = sheet.getImages();
  RUN_BUTTONS.forEach((b) => {
    if (images.some((img) => img.getAnchorCell().getRow() === b.row)) return;
    const blob = Utilities.newBlob(Utilities.base64Decode(b.png()), 'image/png', `${b.fn}.png`);
    // 画像に関数を割り当てると、クリックで実行されるボタンになる
    sheet.insertImage(blob, 2, b.row).setWidth(280).setHeight(60).assignScript(b.fn);
  });
  return sheet;
}

/** 「操作」シートに前回の結果を残す（消える通知だけでなく、画面にも残す） */
function setRunResult_(cell, text) {
  try {
    const sheet = ensureRunSheet_();
    const stamp = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd HH:mm');
    sheet.getRange(cell).setValue(`${stamp}　${text}`);
  } catch (e) {
    Logger.log(e.stack || e);
  }
}

// ---- 設定 ----

/**
 * 「設定」シートを開く（なければ作る）。消された項目の行は、下に足し直す
 * style が true なら、説明の文・色・プルダウンもそろえる（開いたとき・ボタンを押したときの1回だけ。読むたびにすると遅い）
 */
function ensureSettingsSheet_(style) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET.SETTINGS);
  let created = false;
  if (!sheet) {
    const active = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET.SETTINGS, sheetIndex_(SHEET.SETTINGS));
    sheet.getRange(SETTINGS_HEADER_ROW, 1, 1, 3).setValues([SETTINGS_HEADER]);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 360);
    sheet.setColumnWidth(3, 520);
    created = true;
    if (active) active.activate();
  }
  const headerRow = findHeaderRow_(sheet, SETTINGS_HEADER[0]);
  if (!headerRow) throw new AdminError('E24', '「設定」シートの見出しの行が見つかりません', '見出しの「項目」を消すか書き換えると出ます。「編集 → 元に戻す」か「ファイル → 変更履歴」で戻してください');
  const rows = settingRows_(sheet, headerRow);
  let added = false;
  SETTINGS.forEach((s) => {
    if (rows[s.label]) return;
    const row = Math.max(sheet.getLastRow(), headerRow) + 1;
    sheet.getRange(row, 1, 1, 3).setValues([[s.label, s.initial == null ? '' : s.initial, s.note]]);
    rows[s.label] = row;
    added = true;
  });
  if (!style && !created && !added) return { sheet, rows };

  sheet.getRange('A1').setValue('設定').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A2').setValue('薄い青のセルを書きます。項目の名前（A列）は変えないでください').setFontColor(NOTE_COLOR);
  styleHeader_(sheet.getRange(headerRow, 1, 1, 3));
  protectHeader_(sheet, sheet.getRange(headerRow, 1, 1, 3));
  sheet.setFrozenRows(headerRow);
  const last = sheet.getLastRow();
  sheet.getRange(headerRow + 1, 2, last - headerRow, 2).setWrap(true).setVerticalAlignment('top');
  SETTINGS.forEach((s) => {
    const row = rows[s.label];
    sheet.getRange(row, 3).setValue(s.note).setFontColor(NOTE_COLOR); // 説明の文は毎回書き直す
    const cell = sheet.getRange(row, 2);
    if (!s.tool) cell.setBackground(INPUT_COLOR);
    if (s.choices) cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(s.choices, true).setAllowInvalid(false).build());
  });
  return { sheet, rows };
}

/** 項目の名前 → 行 */
function settingRows_(sheet, headerRow) {
  const last = sheet.getLastRow();
  const rows = {};
  if (last <= headerRow) return rows;
  sheet.getRange(headerRow + 1, 1, last - headerRow, 1).getValues().forEach(([v], i) => {
    const label = String(v).trim();
    if (label && !rows[label]) rows[label] = headerRow + 1 + i;
  });
  return rows;
}

/**
 * 設定を読む。戻り値: { 各 key: 値, problems: [問題の文] }
 * 値の形の確かめ（数字・登録番号など）はここで行い、使う側は「どの項目が要るか」だけを見る
 */
function readSettings_() {
  const { sheet, rows } = ensureSettingsSheet_();
  const s = { problems: [], sheet };
  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues(); // 1回でまとめて読む
  SETTINGS.forEach((def) => {
    const v = values[rows[def.label] - 1][0];
    s[def.key] = typeof v === 'string' ? v.trim() : v;
  });
  s.regno = normalizeRegNo_(s.regno);
  if (s.regno && !/^T\d{13}$/.test(s.regno)) s.problems.push('「登録番号」は T と13けたの数字にしてください（例：T1234567890123）');
  s.fee = toNumber_(s.fee);
  if (!(s.fee >= 0)) s.problems.push('「振込手数料として見る差（円）」は0以上の数字にしてください');
  if (!ROUNDINGS.includes(s.rounding)) s.rounding = '切り捨て';
  s.admin = String(s.admin || '');
  if (s.admin && !isEmail_(s.admin)) s.problems.push('「管理者のメールアドレス」の形が正しくありません');
  if (!s.subject) s.subject = SETTINGS.find((d) => d.key === 'subject').initial;
  if (!s.body) s.body = DEFAULT_BODY;
  return s;
}

/** 設定の必須の項目が空なら、どれが空かを出して止める（設定シートを開く） */
function requireSettings_(s, keys) {
  const empty = SETTINGS.filter((d) => keys.includes(d.key) && !s[d.key]).map((d) => `「${d.label}」`);
  const problems = s.problems.slice();
  if (empty.length) problems.unshift(`${empty.join('・')}を書いてください`);
  if (problems.length) {
    s.sheet.activate();
    throw new UserError(`「設定」シートを直して、もう一度押してください\n${problems.join('\n')}`);
  }
}

function writeSetting_(key, value) {
  const { sheet, rows } = ensureSettingsSheet_();
  sheet.getRange(rows[SETTINGS.find((d) => d.key === key).label], 2).setValue(value);
}

// ---- 取引先 ----

/**
 * 「取引先」シートを読む。戻り値: { list: [{ name, person, email, closing, payMonth, payDay, kana, row }], byName: {正規化した名前: 取引先} }
 */
function readClients_() {
  const t = openTable_(SHEET.CLIENTS);
  const list = t.rows().filter((r) => String(r['取引先名']).trim()).map((r) => ({
    name: String(r['取引先名']).trim(),
    person: String(r['先方の担当者']).trim(),
    email: String(r['メールアドレス']).trim(),
    cc: String(r['CCのメールアドレス'] || '').split(/[,、，\s]+/).map((x) => x.trim()).filter(Boolean),
    closing: String(r['締め日']).trim(),
    payMonth: String(r['支払月']).trim(),
    payDay: String(r['支払日']).trim(),
    kana: String(r['振込名義（カナ）']).trim(),
    row: r._row,
  }));
  const byName = {};
  list.forEach((c) => { if (!byName[nameKey_(c.name)]) byName[nameKey_(c.name)] = c; });
  return { list, byName, table: t };
}

/** 取引先の支払条件の問題（なければ空文字） */
function clientTermsProblem_(c) {
  const missing = [];
  if (!CLOSING_DAYS.includes(c.closing)) missing.push('締め日');
  if (!PAY_MONTHS.includes(c.payMonth)) missing.push('支払月');
  if (!CLOSING_DAYS.includes(c.payDay)) missing.push('支払日');
  return missing.length ? `「取引先」シートの「${c.name}」の${missing.join('・')}を選んでください` : '';
}

// ---- 日付・金額の部品 ----

/** 日本時間の年・月・日 */
function jstParts_(date) {
  const d = new Date(date.getTime() + JST_OFFSET_MS);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), w: d.getUTCDay() };
}

/** 日本時間の y年m月d日 0時（m は 13 や 0 でもよい。日は 0 で前の月の末日） */
function jstDate_(y, m, d) {
  return new Date(Date.UTC(y, m - 1, d) - JST_OFFSET_MS);
}

function today_() {
  const p = jstParts_(new Date());
  return jstDate_(p.y, p.m, p.d);
}

function daysInMonth_(y, m) {
  return jstParts_(jstDate_(y, m + 1, 0)).d;
}

/** 「15日」「末日」→ その月の日 */
function dayOf_(label, y, m) {
  const last = daysInMonth_(y, m);
  return label === '末日' ? last : Math.min(parseInt(label, 10), last);
}

function ymd_(date) {
  return Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
}

function slashDate_(date) {
  return Utilities.formatDate(date, TIME_ZONE, 'yyyy/MM/dd');
}

/**
 * セルの値を日付にする（日付・「2026/10/31」などの文字）。できなければ null
 * 銀行のCSVに多い書き方も読む：時刻つき（2026/10/31 10:15）、和暦（R8.10.31・令和8年10月31日）、区切りなし（20261031）
 */
function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    const p = jstParts_(v);
    return jstDate_(p.y, p.m, p.d);
  }
  const s = String(v || '').normalize('NFKC').trim().replace(/[T\s]+\d{1,2}:\d{2}(:\d{2})?$/, '');
  const era = s.match(/^(令和|平成|R|H)\s*(\d{1,2}|元)[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?$/i);
  const m = s.match(/^(\d{4})[/\-.年]?(\d{1,2})[/\-.月]?(\d{1,2})日?$/);
  let y;
  let mo;
  let d;
  if (era) {
    const base = /^(令和|R)$/i.test(era[1]) ? 2018 : 1988;
    [y, mo, d] = [base + (era[2] === '元' ? 1 : Number(era[2])), Number(era[3]), Number(era[4])];
  } else if (m) {
    [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  } else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth_(y, mo)) return null; // 2月31日のような、ない日はじく
  return jstDate_(y, mo, d);
}

/** 数字・「1,280」「１２８０円」などの文字 → 数。数でなければ NaN */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).normalize('NFKC').replace(/[,円¥\\\s+]/g, '').replace(/^[−ー]/, '-');
  if (s === '' || !/^-?\d+(\.\d+)?$/.test(s)) return NaN;
  return Number(s);
}

/** 3けたごとのカンマ */
function yen_(n) {
  return Math.round(n).toLocaleString('ja-JP');
}

/** 端数の処理（マイナスの金額も、0に向けて切り捨てる） */
function roundYen_(n, mode) {
  const sign = n < 0 ? -1 : 1;
  const a = Math.abs(n) + 1e-9; // 0.1 + 0.2 のような小数の誤差で、1円ずれないように
  const r = mode === '四捨五入' ? Math.floor(a + 0.5) : mode === '切り上げ' ? Math.ceil(a - 2e-9) : Math.floor(a);
  return sign * r || 0;
}

/** 登録番号の表記ゆれ（全角・小文字のt・ハイフン・空白）をそろえる */
function normalizeRegNo_(v) {
  return String(v || '').normalize('NFKC').replace(/[\s\-‐ー]/g, '').toUpperCase();
}

/** 取引先の名前を比べるための形（全角/半角・空白・大文字/小文字の違いをなくす） */
function nameKey_(s) {
  return String(s || '').normalize('NFKC').replace(/\s/g, '').toLowerCase();
}

function isEmail_(s) {
  return /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(String(s || '').trim());
}

/** HTMLに入れる文字の、特別な記号を置き換える */
function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * スプレッドシートのタイムゾーンが日本時間でなければ止める
 * スプレッドシートのタイムゾーンはスクリプトとは別で、違うと日付のセルを読み書きしたときに1日ずれる（10/31 が 10/30 になる）
 */
function requireTimeZone_() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  if (tz === TIME_ZONE) return;
  throw new UserError(`このスプレッドシートのタイムゾーンが「${tz}」になっていて、日付が1日ずれてしまいます。「ファイル → 設定」の「タイムゾーン」を「(GMT+09:00) 東京」にしてから、もう一度押してください`);
}

/** 画面のエラーの出し方（使う人が直せるもの・管理者に連絡するもの） */
function showError_(e) {
  Logger.log(e.stack || e);
  SpreadsheetApp.getActive().toast('', '', 1); // 「〜しています」の通知を消す
  const ui = SpreadsheetApp.getUi();
  ui.alert(e instanceof UserError ? '確認してください' : '管理者に連絡してください', errorText_(e), ui.ButtonSet.OK);
}

/** フォルダのURL（またはID）からフォルダを開く。開けなければ code のエラー */
function openFolder_(value, label, code) {
  const m = String(value).match(/folders\/([\w-]+)/) || String(value).match(/[?&]id=([\w-]+)/);
  const id = m ? m[1] : String(value).trim();
  if (!/^[\w-]{10,}$/.test(id)) throw new UserError(`「${label}」には、フォルダを開いたときのアドレス（URL）を貼ってください`);
  try {
    const folder = DriveApp.getFolderById(id);
    folder.getName(); // 開けるかをここで確かめる
    return folder;
  } catch (e) {
    throw new AdminError(code, `「${label}」を開けません`, 'URLが正しいか、このフォルダがボタンを押した人に共有されているか（編集者として）を確かめてください');
  }
}
