/**
 * 請求書を作る・取消す
 *
 * 「請求書を作る」ボタン（createInvoices）:
 *   請求書番号が空の明細を確かめる → 取引先と請求日が同じものを1枚にまとめる → 確認画面
 *   → 1枚ずつ: 番号を決める → PDFを保存 → 「請求書一覧」に1行 → 明細に番号を書く → Gmailの下書き
 *   （番号を先に決めて残すので、途中で止まっても同じ番号を2回使わない。PDFが保存できてから一覧と明細に書くので、PDFのない請求書は残らない）
 * メニュー「選んだ請求書を取消す」（cancelSelectedInvoice）:
 *   一覧を「取消」にして理由を残す → 明細を作り直せる状態に戻す → PDFの名前に「取消」を付け、残っている下書きを消す
 */

// 「明細」の隠し列「管理用」の印。取消した番号は「確認事項」の文ではなくここに持つ（使う人が文を消しても、訂正版とつながるように）
const LINE_MARK = { CANCELED: /取消:(\S+)/g, MONTHLY: '毎月' };
const cancelNoteText_ = (no) => `取消した請求書 ${no} の明細です。直してから「請求書を作る」を押すと、新しい番号で作り直します（訂正版になります）`;
// 自動で入れた請求日・支払期限の印（「明細」の隠し列。例「請求日:2026-10-31 期限:2026-11-30」）
// セルがこの日付のままなら、取引先の条件で入れ直してよい。手で書き換えた日付は変えない
const DATE_MARK_KEYS = { BILL: '請求日', DUE: '期限' };
const CANCELED_SUFFIX = '（取消）'; // 作り直さない取消のとき、明細の「請求書番号」に付ける（作る対象から外れる）
const GROUP_BLOCKED_NOTE = '同じ請求書（取引先と請求日が同じ）の明細に直すところがあるため、まだ作っていません';

/** 「操作」シートのボタンから呼ばれる */
function createInvoices() {
  const started = Date.now(); // GASの6分は実行の始めから数えるので、確認画面を出している間も時間に入れる
  const ss = SpreadsheetApp.getActive();
  // 2回押されても、同じ明細で2枚作らないよう、1つずつ実行する
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) {
    ss.toast('ほかの人が請求書を作っているか、入金を取り込んでいます。終わってから、もう一度押してください', '確認してください', 20);
    return;
  }
  try {
    ss.toast('明細を確かめています。このままお待ちください', '請求書を作る', -1);
    createInvoices_(started);
  } catch (e) {
    showError_(e);
  } finally {
    lock.releaseLock();
  }
}

function createInvoices_(started) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  requireTimeZone_();
  ensureAllSheets_();
  const s = readSettings_();
  requireSettings_(s, ['company', 'regno', 'bank', 'pdfFolder']);
  const folder = openFolder_(s.pdfFolder, '請求書を保存するフォルダ', 'E21');
  retryMonthly_(); // 前に入れられなかった「毎月の請求」の行が直っていれば、ここで明細に入れる（赤い行のままなら何もしない）
  const clients = readClients_();
  const lines = openTable_(SHEET.LINES);
  const plan = planInvoices_(lines.rows(), clients.byName, s.rounding);
  writeLineChecks_(lines, plan);

  if (!plan.groups.length) {
    ss.toast('', '', 1);
    const text = plan.errorRows ? `作れる請求書がありません。「明細」の赤い行（${plan.errorRows}行）の「確認事項」のとおりに直してから、もう一度押してください`
      : '作る請求書がありません（請求書番号が空の明細がありません）';
    setRunResult_(RUN_CELLS.createResult, text);
    if (plan.errorRows) lines.sheet.activate();
    ui.alert('請求書を作りませんでした', text, ui.ButtonSet.OK);
    return;
  }

  const total = plan.groups.reduce((n, g) => n + g.totals.total, 0);
  const list = plan.groups.slice(0, CONFIG.CONFIRM_LIST_MAX).map((g) => `・${g.client.name}（請求日 ${slashDate_(g.date)}）　${yen_(g.totals.total)}円`);
  if (plan.groups.length > CONFIG.CONFIRM_LIST_MAX) list.push(`・ほか ${plan.groups.length - CONFIG.CONFIRM_LIST_MAX}枚`);
  const message = [
    `請求書を ${plan.groups.length}枚（合計 ${yen_(total)}円・税込）作ります。`,
    ...list,
    '',
    'PDFを保存して、Gmailに下書きを作ります（メールは送りません）。',
  ];
  if (plan.errorRows) message.push('', `赤い行（${plan.errorRows}行）の明細は、直すまで作りません。`);
  ss.toast('', '', 1);
  if (ui.alert('請求書を作りますか？', message.join('\n'), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const invoices = openTable_(SHEET.INVOICES);
  const made = [];
  let drafts = 0;
  let noMail = 0;
  for (const g of plan.groups) {
    if (Date.now() - started > CONFIG.TIME_BUDGET_MS) break; // 残りは次に押したとき（確認画面で長く待ったときは、1枚も作らずに止める）
    ss.toast(`${made.length + 1}枚目 / ${plan.groups.length}枚 を作っています（${g.client.name}）`, '請求書を作る', -1);
    try {
      const r = issueInvoice_(g, s, folder, invoices, lines);
      made.push(r);
      if (r.draft) drafts++;
      if (!g.client.email) noMail++;
    } catch (e) {
      const done = made.length ? `\n\nここまでに ${made.length}枚 を作りました（「請求書一覧」にあります）。` : '';
      setRunResult_(RUN_CELLS.createResult, `${made.length}枚 作ったところで止まりました。${errorText_(e)}`);
      e.message += done;
      throw e;
    }
  }
  const left = plan.groups.length - made.length;
  const parts = [`請求書 ${made.length}枚（合計 ${yen_(made.reduce((n, r) => n + r.total, 0))}円）を作りました。下書き ${drafts}通`];
  if (noMail) parts.push(`メールアドレスがなく下書きを作っていない ${noMail}枚（「請求書一覧」の黄色い行）`);
  if (plan.errorRows) parts.push(`直すところのある明細 ${plan.errorRows}行（「明細」の赤い行）`);
  if (left) parts.push(`残り ${left}枚 は、もう一度押してください`);
  const text = parts.join('　');
  setRunResult_(RUN_CELLS.createResult, text);
  ss.toast(text, '請求書を作りました', 30);
}

/**
 * 明細を確かめて、請求書ごとにまとめる（シートには書かない。書くのは writeLineChecks_）
 * rows: 「明細」の行（{ 見出し: 値, _row }）、clients: { 正規化した名前: 取引先 }、rounding: 消費税の端数
 * 戻り値: {
 *   groups: [{ key, client, date, due, lines: [{ row, item, qty, price, rate, amount }], totals, replaces: [取消した番号], rows: [行] }],
 *   （取引先・請求日・支払期限が同じ明細を1枚にする。契約で期限が違う明細は、別の請求書になる）
 *   notes: [{ row, color, note }]（行に書く確認事項。color が null なら消す）,
 *   fills: [{ row, name, value }]（空欄に入れる初期値・数字にそろえた値）,
 *   errorRows: 赤い行の数
 * }
 */
function planInvoices_(rows, clients, rounding) {
  const notes = [];
  const fills = [];
  const byKey = {};
  const order = [];
  const blocked = new Set(); // 直すところのある明細がある「取引先|請求日」（その取引先・請求日の請求書は、どの支払期限のものも作らない）
  rows.forEach((r) => {
    if (String(r['請求書番号']).trim()) return; // 作成済み
    const oldNote = String(r['確認事項'] || '');
    const canceled = [...String(r['管理用'] || '').matchAll(LINE_MARK.CANCELED)].map((m) => m[1]);
    const problems = [];
    const clientName = String(r['取引先']).trim();
    const client = clientName ? clients[nameKey_(clientName)] : null;
    const termsOk = client && !clientTermsProblem_(client);
    let mark = String(r['管理用'] || '');
    let date = toDate_(r['請求日']);
    if (r['請求日'] === '' && termsOk) { // 空の請求日は、次に来る締め日にする
      date = nextClosing_(today_(), client);
      mark = withDateMark_(mark, DATE_MARK_KEYS.BILL, ymd_(date));
      fills.push({ row: r._row, name: '請求日', value: date });
    }
    if (!date) problems.push(r['請求日'] === '' ? '請求日を入れてください' : '請求日を 2026/10/31 の形で入れてください');
    if (!clientName) problems.push('取引先を選んでください');
    else if (!client) problems.push('「取引先」シートにない取引先です。先に「取引先」シートに足してください');
    else if (clientTermsProblem_(client)) problems.push(clientTermsProblem_(client));
    const item = String(r['品目']).trim();
    if (!item) problems.push('品目を入れてください');
    let qty = r['数量'] === '' ? 1 : toNumber_(r['数量']);
    if (r['数量'] === '') fills.push({ row: r._row, name: '数量', value: 1 });
    else if (!(qty > 0)) problems.push('数量は0より大きい数字にしてください');
    else if (typeof r['数量'] !== 'number') fills.push({ row: r._row, name: '数量', value: qty });
    const price = toNumber_(r['単価（円）']);
    if (r['単価（円）'] === '') problems.push('単価を入れてください');
    else if (isNaN(price)) problems.push('単価は数字にしてください（税抜。値引きはマイナス）');
    else if (typeof r['単価（円）'] !== 'number') fills.push({ row: r._row, name: '単価（円）', value: price });
    const rate = r['税率'] === '' ? 10 : taxRateOf_(r['税率']);
    if (r['税率'] === '') fills.push({ row: r._row, name: '税率', value: '10%' });
    else if (!rate) problems.push('税率は 10% か 8% を選んでください');
    else if (r['税率'] !== `${rate}%`) fills.push({ row: r._row, name: '税率', value: `${rate}%` });
    let due = null;
    if (date && termsOk) {
      const cell = r['支払期限'];
      const cellDate = toDate_(cell);
      if (cell === '' || (cellDate && ymd_(cellDate) === dateMark_(mark, DATE_MARK_KEYS.DUE))) {
        due = dueDate_(date, client); // 自動の期限（取引先の条件が変わっていれば、新しい条件で入れ直す）
        if (cell === '' || ymd_(cellDate) !== ymd_(due)) {
          fills.push({ row: r._row, name: '支払期限', value: due });
          mark = withDateMark_(mark, DATE_MARK_KEYS.DUE, ymd_(due));
        }
      } else if (!cellDate) problems.push('支払期限を 2026/11/30 の形で入れてください（空にすると、取引先の条件で入ります）');
      else if (cellDate < date) problems.push('支払期限が請求日より前です');
      else due = cellDate;
    }
    if (mark !== String(r['管理用'] || '')) fills.push({ row: r._row, name: '管理用', value: mark });

    if (problems.length) {
      notes.push({ row: r._row, color: ERROR_ROW_COLOR, note: problems.join('。') });
      if (date && client) blocked.add(`${client.name}|${ymd_(date)}`);
      return;
    }
    const key = `${client.name}|${ymd_(date)}|${ymd_(due)}`;
    if (!byKey[key]) { byKey[key] = { key, client, date, due, lines: [], rows: [], replaces: [] }; order.push(key); }
    const g = byKey[key];
    g.lines.push({ row: r._row, item, qty, price, rate, amount: roundYen_(qty * price, rounding) });
    g.rows.push(r._row);
    canceled.forEach((no) => { if (!g.replaces.includes(no)) g.replaces.push(no); });
    // 前に赤・黄にした行は、直っていれば色を戻す（取消のあとの黄色は、作り直すまで残す）
    if (oldNote && !canceled.some((no) => oldNote === cancelNoteText_(no))) notes.push({ row: r._row, color: null });
  });

  const groups = [];
  order.forEach((key) => {
    const g = byKey[key];
    if (blocked.has(`${g.client.name}|${ymd_(g.date)}`)) {
      g.rows.forEach((row) => notes.push({ row, color: WARN_ROW_COLOR, note: GROUP_BLOCKED_NOTE }));
      return;
    }
    g.totals = invoiceTotals_(g.lines, rounding);
    if (g.totals.total <= 0) {
      g.rows.forEach((row) => notes.push({ row, color: ERROR_ROW_COLOR, note: 'この請求書の合計が0円以下です。数量・単価を確かめてください' }));
      return;
    }
    groups.push(g);
  });
  // 1つの行に、あとから書いた確認事項を優先する（直った印より、まとめた結果の黄色・赤を残す）
  const last = {};
  notes.forEach((n) => { last[n.row] = n; });
  const finalNotes = Object.values(last);
  return { groups, notes: finalNotes, fills, errorRows: finalNotes.filter((n) => n.color === ERROR_ROW_COLOR).length };
}

/** 隠し列「管理用」から、自動で入れた日付（yyyy-MM-dd）を読む。key は DATE_MARK_KEYS。なければ空 */
function dateMark_(value, key) {
  const m = String(value || '').match(new RegExp(`${key}:(\\d{4}-\\d{2}-\\d{2})`));
  return m ? m[1] : '';
}

/** 隠し列「管理用」の、自動で入れた日付の印を書き換える（ほかの印は残す。ymd が空なら印を消す） */
function withDateMark_(value, key, ymd) {
  const rest = String(value || '').replace(new RegExp(`${key}:\\d{4}-\\d{2}-\\d{2}`), '').replace(/\s+/g, ' ').trim();
  return [rest, ymd ? `${key}:${ymd}` : ''].filter(Boolean).join(' ');
}

/** 次に来る締め日（今日が締め日なら今日）。末日締めで 10/18 → 10/31、20日締めで 10/21 → 11/20 */
function nextClosing_(today, client) {
  const p = jstParts_(today);
  const closing = jstDate_(p.y, p.m, dayOf_(client.closing, p.y, p.m));
  if (closing >= today) return closing;
  const n = jstParts_(jstDate_(p.y, p.m + 1, 1));
  return jstDate_(n.y, n.m, dayOf_(client.closing, n.y, n.m));
}

/** 入れた日付が、前に自動で入れたままか（空も自動として扱う） */
function isAutoDate_(cell, mark, key) {
  const d = toDate_(cell);
  return cell === '' || (d && ymd_(d) === dateMark_(mark, key));
}

/**
 * 「明細」で取引先・請求日・支払期限を入れたとき（onEdit から）：取引先の条件で請求日と支払期限を入れる
 * 請求日は次に来る締め日、支払期限は請求日から決まる日
 * 入れ直すのは、空か、前に自動で入れた日付のままのときだけ（手で書き換えた日付は変えない。空にすると自動に戻る）
 */
function fillDates_(range) {
  const t = openTable_(SHEET.LINES);
  const cols = ['請求日', '取引先', '支払期限'].map((h) => t.map[h]);
  if (!cols.some((c) => c >= range.getColumn() && c <= range.getLastColumn())) return;
  const clients = readClients_().byName;
  const first = Math.max(range.getRow(), t.headerRow + 1);
  const last = range.getRow() + range.getNumRows() - 1;
  for (let r = first; r <= last; r++) {
    const row = t.read(r);
    if (String(row['請求書番号']).trim()) continue;
    const client = clients[nameKey_(row['取引先'])];
    if (!client || clientTermsProblem_(client)) continue;
    let mark = String(row['管理用'] || '');
    let date = toDate_(row['請求日']);
    if (isAutoDate_(row['請求日'], mark, DATE_MARK_KEYS.BILL)) {
      date = nextClosing_(today_(), client);
      if (!toDate_(row['請求日']) || ymd_(toDate_(row['請求日'])) !== ymd_(date)) t.write(r, '請求日', date);
      mark = withDateMark_(mark, DATE_MARK_KEYS.BILL, ymd_(date));
    }
    if (date && isAutoDate_(row['支払期限'], mark, DATE_MARK_KEYS.DUE)) {
      const due = dueDate_(date, client);
      t.write(r, '支払期限', due);
      mark = withDateMark_(mark, DATE_MARK_KEYS.DUE, ymd_(due));
    }
    if (mark !== String(row['管理用'] || '')) t.write(r, '管理用', mark);
  }
}

/** 「10%」「10％」「0.1」（セルが％の数字になったとき）「10」 → 10。8% も同じ。それ以外は 0 */
function taxRateOf_(v) {
  let n = typeof v === 'number' ? v : toNumber_(String(v).normalize('NFKC').replace('%', ''));
  if (n > 0 && n < 1) n = Math.round(n * 100);
  return n === 10 || n === 8 ? n : 0;
}

/**
 * 請求書の金額（インボイスの決まり：税率ごとに合計してから、消費税を1回だけ計算する）
 * 戻り値: { byRate: [{ rate, base, tax }]（10%→8%の順。ある税率だけ）, base, tax, total }
 */
function invoiceTotals_(lines, rounding) {
  const byRate = [10, 8].map((rate) => {
    const base = lines.filter((l) => l.rate === rate).reduce((n, l) => n + l.amount, 0);
    return { rate, base, tax: roundYen_((base * rate) / 100, rounding), used: lines.some((l) => l.rate === rate) };
  }).filter((x) => x.used).map(({ rate, base, tax }) => ({ rate, base, tax }));
  const base = byRate.reduce((n, x) => n + x.base, 0);
  const tax = byRate.reduce((n, x) => n + x.tax, 0);
  return { byRate, base, tax, total: base + tax };
}

/**
 * 支払期限：請求日を含む締め（締め日を過ぎた請求日は次の締め）の、支払月の支払日
 * 例：末日締め・翌月・末日 で請求日 10/31 → 11/30。15日締め・翌月・10日 で請求日 10/20 → 12/10
 */
function dueDate_(date, client) {
  const p = jstParts_(date);
  let m = p.m;
  if (p.d > dayOf_(client.closing, p.y, p.m)) m += 1;
  const first = jstParts_(jstDate_(p.y, m + PAY_MONTHS.indexOf(client.payMonth), 1));
  let due = jstDate_(first.y, first.m, dayOf_(client.payDay, first.y, first.m));
  if (due < date) { // 当月払いで支払日が請求日より前（登録の間違いに近い）。次の月の支払日にする
    const n = jstParts_(jstDate_(first.y, first.m + 1, 1));
    due = jstDate_(n.y, n.m, dayOf_(client.payDay, n.y, n.m));
  }
  return due;
}

/** 確かめた結果を「明細」に書く（色と確認事項、空欄の初期値） */
function writeLineChecks_(t, plan) {
  plan.fills.forEach((f) => t.write(f.row, f.name, f.value));
  plan.notes.forEach((n) => (n.color ? t.paintRow(n.row, n.color, n.note) : t.paintRow(n.row, null)));
  if (plan.errorRows) Logger.log(`直すところのある明細: ${plan.errorRows}行`);
}

/**
 * 1枚の請求書を作る。戻り値: { no, total, draft（下書きを作ったか） }
 * 番号 → PDF → 一覧 → 明細 → 下書き の順。下書きが作れなくても、請求書はできている（一覧に黄色で知らせる）
 */
function issueInvoice_(g, s, folder, invoices, lines) {
  const ym = Utilities.formatDate(g.date, TIME_ZONE, 'yyyy-MM');
  const no = reserveNumber_(ym, invoices);
  const inv = { no, date: g.date, client: g.client, lines: g.lines, totals: g.totals, due: g.due, replaces: g.replaces, period: billingPeriod_(g.date, g.client) };
  const fileName = `請求書_${no}_${g.client.name}.pdf`.replace(/[\\/:*?"<>|]/g, '_');
  let file;
  try {
    const blob = Utilities.newBlob(invoiceHtml_(inv, s), 'text/html', 'invoice.html').getAs('application/pdf').setName(fileName);
    file = folder.createFile(blob);
  } catch (e) {
    Logger.log(e.stack || e);
    throw new AdminError('E23', `請求書（${g.client.name}）のPDFを保存できませんでした`, '「請求書を保存するフォルダ」の編集者になっているか、Googleドライブの容量に空きがあるかを確かめてください');
  }

  // 一覧に書く。書く行は、書く直前に決める（途中でほかの人が行を足しても、上書きしない）
  const row = invoices.lastRow() + 1;
  const width = Math.max(...Object.values(invoices.map));
  const values = new Array(width).fill('');
  const set = (h, v) => { values[invoices.map[h] - 1] = v; };
  set('請求書番号', no);
  set('請求日', g.date);
  set('取引先', textCell_(g.client.name));
  set('税抜（円）', g.totals.base);
  set('消費税（円）', g.totals.tax);
  set('請求額（円）', g.totals.total);
  set('支払期限', g.due);
  set('状態', STATUS.UNPAID);
  set('PDF', linkFormula_(file.getUrl(), fileName));
  if (g.replaces.length) set('取消・訂正', `${g.replaces.join('・')} の訂正版`);
  set('管理用', JSON.stringify({ file: file.getId() }));
  invoices.sheet.getRange(row, 1, 1, width).setValues([values]);
  invoices.paintRow(row, null);

  markLinesIssued_(lines, g, no);
  g.replaces.forEach((old) => appendCorrection_(invoices, old, `→ ${no} で作り直し`));

  let draft = false;
  if (!g.client.email) {
    invoices.paintRow(row, WARN_ROW_COLOR, '取引先のメールアドレスがないため、下書きを作っていません。PDFを開いて、ほかの方法で送ってください');
  } else {
    try {
      const mail = draftMail_(inv, s);
      const options = { attachments: [file.getBlob()], name: s.company };
      const cc = g.client.cc.filter(isEmail_);
      const badCc = g.client.cc.filter((x) => !isEmail_(x));
      if (cc.length) options.cc = cc.join(',');
      const d = GmailApp.createDraft(g.client.email, mail.subject, mail.body, options);
      invoices.write(row, '管理用', JSON.stringify({ file: file.getId(), draft: d.getId() }));
      draft = true;
      if (badCc.length) invoices.paintRow(row, WARN_ROW_COLOR, `CCの「${badCc.join('、')}」がメールアドレスの形でないため、外して下書きを作りました。「取引先」のCCを直し、下書きの宛先も確かめてください`);
    } catch (e) {
      Logger.log(e.stack || e);
      invoices.paintRow(row, WARN_ROW_COLOR, 'メールの下書きを作れませんでした。PDFを開いて、Gmailで添付して送ってください');
    }
  }
  return { no, total: g.totals.total, draft };
}

/**
 * インボイスの「取引年月日」：請求日が取引先の締め日なら、締めの期間（前の締め日の翌日〜締め日）。それ以外の日なら null（請求日を取引日として載せる）
 * 例：末日締めで請求日 10/31 → 10/01〜10/31。20日締めで請求日 10/20 → 9/21〜10/20
 */
function billingPeriod_(date, client) {
  const p = jstParts_(date);
  if (p.d !== dayOf_(client.closing, p.y, p.m)) return null;
  const prev = jstParts_(jstDate_(p.y, p.m - 1, 1));
  return { start: new Date(jstDate_(prev.y, prev.m, dayOf_(client.closing, prev.y, prev.m)).getTime() + DAY_MS), end: date };
}

/**
 * 請求書番号を決めて残す（年月-その月の通し番号。例 2026-10-001）
 * スクリプトプロパティの番号と、一覧にある一番大きい番号の、大きいほうの次にする（一覧の行を消しても、番号を使い回さない）
 */
function reserveNumber_(ym, invoices) {
  const props = PropertiesService.getScriptProperties();
  const key = PROP.SEQ_PREFIX + ym;
  let n = Number(props.getProperty(key)) || 0;
  invoices.rows().forEach((r) => {
    const m = String(r['請求書番号']).match(/^(\d{4}-\d{2})-(\d+)$/);
    if (m && m[1] === ym) n = Math.max(n, Number(m[2]));
  });
  n += 1;
  props.setProperty(key, String(n));
  return `${ym}-${String(n).padStart(3, '0')}`;
}

/**
 * 明細の行に請求書番号を書く。PDFを作っている間に並べ替えられても、別の行に書かないよう、中身で行を探し直す
 */
function markLinesIssued_(t, g, no) {
  const current = t.rows();
  const used = new Set();
  g.lines.forEach((l) => {
    const same = (r) => !String(r['請求書番号']).trim() && !used.has(r._row)
      && nameKey_(r['取引先']) === nameKey_(g.client.name) && toDate_(r['請求日']) && ymd_(toDate_(r['請求日'])) === ymd_(g.date)
      && String(r['品目']).trim() === l.item && toNumber_(r['単価（円）']) === l.price && toNumber_(r['数量'] === '' ? 1 : r['数量']) === l.qty;
    const r = current.find((x) => x._row === l.row && same(x)) || current.find(same);
    if (!r) return; // 消された行。請求書はできているので、そのままにする
    used.add(r._row);
    t.write(r._row, '請求書番号', no);
    t.write(r._row, '管理用', ''); // 取消した番号のつながりは、作り直したので要らない
    t.paintRow(r._row, null);
  });
}

/** 一覧の「取消・訂正」に書き足す */
function appendCorrection_(invoices, no, text) {
  const r = invoices.rows().find((x) => x['請求書番号'] === no);
  if (!r) return;
  const before = String(r['取消・訂正'] || '');
  invoices.write(r._row, '取消・訂正', before ? `${before}　${text}` : text);
}

/** 宛名（担当者がいれば「様」、いなければ会社に「御中」） */
function addressee_(client) {
  return client.person ? `${client.name}\n${client.person} 様` : `${client.name} 御中`;
}

/** 下書きの件名と本文（「設定」の文の {〇〇} を置き換える） */
function draftMail_(inv, s) {
  const p = jstParts_(inv.date);
  const words = {
    宛名: addressee_(inv.client),
    会社名: s.company,
    取引先: inv.client.name,
    請求月: `${p.y}年${p.m}月`,
    請求書番号: inv.no,
    請求額: yen_(inv.totals.total),
    支払期限: slashDate_(inv.due),
    振込先: s.bank,
  };
  const fill = (text) => String(text).replace(/\{([^{}]+)\}/g, (all, k) => (k in words ? words[k] : all));
  return { subject: fill(s.subject), body: fill(s.body) };
}

/** 請求書のPDFの元になるHTML（A4・縦。インボイスの項目：発行者・登録番号・取引日・税率ごとの合計と消費税・受け取る人） */
function invoiceHtml_(inv, s) {
  const h = escapeHtml_;
  const br = (text) => h(text).replace(/\n/g, '<br>');
  const p = jstParts_(inv.date);
  const reduced = inv.lines.some((l) => l.rate === 8);
  const qtyText = (q) => (Number.isInteger(q) ? yen_(q) : String(q));
  const items = inv.lines.map((l) => `
      <tr>
        <td>${h(l.item)}${l.rate === 8 ? ' ※' : ''}</td>
        <td class="r">${qtyText(l.qty)}</td>
        <td class="r">${yen_(l.price)}</td>
        <td class="r">${yen_(l.amount)}</td>
      </tr>`).join('');
  const rateRows = inv.totals.byRate.map((x) => `
      <tr><td>${x.rate}%対象（税抜）</td><td class="r">${yen_(x.base)}円</td></tr>
      <tr><td>消費税（${x.rate}%）</td><td class="r">${yen_(x.tax)}円</td></tr>`).join('');
  const issuer = [
    `<div class="company">${h(s.company)}</div>`,
    s.address ? `<div>${br(s.address)}</div>` : '',
    s.tel ? `<div>TEL ${h(s.tel)}</div>` : '',
    `<div>登録番号 ${h(s.regno)}</div>`,
  ].join('');
  const to = inv.client.person
    ? `${h(inv.client.name)}<br>${h(inv.client.person)} 様`
    : `${h(inv.client.name)} 御中`;
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  @page { size: A4; margin: 16mm 15mm; }
  body { font-family: sans-serif; color: #222; font-size: 10pt; line-height: 1.5; }
  table { border-collapse: collapse; width: 100%; }
  td, th { vertical-align: top; }
  .title { text-align: center; font-size: 22pt; font-weight: bold; letter-spacing: 14pt; padding-bottom: 4pt; border-bottom: 2.5pt solid #1f3864; }
  .fix { text-align: center; color: #555; font-size: 9pt; margin-top: 3pt; }
  .to { font-size: 14pt; border-bottom: 1pt solid #222; padding-bottom: 3pt; }
  .meta td { padding: 1pt 0; }
  .meta .k { color: #555; width: 70pt; }
  .company { font-size: 12pt; font-weight: bold; margin-top: 10pt; }
  /* PDFにすると背景の色は出ないので、色は文字と線だけで付ける */
  .amount { margin-top: 12pt; border: 1.5pt solid #1f3864; }
  .amount td { padding: 7pt 10pt; vertical-align: middle; }
  .amount .k { color: #1f3864; font-weight: bold; width: 120pt; }
  .amount .v { font-size: 18pt; font-weight: bold; text-align: right; }
  .items { margin-top: 16pt; }
  .items th { color: #1f3864; font-weight: bold; padding: 5pt 6pt; border: 0.75pt solid #b7b7b7; border-bottom: 1.5pt solid #1f3864; }
  .items td { padding: 5pt 6pt; border: 0.75pt solid #b7b7b7; }
  .r { text-align: right; white-space: nowrap; }
  .sum { width: 52%; margin-left: 48%; margin-top: 8pt; }
  .sum td { padding: 4pt 6pt; border: 0.75pt solid #b7b7b7; }
  .sum .total td { font-weight: bold; font-size: 11pt; border-top: 1.5pt solid #1f3864; }
  .box { border: 0.75pt solid #b7b7b7; padding: 7pt 9pt; margin-top: 12pt; }
  .box .k { font-weight: bold; margin-bottom: 2pt; }
  .small { color: #555; font-size: 8.5pt; margin-top: 4pt; }
</style></head><body>
  <div class="title">請求書</div>
  ${inv.replaces.length ? `<div class="fix">（${h(inv.replaces.join('・'))} の訂正版です。先にお送りした請求書は破棄をお願いいたします）</div>` : ''}
  <table style="margin-top:14pt"><tr>
    <td style="width:56%;padding-right:16pt">
      <div class="to">${to}</div>
      <div style="margin-top:8pt">下記のとおりご請求申し上げます。</div>
      <table class="amount"><tr><td class="k">ご請求金額（税込）</td><td class="v">${yen_(inv.totals.total)} 円</td></tr></table>
      <table class="meta" style="margin-top:6pt">
        <tr><td class="k">お支払期限</td><td>${slashDate_(inv.due)}</td></tr>
        ${inv.period ? `<tr><td class="k">対象期間</td><td>${slashDate_(inv.period.start)}〜${slashDate_(inv.period.end)}</td></tr>`
    : `<tr><td class="k">取引日</td><td>${slashDate_(inv.date)}</td></tr>`}
      </table>
    </td>
    <td style="width:44%">
      <table class="meta">
        <tr><td class="k">請求日</td><td>${p.y}年${p.m}月${p.d}日</td></tr>
        <tr><td class="k">請求書番号</td><td>${h(inv.no)}</td></tr>
      </table>
      ${issuer}
    </td>
  </tr></table>
  <table class="items">
    <tr><th style="text-align:left">品目</th><th style="width:50pt">数量</th><th style="width:80pt">単価（円）</th><th style="width:90pt">金額（円）</th></tr>
    ${items}
  </table>
  <table class="sum">
    <tr><td>小計（税抜）</td><td class="r">${yen_(inv.totals.base)}円</td></tr>
    ${rateRows}
    <tr class="total"><td>合計（税込）</td><td class="r">${yen_(inv.totals.total)}円</td></tr>
  </table>
  ${reduced ? '<div class="small" style="text-align:right">※は軽減税率（8%）の対象です</div>' : ''}
  <div class="box"><div class="k">お振込先</div>${br(s.bank)}</div>
  ${s.remarks ? `<div class="box"><div class="k">備考</div>${br(s.remarks)}</div>` : ''}
</body></html>`;
}

/** セルに文字として入れる（= や + で始まる名前が、数式として動かないように） */
function textCell_(v) {
  const s = String(v);
  return /^[=+\-@]/.test(s) ? `'${s}` : s;
}

/** ファイル名を見せるリンク（URLではなく、人が見て分かる名前で） */
function linkFormula_(url, label) {
  return `=HYPERLINK("${url}","${String(label).replace(/"/g, '""')}")`;
}

// ---- 取消 ----

/** メニュー：選んだ行の請求書を取消して、明細を作り直せる状態に戻す */
function cancelSelectedInvoice() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = ss.getActiveSheet();
    const pick = '「請求書一覧」で、取消す請求書の行（どのセルでも）を選んでから押してください';
    if (sheet.getName() !== SHEET.INVOICES) throw new UserError(pick);
    const invoices = openTable_(SHEET.INVOICES);
    const selected = sheet.getActiveRange().getRow();
    if (selected <= invoices.headerRow) throw new UserError(pick);
    const inv = invoices.read(selected);
    const no = String(inv['請求書番号']).trim();
    if (!no) throw new UserError(pick);
    if (inv['状態'] === STATUS.CANCELED) throw new UserError(`${no} はもう取消してあります`);
    if (inv['状態'] === STATUS.PAID) throw new UserError(`${no} は入金済みです。取消すときは、先に「状態」を「未入金」に戻してください`);

    const answer = ui.prompt(`${no}（${inv['取引先']}・${yen_(toNumber_(inv['請求額（円）']) || 0)}円）を取消しますか？`,
      '取消す理由を書いて「OK」を押してください（例：金額の誤り）。\n\n・請求書のPDFは消さず、名前に「取消」を付けます\n・Gmailの下書きが残っていれば消します\n・次の画面で、明細を作り直すかどうかを選びます',
      ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() !== ui.Button.OK) return;
    const reason = answer.getResponseText().trim();
    if (!reason) throw new UserError('取消す理由を書いてください（あとで一覧を見た人が分かるように）');
    const again = ui.alert('明細を作り直しますか？',
      '「OK」：直して作り直す（金額の誤りなど）。明細の番号が消えて黄色になります。直してから「請求書を作る」を押すと、新しい番号の訂正版ができます\n\n「いいえ」：作り直さない（取引がなくなった、など）。明細は「取消」の印を付けて、作る対象から外します',
      ui.ButtonSet.YES_NO_CANCEL);
    if (again !== ui.Button.YES && again !== ui.Button.NO) return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) throw new UserError('ほかの人が請求書を作っているか、入金を取り込んでいます。終わってから、もう一度押してください');
    let result;
    try {
      result = cancelInvoice_(no, reason, again === ui.Button.YES);
    } finally {
      lock.releaseLock();
    }
    const lines = [again === ui.Button.YES
      ? `明細 ${result.lines}行 を作り直せる状態に戻しました（「明細」の黄色い行）。直してから「請求書を作る」を押してください。`
      : `明細 ${result.lines}行 に「取消」の印を付けました。この明細から請求書は作りません。`];
    lines.push(result.draft === 'deleted' ? 'Gmailの下書きを消しました。'
      : result.draft === 'missing' ? 'Gmailの下書きは見つかりませんでした（送ったあと・消したあと）。送っていた場合は、取引先に取消の連絡をしてください。' : '');
    if (result.fileError) lines.push('PDFの名前を変えられませんでした（フォルダの編集権限を確かめてください）。一覧では取消になっています。');
    ui.alert(`${no} を取消しました`, lines.filter(Boolean).join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    showError_(e);
  }
}

/**
 * 請求書を取消す（確認のあと）。行は番号で探し直す（確認画面を出している間に、並べ替えられても大丈夫なように）
 * 一覧を先に「取消」にする（入金の照らし合わせの対象から、すぐに外すため）
 * rebuild: true なら明細を作り直せる状態に（番号を消し、隠し列に取消した番号を残す）、false なら明細の番号に「（取消）」を付けて作る対象から外す
 * 戻り値: { lines（明細の行数）, draft（'deleted'・'missing'・'none'）, fileError }
 */
function cancelInvoice_(no, reason, rebuild) {
  const invoices = openTable_(SHEET.INVOICES);
  const inv = invoices.rows().find((r) => r['請求書番号'] === no);
  if (!inv) throw new UserError(`${no} が「請求書一覧」に見つかりません`);
  if (inv['状態'] !== STATUS.UNPAID) throw new UserError(`${no} は「${inv['状態']}」になっています。取消せるのは未入金の請求書だけです`);
  const statusCell = invoices.sheet.getRange(inv._row, invoices.map['状態']);
  statusCell.clearDataValidations();
  statusCell.setValue(STATUS.CANCELED);
  const stamp = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd');
  appendCorrection_(invoices, no, `取消：${reason}（${stamp}）`);
  invoices.paintRow(inv._row, null);

  const lines = openTable_(SHEET.LINES);
  let count = 0;
  lines.rows().forEach((r) => {
    if (String(r['請求書番号']).trim() !== no) return;
    if (rebuild) {
      lines.write(r._row, '請求書番号', '');
      lines.write(r._row, '管理用', `取消:${no}`);
      lines.paintRow(r._row, WARN_ROW_COLOR, cancelNoteText_(no));
    } else {
      lines.write(r._row, '請求書番号', `${no}${CANCELED_SUFFIX}`);
      lines.paintRow(r._row, null);
    }
    count++;
  });

  let meta = {};
  try { meta = JSON.parse(inv['管理用'] || '{}'); } catch (e) { meta = {}; }
  let fileError = false;
  if (meta.file) {
    try {
      const file = DriveApp.getFileById(meta.file);
      const name = `【取消】${file.getName()}`;
      file.setName(name);
      invoices.write(inv._row, 'PDF', linkFormula_(file.getUrl(), name));
    } catch (e) {
      Logger.log(e.stack || e);
      fileError = true;
    }
  }
  let draft = 'none';
  if (meta.draft) {
    try {
      GmailApp.getDraft(meta.draft).deleteDraft();
      draft = 'deleted';
    } catch (e) {
      draft = 'missing'; // 送ったあとの下書きは、見つからない
    }
  }
  return { lines: count, draft, fileError };
}
