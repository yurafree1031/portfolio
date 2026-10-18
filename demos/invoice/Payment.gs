/**
 * 入金を取り込んで、請求書と照らし合わせる
 *
 * 「入金を取り込む」ボタン（importPayments）:
 *   決めたフォルダのCSV（まだ読んでいないもの・更新されたもの）を読む → 入金の行だけを「入金」に足す
 *   （期間が重なるCSVでも、同じ入金を2回足さない） → 1件ずつ請求書を探す
 *   → 振込名義と金額で確かなものは「入金済み」に。迷うもの（手数料を引いた・名義が違う・同じ金額が2枚）は、候補を入れて黄色に
 *   順番：先に「入金」に行を書き（確かなものは「反映の途中」の印つき）、そのあと請求書を1件ずつ入金済みにして、行を「入金済みにしました」に変える
 *   （途中で止まっても、請求書だけが入金済みになって入金の記録がない、ということがないように。印の残った行は、次に押したときに続きを反映する）
 *   読めないCSV・日付や金額を読めない行は、ほかの入金を止めずに知らせ、そのCSVは次に押したときにもう一度読む
 * 「入金」の「確認」にチェック（onEdit → confirmPaymentRows_）:
 *   「請求書番号」の請求書を「入金済み」にする。番号が空なら「請求書の支払いではない」。チェックを外すと、反映を取り消す
 */

const PENDING_NOTE = '反映の途中で止まりました。もう一度「入金を取り込む」を押すと、続きを反映します';
const BUSY_NOTE = 'ほかの処理の途中だったため、反映できませんでした。少し待ってから、チェックを入れ直してください';

/** 「操作」シートのボタンから呼ばれる */
function importPayments() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) {
    ss.toast('ほかの人が請求書を作っているか、入金を取り込んでいます。終わってから、もう一度押してください', '確認してください', 20);
    return;
  }
  try {
    ss.toast('入金明細のCSVを読んでいます。このままお待ちください', '入金を取り込む', -1);
    const r = importPayments_(started);
    setRunResult_(RUN_CELLS.paymentResult, [r.text, ...r.problems].join('　'));
    ss.toast('', '', 1);
    if (r.problems.length) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('読めなかったものがあります', [r.text, '', ...r.problems].join('\n'), ui.ButtonSet.OK);
    } else ss.toast(r.text, '入金を取り込みました', 30);
  } catch (e) {
    setRunResult_(RUN_CELLS.paymentResult, `止まりました。${errorText_(e)}`);
    showError_(e);
  } finally {
    lock.releaseLock();
  }
}

/** 戻り値: { text（結果の文）, problems（読めなかったCSV・行の文） } */
function importPayments_(started) {
  requireTimeZone_();
  ensureAllSheets_();
  const s = readSettings_();
  requireSettings_(s, ['csvFolder', 'csvDate', 'csvName', 'csvAmount']);
  const folder = openFolder_(s.csvFolder, '入金明細のCSVを置くフォルダ', 'E22');
  const props = PropertiesService.getScriptProperties();
  let read = {};
  try { read = JSON.parse(props.getProperty(PROP.READ_FILES) || '{}'); } catch (e) { read = {}; }

  const payments = openTable_(SHEET.PAYMENTS);
  const have = {}; // すでに「入金」にある入金の数（日付|名義|金額 ごと）
  payments.rows().forEach((r) => { const k = String(r['管理用']); if (k) have[k] = (have[k] || 0) + 1; });

  const added = [];
  const problems = [];
  const done = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!/\.csv$/i.test(file.getName())) continue;
    const updated = file.getLastUpdated().getTime();
    if (read[file.getId()] === updated) continue;
    const parsed = depositsFromCsv_(decodeCsv_(file.getBlob()), s);
    if (parsed.error) { problems.push(`「${file.getName()}」：${parsed.error}。「設定」の列の名前を、このCSVの1行目（見出し）に合わせてください`); continue; }
    // 同じ入金がこのCSVに何件あるか。「入金」にある数より多い分だけ足す（期間が重なるCSVを続けて置いても、2回足さない）
    const counts = {};
    parsed.deposits.forEach((d) => {
      counts[d.key] = (counts[d.key] || 0) + 1;
      if (counts[d.key] > (have[d.key] || 0)) { added.push(d); have[d.key] = counts[d.key]; }
    });
    if (parsed.unreadable.length) {
      const ex = parsed.unreadable[0];
      problems.push(`「${file.getName()}」：入金の ${parsed.unreadable.length}行 の日付か金額を読めませんでした（例：${ex.line}行目「${ex.value}」）。その入金は「請求書一覧」で手で入金済みにしてください`);
    } else {
      done.push([file.getId(), updated]); // 全部読めたCSVだけ、読み終えた印を付ける（読めない行があるCSVは、次に押したときにもう一度読む）
    }
  }

  // 1. 先に「入金」に書く（確かなものは「反映の途中」の印つき）
  added.sort((a, b) => a.date - b.date);
  const invoices = openTable_(SHEET.INVOICES);
  const clients = readClients_().list;
  const unpaid = unpaidInvoices_(invoices);
  const reserved = new Set(); // ほかの入金の候補にした番号
  const taken = new Set(); // このあと自動で入金済みにする番号（ほかの入金とは照らし合わせない）
  const width = Math.max(...Object.values(payments.map));
  const startRow = payments.lastRow() + 1;
  let wait = 0;
  const values = added.map((d) => {
    const m = matchDeposit_(d, unpaid.filter((i) => !taken.has(i.no)), clients, s.fee, reserved);
    m.nos.forEach((no) => (m.auto ? taken : reserved).add(no));
    if (!m.auto) wait++;
    const v = new Array(width).fill('');
    const set = (h, x) => { v[payments.map[h] - 1] = x; };
    set('入金日', d.date);
    set('振込名義', textCell_(d.name));
    set('入金額（円）', d.amount);
    set('請求書番号', m.nos.join(', '));
    set('結果', RESULT.WAIT);
    set('確認', false);
    set('確認事項', m.auto ? PENDING_NOTE : m.note);
    set('管理用', d.key);
    return v;
  });
  if (values.length) {
    payments.sheet.getRange(startRow, 1, values.length, width).setValues(values);
    setCheckboxes_(payments, startRow, values.length);
    values.forEach((v, i) => payments.paintRow(startRow + i, WARN_ROW_COLOR, v[payments.map['確認事項'] - 1]));
  }
  done.forEach(([id, updated]) => { read[id] = updated; });
  props.setProperty(PROP.READ_FILES, JSON.stringify(trimRead_(read)));

  // 2. 確かなもの（今回の分と、前に途中で止まった分）を、1件ずつ請求書に反映する
  let applied = 0;
  let left = 0;
  payments.rows().filter((r) => r['結果'] === RESULT.WAIT && r['確認事項'] === PENDING_NOTE).forEach((r) => {
    if (Date.now() - started > CONFIG.TIME_BUDGET_MS) { left++; return; }
    const res = confirmPayment_(payments, r, s.fee, true);
    if (res === 'done') applied += String(r['請求書番号']).split(',').filter((x) => x.trim()).length;
  });

  if (!added.length && !applied && !left) {
    return { text: '新しい入金はありませんでした（読んだことのあるCSVと、入金のない行は飛ばします）', problems };
  }
  const parts = [`入金 ${added.length}件 を取り込みました。入金済みにした請求書 ${applied}枚`];
  if (wait) parts.push(`確認待ち ${wait}件（「入金」の黄色い行）`);
  if (left) parts.push(`残り ${left}件 は、もう一度押すと反映します`);
  return { text: parts.join('　'), problems };
}

/** 読んだCSVの記録は、新しいほうから500件だけ残す（記録の大きさの上限を超えないように） */
function trimRead_(read) {
  const entries = Object.entries(read).sort((a, b) => b[1] - a[1]).slice(0, 500);
  return Object.fromEntries(entries);
}

/** CSVの文字コード：UTF-8 で読めなければ、日本の銀行に多い Shift_JIS で読む */
function decodeCsv_(blob) {
  let text = blob.getDataAsString('UTF-8');
  if (text.includes('�')) text = blob.getDataAsString('Shift_JIS');
  return text.replace(/^﻿/, '');
}

/** CSVの文を行と列に分ける（"" で囲んだ列の中のカンマ・改行も扱う）。戻り値の行には、元の行番号（line）を付ける */
function parseCsv_(text) {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;
  let line = 1;
  let start = 1;
  const push = () => { row.push(cell); row.line = start; rows.push(row); row = []; cell = ''; start = line; };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') quoted = false; else { if (ch === '\n') line++; cell += ch; }
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(cell); cell = ''; } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      line++;
      push();
    } else cell += ch;
  }
  if (cell !== '' || row.length) push();
  return rows.filter((r) => r.some((c) => c.trim() !== ''));
}

/**
 * CSVから入金だけを取り出す（出金・金額が空の行は飛ばす）
 * 銀行のCSVは、見出しの前に口座名などの行があることがあるので、上から10行の中で見出しの行を探す
 * 入金額の列に何か入っているのに、日付か金額を読めない行は、読み飛ばさずに unreadable で返す
 * 戻り値: { deposits: [{ date, name, amount, key }], unreadable: [{ line, value }] } か { error }
 */
function depositsFromCsv_(text, s) {
  const rows = parseCsv_(text);
  const norm = (x) => String(x).normalize('NFKC').trim();
  const want = [s.csvDate, s.csvName, s.csvAmount].map(norm);
  const headerIndex = rows.slice(0, 10).findIndex((r) => want.every((w) => r.map(norm).includes(w)));
  if (headerIndex < 0) return { error: `「${s.csvDate}」「${s.csvName}」「${s.csvAmount}」の列が見つかりません` };
  const header = rows[headerIndex].map(norm);
  const [di, ni, ai] = want.map((w) => header.indexOf(w));
  const deposits = [];
  const unreadable = [];
  rows.slice(headerIndex + 1).forEach((r) => {
    const rawAmount = String(r[ai] == null ? '' : r[ai]).trim();
    if (rawAmount === '' || /^[0０]+$/.test(rawAmount)) return; // 出金の行
    const amount = toNumber_(rawAmount);
    const date = toDate_(r[di]);
    if (!(amount > 0) || !date) { unreadable.push({ line: r.line, value: !date ? r[di] : rawAmount }); return; }
    const name = String(r[ni] || '').trim();
    deposits.push({ date, name, amount, key: `${ymd_(date)}|${kanaKey_(name)}|${amount}` });
  });
  return { deposits, unreadable };
}

/** 未入金の請求書（期限の早い順）。{ no, client, total, due, row } */
function unpaidInvoices_(invoices) {
  return invoices.rows()
    .filter((r) => r['状態'] === STATUS.UNPAID && String(r['請求書番号']).trim())
    .map((r) => ({ no: String(r['請求書番号']).trim(), client: String(r['取引先']).trim(), total: toNumber_(r['請求額（円）']), due: toDate_(r['支払期限']) || new Date(8640000000000000), row: r._row }))
    .sort((a, b) => a.due - b.due);
}

/**
 * 振込名義を比べるための形
 * 半角カナ→全角、ひらがな→カタカナ、小さい字→大きい字（通帳では「シヨウジ」と書く）、空白・記号を除く
 * 会社の種類の略（カ)＝株式会社、ユ)＝有限会社、ド)＝合同会社 など）と、先頭の「振込」を除く
 */
function kanaKey_(s) {
  let t = String(s || '').normalize('NFKC');
  t = t.replace(/[ぁ-ゖ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60));
  t = t.replace(/^(振込|フリコミ)[*＊\s]*/, '');
  t = t.replace(/[(（](カ|ユ|ド|ザイ|シヤ|シャ|イ|ガク|ゴ|メ)[)）]?/g, '').replace(/(カ|ユ|ド|ザイ|シヤ|シャ|イ|ガク|ゴ|メ)[)）]/g, '');
  const small = { ァ: 'ア', ィ: 'イ', ゥ: 'ウ', ェ: 'エ', ォ: 'オ', ッ: 'ツ', ャ: 'ヤ', ュ: 'ユ', ョ: 'ヨ', ヮ: 'ワ', ヵ: 'カ', ヶ: 'ケ' };
  t = t.replace(/[ァィゥェォッャュョヮヵヶ]/g, (c) => small[c]);
  return t.replace(/[\s\-‐－ー―・.,、。()（）]/g, '').toUpperCase();
}

/** 入金の名義が、取引先の振込名義と合うか（銀行が長い名前を途中で切ることがあるので、先頭が合えばよい） */
function nameMatches_(depositName, kana) {
  const d = kanaKey_(depositName);
  const c = kanaKey_(kana);
  if (!d || !c) return false;
  return d.includes(c) || (d.length >= 6 && c.startsWith(d));
}

/**
 * 1件の入金に合う請求書を探す
 * unpaid: 未入金の請求書（期限の早い順）、clients: 取引先の一覧、fee: 手数料として見る差、reserved: この取り込みで、ほかの入金の候補にした番号
 * 戻り値: { auto（確かなので自動で入金済みにする）, nos: [候補の番号], note（確認待ちのときの説明） }
 */
function matchDeposit_(dep, unpaid, clients, fee, reserved) {
  const list = (xs) => xs.slice(0, 3).map((i) => `${i.no} ${yen_(i.total)}円`).join('、') + (xs.length > 3 ? ` ほか${xs.length - 3}枚` : '');
  const free = (xs) => xs.filter((i) => !reserved.has(i.no));
  const named = new Set(clients.filter((c) => c.kana && nameMatches_(dep.name, c.kana)).map((c) => nameKey_(c.name)));
  const pool = unpaid.filter((i) => named.has(nameKey_(i.client)));

  if (pool.length) {
    const who = [...new Set(pool.map((i) => i.client))].join('・');
    const exact = pool.filter((i) => i.total === dep.amount);
    if (exact.length === 1 && free(exact).length === 1) return { auto: true, nos: [exact[0].no], note: '' };
    if (exact.length >= 1) {
      const c = free(exact)[0] || exact[0];
      return { auto: false, nos: [c.no], note: `${who}の、同じ金額の未入金の請求書が${exact.length}枚あります（${list(exact)}）。期限のいちばん早い ${c.no} を候補にしました。合っていれば「確認」にチェックを入れてください` };
    }
    const sets = combinations_(free(pool).slice(0, CONFIG.MAX_COMBINE), (sum) => sum === dep.amount);
    if (sets.length === 1) return { auto: true, nos: sets[0].map((i) => i.no), note: '' };
    if (sets.length > 1) {
      return { auto: false, nos: sets[0].map((i) => i.no), note: `${who}の請求書の、何枚かをまとめた入金のようです。組み合わせが${sets.length}通りあるので、そのうち1つを候補にしました。合っていれば「確認」にチェックを入れてください` };
    }
    const short = free(pool).filter((i) => i.total - dep.amount > 0 && i.total - dep.amount <= fee);
    if (short.length) {
      const c = short[0];
      return { auto: false, nos: [c.no], note: `${c.no}（${yen_(c.total)}円）より ${yen_(c.total - dep.amount)}円 少ない入金です。振込手数料を引いた入金なら、「確認」にチェックを入れてください` };
    }
    const shortSets = combinations_(free(pool).slice(0, CONFIG.MAX_COMBINE), (sum) => sum - dep.amount > 0 && sum - dep.amount <= fee);
    if (shortSets.length) {
      const c = shortSets[0];
      const sum = c.reduce((n, i) => n + i.total, 0);
      return { auto: false, nos: c.map((i) => i.no), note: `${c.map((i) => i.no).join('・')} の合計（${yen_(sum)}円）より ${yen_(sum - dep.amount)}円 少ない入金です。まとめて振込手数料を引いた入金なら、「確認」にチェックを入れてください` };
    }
    return { auto: false, nos: [], note: `${who}からの入金ですが、金額の合う未入金の請求書がありません（未入金：${list(pool)}）。どの請求書の支払いか確かめて、「請求書番号」を入れてから「確認」にチェックを入れてください（金額に差があっても入金済みにでき、差は「確認事項」に残ります）` };
  }

  const exactAll = unpaid.filter((i) => i.total === dep.amount);
  if (exactAll.length === 1) {
    const c = exactAll[0];
    return { auto: false, nos: [c.no], note: `振込名義が「取引先」の振込名義（カナ）と合いませんが、金額が合う請求書があります（${c.client}）。合っていれば「確認」にチェックを入れてください。「取引先」の振込名義（カナ）を直しておくと、次から自動で照らし合わせます` };
  }
  if (exactAll.length > 1) {
    return { auto: false, nos: [], note: `金額が合う請求書が${exactAll.length}枚あります（${list(exactAll)}）。振込名義から取引先を確かめて、「請求書番号」を入れてから「確認」にチェックを入れてください` };
  }
  return { auto: false, nos: [], note: '合う請求書が見つかりません。請求書の支払いなら「請求書番号」を入れて、支払いでなければ（利息など）空のまま、「確認」にチェックを入れてください' };
}

/** 2枚以上の請求書の組み合わせのうち、合計が条件に合うもの（見る枚数は CONFIG.MAX_COMBINE まで） */
function combinations_(items, ok) {
  const out = [];
  const n = items.length;
  for (let mask = 1; mask < (1 << n); mask++) {
    const pick = items.filter((_, i) => mask & (1 << i));
    if (pick.length < 2) continue;
    if (ok(pick.reduce((s, i) => s + i.total, 0))) out.push(pick);
  }
  return out.sort((a, b) => a.length - b.length);
}

/**
 * 請求書を入金済みにする。入金額は1枚ずつ請求額で割り振り、差（振込手数料など）は最後の1枚で調整する
 * invs: [{ no, total }]
 */
function applyPayment_(invoices, invs, date, amount) {
  let left = amount;
  invs.forEach((inv, i) => {
    const r = invoices.rows().find((x) => String(x['請求書番号']).trim() === inv.no); // 行は番号で探し直す
    if (!r) return;
    const part = i === invs.length - 1 ? left : Math.min(inv.total, left);
    left -= part;
    invoices.write(r._row, '状態', STATUS.PAID);
    invoices.write(r._row, '入金日', date);
    invoices.write(r._row, '入金額（円）', part);
  });
}

/** 「請求書番号」の欄の番号（カンマ・読点・空白区切り） */
function paymentNos_(value) {
  return String(value).split(/[,、，\s・]+/).map((x) => x.trim()).filter(Boolean);
}

/**
 * 「入金」の「確認」が変わった行を反映する（onEdit から）
 * チェックを入れた確認待ちの行 → 請求書に反映。チェックを外した反映済みの行 → 反映を取り消す
 * ほかの処理（入金の取り込み・請求書の作成）の途中なら、黙って何もしないのではなく、チェックを戻して理由を書く
 */
function confirmPaymentRows_(sheet, range) {
  const t = openTable_(SHEET.PAYMENTS);
  const col = t.map['確認'];
  if (range.getColumn() > col || range.getLastColumn() < col) return;
  const first = Math.max(range.getRow(), t.headerRow + 1);
  const last = range.getRow() + range.getNumRows() - 1;
  const targets = [];
  for (let r = first; r <= last; r++) {
    const row = t.read(r);
    if (row['確認'] === true && row['結果'] === RESULT.WAIT) targets.push({ row, undo: false });
    if (row['確認'] === false && (row['結果'] === RESULT.DONE || row['結果'] === RESULT.NONE)) targets.push({ row, undo: true });
  }
  if (!targets.length) return;
  // ボタンの処理と同じロックを使う（取り込みの最中に手で確認した入金が、二重に反映されないように）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    targets.forEach(({ row, undo }) => {
      t.write(row._row, '確認', undo);
      if (!undo) t.paintRow(row._row, WARN_ROW_COLOR, BUSY_NOTE);
    });
    return;
  }
  try {
    let fee = 880;
    try { fee = readSettings_().fee; } catch (e) { Logger.log(e.stack || e); }
    targets.forEach(({ row, undo }) => {
      try {
        if (undo) undoPayment_(t, row, row['結果'] === RESULT.NONE ? 'チェックを外したので、確認待ちに戻しました'
          : 'チェックを外したので、反映を取り消しました（請求書を未入金に戻しました）。合う請求書を確かめて、もう一度「確認」にチェックを入れてください');
        else confirmPayment_(t, row, fee, false);
      } catch (e) {
        Logger.log(e.stack || e);
        t.write(row._row, '確認', undo);
        t.paintRow(row._row, WARN_ROW_COLOR, '反映できませんでした。もう一度チェックを入れ直し、それでも出るときは管理者に連絡してください');
      }
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1行を請求書に反映する。戻り値: 'done'・'none'（請求書の支払いではない）・'fail'
 * 番号がない・未入金でないときは、チェックを外して理由を書く
 * 金額に差があっても反映する（振込手数料の範囲なら色なし、それより大きい差は黄色で差額を残す）
 * auto: 取り込みのときの自動の反映（差のある入金は来ない）
 */
function confirmPayment_(t, row, fee, auto) {
  const nos = paymentNos_(row['請求書番号']);
  const fail = (note) => {
    t.write(row._row, '確認', false);
    t.write(row._row, '結果', RESULT.WAIT);
    t.paintRow(row._row, WARN_ROW_COLOR, note);
    return 'fail';
  };
  if (!nos.length) {
    if (auto) return fail('合う請求書が見つかりません。請求書の支払いなら「請求書番号」を入れて「確認」にチェックを入れてください');
    t.write(row._row, '結果', RESULT.NONE);
    t.paintRow(row._row, null);
    return 'none';
  }
  const invoices = openTable_(SHEET.INVOICES);
  const unpaid = unpaidInvoices_(invoices);
  const all = invoices.rows();
  const invs = [];
  for (const no of nos) {
    const inv = unpaid.find((i) => i.no === no);
    if (!inv) {
      const other = all.find((x) => String(x['請求書番号']).trim() === no);
      return fail(other ? `${no} は「${other['状態']}」です。未入金の請求書の番号を入れてください` : `${no} という請求書がありません。「請求書一覧」の番号を入れてください`);
    }
    invs.push(inv);
  }
  const amount = toNumber_(row['入金額（円）']);
  const sum = invs.reduce((n, i) => n + i.total, 0);
  const diff = sum - amount;
  applyPayment_(invoices, invs, toDate_(row['入金日']) || today_(), amount);
  t.write(row._row, '結果', RESULT.DONE);
  t.write(row._row, '確認', true);
  t.paintRow(row._row, null);
  if (diff > 0 && diff <= fee) t.write(row._row, '確認事項', `振込手数料 ${yen_(diff)}円 を引いた入金として反映しました`);
  else if (diff !== 0) {
    t.paintRow(row._row, WARN_ROW_COLOR, `請求額の合計（${yen_(sum)}円）より ${yen_(Math.abs(diff))}円 ${diff > 0 ? '少ない' : '多い'}入金のまま、入金済みにしました。差の扱い（残りの請求・返金など）を確かめ、済んだらこの文を消してください。番号を間違えたときは、チェックを外すと取り消せます`);
  }
  return 'done';
}

/** 反映した行を取り消す：請求書を未入金に戻し（入金日・入金額を消す）、行を確認待ちに戻す */
function undoPayment_(t, row, note) {
  const nos = paymentNos_(row['請求書番号']);
  if (row['結果'] === RESULT.DONE && nos.length) {
    const invoices = openTable_(SHEET.INVOICES);
    invoices.rows().forEach((r) => {
      if (!nos.includes(String(r['請求書番号']).trim()) || r['状態'] !== STATUS.PAID) return;
      invoices.write(r._row, '状態', STATUS.UNPAID);
      invoices.write(r._row, '入金日', '');
      invoices.write(r._row, '入金額（円）', '');
    });
  }
  t.write(row._row, '確認', false);
  t.write(row._row, '結果', RESULT.WAIT);
  t.paintRow(row._row, WARN_ROW_COLOR, note);
}

/** 請求書 no で反映した「入金」の行を、確認待ちに戻す（「請求書一覧」で未入金に戻されたとき） */
function reopenPaymentsOf_(no, note) {
  if (!no) return;
  const t = openTable_(SHEET.PAYMENTS);
  t.rows().filter((r) => r['結果'] === RESULT.DONE && paymentNos_(r['請求書番号']).includes(no)).forEach((r) => undoPayment_(t, r, note));
}
