/**
 * 毎朝の自動の処理（トリガー。「自動の処理を始める」で入り、「自動の処理を止める」で外れる）
 *
 * dailyRun（毎朝 CONFIG.DAILY_HOUR 時ごろ）:
 *   - その月の「毎月の請求」の明細を、まだ入れていなければ「明細」に入れる（請求日は、その月の締め日）
 *     （1日に動けなかったときも、次に動いたときに入れる。入れた月はスクリプトプロパティに残し、2回入れない）
 *   - 前に入れられなかった「毎月の請求」の行が直っていれば、その月の明細として入れる
 * エラーは画面に出せないので、管理者にメールで知らせ、「明細」の上の帯にも書く。メールで知らせられないときだけ、エラーにして終える
 * （トリガーの持ち主に、Googleから失敗のお知らせが届くように）
 */

// ---- メニュー ----

/**
 * 管理者用メニュー：自動の処理を始める（再開する）
 * 今月の毎月の明細をまだ入れていなければ、今入れるかを選んでもらう（月の途中で始めた・止めていた間に月が替わった、など）
 */
function startAuto() {
  const ui = SpreadsheetApp.getUi();
  try {
    requireTimeZone_();
    ensureAllSheets_();
    const s = readSettings_();
    requireSettings_(s, ['admin']);
    const props = PropertiesService.getScriptProperties();
    const me = Session.getEffectiveUser().getEmail();
    const sender = props.getProperty(PROP.SENDER);
    const running = sender && props.getProperty(PROP.STOPPED) !== '1';
    if (running && sender !== me) {
      const answer = ui.alert('送信元を替えますか？',
        `今は ${sender} から、お知らせのメールを送っています。\n替えると、これからは ${me}（あなた）から送ります。\n${sender} のアカウントがまだ使えるなら、先にそのアカウントで「自動の処理を止める」を押してください。`,
        ui.ButtonSet.YES_NO);
      if (answer !== ui.Button.YES) return;
    }

    const now = new Date();
    const p = jstParts_(now);
    const ym = Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM');
    let prepared = null;
    if (props.getProperty(PROP.MONTHLY_DONE) !== ym) {
      const count = openTable_(SHEET.MONTHLY).rows().length;
      if (count) {
        const answer = ui.alert(`${p.m}月の毎月の明細を入れますか？`,
          `「毎月の請求」の ${count}行 を、${p.m}月の明細として「明細」に入れます（請求日は、取引先ごとの${p.m}月の締め日）。\n\n「OK」で今入れます。「いいえ」なら、${p.m}月の分は入れず、来月1日から入れます（すでに${p.m}月分を手で入れたときは「いいえ」）。`,
          ui.ButtonSet.YES_NO_CANCEL);
        if (answer !== ui.Button.YES && answer !== ui.Button.NO) return;
        if (answer === ui.Button.YES) prepared = prepareMonthly_(p.y, p.m);
      }
      props.setProperty(PROP.MONTHLY_DONE, ym);
    }

    props.setProperty(PROP.SENDER, me);
    props.deleteProperty(PROP.STOPPED);
    removeMyTriggers_();
    ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1).atHour(CONFIG.DAILY_HOUR).inTimezone(TIME_ZONE).create();
    writeSetting_('state', `動いています（毎朝${CONFIG.DAILY_HOUR}時ごろ。お知らせは ${me} から送ります）`);
    setLinesAlert_('');

    const lines = [];
    if (prepared) lines.push(`${p.m}月の毎月の明細を ${prepared.added}行 入れました。${prepared.problems ? `入れられなかった行が ${prepared.problems}行 あります（「毎月の請求」の赤い行）。` : ''}`, '');
    lines.push(
      `毎朝${CONFIG.DAILY_HOUR}時ごろに動きます。`,
      '・毎月1日：「毎月の請求」の明細を「明細」に入れます',
      `エラーのお知らせは ${me}（あなた）から管理者に送ります。`,
    );
    ui.alert('自動の処理を始めました', lines.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    showError_(e);
  }
}

/** 管理者用メニュー：自動の処理を止める */
function stopAuto() {
  const ui = SpreadsheetApp.getUi();
  try {
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty(PROP.SENDER) || props.getProperty(PROP.STOPPED) === '1') {
      ui.alert('自動の処理は止まっています', '始めるときは「自動の処理を始める（管理者）」を押してください。', ui.ButtonSet.OK);
      return;
    }
    const answer = ui.alert('自動の処理を止めますか？',
      '止めている間は、毎月の明細が入りません。\n再開するときに、その月の毎月の明細を入れるかどうかを選べます（止めていた間のお知らせは送りません）。',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
    // 止める印は全員で共有する（ほかの人のトリガーが残っていても、印を見て何もしない）
    props.setProperty(PROP.STOPPED, '1');
    removeMyTriggers_();
    writeSetting_('state', '止まっています');
    setLinesAlert_('');
    ui.alert('自動の処理を止めました', '再開するときは「自動の処理を始める（管理者）」を押してください。', ui.ButtonSet.OK);
  } catch (e) {
    showError_(e);
  }
}

/** このアカウントのトリガーを外す（ほかの人のトリガーは、Googleの仕組みで見えないので外せない） */
function removeMyTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => TRIGGER_HANDLERS.includes(t.getHandlerFunction()))
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

// ---- 毎朝 ----

function dailyRun() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP.STOPPED) === '1' || !props.getProperty(PROP.SENDER)) return;
  const me = Session.getEffectiveUser().getEmail();
  if (me && me !== props.getProperty(PROP.SENDER)) {
    removeMyTriggers_(); // 送信元が替わった。前の送信元に残っていたトリガーを片づける
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) return; // ボタンの処理の途中。明日、まとめて確かめる（毎月の明細と、過ぎたお知らせの日も拾う）
  let s = { admin: '' };
  try {
    s = readSettings_();
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    if (tz !== TIME_ZONE) throw new AdminError('E28', `スプレッドシートのタイムゾーンが「${tz}」になっています`, '日付が1日ずれるので、「ファイル → 設定」の「タイムゾーン」を「(GMT+09:00) 東京」にしてください');
    if (s.problems.length) throw new AdminError('E27', '「設定」シートを直してください', s.problems.join('　'));
    const now = new Date();
    const p = jstParts_(now);
    const ym = Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM');
    let monthly = null;
    if (props.getProperty(PROP.MONTHLY_DONE) !== ym) {
      monthly = prepareMonthly_(p.y, p.m);
      props.setProperty(PROP.MONTHLY_DONE, ym);
    }
    // 前に入れられなかった行が直っていれば、その月の分として入れる（今入れた月の赤い行は、今日はもう見ない）
    const failed = monthly ? monthly.problems : retryMonthly_().problems;
    if (failed) {
      throw new AdminError('E27', `「毎月の請求」の ${failed}行 を、明細に入れられませんでした`, '「毎月の請求」の赤い行の「確認事項」のとおりに直してください。直すと、次の朝（または「請求書を作る」を押したとき）に、その月の明細として入ります');
    }
    setLinesAlert_('');
  } catch (e) {
    Logger.log(e.stack || e);
    setLinesAlert_(errorText_(e));
    // 管理者へのメールで知らせられなかったときだけ、エラーにして終える（Googleから失敗のお知らせが届くように）
    const alerted = sendAdminAlert_({
      to: s.admin, key: e.code || 'E99', subject: '【請求書ツール】自動の処理でエラーが起きました', senderName: CONFIG.MAIL_SENDER_NAME,
      lines: [errorText_(e), ...(e.raw ? ['', `詳しい内容: ${e.raw}`] : [])],
    });
    if (!alerted) throw e;
  } finally {
    lock.releaseLock();
  }
}

/** 「明細」の上の帯に、自動の処理のエラーを書く（空文字なら消す） */
function setLinesAlert_(text) {
  try {
    const t = openTable_(SHEET.LINES);
    setAlertBand_(t.sheet, t.spec.alertRow, text);
  } catch (e) {
    Logger.log(e.stack || e);
  }
}

const MONTHLY_FAIL = /（(\d{4})年(\d{1,2})月の明細に入れられませんでした）$/;

/**
 * 「毎月の請求」の明細を、y年m月の明細として「明細」に入れる
 * 請求日は取引先ごとのその月の締め日。品目の {月} は「10月」に置き換える。入れた行には、隠し列に「毎月」の印を付ける
 * 入れられない行は赤くし、「（2026年10月の明細に入れられませんでした）」と書く（直すと retryMonthly_ でその月の分として入る）
 * 戻り値: { added（入れた行数）, problems（入れられなかった行数） }
 */
function prepareMonthly_(y, m) {
  return addMonthlyLines_(openTable_(SHEET.MONTHLY).rows().map((r) => ({ r, y, m })));
}

/** 前に入れられなかった「毎月の請求」の行のうち、直ったものを、その月の明細として入れる（毎朝と「請求書を作る」のときに呼ぶ） */
function retryMonthly_() {
  const targets = [];
  openTable_(SHEET.MONTHLY).rows().forEach((r) => {
    const mark = String(r['確認事項']).match(MONTHLY_FAIL);
    if (mark) targets.push({ r, y: Number(mark[1]), m: Number(mark[2]) });
  });
  return targets.length ? addMonthlyLines_(targets) : { added: 0, problems: 0 };
}

function addMonthlyLines_(targets) {
  const monthly = openTable_(SHEET.MONTHLY);
  const lines = openTable_(SHEET.LINES);
  const clients = readClients_().byName;
  const out = [];
  let problems = 0;
  targets.forEach(({ r, y, m }) => {
    const name = String(r['取引先']).trim();
    const c = clients[nameKey_(name)];
    const qty = r['数量'] === '' ? 1 : toNumber_(r['数量']);
    const price = toNumber_(r['単価（円）']);
    const rate = r['税率'] === '' ? 10 : taxRateOf_(r['税率']);
    const item = String(r['品目']).trim().replace(/\{月\}/g, `${m}月`);
    const problem = !name ? '取引先を選んでください'
      : !c ? '「取引先」シートにない取引先です'
        : clientTermsProblem_(c) || (!item ? '品目を入れてください' : !(qty > 0) ? '数量は0より大きい数字にしてください'
          : isNaN(price) ? '単価は数字にしてください' : !rate ? '税率は 10% か 8% を選んでください' : '');
    if (problem) {
      monthly.paintRow(r._row, ERROR_ROW_COLOR, `${problem}（${y}年${m}月の明細に入れられませんでした）`);
      problems++;
      return;
    }
    if (r['確認事項'] !== '') monthly.paintRow(r._row, null);
    const date = jstDate_(y, m, dayOf_(c.closing, y, m));
    out.push({ date, due: dueDate_(date, c), client: c.name, item, qty, price, rate: `${rate}%` });
  });
  if (out.length) {
    const start = lines.lastRow() + 1;
    const width = Math.max(...Object.values(lines.map));
    const values = out.map((o) => {
      const v = new Array(width).fill('');
      v[lines.map['請求日'] - 1] = o.date;
      v[lines.map['取引先'] - 1] = textCell_(o.client);
      v[lines.map['品目'] - 1] = textCell_(o.item);
      v[lines.map['数量'] - 1] = o.qty;
      v[lines.map['単価（円）'] - 1] = o.price;
      v[lines.map['税率'] - 1] = o.rate;
      v[lines.map['支払期限'] - 1] = o.due;
      v[lines.map['管理用'] - 1] = withDateMark_(LINE_MARK.MONTHLY, DATE_MARK_KEYS.DUE, ymd_(o.due));
      return v;
    });
    lines.sheet.getRange(start, 1, values.length, width).setValues(values);
    for (let i = 0; i < values.length; i++) lines.paintRow(start + i, null);
  }
  return { added: out.length, problems };
}
