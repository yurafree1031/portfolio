/**
 * 自動で動く処理（トリガー）のエラーを知らせる部品（どのデモでも共通）
 * トリガーのエラーは画面に出せないので、管理者へのメールと、シートの上の赤い帯で知らせる
 * 使う側のデモで、TIME_ZONE（日付を決めるタイムゾーン）と ERROR_ROW_COLOR（薄い赤）を決めておく
 */

const ALERT_SENT_PREFIX_ = 'ALERT_SENT_'; // 同じエラーのメールを1日1通にするための記録（スクリプトプロパティ）

/**
 * 管理者にメールで知らせる。key が同じものは1日1通まで（key が null なら毎回送る）
 * 送れたら true。管理者のメールアドレスがない・送信数の上限・送信の失敗なら false（呼んだ側で、ほかの方法で知らせる）
 */
function sendAdminAlert_({ to, key, subject, lines, senderName }) {
  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
    if (key && props.getProperty(ALERT_SENT_PREFIX_ + key) === today) return true;
    const admin = String(to || '').trim();
    if (!/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(admin) || MailApp.getRemainingDailyQuota() < 1) return false;
    const url = SpreadsheetApp.getActive().getUrl();
    MailApp.sendEmail({
      to: admin,
      subject,
      body: lines.concat(['', `スプレッドシート: ${url}`, 'エラー番号ごとの直し方は、説明書の「エラー番号の一覧」を見てください。']).join('\n'),
      name: senderName,
    });
    if (key) props.setProperty(ALERT_SENT_PREFIX_ + key, today);
    return true;
  } catch (e) {
    Logger.log(`管理者に知らせられません: ${e.stack || e}`);
    return false;
  }
}

/** シートの row 行目に、日時つきの赤い帯でエラーを書く（空文字なら消す） */
function setAlertBand_(sheet, row, text) {
  try {
    const band = sheet.getRange(row, 1, 1, sheet.getMaxColumns());
    if (!text) {
      if (sheet.getRange(row, 1).getValue() !== '') band.clearContent().setBackground(null);
      return;
    }
    const stamp = Utilities.formatDate(new Date(), TIME_ZONE, 'M/d HH:mm');
    sheet.getRange(row, 1).setValue(`${stamp}　${text.replace(/\n/g, '　')}`);
    band.setBackground(ERROR_ROW_COLOR);
  } catch (e) {
    Logger.log(e.stack || e); // 書けなくても、管理者へのメールで分かる
  }
}
