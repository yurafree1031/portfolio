/**
 * スプレッドシートの共通部品（どのデモでも共通）
 * 使う側のデモで、次の定数を決めておく
 *   HEADER_COLOR: 表の見出しの背景（見た目ルールの薄いグレー）
 *   PROTECT_NOTE: 見出しの行に付ける保護（消そうとすると警告が出る）の名前
 */

/** 表の見出しの見た目（太字・薄いグレー・下に罫線） */
function styleHeader_(range) {
  range.setFontWeight('bold').setBackground(HEADER_COLOR).setBorder(null, null, true, null, null, null);
}

/** 見出しの行を探す（上から20行の中で、label と同じ文字のセルがある行）。なければ 0 */
function findHeaderRow_(sheet, label) {
  const rows = Math.min(sheet.getLastRow(), 20);
  const cols = sheet.getLastColumn();
  if (rows === 0 || cols === 0) return 0;
  const i = sheet.getRange(1, 1, rows, cols).getValues().findIndex((r) => r.some((v) => String(v).trim() === label));
  return i + 1;
}

/** 見出しの名前 → 列番号 */
function headerMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const map = {};
  sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].forEach((v, i) => {
    const name = String(v).trim();
    if (name && !map[name]) map[name] = i + 1;
  });
  return map;
}

/** 見出しの行に、消そうとすると警告が出る保護を付ける（1つだけ。行が動いても保護は一緒に動く） */
function protectHeader_(sheet, range) {
  try {
    const exists = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).some((p) => p.getDescription() === PROTECT_NOTE);
    if (!exists) range.protect().setDescription(PROTECT_NOTE).setWarningOnly(true);
  } catch (e) {
    Logger.log(e.stack || e); // 保護を付けられなくても、動きには関係しない
  }
}
