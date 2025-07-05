/**
 * スクリプトプロパティから値を取得
 * @param {string} key - プロパティキー
 * @return {string} プロパティ値
 * @throws {Error} プロパティが存在しない場合
 */
function getScriptProperty(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value === null) {
    throw new Error(`Script property not found: ${key}`);
  }
  return value;
}
/**
 * プレーンテキストを取得
 */
function fetchText(url) {
  const res  = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true});
  const code = res.getResponseCode();
  if (code !== 200) throw new Error(`Markdown 取得失敗 (${code}) → ${url}`);
  return res.getContentText();
}

/* ---------- ユーティリティ ---------- */

/**
 * Base64 → UTF-8 文字列（BOM 除去）
 */
function base64ToUtf8(b64) {
  return Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('utf-8').replace(/^\uFEFF/, '').trim();
}

/**
 * 長い文字列を固定幅で分割
 * @param {string} str
 * @param {number} size
 * @return {string[]}
 */
function chunkString(str, size) {
  const out = [];
  for (let i = 0; i < str.length; i += size) out.push(str.slice(i, i + size));
  return out;
}

/**
 * 今月の残営業日（Business Days）を返す
 * @param {Boolean} [includeToday=true]  trueなら今日を営業日に含める
 * @return {Number}  残営業日数
 *
 * スプレッドシート関数としても
 * =REMAINING_BIZ_DAYS()            // 今日を含む
 * =REMAINING_BIZ_DAYS(FALSE)       // 今日を含めない
 * のように呼び出せる
 * todayのサンプル
 * const today = new Date('2025-04-29T00:00:00+09:00'); // 時間を 00:00:00 に丸める
 * const today = new Date(); // 時間を 00:00:00 に丸める
 */
function getRemainingBizDays(baseDate, includeToday = true) {
  // ---- ここが追加ポイント ----
  const start = new Date(baseDate);
  start.setHours(0, 0, 0, 0);          // どんな入力でも 00:00 JST に揃える
  // --------------------------------

  const y = start.getFullYear();
  const m = start.getMonth();
  const last = new Date(y, m + 1, 0);  // 月末 00:00

  const cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  let biz = 0;

  for (let d = new Date(start); d <= last; d.setDate(d.getDate() + 1)) {
    if (!includeToday && d.getTime() === start.getTime()) continue;
    if (d.getDay() === 0 || d.getDay() === 6) continue;
    if (cal.getEventsForDay(d).length) continue;
    biz++;
  }
  return biz;
}
function getBizDaysThisMonth(anyDay = new Date()) {
  // 月初・月末
  const y = anyDay.getFullYear();
  const m = anyDay.getMonth();               // 0 = Jan
  const first = new Date(y, m, 1);           // 00:00 JST
  const last  = new Date(y, m + 1, 0);       // 00:00 JST

  // 日本の祝日カレンダー（読み取り専用）
  const cal = CalendarApp.getCalendarById(
               'ja.japanese#holiday@group.v.calendar.google.com');

  let biz = 0;
  for (let d = new Date(first); d <= last; d.setDate(d.getDate() + 1)) {
    const dow = d.getDay();                   // 0 Sun … 6 Sat
    if (dow === 0 || dow === 6) continue;     // 土日
    if (cal.getEventsForDay(d).length) continue; // 祝日
    biz++;
  }
  return biz;
}

// 日本時間の "本日" を「YYYY-MM-DD」形式の文字列で返す
function getTodayJSTString(date) {
  return date.toLocaleDateString('ja-JP', {
    timeZone: 'Asia/Tokyo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).replace(/\//g, '-');   // 2025/04/29 → 2025-04-29
} 

/**
 * 任意の日付（YYYY, MM, DD）を JST 00:00:00 の Date で返す
 * @param {number} y - 西暦年
 * @param {number} m - 月 (1-12)
 * @param {number} d - 日 (1-31)
 */
function makeJstDate(y, m, d) {
  // UTC ベースで「JST の真夜中」に相当する時刻を作成
  return new Date(Date.UTC(y, m - 1, d, 0, 0, 0)); // +09:00 が内蔵 UTC に差し引かれる
}


/**
 * ドキュメントをリフレッシュ → 新しいテキストで上書き
 */
function refreshDocAndWrite(text) {
  const DOC_ID = '1cNlEGtCtyWqlMtazvjc2SiMt2Tgs0mg1iKjLNegU7Zc';  // 対象ドキュメント

  const doc  = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();

  // 1) 既存内容をまるごと消す
  body.clear();

  body.appendParagraph(text);

  doc.saveAndClose();   // 念のため保存して閉じる
}

// ── 共通フォーマッタ（JST 固定で YYYY-MM-DD）─────────────
const fmtJST = d =>
  d.toLocaleDateString('ja-JP', {
    timeZone: 'Asia/Tokyo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).replace(/\//g, '-');

// ── ① きょうの月初（今月 1 日）──────────────────────
function thisMonthStartJST(date = new Date()) {
  const d = new Date(date);            // ← ★コピー
  const y = d.getFullYear();
  const m = d.getMonth();            // ±n か月
  const jstFirst = new Date(Date.UTC(y, m, 1));  ;
  return fmtJST(jstFirst);
}

// ── ② 前月の月初（先月 1 日）──────────────────────
function lastMonthStartJST(date = new Date()) {
  const d = new Date(date);            // ← ★コピー
  const y = d.getFullYear();
  const m = d.getMonth() - 1;            // ±n か月
  const jstFirst = new Date(Date.UTC(y, m, 1));  ;
  return fmtJST(jstFirst);
}

// ── ③ 前前月の月初（先々月 1 日）─────────────────
function twoMonthsAgoStartJST(date = new Date()) {
  const d = new Date(date);            // ← ★コピー
  const y = d.getFullYear();
  const m = d.getMonth() - 2;            // ±n か月
  const jstFirst = new Date(Date.UTC(y, m, 1));
  return fmtJST(jstFirst);
}

/**
 * スプレッドシートのセルから日付をパースします。
 * 以下の形式に対応します：
 * - yyyy/MM/dd 形式の文字列
 * - Excelのシリアル値
 * - Dateオブジェクト
 * JSTとして解釈します。
 * @param {string|number|Date} dateValue - 日付の文字列、数値、またはDateオブジェクト。
 * @return {Date|null} パースされたDateオブジェクト。無効な場合はnull。
 */
function parseDateFromSheetValue(dateValue) {
  if (dateValue === null || dateValue === undefined || dateValue === '') {
    return null;
  }

  // Dateオブジェクトの場合はそのまま返す
  if (dateValue instanceof Date) {
    return dateValue;
  }

  if (typeof dateValue === 'string') {
    // 日付文字列の正規化（スラッシュを統一）
    const normalizedDate = dateValue.replace(/[-\/]/g, '/');
    
    // yyyy/MM/dd 形式
    if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(normalizedDate)) {
      const parts = normalizedDate.split('/');
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const day = parseInt(parts[2], 10);
      
      // 日付の妥当性チェック
      const date = new Date(year, month, day);
      if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
        return date;
      }
    }
    return null;
  }

  if (typeof dateValue === 'number') {
    // Excelシリアル値の場合 (1899/12/30 が基準日)
    const excelEpoch = new Date(1899, 11, 30); // 1899年12月30日
    const millisecondsPerDay = 24 * 60 * 60 * 1000;
    const date = new Date(excelEpoch.getTime() + dateValue * millisecondsPerDay);
    
    // 日付の妥当性チェック
    if (!isNaN(date.getTime())) {
      return date;
    }
  }

  return null;
}