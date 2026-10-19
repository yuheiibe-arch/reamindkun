/**
 * =================================================================
 * 【5_2_キュー掃除_トリガー管理.gs】（全体トリガー管理 ＆ キュー・履歴掃除）
 * ・連携キュー（⚙️_連携キュー）の処理
 * ・「⚙️_連携キュー」「⚙️_送信履歴」の過去48時間前データの自動削除（2行目以下）
 *   ※「⚙️_新規お久しぶり管理」は永続台帳として削除せず保持
 * ・定期実行トリガー（5分おき / 朝9時 / 夕方17時）の一括管理
 * =================================================================
 */

/**
 * 連携キュー（⚙️_連携キュー）を処理する関数（5分おきトリガー等で処理）
 * ＋ 連携キュー・送信履歴の48時間経過データを自動削除
 */
function processEventQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);
    const hasTargetQueueDef = Boolean(CONFIG.SHEET && CONFIG.SHEET.TARGET_QUEUE);
    if (
      !ss.getSheetByName(CONFIG.SHEET.EVENT_QUEUE) ||
      !ss.getSheetByName(CONFIG.SHEET.LOG_HISTORY) ||
      (hasTargetQueueDef && !ss.getSheetByName(CONFIG.SHEET.TARGET_QUEUE))
    ) {
      initialSetup();
    }

    const opened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.EVENT_QUEUE, 1);
    const { sheet, values, headerMap } = opened;

    if (values.length > 1) {
      const idxType = requireColumnAny(headerMap, ['イベント種別'], '⚙️_連携キュー');
      const idxStatus = requireColumnAny(headerMap, ['処理ステータス'], '⚙️_連携キュー');
      const idxDoneAt = requireColumnAny(headerMap, ['処理完了日時'], '⚙️_連携キュー');

      for (let i = 1; i < values.length; i++) {
        const status = String(values[i][idxStatus] || '').trim();
        if (status !== '未処理') continue;

        const eventType = String(values[i][idxType] || '').trim();
        try {
          if (eventType === 'MIJYUSOKU_POST') {
            runWeeklyAbsencePost();
          } else if (eventType === 'SAIYO_POST') {
            runNewHirePostFromUrlSheet();
          }
          const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
          sheet.getRange(i + 1, idxStatus + 1).setValue('処理済');
          sheet.getRange(i + 1, idxDoneAt + 1).setValue(nowStr);
        } catch (err) {
          sheet.getRange(i + 1, idxStatus + 1).setValue(`エラー: ${err.message}`);
        }
      }
    }

    // 連携キュー＆送信履歴の48時間前データの自動掃除（※⚙️_新規お久しぶり管理は削除しない）
    cleanupOldManagementRows_();

  } finally {
    lock.releaseLock();
  }
}

/**
 * 「⚙️_連携キュー」および「⚙️_送信履歴」の2行目以下で、
 * 過去48時間より前のデータを自動削除するヘルパー関数
 * ※「⚙️_新規お久しぶり管理」は永続台帳のため削除しない
 */
function cleanupOldManagementRows_() {
  const cutoffMs = Date.now() - (48 * 60 * 60 * 1000); // 48時間前

  // 日時セル（Date型または文字列）をミリ秒に変換するヘルパー
  const toTimestampMs = (val) => {
    if (!val) return NaN;
    if (val instanceof Date) return val.getTime();
    const d = new Date(String(val).trim().replace(/\//g, '-'));
    return d.getTime();
  };

  // --- 1. ⚙️_連携キュー の掃除（2行目以下） ---
  const qOpened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.EVENT_QUEUE, 1);
  if (qOpened.values.length > 1) {
    const idxReqAt = requireColumnAny(qOpened.headerMap, ['リクエスト日時'], '⚙️_連携キュー');
    const idxStatus = requireColumnAny(qOpened.headerMap, ['処理ステータス'], '⚙️_連携キュー');
    const idxDoneAt = requireColumnAny(qOpened.headerMap, ['処理完了日時'], '⚙️_連携キュー');

    for (let i = qOpened.values.length - 1; i >= 1; i--) {
      const status = String(qOpened.values[i][idxStatus] || '').trim();
      if (status === '未処理') continue; // 未処理のものは残す

      const rowMs = toTimestampMs(qOpened.values[i][idxReqAt] || qOpened.values[i][idxDoneAt]);
      if (!isNaN(rowMs) && rowMs < cutoffMs) {
        qOpened.sheet.deleteRow(i + 1);
      }
    }
  }

  // --- 2. ⚙️_送信履歴 の掃除（2行目以下） ---
  const lOpened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.LOG_HISTORY, 1);
  if (lOpened.values.length > 1) {
    const idxSentAt = requireColumnAny(lOpened.headerMap, ['送信日時'], '⚙️_送信履歴');
    const idxCat = requireColumnAny(lOpened.headerMap, ['種別'], '⚙️_送信履歴');
    const idxWorkDate = optionalColumnAny(lOpened.headerMap, ['勤務日']);

    for (let i = lOpened.values.length - 1; i >= 1; i--) {
      const sentMs = toTimestampMs(lOpened.values[i][idxSentAt]);
      if (isNaN(sentMs) || sentMs >= cutoffMs) continue; // 送信から48時間未満は残す

      // 安全策：「お久しぶり_初回」だけは、勤務日がまだ来ていない（未来の）場合は消すと5分後に再送されるため、
      // 勤務日からも48時間経過するまで保持する
      const cat = String(lOpened.values[i][idxCat] || '').trim();
      if (cat === 'お久しぶり_初回' && idxWorkDate !== null) {
        const workDateObj = parseDateToSafeDateObj(lOpened.values[i][idxWorkDate]);
        if (workDateObj && workDateObj.getTime() >= cutoffMs) {
          continue;
        }
      }

      lOpened.sheet.deleteRow(i + 1);
    }
  }
}

/**
 * 定期実行トリガーの一括登録（初回のみ実行）
 */
function setupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // 1. 連携キュー＆お久しぶり勤務の初回検知＋48h自動削除（5分おき）
  ScriptApp.newTrigger('periodicDaemon5Min')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 2. 毎日朝9時〜10時：1週間前リマインド ＆（月曜なら週次リマインド）
  ScriptApp.newTrigger('dailyReminderBatch')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // 3. 毎日夕方17時〜18時：拠点間移動通知 ＆ 明日の新規・お久しぶり前日リマインド
  ScriptApp.newTrigger('dailyMorningBatch')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();

  Logger.log('トリガーの設定が完了しました。（5分おき監視 ／ 朝9時リマインド ／ 17時拠点間移動＆前日リマインド）');
  try {
    SpreadsheetApp.getUi().alert(
      'トリガーの設定が完了しました。\n' +
      '・5分おき：DS部投稿連動＆お久しぶり勤務検知（＋48時間前データの自動削除）\n' +
      '・毎日 9時：1週間前リマインド・月曜週次リマインド\n' +
      '・毎日17時：拠点間移動通知・明日の新規/お久しぶり前日リマインド'
    );
  } catch (e) {
    // スクリプトエディタから直接実行した場合はUIアラートをスキップ
  }
}

function periodicDaemon5Min() {
  try { processEventQueue(); } catch (e) { Logger.log('Queue Error: ' + e.message); }
  try { runOhisashiburiInitialCheck(); } catch (e) { Logger.log('Ohisa Error: ' + e.message); }
}

/**
 * お久しぶり勤務・新規採用医師のリマインド用バッチ（毎日朝9時トリガー等）
 */
function dailyReminderBatch() {
  try { runOneWeekBeforeReminder(); } catch (e) { Logger.log('7D Remind Error: ' + e.message); }
  // 月曜日（getDay() === 1）なら今週のまとめリマインドを実行
  if (new Date().getDay() === 1) {
    try { runMondayWeeklyReminder(); } catch (e) { Logger.log('Monday Remind Error: ' + e.message); }
  }
}

/**
 * 拠点間移動通知 ＆ 明日の新規・お久しぶり前日リマインド用バッチ（毎日17時〜18時トリガー）
 */
function dailyMorningBatch() {
  try { runClinicTransferPost(); } catch (e) { Logger.log('Transfer Error: ' + e.message); }
  try { runTomorrowNewAndOhisaReminder(); } catch (e) { Logger.log('1D Tomorrow Remind Error: ' + e.message); }
}