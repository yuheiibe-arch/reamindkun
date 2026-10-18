/**
 * =================================================================
 * 【5_拠点間移動_トリガー管理.gs】（要件6 ＆ 全体トリガー管理）
 * ・通常平日（翌日が平日）：翌日（明日）1日分のみを通知
 * ・金曜日・祝前日：直後の連続する「土日・祝日」の分のみをまとめて通知（休み明けの平日は含めない）
 * ・土日・祝日当日：デイリー分として翌日（明日）1日分を通知（土曜なら日曜分、日曜なら月曜分を通知）
 * ・祝日判定は内閣府（https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html）から動的取得
 * =================================================================
 */

/**
 * 実行日（今日）が通知すべき対象日の配列を返す
 * 1. 今日が「平日（営業日）」かつ「明日が土日または祝日（連休）」の場合（＝金曜や祝前日）：
 *    👉 明日からの連続する「土日・祝日」のみをすべて返す（休み明けの平日は含めない）
 * 2. それ以外の日（通常の月〜木、または土日・祝日当日）：
 *    👉 デイリー分として「明日（翌日）」の1日分のみを返す
 */
function getTargetDatesForClinicTransfer_(baseDate = new Date()) {
  const today = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
  const isNonBusinessDay = (d) => d.getDay() === 0 || d.getDay() === 6 || isJapaneseHoliday(d);

  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  // 今日が平日（営業日）で、明日が休日（土日・祝日）の場合＝金曜・祝前日
  if (!isNonBusinessDay(today) && isNonBusinessDay(tomorrow)) {
    const holidayTargets = [];
    const cursor = new Date(tomorrow);
    // 連続する土日・祝日だけをすべて追加（月曜などの平日になったらループ終了）
    while (isNonBusinessDay(cursor)) {
      holidayTargets.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    return holidayTargets;
  }

  // 通常平日（翌日が平日）および土日・祝日当日は、デイリーとして「明日（翌日）」の1日分のみ
  return [tomorrow];
}

/**
 * 要件6：拠点間移動が生じる医師の抽出＆投稿
 */
function runClinicTransferPost() {
  const targetDates = getTargetDatesForClinicTransfer_(new Date());
  if (targetDates.length === 0) return;

  const execDateKey = fastFormatDate(new Date());
  const targetDateKeys = new Set(targetDates.map(d => fastFormatDate(d)));

  const opened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.PASTE, 1);
  const { values, headerMap } = opened;

  const idxName = requireColumnAny(headerMap, ['名前'], '貼付用');
  const idxClinic = requireColumnAny(headerMap, ['クリニック名'], '貼付用');
  const idxDept = requireColumnAny(headerMap, ['診療科'], '貼付用');
  const idxDate = requireColumnAny(headerMap, ['勤務日'], '貼付用');
  const idxStart = requireColumnAny(headerMap, ['勤務開始時間'], '貼付用');
  const idxEnd1 = optionalColumnAny(headerMap, ['終了時間1']);
  const idxEnd2 = optionalColumnAny(headerMap, ['終了時間2']);
  const idxEnd3 = optionalColumnAny(headerMap, ['終了時間3']);
  const idxEnd = requireColumnAny(headerMap, ['勤務終了時間'], '貼付用');

  const EXCLUDED = ['有給', '欠勤', '院外勤務', 'バックアップ', '医師会', '嘱託医', '休館'];

  // 1行から中抜け考慮のインターバル配列 [{start, end}] を抽出
  const extractIntervals = (row) => {
    const cols = [idxStart, idxEnd1, idxEnd2, idxEnd3, idxEnd].filter(c => c !== null);
    const times = cols.map(c => safeParseTime(row[c])).filter(v => !isNaN(v));
    if (times.length === 4) {
      const res = [];
      if (times[0] < times[1]) res.push({ start: times[0], end: times[1] });
      if (times[2] < times[3]) res.push({ start: times[2], end: times[3] });
      return res;
    } else if (times.length >= 2) {
      const s = times[0], e = times[times.length - 1];
      if (s < e) return [{ start: s, end: e }];
    }
    return [];
  };

  // 日付ごとの全シフト記録（人数カウント＆医師別シフト集計用）
  const shiftsByDate = new Map();

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const docName = String(row[idxName] || '').trim();
    const clinic = String(row[idxClinic] || '').trim();
    const dept = String(row[idxDept] || '').trim();
    if (!docName || !clinic) continue;
    if (EXCLUDED.some(ex => docName.includes(ex) || clinic.includes(ex)) || dept.includes('ワクチン専任')) continue;

    const dObj = parseDateToSafeDateObj(row[idxDate]);
    if (!dObj) continue;
    const dKey = fastFormatDate(dObj);
    if (!targetDateKeys.has(dKey)) continue;

    const intervals = extractIntervals(row);
    if (intervals.length === 0) continue;

    if (!shiftsByDate.has(dKey)) {
      shiftsByDate.set(dKey, { dateObj: dObj, allIntervals: [], byDoctor: new Map() });
    }
    const dayGroup = shiftsByDate.get(dKey);

    intervals.forEach(inv => {
      dayGroup.allIntervals.push({ docName, clinic, dept, start: inv.start, end: inv.end });
      if (!dayGroup.byDoctor.has(docName)) dayGroup.byDoctor.set(docName, []);
      dayGroup.byDoctor.get(docName).push({ clinic, dept, start: inv.start, end: inv.end });
    });
  }

  // 日付ごとに該当する移動医師をまとめ上げる
  const sortedDateKeys = Array.from(shiftsByDate.keys()).sort();
  const clinicsForMention = new Set();
  const bodyLines = [];
  const pendingHistoryRecords = [];

  sortedDateKeys.forEach(dKey => {
    const dayGroup = shiftsByDate.get(dKey);
    const dayDoctorBlocks = [];

    // 同時刻・同拠点・同診療科の医師数をカウントして「※1診 / ※2診」を判定
    const countDoctorsInSlot = (targetShift) => {
      const docs = new Set();
      dayGroup.allIntervals.forEach(other => {
        if (other.clinic === targetShift.clinic && other.dept === targetShift.dept) {
          if (Math.max(other.start, targetShift.start) < Math.min(other.end, targetShift.end)) {
            docs.add(other.docName);
          }
        }
      });
      return docs.size >= 2 ? '※2診' : '※1診';
    };

    const formatClinicLabel = (s) => (CONFIG.DEPT_LABEL_CLINICS.includes(s.clinic) && s.dept) ? `【${s.clinic}(${s.dept})】` : `【${s.clinic}】`;

    dayGroup.byDoctor.forEach((docShifts, docName) => {
      // 午前: 09:00(540)〜13:00(780)
      const amShift = docShifts.find(s => s.start === 540 && s.end === 780);
      // 午後: 15:00(900)開始
      const pmShift = docShifts.find(s => s.start === 900 && s.end > 900);

      if (!amShift || !pmShift) return;
      if (amShift.clinic === pmShift.clinic) return; // 同一拠点なら移動なし

      // 実行日(execDateKey)をキーに含めることで、金曜の土日まとめ送信後も土曜当日の日曜デイリー送信がスキップされないようにする
      const uniqueKey = `TRANSFER_EXEC_${execDateKey}_TARGET_${dKey}_${docName.replace(/\s+/g, '')}_${amShift.clinic}_${pmShift.clinic}`;
      if (hasSentHistory('拠点間移動', uniqueKey)) return;

      const amShin = countDoctorsInSlot(amShift);
      const pmShin = countDoctorsInSlot(pmShift);
      const pmTimeStr = `${safeFormatTime(pmShift.start)}-${safeFormatTime(pmShift.end)}`;
      const cleanDocName = docName.replace(/先生|医師/g, '').trim() + '医師';

      // 午前拠点・午後拠点の双方をメンション対象に追加
      clinicsForMention.add(amShift.clinic);
      clinicsForMention.add(pmShift.clinic);

      dayDoctorBlocks.push(
        `${cleanDocName}\n` +
        `09:00-13:00: ${formatClinicLabel(amShift)}${amShin}\n` +
        `${pmTimeStr}: ${formatClinicLabel(pmShift)}${pmShin}[hr]`
      );

      pendingHistoryRecords.push({
        uniqueKey: uniqueKey,
        doctorName: cleanDocName,
        clinicPair: `${amShift.clinic}->${pmShift.clinic}`,
        dateKey: dKey
      });
    });

    if (dayDoctorBlocks.length > 0) {
      bodyLines.push(`[code]${formatJPDateWithDay(dayGroup.dateObj, true)}[/code]`);
      bodyLines.push(dayDoctorBlocks.join('\n'));
    }
  });

  if (bodyLines.length === 0) {
    Logger.log(`対象日 (${Array.from(targetDateKeys).join(', ')}) に拠点間移動（午前09:00-13:00 → 午後15:00開始で別拠点）の該当医師はいませんでした。`);
    return;
  }

  const mentionHeader = buildClinicMentionsText(Array.from(clinicsForMention), false);
  const message = `${mentionHeader}[info][title]拠点間移動発生医師[/title]
${bodyLines.join('\n')}
移動が発生する場合において、休憩取得ができるようにご配慮お願いいたします。
1時間の休憩が必須となります。
[/info]`;

  sendChatworkMessage(message);

  pendingHistoryRecords.forEach(rec => {
    recordSentHistory('拠点間移動', rec.uniqueKey, rec.doctorName, rec.clinicPair, rec.dateKey);
  });
}

/**
 * 連携キュー（⚙️_連携キュー）を処理する関数（5分おきトリガー等で処理）
 */
function processEventQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);
    if (!ss.getSheetByName(CONFIG.SHEET.EVENT_QUEUE)) {
      initialSetup();
    }
    const opened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.EVENT_QUEUE, 1);
    const { sheet, values, headerMap } = opened;
    if (values.length <= 1) return;

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
  } finally {
    lock.releaseLock();
  }
}

/**
 * 定期実行トリガーの一括登録（初回のみ実行）
 */
function setupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // 1. 連携キュー＆お久しぶり勤務の初回検知（5分おき）
  ScriptApp.newTrigger('periodicDaemon5Min')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 2. 毎日朝9時：1週間前リマインド ＆ 拠点間移動通知 ＆（月曜なら週次リマインド）
  ScriptApp.newTrigger('dailyMorningBatch')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('✅ トリガーの設定が完了しました！（5分おき監視 ＆ 毎日朝9時バッチ）');
  try {
    SpreadsheetApp.getUi().alert('✅ トリガーの設定が完了しました！\n・5分おき：DS部投稿連動＆お久しぶり勤務検知\n・毎日朝9時：1週間前リマインド・月曜週次リマインド・拠点間移動通知');
  } catch (e) {
    // スクリプトエディタから直接実行した場合はUIアラートをスキップ
  }
}

function periodicDaemon5Min() {
  try { processEventQueue(); } catch (e) { Logger.log('Queue Error: ' + e.message); }
  try { runOhisashiburiInitialCheck(); } catch (e) { Logger.log('Ohisa Error: ' + e.message); }
}

function dailyMorningBatch() {
  try { runOneWeekBeforeReminder(); } catch (e) { Logger.log('7D Remind Error: ' + e.message); }
  try { runClinicTransferPost(); } catch (e) { Logger.log('Transfer Error: ' + e.message); }
  // 月曜日（getDay() === 1）なら今週のまとめリマインドを実行
  if (new Date().getDay() === 1) {
    try { runMondayWeeklyReminder(); } catch (e) { Logger.log('Monday Remind Error: ' + e.message); }
  }
}