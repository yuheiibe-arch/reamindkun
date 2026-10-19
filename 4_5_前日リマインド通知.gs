/**
 * =================================================================
 * 【4_5_前日リマインド通知.gs】
 * ・機能：「明日の新規勤務医師・お久しぶり勤務医師」を抽出し、本番ルーム（449648673）へ投稿
 * ・既存の syncTargetDoctorsQueue() と完全連携（フォーム3種[正] ＋ 貼付用スタッフコメント[保険] ＋ 2診判定）
 * ・送信履歴シート（⚙️_送信履歴）による同日重複送信防止つき
 * =================================================================
 */

/**
 * 翌日（明日）の新規勤務・お久しぶり勤務医師リマインドを実行するメイン関数
 */
function runTomorrowNewAndOhisaReminder() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowKey = fastFormatDate(tomorrow);
  const execDateKey = fastFormatDate(today);

  // 永続管理キュー（フォーム3種[正] ＋ 貼付用スタッフコメント[保険] ＋ 2診判定）を最新同期して取得
  const allEntries = syncTargetDoctorsQueue();
  const tomorrowTargets = allEntries
    .filter(e => e.dateKey === tomorrowKey)
    .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

  if (tomorrowTargets.length === 0) {
    Logger.log(`明日 (${tomorrowKey}) の新規勤務・お久しぶり勤務医師はいませんでした。`);
    return;
  }

  // 既に本日実行分で送信済みの医師を除外（手動・トリガー重複実行時の2重投稿防止）
  const unsentTargets = tomorrowTargets.filter(item => {
    const historyKey = `REMIND_1D_EXEC_${execDateKey}_TARGET_${item.uniqueKey}_${item.dateKey}`;
    return !hasSentHistory('前日リマインド', historyKey);
  });

  if (unsentTargets.length === 0) {
    Logger.log(`明日 (${tomorrowKey}) の対象医師はすべて本日送信済みのためスキップしました。`);
    return;
  }

  const byDate = new Map();
  const clinicsForMention = new Set();

  unsentTargets.forEach(item => {
    clinicsForMention.add(item.clinic);
    if (!byDate.has(item.dateKey)) {
      byDate.set(item.dateKey, { dateObj: item.dateObj, items: [] });
    }
    byDate.get(item.dateKey).items.push(item);
  });

  const bodyLines = [];
  byDate.forEach(group => {
    bodyLines.push(`[code]${formatJPDateWithDay(group.dateObj, true)}[/code]`);
    group.items.forEach(item => {
      const clinicDisplay = (CONFIG.DEPT_LABEL_CLINICS.includes(item.clinic) && item.dept)
        ? `【${item.clinic}/${item.dept}】`
        : (item.type === '新規' && item.dept ? `【${item.clinic}/${item.dept}】` : `【${item.clinic}】`);

      bodyLines.push(`${clinicDisplay}：${item.timeStr}`);
      if (item.type === '新規') {
        bodyLines.push(`(新規)${item.shortDoctorName}`);
      } else {
        bodyLines.push(`(お久しぶり)：${item.shortDoctorName}`);
      }
      bodyLines.push(`診療体制：${item.shinTaisei}[hr]`);
    });
  });

  const mentionHeader = buildClinicMentionsText(Array.from(clinicsForMention), false);
  const message = `${mentionHeader}[info][title]明日の新規勤務医師・お久しぶり勤務医師[/title]
${bodyLines.join('\n')}
オリエンテーションの設定のほどお願い申し上げます。
上記はあくまで予定になりますので、変更の場合は速やかにご連絡させていただきます。
[/info]`;

  // 本番ルーム（CONFIG.getTargetRoomId() -> 449648673）へ送信
  sendChatworkMessage(message);

  // 送信履歴に記録
  unsentTargets.forEach(item => {
    const historyKey = `REMIND_1D_EXEC_${execDateKey}_TARGET_${item.uniqueKey}_${item.dateKey}`;
    recordSentHistory('前日リマインド', historyKey, item.doctorName, item.clinic, item.dateKey);
  });
}