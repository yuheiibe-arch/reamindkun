/**
 * =================================================================
 * 【4_4_リマインド通知実行.gs】（要件4・5のリマインド通知実行）
 * 1. 初回勤務日の1週間前リマインド（runOneWeekBeforeReminder）
 * 2. 毎週月曜日の「今週の新規勤務医師・お久しぶり勤務医師」まとめリマインド（runMondayWeeklyReminder）
 * =================================================================
 */

/**
 * 1. 初回勤務日の1週間前リマインド（毎日朝実行）
 */
function runOneWeekBeforeReminder() {
  const targetDate = new Date();
  targetDate.setHours(0, 0, 0, 0);
  targetDate.setDate(targetDate.getDate() + 7);
  const targetDateKey = fastFormatDate(targetDate);

  // 貼付用シートと最新同期した上で抽出
  const allEntries = syncTargetDoctorsQueue();
  const targets = allEntries.filter(e => e.dateKey === targetDateKey && !e.remind7Sent);

  if (targets.length === 0) {
    Logger.log(`1週間後 (${targetDateKey}) に該当する未送信の新規・お久しぶり医師はいませんでした。`);
    return;
  }

  targets.forEach(item => {
    const historyKey = `REMIND_7D_${item.uniqueKey}_${item.dateKey}`;
    if (hasSentHistory('1週間前リマインド', historyKey)) {
      markTargetQueueFlag_(item.uniqueKey, '1週間前リマインド済');
      return;
    }

    const mentionHeader = buildClinicMentionsText([item.clinic], false);
    const w = ['日', '月', '火', '水', '木', '金', '土'][item.dateObj.getDay()];
    let message = '';

    if (item.type === '新規') {
      const clinicDisplay = item.dept ? `【${item.clinic}/${item.dept}】` : `【${item.clinic}】`;
      message = `${mentionHeader}[info][title]新規医師採用報告（1週間前リマインド）[/title]
勤務拠点：${clinicDisplay}
勤務日：${item.dateKey}(${w}) ${item.timeStr}
医師名：${item.doctorName}
性別：${item.gender}
専門科目：${item.specialty}
診療体制：${item.shinTaisei}
[/info]`;
    } else {
      const clinicDisplay = (CONFIG.DEPT_LABEL_CLINICS.includes(item.clinic) && item.dept) ? `【${item.clinic}/${item.dept}】` : `【${item.clinic}】`;
      message = `${mentionHeader}[info][title]お久しぶり勤務&オリエンテーション依頼（1週間前リマインド）[/title]
対象拠点：${clinicDisplay}
勤務日：${item.dateKey}(${w})
対象時間：${item.timeStr}
医師名：${item.doctorName}
診療体制：${item.shinTaisei}
[/info]`;
    }

    sendChatworkMessage(message);
    recordSentHistory('1週間前リマインド', historyKey, item.doctorName, item.clinic, item.dateKey);
    markTargetQueueFlag_(item.uniqueKey, '1週間前リマインド済');
    Utilities.sleep(1500);
  });
}

/**
 * 2. 毎週月曜日の「今週の新規勤務医師・お久しぶり勤務医師」まとめリマインド
 */
function runMondayWeeklyReminder() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endOfWeek = new Date(today);
  endOfWeek.setDate(today.getDate() + 6); // 月曜〜日曜の7日間

  const startKey = fastFormatDate(today);
  const endKey = fastFormatDate(endOfWeek);

  // 貼付用シートと最新同期した上で抽出（前倒し済みの人は dateKey が過去になるため今週分から自動で外れる）
  const allEntries = syncTargetDoctorsQueue();
  const weeklyTargets = allEntries
    .filter(e => e.dateKey >= startKey && e.dateKey <= endKey)
    .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

  if (weeklyTargets.length === 0) {
    Logger.log(`期間 (${startKey}〜${endKey}) の新規勤務・お久しぶり勤務医師はいませんでした。`);
    return;
  }

  const byDate = new Map();
  const clinicsForMention = new Set();

  weeklyTargets.forEach(item => {
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
  const message = `${mentionHeader}[info][title]今週の新規勤務医師・お久しぶり勤務医師[/title]
${bodyLines.join('\n')}
オリエンテーションの設定のほどお願い申し上げます。
上記はあくまで予定になりますので、変更の場合は速やかにご連絡させていただきます。
[/info]`;

  sendChatworkMessage(message);

  weeklyTargets.forEach(item => {
    markTargetQueueFlag_(item.uniqueKey, '週次リマインド済');
  });
}