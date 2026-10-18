/**
 * =================================================================
 * 【4_リマインド処理.gs】（要件4・5のリマインド）
 * ★重要仕様：「貼付用」シートに、採用報告やお久しぶり勤務の日付より
 *   前の日付でシフトが登場している場合は、そちらの日付・拠点・時間を「正」とする。
 * 1. 勤務日1週間前リマインド（毎日実行）
 * 2. 今週の新規勤務医師・お久しぶり勤務医師（毎週月曜実行）
 * =================================================================
 */

/**
 * 「貼付用」シートから全医師の今後のシフト一覧を読み込み、
 * 医籍番号および正規化氏名から「最も早いシフト」を引けるインデックスを作成する
 */
function buildEarliestShiftLookupFromPaste_() {
  const opened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.PASTE, 1);
  const { values, headerMap } = opened;

  const idxName = requireColumnAny(headerMap, ['名前'], '貼付用');
  const idxIki = requireColumnAny(headerMap, ['医籍番号'], '貼付用');
  const idxClinic = requireColumnAny(headerMap, ['クリニック名'], '貼付用');
  const idxDept = requireColumnAny(headerMap, ['診療科'], '貼付用');
  const idxDate = requireColumnAny(headerMap, ['勤務日'], '貼付用');
  const idxStart = requireColumnAny(headerMap, ['勤務開始時間'], '貼付用');
  const idxEnd = requireColumnAny(headerMap, ['勤務終了時間'], '貼付用');

  const byIki = new Map();
  const byName = new Map();

  const normNameFn = (n) => String(n || '').split(/[（(]/)[0].replace(/先生|医師|[\s ]/g, '').replace(/髙/g, '高').replace(/﨑/g, '崎').trim();

  // 貼付用シートは3行目(インデックス2)から実データ
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const docName = String(row[idxName] || '').trim();
    const clinic = String(row[idxClinic] || '').trim();
    if (!docName || !clinic || docName.includes('休館') || clinic.includes('バックアップ') || clinic.includes('有給')) continue;

    const dObj = parseDateToSafeDateObj(row[idxDate]);
    if (!dObj) continue;

    const sMin = safeParseTime(row[idxStart]);
    const eMin = safeParseTime(row[idxEnd]);
    const timeStr = (!isNaN(sMin) && !isNaN(eMin)) ? `${safeFormatTime(sMin)}-${safeFormatTime(eMin)}` : '';

    const shiftEntry = {
      dateObj: dObj,
      dateKey: fastFormatDate(dObj),
      clinic: clinic,
      dept: String(row[idxDept] || '').trim(),
      timeStr: timeStr
    };

    const iki = String(row[idxIki] || '').trim();
    if (iki) {
      if (!byIki.has(iki)) byIki.set(iki, []);
      byIki.get(iki).push(shiftEntry);
    }

    const nKey = normNameFn(docName);
    if (nKey) {
      if (!byName.has(nKey)) byName.set(nKey, []);
      byName.get(nKey).push(shiftEntry);
    }
  }

  return { byIki, byName, normNameFn };
}

/**
 * 〈採用報告〉と「お久しぶり勤務」の全エントリーを読み込み、
 * 貼付用シートに前の日付があればそちらを「正」として補正したリストを返す
 */
function collectAllTargetDoctorsWithTrueDate_() {
  const shiftLookup = buildEarliestShiftLookupFromPaste_();
  const results = [];

  // 貼付用シートから「登録日以降で、フォームの勤務日より前のシフト」があれば置き換えるヘルパー
  const resolveTrueShift = (iki, rawName, formDateObj, timestampDateObj, defaultClinic, defaultDept, defaultTime) => {
    const candidates = [];
    if (iki && shiftLookup.byIki.has(iki)) candidates.push(...shiftLookup.byIki.get(iki));
    const nKey = shiftLookup.normNameFn(rawName);
    if (nKey && shiftLookup.byName.has(nKey)) candidates.push(...shiftLookup.byName.get(nKey));

    if (candidates.length === 0) {
      return { dateObj: formDateObj, clinic: defaultClinic, dept: defaultDept, timeStr: defaultTime };
    }

    // フォーム送信日（なければフォーム勤務日の60日前）以降で最も早いシフトを探す
    const minValidDate = timestampDateObj ? new Date(timestampDateObj.getFullYear(), timestampDateObj.getMonth(), timestampDateObj.getDate()) : new Date(formDateObj.getTime() - 60 * 86400000);
    const validShifts = candidates
      .filter(s => s.dateObj.getTime() >= minValidDate.getTime())
      .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

    if (validShifts.length > 0 && validShifts[0].dateObj.getTime() < formDateObj.getTime()) {
      return {
        dateObj: validShifts[0].dateObj,
        clinic: validShifts[0].clinic,
        dept: validShifts[0].dept || defaultDept,
        timeStr: validShifts[0].timeStr || defaultTime
      };
    }
    return { dateObj: formDateObj, clinic: defaultClinic, dept: defaultDept, timeStr: defaultTime };
  };

  // --- A. 〈採用報告〉シート ---
  const saiyoOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.SAIYO_REPORT, 1);
  const sVals = saiyoOpened.values;
  const sMap = saiyoOpened.headerMap;
  const sIdxTs = requireColumnAny(sMap, ['タイムスタンプ'], '〈採用報告〉');
  const sIdxClinic = requireColumnAny(sMap, ['拠点'], '〈採用報告〉');
  const sIdxDept = requireColumnAny(sMap, ['診療科'], '〈採用報告〉');
  const sIdxDate = requireColumnAny(sMap, ['勤務日'], '〈採用報告〉');
  const sIdxTime1 = requireColumnAny(sMap, ['勤務時間'], '〈採用報告〉');
  const sIdxTime2 = optionalColumnAny(sMap, ['勤務時間が上記に当てはまらい場合']);
  const sIdxName = requireColumnAny(sMap, ['氏名'], '〈採用報告〉');
  const sIdxGender = requireColumnAny(sMap, ['性別'], '〈採用報告〉');
  const sIdxSpec = requireColumnAny(sMap, ['専門科目'], '〈採用報告〉');
  const sIdxIki = optionalColumnAny(sMap, ['医籍番号']);
  const sIdxHist = optionalColumnAny(sMap, ['学歴・職歴']);

  for (let i = 1; i < sVals.length; i++) {
    const row = sVals[i];
    const formDate = parseDateToSafeDateObj(row[sIdxDate]);
    const rawName = String(row[sIdxName] || '').trim();
    if (!formDate || !rawName) continue;

    const tsDate = parseDateToSafeDateObj(row[sIdxTs]);
    let iki = sIdxIki !== null ? String(row[sIdxIki] || '').trim() : '';
    if (!iki && sIdxHist !== null) {
      const m = String(row[sIdxHist] || '').match(/第\s*(\d+)\s*号/);
      if (m) iki = m[1];
    }

    const rawTime = String((sIdxTime2 !== null && row[sIdxTime2]) ? row[sIdxTime2] : row[sIdxTime1] || '').trim();
    const resolved = resolveTrueShift(
      iki, rawName, formDate, tsDate,
      String(row[sIdxClinic] || '').trim(),
      String(row[sIdxDept] || '').trim(),
      rawTime
    );

    results.push({
      type: '新規',
      doctorName: rawName.endsWith('先生') ? rawName : `${rawName}先生`,
      shortDoctorName: rawName.split(/[（(]/)[0].replace(/\s+/g, '').replace(/先生$/, '') + '先生',
      gender: String(row[sIdxGender] || '').trim(),
      specialty: String(row[sIdxSpec] || '').trim(),
      dateObj: resolved.dateObj,
      dateKey: fastFormatDate(resolved.dateObj),
      clinic: resolved.clinic,
      dept: resolved.dept,
      timeStr: resolved.timeStr.replace(/[~〜～]/g, '-')
    });
  }

  // --- B. 「お久しぶり勤務」シート ---
  const ohisaOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.OHISASHIBURI, 1);
  const oVals = ohisaOpened.values;
  const oMap = ohisaOpened.headerMap;
  const oIdxTs = requireColumnAny(oMap, ['タイムスタンプ'], 'お久しぶり勤務');
  const oIdxClinic = requireColumnAny(oMap, ['拠点'], 'お久しぶり勤務');
  const oIdxDate = requireColumnAny(oMap, ['勤務日'], 'お久しぶり勤務');
  const oIdxTime1 = requireColumnAny(oMap, ['勤務時間（自動で30分前に補正します。）', '勤務時間'], 'お久しぶり勤務');
  const oIdxTime2 = optionalColumnAny(oMap, ['勤務時間が上記に当てはまらい場合']);
  const oIdxName = requireColumnAny(oMap, ['氏名'], 'お久しぶり勤務');
  const oIdxIki = optionalColumnAny(oMap, ['医籍番号']);
  const oIdxDept = optionalColumnAny(oMap, ['診療科']);
  const oIdxOrien = requireColumnAny(oMap, ['30分前オリエンテーションの有無'], 'お久しぶり勤務');

  for (let i = 1; i < oVals.length; i++) {
    const row = oVals[i];
    const formDate = parseDateToSafeDateObj(row[oIdxDate]);
    const rawName = String(row[oIdxName] || '').trim();
    if (!formDate || !rawName) continue;

    const tsDate = parseDateToSafeDateObj(row[oIdxTs]);
    const iki = oIdxIki !== null ? String(row[oIdxIki] || '').trim() : '';
    const rawTime = String((oIdxTime2 !== null && row[oIdxTime2]) ? row[oIdxTime2] : row[oIdxTime1] || '').trim();
    const isOrienRequired = String(row[oIdxOrien] || '').trim() === '必須';

    const resolved = resolveTrueShift(
      iki, rawName, formDate, tsDate,
      String(row[oIdxClinic] || '').trim(),
      oIdxDept !== null ? String(row[oIdxDept] || '').trim() : '',
      rawTime
    );

    const adjustedTime = shiftStartTimeBack30Min(resolved.timeStr, isOrienRequired);

    results.push({
      type: 'お久しぶり',
      doctorName: rawName.endsWith('先生') ? rawName : `${rawName}先生`,
      shortDoctorName: rawName.split(/[（(]/)[0].replace(/\s+/g, '').replace(/先生$/, '') + '先生',
      dateObj: resolved.dateObj,
      dateKey: fastFormatDate(resolved.dateObj),
      clinic: resolved.clinic,
      dept: resolved.dept,
      timeStr: adjustedTime
    });
  }

  return results;
}

/**
 * ① 初回勤務日の1週間前リマインド（毎日朝実行）
 */
function runOneWeekBeforeReminder() {
  const targetDate = new Date();
  targetDate.setHours(0, 0, 0, 0);
  targetDate.setDate(targetDate.getDate() + 7);
  const targetDateKey = fastFormatDate(targetDate);

  const allEntries = collectAllTargetDoctorsWithTrueDate_();
  const targets = allEntries.filter(e => e.dateKey === targetDateKey);

  if (targets.length === 0) {
    Logger.log(`1週間後 (${targetDateKey}) に該当する新規・お久しぶり医師はいませんでした。`);
    return;
  }

  targets.forEach(item => {
    const uniqueKey = `REMIND_7D_${item.type}_${item.clinic}_${item.dateKey}_${item.shortDoctorName}`;
    if (hasSentHistory('1週間前リマインド', uniqueKey)) return;

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
[/info]`;
    } else {
      const clinicDisplay = (CONFIG.DEPT_LABEL_CLINICS.includes(item.clinic) && item.dept) ? `【${item.clinic}/${item.dept}】` : `【${item.clinic}】`;
      message = `${mentionHeader}[info][title]お久しぶり勤務&オリエンテーション依頼（1週間前リマインド）[/title]
対象拠点：${clinicDisplay}
勤務日：${item.dateKey}(${w})
対象時間：${item.timeStr}
医師名：${item.doctorName}
[/info]`;
    }

    sendChatworkMessage(message);
    recordSentHistory('1週間前リマインド', uniqueKey, item.doctorName, item.clinic, item.dateKey);
    Utilities.sleep(1500);
  });
}

/**
 * ② 毎週月曜日の「今週の新規勤務医師・お久しぶり勤務医師」まとめリマインド
 */
function runMondayWeeklyReminder() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endOfWeek = new Date(today);
  endOfWeek.setDate(today.getDate() + 6); // 月曜〜日曜の7日間

  const startKey = fastFormatDate(today);
  const endKey = fastFormatDate(endOfWeek);

  const allEntries = collectAllTargetDoctorsWithTrueDate_();
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
        bodyLines.push(`(新規)${item.shortDoctorName}[hr]`);
      } else {
        bodyLines.push(`(お久しぶり)：${item.shortDoctorName}[hr]`);
      }
    });
  });

  const mentionHeader = buildClinicMentionsText(Array.from(clinicsForMention), false);
  const message = `${mentionHeader}[info][title]今週の新規勤務医師・お久しぶり勤務医師[/title]
${bodyLines.join('\n')}
オリエンテーションの設定のほどお願い申し上げます。
上記はあくまで予定になりますので、変更の場合は速やかにご連絡させていただきます。
[/info]`;

  sendChatworkMessage(message);
}