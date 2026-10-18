/**
 * =================================================================
 * 【2_未充足不在拠点_投稿.gs】（要件1・2）
 * 実シートの見出し構造に100%合致させた完全動的取得版
 * ★メンションは [toall] のみ付与
 * =================================================================
 */

/**
 * 本番・連動用メイン関数
 */
function runWeeklyAbsencePost() {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID.MIJYUSOKU);
  const autoSheet = ss.getSheetByName(CONFIG.SHEET.AUTO_TEXT);
  if (!autoSheet) throw new Error('「文章自動作成」シートが見つかりません。');

  // 「文章自動作成」シートからラベル「作成期日」「終了期日」を動的に探して右隣の値を取得
  const autoVals = autoSheet.getDataRange().getDisplayValues();
  let rawStart = '';
  let rawEnd = '';
  for (let r = 0; r < autoVals.length; r++) {
    for (let c = 0; c < autoVals[r].length; c++) {
      const label = String(autoVals[r][c]).trim();
      if (label === '作成期日' && c + 1 < autoVals[r].length) rawStart = autoVals[r][c + 1];
      if (label === '終了期日' && c + 1 < autoVals[r].length) rawEnd = autoVals[r][c + 1];
    }
  }

  let startDate = parseDateToSafeDateObj(rawStart);
  let endDate = parseDateToSafeDateObj(rawEnd);
  if (!startDate || !endDate || startDate > endDate) {
    startDate = new Date();
    startDate.setHours(0, 0, 0, 0);
    endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 6);
  }

  // 1. 拠点マスタ（開院日）読み込み
  const { openDateMap } = getClinicMentionDictionary();

  // 2. 休館日シート（1行目: 日付 | 曜日（自動） | 診療科 | 拠点名 | 時間 | 備考 | チェック）
  const closedDataMap = new Map();
  const closedSheet = ss.getSheetByName(CONFIG.SHEET.CLOSED);
  if (closedSheet && closedSheet.getLastRow() > 1) {
    const cValues = closedSheet.getDataRange().getValues();
    const cMap = buildHeaderMap(cValues[0]);

    const idxDate = requireColumnAny(cMap, ['日付'], '休館日');
    const idxDept = requireColumnAny(cMap, ['診療科'], '休館日');
    const idxLoc  = requireColumnAny(cMap, ['拠点名'], '休館日');
    const idxTime = requireColumnAny(cMap, ['時間'], '休館日');

    for (let i = 1; i < cValues.length; i++) {
      const dObj = parseDateToSafeDateObj(cValues[i][idxDate]);
      const loc = String(cValues[i][idxLoc] || '').trim();
      const dept = String(cValues[i][idxDept] || '').trim();
      const cTime = String(cValues[i][idxTime] || '全日').trim();
      if (dObj && loc) {
        const dKey = fastFormatDate(dObj);
        const normLoc = loc.replace(/[（(]小児科[）)]/, '').trim();
        if (loc === '全拠点') closedDataMap.set(`${dKey}_全拠点`, cTime);
        else if (dept === '内科') {
          closedDataMap.set(`${dKey}_${normLoc}_内科`, cTime);
          closedDataMap.set(`${dKey}_${loc}_内科`, cTime);
        } else {
          closedDataMap.set(`${dKey}_${normLoc}`, cTime);
        }
      }
    }
  }

  // 3. 変則営業シート（1行目: 対象日 | 営業時間 | 拠点名）
  const irregularMap = {};
  const irrSheet = ss.getSheetByName(CONFIG.SHEET.IRREGULAR);
  if (irrSheet && irrSheet.getLastRow() > 1) {
    const iValues = irrSheet.getDataRange().getValues();
    const iMap = buildHeaderMap(iValues[0]);

    const idxDate = requireColumnAny(iMap, ['対象日'], '変則営業');
    const idxTime = requireColumnAny(iMap, ['営業時間'], '変則営業');
    const idxLoc  = requireColumnAny(iMap, ['拠点名'], '変則営業');

    for (let i = 1; i < iValues.length; i++) {
      const dObj = parseDateToSafeDateObj(iValues[i][idxDate]);
      const tStr = String(iValues[i][idxTime] || '').trim();
      const loc = String(iValues[i][idxLoc] || '').trim();
      if (dObj && tStr && loc) {
        const dKey = fastFormatDate(dObj);
        const validRanges = [];
        tStr.split(/[,/、]/).forEach(rStr => {
          const pts = rStr.split('-');
          if (pts.length === 2) {
            const o = safeParseTime(pts[0]), c = safeParseTime(pts[1]);
            if (!isNaN(o) && !isNaN(c) && o < c) validRanges.push({ open: o, close: c });
          }
        });
        if (validRanges.length > 0) {
          const norm = loc.replace(/[（(]小児科[）)]/, '').trim();
          if (norm === '全拠点') {
            irregularMap[`${dKey}_全拠点`] = validRanges;
          } else {
            irregularMap[`${dKey}_${norm}`] = validRanges;
            irregularMap[`${dKey}_${loc}`] = validRanges;
          }
        }
      }
    }
  }

  // 4. 確認用シート＆医師不在拠点シート
  const kakuninOpened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.KAKUNIN, 1);
  const kValues = kakuninOpened.values;
  const kMap = kakuninOpened.headerMap;
  const idxKClinic = requireColumnAny(kMap, ['拠点名'], '確認用');
  const idxKDate = requireColumnAny(kMap, ['勤務日'], '確認用');
  const idxDocA = requireColumnAny(kMap, ['Aシフト医師 09:00-13:00', 'Aシフト医師'], '確認用');
  const idxDocB = requireColumnAny(kMap, ['Bシフト医師 15:00-18:00', 'Bシフト医師'], '確認用');
  const idxDocC = requireColumnAny(kMap, ['Cシフト医師 18:00-21:00', 'Cシフト医師'], '確認用');

  const EXCLUDED_KEYWORDS = ['有給', '欠勤', '院外勤務', 'バックアップ', '医師会', '嘱託医', '出張インフルエンザワクチン'];
  const allClinics = new Set();
  const backupInfoMap = {};
  const rawNameMap = {};
  const workingClinics = {};

  for (let i = 1; i < kValues.length; i++) {
    const row = kValues[i];
    const name = String(row[idxKClinic] || '').trim();
    const dObj = parseDateToSafeDateObj(row[idxKDate]);
    if (!dObj) continue;
    const dKey = fastFormatDate(dObj);

    if (name === '【関東】バックアップシフト') {
      const slotCols = [
        { label: '09:00~13:00', idx: idxDocA },
        { label: '15:00~18:00', idx: idxDocB },
        { label: '18:00~21:00', idx: idxDocC }
      ];
      const backups = [];
      slotCols.forEach(s => {
        const doc = String(row[s.idx] || '').trim();
        if (doc) backups.push(`${s.label}：${doc}先生（全拠点）`);
      });
      if (backups.length > 0) backupInfoMap[dKey] = `【バックアップ】${backups.join('、')}`;
      continue;
    }

    if (!name || EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;
    const normName = name.replace(/[（(]小児科[）)]/, '').trim();
    allClinics.add(normName);
    rawNameMap[normName] = name;
    if (!workingClinics[dKey]) workingClinics[dKey] = new Set();
    workingClinics[dKey].add(normName);
  }

  const fuzaiOpened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.FUZAI, 1);
  const fValues = fuzaiOpened.values;
  const fMap = fuzaiOpened.headerMap;
  const idxFDate = requireColumnAny(fMap, ['日付'], '医師不在拠点');
  const idxFClinic = requireColumnAny(fMap, ['拠点名'], '医師不在拠点');
  const idxFTime = requireColumnAny(fMap, ['不在時間'], '医師不在拠点');

  const fuzaiRecords = [];
  for (let i = 1; i < fValues.length; i++) {
    const row = fValues[i];
    const dObj = parseDateToSafeDateObj(row[idxFDate]);
    const name = String(row[idxFClinic] || '').trim();
    const time = String(row[idxFTime] || '').trim();
    if (!dObj || !name || !time) continue;
    if (EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;

    const normName = name.replace(/[（(]小児科[）)]/, '').trim();
    allClinics.add(normName);
    rawNameMap[normName] = name;
    fuzaiRecords.push({
      dateKey: fastFormatDate(dObj),
      name: name,
      normName: normName,
      timeRaw: time
    });
  }

  const sortedClinicList = Array.from(allClinics).sort();
  const dailyBlocks = [];

  // 5. 1週間分の各日ループ
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateKey = fastFormatDate(d);
    const dateTitle = formatJPDateWithDay(d, false);
    const isSpecialDay = (d.getMonth() + 1 === 12 && d.getDate() === 31) || (d.getMonth() + 1 === 1 && d.getDate() <= 3);

    let totalRequiredMinutes = 0;
    sortedClinicList.forEach(clinic => {
      if (openDateMap.has(clinic)) {
        const openD = new Date(openDateMap.get(clinic));
        openD.setHours(0, 0, 0, 0);
        if (d < openD) return;
      }
      if (clinic.includes('内科')) return;
      const closedTime = closedDataMap.get(`${dateKey}_${clinic}`) || closedDataMap.get(`${dateKey}_全拠点`);
      if (closedTime === '全日') return;

      let reqMin = isSpecialDay ? 420 : (clinic.includes('北葛西') ? 540 : 600);
      if (closedTime === '午前') reqMin -= 240;
      else if (closedTime === '午後' || closedTime === '夜間') reqMin -= 180;
      else if (closedTime === '午後夜間') reqMin -= 360;

      const irr = irregularMap[`${dateKey}_${clinic}`] || irregularMap[`${dateKey}_全拠点`];
      if (irr) {
        reqMin = 0;
        irr.forEach(r => { reqMin += (r.close - r.open); });
      }
      totalRequiredMinutes += Math.max(0, reqMin);
    });

    if (totalRequiredMinutes === 0) continue;

    const dailyRecords = fuzaiRecords.filter(r => r.dateKey === dateKey);
    const groupedFuzai = {};
    dailyRecords.forEach(r => {
      if (!groupedFuzai[r.normName]) groupedFuzai[r.normName] = [];
      groupedFuzai[r.normName].push(r.timeRaw);
    });

    sortedClinicList.forEach(clinic => {
      if (openDateMap.has(clinic)) {
        const openD = new Date(openDateMap.get(clinic));
        openD.setHours(0, 0, 0, 0);
        if (d < openD) return;
      }
      if (clinic.includes('内科')) return;
      const closedTime = closedDataMap.get(`${dateKey}_${clinic}`) || closedDataMap.get(`${dateKey}_全拠点`);
      if (closedTime === '全日') return;

      const hasFuzai = groupedFuzai[clinic] && groupedFuzai[clinic].length > 0;
      const isWorking = workingClinics[dateKey] && workingClinics[dateKey].has(clinic);
      if (!hasFuzai && !isWorking) {
        groupedFuzai[clinic] = isSpecialDay ? ['10:00-19:00'] : (clinic.includes('北葛西') ? ['09:00-13:00', '15:00-20:00'] : ['09:00-13:00', '15:00-21:00']);
      }
    });

    let totalGapMinutes = 0;
    const dailyOutputLines = [];

    for (const normName in groupedFuzai) {
      if (openDateMap.has(normName)) {
        const openD = new Date(openDateMap.get(normName));
        openD.setHours(0, 0, 0, 0);
        if (d < openD) continue;
      }
      const closedTime = closedDataMap.get(`${dateKey}_${normName}`) || closedDataMap.get(`${dateKey}_${normName}_内科`) || closedDataMap.get(`${dateKey}_全拠点`);
      if (closedTime === '全日') continue;

      let finalTimeStr = mergeAbsenceIntervals_(groupedFuzai[normName], isSpecialDay, closedTime, irregularMap[`${dateKey}_${normName}`] || irregularMap[`${dateKey}_全拠点`]);
      if (!finalTimeStr) continue;

      if (!isSpecialDay && normName.includes('北葛西')) {
        finalTimeStr = finalTimeStr.replace('21:00', '20:00');
      }

      if (!normName.includes('内科')) {
        finalTimeStr.split('/').forEach(p => {
          const [s, e] = p.split('-').map(safeParseTime);
          if (!isNaN(s) && !isNaN(e)) totalGapMinutes += (e - s);
        });
      }

      let displayTimeStr = finalTimeStr
        .replace(/^09:00-13:00\/15:00-/, '09:00-')
        .replace(/^10:00-13:00\/15:00-/, '10:00-');

      const record = dailyRecords.find(r => r.normName === normName);
      const displayName = record ? record.name : (rawNameMap[normName] || normName);
      dailyOutputLines.push(`【${displayName}】${displayTimeStr}`);
    }

    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;
    const requiredHours = Math.round(totalRequiredMinutes / 60);
    const filledHours = Math.round(totalFilledMinutes / 60);

    let block = `[info][title]${dateTitle}[/title]`;
    if (backupInfoMap[dateKey]) block += backupInfoMap[dateKey];
    block += `[hr]\n`;
    block += `小児科1診目充足率：${rate}%（充足済み：${filledHours}h/募集：${requiredHours}h）\n`;
    if (dailyOutputLines.length > 0) {
      dailyOutputLines.sort();
      block += `<医師不在拠点>\n` + dailyOutputLines.join('\n') + `\n`;
    } else {
      block += `充足\n`;
    }
    block += `[/info]`;
    dailyBlocks.push(block);
  }

  if (dailyBlocks.length === 0) {
    Logger.log('対象期間のデータがありませんでした。');
    return;
  }

  // 6. メンションは [toall] のみを付与して送信
  const fullMessage = '[toall]\n' + dailyBlocks.join('\n');
  sendChatworkMessage(fullMessage);
}

/**
 * 不在時間のマージ＆休館・変則営業除外ヘルパー
 */
function mergeAbsenceIntervals_(timeStrings, isSpecialDay, closedTime, irregularRule) {
  const OPEN = isSpecialDay ? 600 : 540;
  const CLOSE = isSpecialDay ? 1140 : 1260;
  const intervals = [];

  timeStrings.forEach(str => {
    String(str).split(/[,/、]/).forEach(sub => {
      const pts = sub.trim().split(/[-~〜～]/);
      if (pts.length !== 2) return;
      let s = Math.max(safeParseTime(pts[0]), OPEN);
      let e = Math.min(safeParseTime(pts[1]), CLOSE);
      if (!isNaN(s) && !isNaN(e) && s < e) intervals.push({ start: s, end: e });
    });
  });
  if (intervals.length === 0) return '';

  intervals.sort((a, b) => a.start - b.start);
  const merged = [intervals[0]];
  for (let i = 1; i < intervals.length; i++) {
    const cur = merged[merged.length - 1];
    if (cur.end >= intervals[i].start) cur.end = Math.max(cur.end, intervals[i].end);
    else merged.push(intervals[i]);
  }

  let ranges = merged.map(m => `${safeFormatTime(m.start)}-${safeFormatTime(m.end)}`);

  if (closedTime) {
    const cRanges = [];
    if (closedTime === '午前') cRanges.push({ s: 540, e: 780 });
    else if (closedTime === '午後') cRanges.push({ s: 900, e: 1080 });
    else if (closedTime === '夜間') cRanges.push({ s: 1080, e: 1260 });
    else if (closedTime === '午後夜間') cRanges.push({ s: 900, e: 1260 });

    const nextRanges = [];
    ranges.forEach(r => {
      let [s, e] = r.split('-').map(safeParseTime);
      for (const c of cRanges) {
        if (s >= e) break;
        if (c.s <= s && c.e >= e) s = e;
        else if (c.s <= s && c.e > s) s = c.e;
        else if (c.s < e && c.e >= e) e = c.s;
        else if (c.s > s && c.e < e) {
          nextRanges.push(`${safeFormatTime(s)}-${safeFormatTime(c.s)}`);
          s = c.e;
        }
      }
      if (s < e) nextRanges.push(`${safeFormatTime(s)}-${safeFormatTime(e)}`);
    });
    ranges = nextRanges;
  }

  if (irregularRule && ranges.length > 0) {
    const valid = [];
    ranges.forEach(r => {
      const [s, e] = r.split('-').map(safeParseTime);
      irregularRule.forEach(ir => {
        const as = Math.max(s, ir.open), ae = Math.min(e, ir.close);
        if (as < ae) valid.push(`${safeFormatTime(as)}-${safeFormatTime(ae)}`);
      });
    });
    ranges = valid;
  }

  return ranges.join('/');
}