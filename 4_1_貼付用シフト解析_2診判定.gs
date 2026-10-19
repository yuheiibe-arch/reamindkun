/**
 * =================================================================
 * 【4_1_貼付用シフト解析_2診判定.gs】
 * ・医籍番号・医師名の正規化ヘルパー
 * ・「貼付用」シート（当月・翌月）＋「確定シフト」シート（未来分）のシフトインデックス作成
 * ・医籍番号がない場合の保険用「勤務日＋氏名（またはカナ）」インデックス作成
 * ・同日・同拠点・同診療科の他医師シフトからの「診療体制（2診/1診）」自動算出
 * =================================================================
 */

/**
 * 医籍番号の正規化ヘルパー（「第427465号」や全角数字から半角数字のみを抽出）
 */
function normalizeIkiNumber_(rawVal) {
  if (rawVal === null || rawVal === undefined) return '';
  const halfWidth = String(rawVal).replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  const match = halfWidth.match(/\d{4,8}/);
  return match ? match[0] : halfWidth.replace(/[^\d]/g, '').trim();
}

/**
 * 医師名の正規化ヘルパー（ふりがなカッコ・敬称・空白・異体字を除去）
 */
function normalizeDoctorName_(rawName) {
  return String(rawName || '')
    .split(/[（(【\[「]/)[0]
    .replace(/先生|医師|Dr\.?|[\s ]/gi, '')
    .replace(/髙/g, '高')
    .replace(/﨑|嵜/g, '崎')
    .replace(/齋|齊/g, '斎')
    .replace(/澤/g, '沢')
    .replace(/國/g, '国')
    .replace(/邊|邉/g, '辺')
    .replace(/濱|濵/g, '浜')
    .replace(/廣/g, '広')
    .replace(/櫻/g, '桜')
    .replace(/柳/g, '栁')
    .trim();
}

/**
 * カッコ内またはカタカナ文字列からのカナ正規化ヘルパー
 * （例：「中田 竜介（ナカダ リュウスケ）」→「ナカダリュウスケ」）
 */
function extractAndNormalizeKana_(rawText) {
  if (!rawText) return '';
  const str = String(rawText);
  const parenMatch = str.match(/[（(]([^）)]+)[）)]/);
  const target = parenMatch ? parenMatch[1] : str;
  const katakana = target.replace(/[\u3041-\u3096]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
  return katakana.replace(/[^ァ-ヶー]/g, '').trim();
}

/**
 * 「貼付用」シート（および来月以降の「確定シフト」シート）から全シフトを読み込み、
 * 医籍番号・正規化氏名・「勤務日＋氏名（保険用）」から引けるインデックス＋日別クリニックシフト一覧を作成する（列ハードコードなし）
 */
function buildCurrentAndNextMonthShiftLookup_() {
  const { normalizeKey } = getClinicMentionDictionary();
  const opened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.PASTE, 1);
  const { values, headerMap } = opened;

  const idxName   = requireColumnAny(headerMap, ['名前', '氏名'], '貼付用');
  const idxKana   = optionalColumnAny(headerMap, ['名前（カタカナ）', 'フリガナ', 'カナ']);
  const idxIki    = requireColumnAny(headerMap, ['医籍番号'], '貼付用');
  const idxClinic = requireColumnAny(headerMap, ['クリニック名', '拠点名'], '貼付用');
  const idxDept   = optionalColumnAny(headerMap, ['診療科']);
  const idxDate   = requireColumnAny(headerMap, ['勤務日'], '貼付用');
  const idxStart  = requireColumnAny(headerMap, ['勤務開始時間'], '貼付用');
  const idxEnd1   = optionalColumnAny(headerMap, ['終了時間1']);
  const idxEnd2   = optionalColumnAny(headerMap, ['終了時間2']);
  const idxEnd3   = optionalColumnAny(headerMap, ['終了時間3']);
  const idxEnd    = requireColumnAny(headerMap, ['勤務終了時間'], '貼付用');

  const today = new Date();
  const startOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const endOfNextMonth = new Date(today.getFullYear(), today.getMonth() + 2, 0);

  const byIki = new Map();
  const byName = new Map();
  const byKana = new Map();
  const byDateAndName = new Map();
  const byDateAndKana = new Map();
  const ikiByKanji = new Map();
  const ikiByKana = new Map();
  const dailyClinicShiftsMap = new Map();

  const EXCLUDED = ['有給', '欠勤', '院外勤務', 'バックアップ', '医師会', '嘱託医', '休館'];

  // 中抜けを考慮したインターバル抽出ヘルパー（2診判定用）
  const extractPasteIntervals = (row) => {
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

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const docName = String(row[idxName] || '').trim();
    const clinic  = String(row[idxClinic] || '').trim();
    const dept    = idxDept !== null ? String(row[idxDept] || '').trim() : '';
    if (!docName || !clinic || docName.includes('休館') || clinic.includes('バックアップ') || clinic.includes('有給') || clinic.includes('欠勤')) continue;

    const iki = normalizeIkiNumber_(row[idxIki]);
    const nKey = normalizeDoctorName_(docName);
    const rawKana = (idxKana !== null && row[idxKana]) ? String(row[idxKana]).trim() : docName;
    const kanaKey = extractAndNormalizeKana_(rawKana);

    // 範囲外の日付であっても「氏名→医籍番号」の辞書には登録しておく
    if (iki) {
      if (nKey && !ikiByKanji.has(nKey)) ikiByKanji.set(nKey, iki);
      if (kanaKey && !ikiByKana.has(kanaKey)) ikiByKana.set(kanaKey, iki);
    }

    const dObj = parseDateToSafeDateObj(row[idxDate]);
    if (!dObj) continue;

    // 当月1日〜翌月末日の範囲に絞る
    if (dObj < startOfCurrentMonth || dObj > endOfNextMonth) continue;

    const sMin = safeParseTime(row[idxStart]);
    const eMin = safeParseTime(row[idxEnd]);
    const timeStr = (!isNaN(sMin) && !isNaN(eMin)) ? `${safeFormatTime(sMin)}-${safeFormatTime(eMin)}` : '';
    const dKey = fastFormatDate(dObj);

    const shiftEntry = {
      source: '貼付用',
      dateObj: dObj,
      dateKey: dKey,
      clinic: clinic,
      dept: dept,
      timeStr: timeStr,
      iki: iki,
      rowNumber: i + 1
    };

    if (iki) {
      if (!byIki.has(iki)) byIki.set(iki, []);
      byIki.get(iki).push(shiftEntry);
    }

    if (nKey) {
      if (!byName.has(nKey)) byName.set(nKey, []);
      byName.get(nKey).push(shiftEntry);

      // ★保険用：「勤務日＋漢字氏名」インデックス（時間がずれていても日付と氏名で一致させる）
      const dateNameKey = `${dKey}_${nKey}`;
      if (!byDateAndName.has(dateNameKey)) byDateAndName.set(dateNameKey, []);
      byDateAndName.get(dateNameKey).push(shiftEntry);
    }

    if (kanaKey) {
      if (!byKana.has(kanaKey)) byKana.set(kanaKey, []);
      byKana.get(kanaKey).push(shiftEntry);

      // ★保険用：「勤務日＋カナ氏名」インデックス
      const dateKanaKey = `${dKey}_${kanaKey}`;
      if (!byDateAndKana.has(dateKanaKey)) byDateAndKana.set(dateKanaKey, []);
      byDateAndKana.get(dateKanaKey).push(shiftEntry);
    }

    // 2診判定用リストへの格納（ワクチン専任等の除外キーワードもチェック）
    if (!EXCLUDED.some(ex => docName.includes(ex) || clinic.includes(ex)) && !dept.includes('ワクチン専任')) {
      const intervals = extractPasteIntervals(row);
      const shortDocName = docName.split(/[（(]/)[0].replace(/先生|医師|[\s ]/g, '') + '先生';
      if (!dailyClinicShiftsMap.has(dKey)) dailyClinicShiftsMap.set(dKey, []);
      const dayList = dailyClinicShiftsMap.get(dKey);
      intervals.forEach(inv => {
        dayList.push({
          docName: docName,
          shortDocName: shortDocName,
          normName: nKey,
          iki: iki,
          clinic: clinic,
          normClinic: normalizeKey(clinic),
          dept: dept,
          start: inv.start,
          end: inv.end
        });
      });
    }
  }

  // 来月以降の未来分が含まれる「確定シフト」シートもスキャンしてインデックスを補完（動的ヘッダー検出）
  if (CONFIG.SS_ID.SHIFT && CONFIG.SHEET.KAKUTEI_SHIFT) {
    try {
      const kSs = SpreadsheetApp.openById(CONFIG.SS_ID.SHIFT);
      const kSheet = kSs.getSheetByName(CONFIG.SHEET.KAKUTEI_SHIFT);
      if (kSheet) {
        const kVals = kSheet.getDataRange().getValues();
        let hRowIdx = 0;
        let maxScore = -1;
        const kwList = ['名前', '氏名', '医師', '医籍', '拠点', 'クリニック', '勤務日', '日付'];
        for (let r = 0; r < Math.min(5, kVals.length); r++) {
          let sc = 0;
          kVals[r].forEach(c => {
            const s = String(c).replace(/\r?\n/g, '').trim();
            if (kwList.some(kw => s.includes(kw))) sc++;
          });
          if (sc > maxScore) { maxScore = sc; hRowIdx = r; }
        }
        if (maxScore > 0) {
          const kMap = buildHeaderMap(kVals[hRowIdx]);
          const kIdxName   = optionalColumnAny(kMap, ['名前', '氏名', '医師名', '医師氏名']);
          const kIdxLast   = optionalColumnAny(kMap, ['姓', '氏']);
          const kIdxFirst  = optionalColumnAny(kMap, ['名']);
          const kIdxKana   = optionalColumnAny(kMap, ['名前（カタカナ）', 'フリガナ', 'カナ']);
          const kIdxIki    = optionalColumnAny(kMap, ['医籍番号', '医籍']);
          const kIdxClinic = optionalColumnAny(kMap, ['クリニック名', '拠点名', '拠点']);
          const kIdxDept   = optionalColumnAny(kMap, ['診療科']);
          const kIdxDate   = optionalColumnAny(kMap, ['勤務日', '日付']);
          const kIdxStart  = optionalColumnAny(kMap, ['勤務開始時間', '開始時間']);
          const kIdxEnd    = optionalColumnAny(kMap, ['勤務終了時間', '終了時間']);

          for (let i = hRowIdx + 1; i < kVals.length; i++) {
            const row = kVals[i];
            let docName = kIdxName !== null ? String(row[kIdxName] || '').trim() : '';
            if (!docName && kIdxLast !== null && kIdxFirst !== null) {
              docName = `${String(row[kIdxLast] || '').trim()} ${String(row[kIdxFirst] || '').trim()}`.trim();
            }
            if (!docName || docName.includes('休館') || docName.includes('バックアップ') || docName.includes('有給') || docName.includes('欠勤')) continue;

            const nKey = normalizeDoctorName_(docName);
            const rawKana = (kIdxKana !== null && row[kIdxKana]) ? String(row[kIdxKana]).trim() : docName;
            const kanaKey = extractAndNormalizeKana_(rawKana);
            const iki = (kIdxIki !== null ? normalizeIkiNumber_(row[kIdxIki]) : '') || ikiByKanji.get(nKey) || '';

            if (iki) {
              if (nKey && !ikiByKanji.has(nKey)) ikiByKanji.set(nKey, iki);
              if (kanaKey && !ikiByKana.has(kanaKey)) ikiByKana.set(kanaKey, iki);
            }

            const dObj = kIdxDate !== null ? parseDateToSafeDateObj(row[kIdxDate]) : null;
            const clinic = kIdxClinic !== null ? String(row[kIdxClinic] || '').trim() : '';
            if (!dObj || !clinic || dObj < startOfCurrentMonth) continue;

            const dept = kIdxDept !== null ? String(row[kIdxDept] || '').trim() : '';
            const sMin = kIdxStart !== null ? safeParseTime(row[kIdxStart]) : NaN;
            const eMin = kIdxEnd !== null ? safeParseTime(row[kIdxEnd]) : NaN;
            const timeStr = (!isNaN(sMin) && !isNaN(eMin)) ? `${safeFormatTime(sMin)}-${safeFormatTime(eMin)}` : '';
            const dKey = fastFormatDate(dObj);

            // すでに貼付用で同日・同医師・同拠点が登録されていれば重複追加しない
            const dateNameKey = `${dKey}_${nKey}`;
            if (byDateAndName.has(dateNameKey) && byDateAndName.get(dateNameKey).some(x => x.clinic === clinic)) continue;

            const shiftEntry = {
              source: '確定シフト',
              dateObj: dObj,
              dateKey: dKey,
              clinic: clinic,
              dept: dept,
              timeStr: timeStr,
              iki: iki,
              rowNumber: i + 1
            };

            if (iki) {
              if (!byIki.has(iki)) byIki.set(iki, []);
              byIki.get(iki).push(shiftEntry);
            }
            if (nKey) {
              if (!byName.has(nKey)) byName.set(nKey, []);
              byName.get(nKey).push(shiftEntry);
              if (!byDateAndName.has(dateNameKey)) byDateAndName.set(dateNameKey, []);
              byDateAndName.get(dateNameKey).push(shiftEntry);
            }
            if (kanaKey) {
              if (!byKana.has(kanaKey)) byKana.set(kanaKey, []);
              byKana.get(kanaKey).push(shiftEntry);
              const dateKanaKey = `${dKey}_${kanaKey}`;
              if (!byDateAndKana.has(dateKanaKey)) byDateAndKana.set(dateKanaKey, []);
              byDateAndKana.get(dateKanaKey).push(shiftEntry);
            }

            if (!EXCLUDED.some(ex => docName.includes(ex) || clinic.includes(ex)) && !dept.includes('ワクチン専任') && !isNaN(sMin) && !isNaN(eMin) && sMin < eMin) {
              const shortDocName = docName.split(/[（(]/)[0].replace(/先生|医師|[\s ]/g, '') + '先生';
              if (!dailyClinicShiftsMap.has(dKey)) dailyClinicShiftsMap.set(dKey, []);
              dailyClinicShiftsMap.get(dKey).push({
                docName: docName,
                shortDocName: shortDocName,
                normName: nKey,
                iki: iki,
                clinic: clinic,
                normClinic: normalizeKey(clinic),
                dept: dept,
                start: sMin,
                end: eMin
              });
            }
          }
        }
      }
    } catch (e) {
      Logger.log('確定シフトシート補完スキャンスキップ: ' + e.message);
    }
  }

  byIki.forEach(list => list.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime()));
  byName.forEach(list => list.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime()));
  byKana.forEach(list => list.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime()));

  return {
    byIki,
    byName,
    byKana,
    byDateAndName,
    byDateAndKana,
    ikiByKanji,
    ikiByKana,
    dailyClinicShiftsMap
  };
}

/**
 * 対象医師のシフトに対し、同日・同拠点・同診療科の他医師のシフトと照合して
 * 「〇〇先生と２診（09:00-13:00）」や「〇〇先生と２診（09:00-12:00）/１診体制」または「１診体制」を生成するヘルパー
 */
function buildShinTaiseiText_(targetItem, dailyClinicShiftsMap) {
  const { normalizeKey } = getClinicMentionDictionary();
  const normClinic = normalizeKey(targetItem.clinic);
  const targetDept = String(targetItem.dept || '小児科').trim();

  // 比較元の時間文字列を取得
  const baseTimeStr = targetItem.rawTimeStr || targetItem.timeStr || '';
  const match = String(baseTimeStr).match(/(\d{1,2}:\d{2})\s*[-~〜～]\s*(\d{1,2}:\d{2})/);
  if (!match) return '１診体制';

  let targetStart = safeParseTime(match[1]);
  const targetEnd = safeParseTime(match[2]);
  if (isNaN(targetStart) || isNaN(targetEnd) || targetStart >= targetEnd) return '１診体制';

  // オリエンテーション30分前倒し（08:30開始, 14:30開始, 17:30開始など）が含まれている場合、
  // 実診療の開始時刻（09:00, 15:00, 18:00）を基準にして「完全2診か／一部1診か」を判定する
  if (targetStart % 60 === 30 && (targetEnd - targetStart) > 60) {
    targetStart += 30;
  }

  const dayShifts = dailyClinicShiftsMap.get(targetItem.dateKey) || [];
  const normSelfName = normalizeDoctorName_(targetItem.doctorName);

  // 同日・同拠点・同診療科の「本人以外」のシフトを抽出し、重なり区間を計算
  const overlaps = [];
  dayShifts.forEach(other => {
    if (other.normClinic !== normClinic) return;
    const otherDept = other.dept || '小児科';
    if (targetDept && otherDept && targetDept !== otherDept) return;

    // 本人除外（医籍番号一致 または 正規化氏名一致）
    if (targetItem.iki && other.iki && targetItem.iki === other.iki) return;
    if (normSelfName && other.normName === normSelfName) return;

    const overlapStart = Math.max(targetStart, other.start);
    const overlapEnd = Math.min(targetEnd, other.end);
    if (overlapStart < overlapEnd) {
      overlaps.push({
        shortDocName: other.shortDocName,
        start: overlapStart,
        end: overlapEnd,
        rangeStr: `${safeFormatTime(overlapStart)}-${safeFormatTime(overlapEnd)}`
      });
    }
  });

  if (overlaps.length === 0) {
    return '１診体制';
  }

  // 開始時間順にソート
  overlaps.sort((a, b) => a.start - b.start || a.end - b.end);
  const twoShinParts = overlaps.map(o => `${o.shortDocName}と２診（${o.rangeStr}）`);

  // 対象の実診療時間（targetStart〜targetEnd）のうち、他医師と重なっている合計時間を算出
  let coveredMinutes = 0;
  const mergedOverlaps = [{ start: overlaps[0].start, end: overlaps[0].end }];
  for (let i = 1; i < overlaps.length; i++) {
    const cur = mergedOverlaps[mergedOverlaps.length - 1];
    if (cur.end >= overlaps[i].start) {
      cur.end = Math.max(cur.end, overlaps[i].end);
    } else {
      mergedOverlaps.push({ start: overlaps[i].start, end: overlaps[i].end });
    }
  }
  mergedOverlaps.forEach(m => { coveredMinutes += (m.end - m.start); });

  const totalTargetMinutes = targetEnd - targetStart;
  if (coveredMinutes < totalTargetMinutes) {
    // 一部時間帯は2診、残りの時間帯は1診の場合
    return `${twoShinParts.join('・')}/１診体制`;
  }

  return twoShinParts.join('・');
}