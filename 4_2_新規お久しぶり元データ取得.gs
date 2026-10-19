/**
 * =================================================================
 * 【4_2_新規お久しぶり元データ取得.gs】
 * ・〈採用報告〉シートと「お久しぶり勤務」シートからの全エントリー動的取得
 * ・「DS直接応募初回勤務」シートからの全エントリー動的取得（正）
 * ・【保険】フォーム未登録の初回・お久しぶり医師を「貼付用」の「スタッフコメント」列から自動補完
 * ・貼付用シートのシフトインデックスを使った前倒し・日程変更・「勤務日＋氏名」一致解決ヘルパー
 * =================================================================
 */

/**
 * 貼付用シートから「前倒しシフト（または同日の実シフト時間補正）」を解決するヘルパー
 * ※2ヶ月前など過去に初回勤務を終えた医師が、当月・翌月の別シフトで再度「初回」と
 *   誤判定されないよう、過去日のシフトへの後ろ倒し上書きは一切行わない。
 */
function resolveShiftWithPaste_(shiftLookup, today, iki, rawName, tsDateObj, formDateObj, currentTrueDateObj, defaultClinic, defaultDept, defaultTime, currentStatus) {
  // すでに確定勤務日が今日より過去で確定済みの場合は、過去の初回勤務が済んでいるため後続シフトで上書きしない
  if (currentTrueDateObj && currentTrueDateObj < today && currentStatus && currentStatus.includes('貼付用')) {
    return null;
  }

  const nKey = normalizeDoctorName_(rawName);
  const kanaKey = (typeof extractAndNormalizeKana_ === 'function') ? extractAndNormalizeKana_(rawName) : '';
  const fKey = fastFormatDate(formDateObj);

  let candidates = [];
  if (iki && shiftLookup.byIki.has(iki)) {
    candidates = shiftLookup.byIki.get(iki);
  } else if (nKey && shiftLookup.byName.has(nKey)) {
    candidates = shiftLookup.byName.get(nKey);
  } else if (kanaKey && shiftLookup.byKana && shiftLookup.byKana.has(kanaKey)) {
    candidates = shiftLookup.byKana.get(kanaKey);
  }

  // 採用報告のタイムスタンプ（登録日）の7日前以降〜フォーム勤務日以下のシフトに絞る
  // ※これにより「2ヶ月前に初回勤務を終えた人の来月の2回目シフト」を拾ってしまう誤検知を防止し、
  //   かつ「フォーム勤務日より前倒しになった初回シフト（例: 10/19→10/13）」を確実に捕捉する
  if (candidates.length > 0) {
    const minValidMs = tsDateObj
      ? (tsDateObj.getTime() - 7 * 86400000)
      : (formDateObj.getTime() - 45 * 86400000);
    const maxValidMs = currentTrueDateObj ? currentTrueDateObj.getTime() : formDateObj.getTime();

    const earlierOrSameShifts = candidates.filter(s => s.dateObj.getTime() >= minValidMs && s.dateObj.getTime() <= maxValidMs);
    if (earlierOrSameShifts.length > 0) {
      const firstShift = earlierOrSameShifts[0];
      const isChanged = (firstShift.dateKey !== fKey || firstShift.clinic !== defaultClinic);
      return {
        dateObj: firstShift.dateObj,
        dateKey: firstShift.dateKey,
        clinic: firstShift.clinic || defaultClinic,
        dept: firstShift.dept || defaultDept,
        timeStr: firstShift.timeStr || defaultTime,
        statusText: isChanged
          ? `★貼付用で前倒し補正 [元:${fKey} ${defaultClinic}(${defaultTime}) → 実:${firstShift.dateKey} ${firstShift.clinic}(${firstShift.timeStr})]`
          : `貼付用と一致確認済 (${firstShift.timeStr})`
      };
    }
  }

  // ★【保険】医籍番号がない（またはタイムスタンプ範囲外だった）場合でも、
  // 「フォームの勤務日 ＋ 氏名（またはカナ）」が一致するシフトがあれば、時間がずれていても同一シフトとみなして採用する
  const dateNameKey = `${fKey}_${nKey}`;
  const dateKanaKey = kanaKey ? `${fKey}_${kanaKey}` : '';
  const sameDayMatches = (shiftLookup.byDateAndName && shiftLookup.byDateAndName.get(dateNameKey))
    || (dateKanaKey && shiftLookup.byDateAndKana && shiftLookup.byDateAndKana.get(dateKanaKey))
    || [];

  if (sameDayMatches.length > 0) {
    const exactClinicShift = sameDayMatches.find(s => s.clinic === defaultClinic) || sameDayMatches[0];
    return {
      dateObj: exactClinicShift.dateObj,
      dateKey: exactClinicShift.dateKey,
      clinic: exactClinicShift.clinic || defaultClinic,
      dept: exactClinicShift.dept || defaultDept,
      timeStr: exactClinicShift.timeStr || defaultTime,
      statusText: `★保険(勤務日＋氏名)で一致確認 (${exactClinicShift.timeStr || defaultTime})`
    };
  }

  // もしフォーム勤務日当日にシフトがなく、フォーム勤務日より後（直近）に日程変更された場合（ただしフォーム日が今月・来月内の場合のみ）
  const startOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  if (candidates.length > 0 && formDateObj >= startOfCurrentMonth) {
    const minValidMs = tsDateObj
      ? (tsDateObj.getTime() - 7 * 86400000)
      : (formDateObj.getTime() - 45 * 86400000);
    const futureShifts = candidates.filter(s => s.dateObj.getTime() >= minValidMs);
    if (futureShifts.length > 0) {
      const firstShift = futureShifts[0];
      return {
        dateObj: firstShift.dateObj,
        dateKey: firstShift.dateKey,
        clinic: firstShift.clinic || defaultClinic,
        dept: firstShift.dept || defaultDept,
        timeStr: firstShift.timeStr || defaultTime,
        statusText: `★貼付用で日程変更補正 [元:${fKey} ${defaultClinic}(${defaultTime}) → 実:${firstShift.dateKey} ${firstShift.clinic}(${firstShift.timeStr})]`
      };
    }
  }

  return null;
}

/**
 * 〈採用報告〉シート・「お久しぶり勤務」シート・「DS直接応募初回勤務」シートの全行をスキャンし、
 * さらにフォーム未登録分を「貼付用」の「スタッフコメント」列から保険で補完して
 * 正規化した元データ配列（sourceEntries）を返す関数（動的ヘッダー取得）
 */
function fetchSourceEntriesFromSaiyoAndOhisa_(shiftLookup) {
  const sourceEntries = [];
  const registeredIkiSet = new Set();
  const registeredNameSet = new Set();

  // 医籍番号が空欄の場合にシフト側辞書から氏名一致で医籍番号を引くヘルパー
  const resolveIkiByNameFallback = (rawName) => {
    if (!shiftLookup) return '';
    const nKey = normalizeDoctorName_(rawName);
    if (nKey && shiftLookup.ikiByKanji && shiftLookup.ikiByKanji.has(nKey)) {
      return shiftLookup.ikiByKanji.get(nKey);
    }
    const kanaKey = (typeof extractAndNormalizeKana_ === 'function') ? extractAndNormalizeKana_(rawName) : '';
    if (kanaKey && shiftLookup.ikiByKana && shiftLookup.ikiByKana.has(kanaKey)) {
      return shiftLookup.ikiByKana.get(kanaKey);
    }
    return '';
  };

  // -----------------------------------------------------------------
  // A. 〈採用報告〉シートの全行スキャン（動的ヘッダー取得）
  // -----------------------------------------------------------------
  const saiyoOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.SAIYO_REPORT, 1);
  const sVals = saiyoOpened.values;
  const sMap = saiyoOpened.headerMap;

  const sIdxTs     = optionalColumnAny(sMap, ['タイムスタンプ']);
  const sIdxClinic = requireColumnAny(sMap, ['拠点'], '〈採用報告〉');
  const sIdxDept   = requireColumnAny(sMap, ['採用したキャップスクリニックの診療科', '診療科'], '〈採用報告〉');
  const sIdxDate   = requireColumnAny(sMap, ['勤務日'], '〈採用報告〉');
  const sIdxTime1  = requireColumnAny(sMap, ['勤務時間'], '〈採用報告〉');
  const sIdxTime2  = optionalColumnAny(sMap, ['勤務時間が上記に当てはまらい場合', '当てはまらない場合']);
  const sIdxName   = requireColumnAny(sMap, ['氏名'], '〈採用報告〉');
  const sIdxGender = requireColumnAny(sMap, ['性別'], '〈採用報告〉');
  const sIdxSpec   = requireColumnAny(sMap, ['専門科目'], '〈採用報告〉');
  const sIdxIki    = optionalColumnAny(sMap, ['医籍番号']);
  const sIdxHist   = optionalColumnAny(sMap, ['学歴・職歴']);

  for (let i = 1; i < sVals.length; i++) {
    const row = sVals[i];
    const rawName = String(row[sIdxName] || '').trim();
    const formDateObj = parseDateToSafeDateObj(row[sIdxDate]);
    if (!rawName || !formDateObj) continue;

    const tsDateObj = sIdxTs !== null ? parseDateToSafeDateObj(row[sIdxTs]) : null;
    const tsStr = tsDateObj ? fastFormatDate(tsDateObj) : '';
    let iki = sIdxIki !== null ? normalizeIkiNumber_(row[sIdxIki]) : '';
    if (!iki && sIdxHist !== null) {
      const m = String(row[sIdxHist] || '').match(/第?\s*([0-9０-９]{4,8})\s*号/);
      if (m) iki = normalizeIkiNumber_(m[1]);
    }
    if (!iki) {
      iki = resolveIkiByNameFallback(rawName);
    }

    const nKey = normalizeDoctorName_(rawName);
    if (iki) registeredIkiSet.add(iki);
    if (nKey) registeredNameSet.add(nKey);

    const formClinic = String(row[sIdxClinic] || '').trim();
    const formDept = String(row[sIdxDept] || '').trim();
    const formTime = String((sIdxTime2 !== null && row[sIdxTime2]) ? row[sIdxTime2] : row[sIdxTime1] || '').trim().replace(/[~〜～]/g, '-');
    const formDateKey = fastFormatDate(formDateObj);

    const uniqueKey = `SAIYO_${iki || nKey}_${formDateKey}_${formClinic}`;

    sourceEntries.push({
      uniqueKey: uniqueKey,
      type: '新規',
      tsDateObj: tsDateObj,
      tsStr: tsStr,
      iki: iki,
      rawName: rawName,
      gender: String(row[sIdxGender] || '').trim(),
      specialty: String(row[sIdxSpec] || '').trim(),
      formClinic: formClinic,
      formDept: formDept,
      formDateObj: formDateObj,
      formDateKey: formDateKey,
      rawFormTime: formTime,
      formTime: formTime,
      isOrienRequired: false
    });
  }

  // -----------------------------------------------------------------
  // B. 「お久しぶり勤務」シートの全行スキャン（動的ヘッダー取得）
  // -----------------------------------------------------------------
  const ohisaOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.OHISASHIBURI, 1);
  const oVals = ohisaOpened.values;
  const oMap = ohisaOpened.headerMap;

  const oIdxTs     = optionalColumnAny(oMap, ['タイムスタンプ']);
  const oIdxClinic = requireColumnAny(oMap, ['拠点'], 'お久しぶり勤務');
  const oIdxDate   = requireColumnAny(oMap, ['勤務日'], 'お久しぶり勤務');
  const oIdxTime1  = requireColumnAny(oMap, ['勤務時間（自動で30分前に補正します。）', '勤務時間'], 'お久しぶり勤務');
  const oIdxTime2  = optionalColumnAny(oMap, ['勤務時間が上記に当てはまらい場合', '当てはまらない場合']);
  const oIdxName   = requireColumnAny(oMap, ['氏名'], 'お久しぶり勤務');
  const oIdxIki    = optionalColumnAny(oMap, ['医籍番号']);
  const oIdxDept   = optionalColumnAny(oMap, ['診療科']);
  const oIdxOrien  = requireColumnAny(oMap, ['30分前オリエンテーションの有無'], 'お久しぶり勤務');

  for (let i = 1; i < oVals.length; i++) {
    const row = oVals[i];
    const rawName = String(row[oIdxName] || '').trim();
    const formDateObj = parseDateToSafeDateObj(row[oIdxDate]);
    if (!rawName || !formDateObj) continue;

    const tsDateObj = oIdxTs !== null ? parseDateToSafeDateObj(row[oIdxTs]) : null;
    const tsStr = tsDateObj ? fastFormatDate(tsDateObj) : '';
    let iki = oIdxIki !== null ? normalizeIkiNumber_(row[oIdxIki]) : '';
    if (!iki) {
      iki = resolveIkiByNameFallback(rawName);
    }

    const nKey = normalizeDoctorName_(rawName);
    if (iki) registeredIkiSet.add(iki);
    if (nKey) registeredNameSet.add(nKey);

    const formClinic = String(row[oIdxClinic] || '').trim();
    const formDept = oIdxDept !== null ? String(row[oIdxDept] || '').trim() : '';
    const rawTime = String((oIdxTime2 !== null && row[oIdxTime2]) ? row[oIdxTime2] : row[oIdxTime1] || '').trim().replace(/[~〜～]/g, '-');
    const isOrienRequired = String(row[oIdxOrien] || '').trim() === '必須';
    const adjustedFormTime = shiftStartTimeBack30Min(rawTime, isOrienRequired);
    const formDateKey = fastFormatDate(formDateObj);

    const uniqueKey = `OHISA_${iki || nKey}_${formDateKey}_${formClinic}`;

    sourceEntries.push({
      uniqueKey: uniqueKey,
      type: 'お久しぶり',
      tsDateObj: tsDateObj,
      tsStr: tsStr,
      iki: iki,
      rawName: rawName,
      gender: '',
      specialty: '',
      formClinic: formClinic,
      formDept: formDept,
      formDateObj: formDateObj,
      formDateKey: formDateKey,
      rawFormTime: rawTime,
      formTime: adjustedFormTime,
      isOrienRequired: isOrienRequired
    });
  }

  // -----------------------------------------------------------------
  // C. 「DS直接応募初回勤務」シートの全行スキャン（正・動的ヘッダー取得）
  // -----------------------------------------------------------------
  try {
    const saiyoSs = SpreadsheetApp.openById(CONFIG.SS_ID.SAIYO);
    const directSheetName = (CONFIG.SHEET && CONFIG.SHEET.DIRECT_FIRST) ? CONFIG.SHEET.DIRECT_FIRST : 'DS直接応募初回勤務';
    let directSheet = saiyoSs.getSheetByName(directSheetName);
    if (!directSheet) {
      directSheet = saiyoSs.getSheets().find(sh => sh.getName().includes('直接応募')) || null;
    }

    if (directSheet && directSheet.getLastRow() > 1) {
      const dVals = directSheet.getDataRange().getValues();
      const dMap = buildHeaderMap(dVals[0]);

      const dIdxTs     = optionalColumnAny(dMap, ['タイムスタンプ']);
      const dIdxClinic = optionalColumnAny(dMap, ['拠点', 'クリニック']);
      const dIdxDept   = optionalColumnAny(dMap, ['採用したキャップスクリニックの診療科', '診療科']);
      const dIdxDate   = optionalColumnAny(dMap, ['勤務日', '初回勤務日', '日付']);
      const dIdxTime1  = optionalColumnAny(dMap, ['勤務時間（自動で30分前に補正します。）', '勤務時間', '時間']);
      const dIdxTime2  = optionalColumnAny(dMap, ['勤務時間が上記に当てはまらい場合', '当てはまらない場合']);
      const dIdxName   = optionalColumnAny(dMap, ['氏名', '名前', '医師名']);
      const dIdxGender = optionalColumnAny(dMap, ['性別']);
      const dIdxSpec   = optionalColumnAny(dMap, ['専門科目', '専門']);
      const dIdxIki    = optionalColumnAny(dMap, ['医籍番号', '医籍']);
      const dIdxOrien  = optionalColumnAny(dMap, ['30分前オリエンテーションの有無', 'オリエンテーション']);

      if (dIdxClinic !== null && dIdxDate !== null && dIdxName !== null) {
        for (let i = 1; i < dVals.length; i++) {
          const row = dVals[i];
          const rawName = String(row[dIdxName] || '').trim();
          const formDateObj = parseDateToSafeDateObj(row[dIdxDate]);
          if (!rawName || !formDateObj) continue;

          const tsDateObj = dIdxTs !== null ? parseDateToSafeDateObj(row[dIdxTs]) : null;
          const tsStr = tsDateObj ? fastFormatDate(tsDateObj) : '';
          let iki = dIdxIki !== null ? normalizeIkiNumber_(row[dIdxIki]) : '';
          if (!iki) {
            iki = resolveIkiByNameFallback(rawName);
          }

          const nKey = normalizeDoctorName_(rawName);
          if (iki) registeredIkiSet.add(iki);
          if (nKey) registeredNameSet.add(nKey);

          const formClinic = String(row[dIdxClinic] || '').trim();
          const formDept = dIdxDept !== null ? String(row[dIdxDept] || '').trim() : '';
          const rawTime = String((dIdxTime2 !== null && row[dIdxTime2]) ? row[dIdxTime2] : (dIdxTime1 !== null ? row[dIdxTime1] : '')).trim().replace(/[~〜～]/g, '-');
          const isOrienRequired = dIdxOrien !== null ? (String(row[dIdxOrien] || '').trim() === '必須') : false;
          const adjustedFormTime = shiftStartTimeBack30Min(rawTime, isOrienRequired);
          const formDateKey = fastFormatDate(formDateObj);

          const uniqueKey = `DIRECT_${iki || nKey}_${formDateKey}_${formClinic}`;

          sourceEntries.push({
            uniqueKey: uniqueKey,
            type: '新規',
            tsDateObj: tsDateObj,
            tsStr: tsStr,
            iki: iki,
            rawName: rawName,
            gender: dIdxGender !== null ? String(row[dIdxGender] || '').trim() : '',
            specialty: dIdxSpec !== null ? String(row[dIdxSpec] || '').trim() : '',
            formClinic: formClinic,
            formDept: formDept,
            formDateObj: formDateObj,
            formDateKey: formDateKey,
            rawFormTime: rawTime,
            formTime: adjustedFormTime,
            isOrienRequired: isOrienRequired
          });
        }
      }
    }
  } catch (e) {
    Logger.log('DS直接応募初回勤務シート読み込みスキップ: ' + e.message);
  }

  // -----------------------------------------------------------------
  // D. 【保険】フォーム未登録の医師を「貼付用」シートの「スタッフコメント」列から自動補完
  // -----------------------------------------------------------------
  try {
    const pasteOpened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.PASTE, 1);
    const pVals = pasteOpened.values;
    const pMap = pasteOpened.headerMap;
    const rawHeaderRow = pVals[0] || [];

    const pIdxName   = requireColumnAny(pMap, ['名前', '氏名'], '貼付用');
    const pIdxIki    = requireColumnAny(pMap, ['医籍番号'], '貼付用');
    const pIdxClinic = requireColumnAny(pMap, ['クリニック名', '拠点名'], '貼付用');
    const pIdxDept   = optionalColumnAny(pMap, ['診療科']);
    const pIdxDate   = requireColumnAny(pMap, ['勤務日'], '貼付用');
    const pIdxStart  = requireColumnAny(pMap, ['勤務開始時間'], '貼付用');
    const pIdxEnd    = requireColumnAny(pMap, ['勤務終了時間'], '貼付用');

    // 見出しに「スタッフコメント」を含む列（スタッフコメント1〜5）のみを動的取得
    const staffCommentColIndices = [];
    rawHeaderRow.forEach((h, idx) => {
      const hName = String(h || '').replace(/\r?\n/g, '').trim();
      if (hName.includes('スタッフコメント')) {
        staffCommentColIndices.push(idx);
      }
    });

    const FIRST_WORK_KEYWORDS = ['ナイズ初', 'オリエン30分', 'オリエン３０分', 'オリエン', '初回勤務', '初勤務', '初回'];
    const today = new Date();
    const startOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const endOfNextMonth = new Date(today.getFullYear(), today.getMonth() + 2, 0);

    const commentFallbackByDoc = new Map();

    for (let i = 1; i < pVals.length; i++) {
      const row = pVals[i];
      const dObj = parseDateToSafeDateObj(row[pIdxDate]);
      if (!dObj || dObj < startOfCurrentMonth || dObj > endOfNextMonth) continue;

      const docName = String(row[pIdxName] || '').trim();
      const clinic  = String(row[pIdxClinic] || '').trim();
      if (!docName || !clinic || docName.includes('休館') || clinic.includes('バックアップ') || clinic.includes('有給') || clinic.includes('欠勤')) continue;

      const matchedComments = [];
      staffCommentColIndices.forEach(cIdx => {
        const text = String(row[cIdx] || '').trim();
        if (text && FIRST_WORK_KEYWORDS.some(kw => text.includes(kw))) {
          matchedComments.push(text);
        }
      });
      if (matchedComments.length === 0) continue;

      const iki = normalizeIkiNumber_(row[pIdxIki]) || resolveIkiByNameFallback(docName);
      const nKey = normalizeDoctorName_(docName);

      // すでにフォーム側（A・B・C）で登録されている医師はフォーム側を正とするためスキップ
      if ((iki && registeredIkiSet.has(iki)) || (nKey && registeredNameSet.has(nKey))) continue;

      const docKey = iki || nKey;
      if (!docKey) continue;

      const sMin = safeParseTime(row[pIdxStart]);
      const eMin = safeParseTime(row[pIdxEnd]);
      const timeStr = (!isNaN(sMin) && !isNaN(eMin)) ? `${safeFormatTime(sMin)}-${safeFormatTime(eMin)}` : '';
      const dept = pIdxDept !== null ? String(row[pIdxDept] || '').trim() : '';
      const dKey = fastFormatDate(dObj);

      // コメント内に「ぶり」「久しぶり」が含まれ、かつ「ナイズ初」ではない場合は「お久しぶり」、それ以外は「新規」
      const joinedComment = matchedComments.join(' ');
      const isOhisaByComment = (joinedComment.includes('ぶり') || joinedComment.includes('久しぶり')) && !joinedComment.includes('ナイズ初');

      if (!commentFallbackByDoc.has(docKey) || dObj < commentFallbackByDoc.get(docKey).formDateObj) {
        commentFallbackByDoc.set(docKey, {
          uniqueKey: `COMMENT_${iki || nKey}_${dKey}_${clinic}`,
          type: isOhisaByComment ? 'お久しぶり' : '新規',
          tsDateObj: dObj,
          tsStr: dKey,
          iki: iki,
          rawName: docName,
          gender: '',
          specialty: '',
          formClinic: clinic,
          formDept: dept,
          formDateObj: dObj,
          formDateKey: dKey,
          rawFormTime: timeStr,
          formTime: timeStr,
          isOrienRequired: false
        });
      }
    }

    commentFallbackByDoc.forEach(entry => {
      sourceEntries.push(entry);
    });
  } catch (e) {
    Logger.log('貼付用スタッフコメント保険スキャンスキップ: ' + e.message);
  }

  return sourceEntries;
}