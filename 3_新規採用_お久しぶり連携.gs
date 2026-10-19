/**
 * =================================================================
 * 【3_新規採用_お久しぶり連携.gs】（要件4・5）
 * ・機能②：DS部に〈採用報告〉が投稿されたタイミングで個人情報カット版を拠点へ投稿
 *   （※貼付用シートに前倒しのシフトがあればそちらを「正」として通知＆永続管理キューに記録）
 * ・機能③：「お久しぶり勤務」シートへのフォーム追加時に30分前補正して拠点へ投稿
 * =================================================================
 */

/**
 * 機能②：URLシートの「生成された文章」（または〈採用報告〉シートの最新行）から
 * 個人情報をカットした「新規医師採用報告」を作成して拠点へ投稿する
 */
function runNewHirePostFromUrlSheet() {
  const urlOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.URL_SHEET, 1);
  const urlMap = urlOpened.headerMap;
  const idxGenText = requireColumnAny(urlMap, ['生成された文章'], 'URL');
  const c2Message = String(urlOpened.values[1] ? urlOpened.values[1][idxGenText] : '').trim();

  // 〈採用報告〉以外は対象外（〈不採用報告〉や〈DS承認報告〉は絶対に送らない）
  const isSaiyoReport = (c2Message.includes('〈採用報告〉') || c2Message.includes('<採用報告>'))
    && !c2Message.includes('不採用報告')
    && !c2Message.includes('DS承認報告');

  if (!isSaiyoReport) {
    Logger.log('URLシートの文章が〈採用報告〉ではないため、拠点への投稿をスキップしました。');
    return;
  }

  // 永続管理シート（⚙️_新規お久しぶり管理）があれば同期を実行
  let syncedRecords = [];
  if (typeof syncTargetDoctorsQueue === 'function') {
    syncedRecords = syncTargetDoctorsQueue().filter(r => r.type === '新規');
  }

  // 〈採用報告〉シートから該当医師データを動的に取得（列ハードコードなし）
  const saiyoOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.SAIYO_REPORT, 1);
  const { values, headerMap } = saiyoOpened;
  if (values.length <= 1) return;

  const idxClinic = requireColumnAny(headerMap, ['拠点'], '〈採用報告〉');
  const idxDept   = requireColumnAny(headerMap, ['採用したキャップスクリニックの診療科', '診療科'], '〈採用報告〉');
  const idxDate   = requireColumnAny(headerMap, ['勤務日'], '〈採用報告〉');
  const idxTime1  = requireColumnAny(headerMap, ['勤務時間'], '〈採用報告〉');
  const idxTime2  = optionalColumnAny(headerMap, ['勤務時間が上記に当てはまらい場合', '当てはまらない場合']);
  const idxName   = requireColumnAny(headerMap, ['氏名'], '〈採用報告〉');
  const idxGender = requireColumnAny(headerMap, ['性別'], '〈採用報告〉');
  const idxSpec   = requireColumnAny(headerMap, ['専門科目'], '〈採用報告〉');
  const idxIki    = optionalColumnAny(headerMap, ['医籍番号']);
  const idxHist   = optionalColumnAny(headerMap, ['学歴・職歴']);

  // 生成された文章に含まれる氏名と一致する最新行を下から探す（見つからなければ有効な最終行）
  let targetRow = null;
  for (let i = values.length - 1; i >= 1; i--) {
    const rawName = String(values[i][idxName] || '').trim();
    if (!rawName) continue;
    if (!targetRow) targetRow = values[i];
    const cleanName = rawName.split(/[（(]/)[0].replace(/\s+/g, '');
    if (cleanName && c2Message.replace(/\s+/g, '').includes(cleanName)) {
      targetRow = values[i];
      break;
    }
  }

  if (!targetRow) return;

  let clinic = String(targetRow[idxClinic] || '').trim();
  let dept = String(targetRow[idxDept] || '').trim();
  let dateObj = parseDateToSafeDateObj(targetRow[idxDate]);
  let timeStr = String((idxTime2 !== null && targetRow[idxTime2]) ? targetRow[idxTime2] : targetRow[idxTime1] || '').trim();
  const rawName = String(targetRow[idxName] || '').trim();
  const doctorName = rawName.endsWith('先生') ? rawName : `${rawName}先生`;
  const gender = String(targetRow[idxGender] || '').trim();
  const specialty = String(targetRow[idxSpec] || '').trim();

  // 医籍番号の抽出
  let iki = idxIki !== null ? String(targetRow[idxIki] || '').replace(/[^\d]/g, '').trim() : '';
  if (!iki && idxHist !== null) {
    const m = String(targetRow[idxHist] || '').match(/第?\s*([0-9０-９]{4,8})\s*号/);
    if (m) iki = m[1].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  }

  // 永続管理シート（貼付用シート突合済データ）に該当医師がいれば、確定した勤務日・拠点・時間を「正」として採用
  let matchedQueueRec = null;
  if (syncedRecords.length > 0) {
    const cleanTargetName = rawName.split(/[（(]/)[0].replace(/\s+/g, '');
    for (let i = syncedRecords.length - 1; i >= 0; i--) {
      const r = syncedRecords[i];
      if ((iki && r.iki === iki) || r.shortDoctorName.replace(/先生$/, '') === cleanTargetName) {
        matchedQueueRec = r;
        clinic = r.clinic || clinic;
        dept = r.dept || dept;
        dateObj = r.dateObj || dateObj;
        timeStr = r.timeStr || timeStr;
        break;
      }
    }
  }

  // 診療体制（〇〇先生と２診 / １診体制）の取得
  let shinTaisei = (matchedQueueRec && matchedQueueRec.shinTaisei) ? matchedQueueRec.shinTaisei : '';
  if (!shinTaisei && typeof buildCurrentAndNextMonthShiftLookup_ === 'function' && typeof buildShinTaiseiText_ === 'function' && dateObj) {
    const lookup = buildCurrentAndNextMonthShiftLookup_();
    shinTaisei = buildShinTaiseiText_({
      clinic: clinic,
      dept: dept,
      dateKey: fastFormatDate(dateObj),
      doctorName: doctorName,
      iki: iki,
      rawTimeStr: timeStr,
      timeStr: timeStr
    }, lookup.dailyClinicShiftsMap);
  }
  if (!shinTaisei) shinTaisei = '１診体制';

  const w = dateObj ? ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()] : '';
  const dateFormatted = dateObj ? `${fastFormatDate(dateObj)}(${w})` : String(targetRow[idxDate] || '').trim();
  const uniqueKey = matchedQueueRec
    ? `NEW_INITIAL_${matchedQueueRec.uniqueKey}`
    : `NEW_INITIAL_${clinic}_${fastFormatDate(dateObj || new Date())}_${doctorName}`;

  if ((matchedQueueRec && matchedQueueRec.initialSent) || hasSentHistory('新規採用_初回', uniqueKey)) {
    Logger.log(`送信済みのためスキップ: ${uniqueKey}`);
    return;
  }

  const clinicDisplay = dept ? `【${clinic}/${dept}】` : `【${clinic}】`;
  const mentionHeader = buildClinicMentionsText([clinic], false);

  const message = `${mentionHeader}[info][title]新規医師採用報告[/title]
勤務拠点：${clinicDisplay}
勤務日：${dateFormatted} ${timeStr}
医師名：${doctorName}
性別：${gender}
専門科目：${specialty}
診療体制：${shinTaisei}
[/info]`;

  sendChatworkMessage(message);
  recordSentHistory('新規採用_初回', uniqueKey, doctorName, clinic, fastFormatDate(dateObj || new Date()));

  if (matchedQueueRec && typeof markTargetQueueFlag_ === 'function') {
    markTargetQueueFlag_(matchedQueueRec.uniqueKey, '初回通知済');
  }
}

/**
 * 機能③：「お久しぶり勤務」シートの未送信行をチェックし、
 * オリエンテーション必須なら30分前補正して投稿する
 */
function runOhisashiburiInitialCheck() {
  // 永続管理シート（⚙️_新規お久しぶり管理）があれば先に同期を実行
  let syncedOhisaMap = new Map();
  if (typeof syncTargetDoctorsQueue === 'function') {
    const synced = syncTargetDoctorsQueue().filter(r => r.type === 'お久しぶり');
    synced.forEach(r => {
      syncedOhisaMap.set(`${r.clinic}_${r.dateKey}_${r.shortDoctorName.replace(/先生$/, '')}`, r);
    });
  }

  let shiftLookupForFallback = null;

  const opened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.OHISASHIBURI, 1);
  const { values, headerMap } = opened;
  if (values.length <= 1) return;

  const idxClinic = requireColumnAny(headerMap, ['拠点'], 'お久しぶり勤務');
  const idxDate   = requireColumnAny(headerMap, ['勤務日'], 'お久しぶり勤務');
  const idxTime1  = requireColumnAny(headerMap, ['勤務時間（自動で30分前に補正します。）', '勤務時間'], 'お久しぶり勤務');
  const idxTime2  = optionalColumnAny(headerMap, ['勤務時間が上記に当てはまらい場合', '当てはまらない場合']);
  const idxName   = requireColumnAny(headerMap, ['氏名'], 'お久しぶり勤務');
  const idxIki    = optionalColumnAny(headerMap, ['医籍番号']);
  const idxDept   = optionalColumnAny(headerMap, ['診療科']);
  const idxOrien  = requireColumnAny(headerMap, ['30分前オリエンテーションの有無'], 'お久しぶり勤務');

  // 今日の日付（0時0分）を取得し、すでに勤務日が過ぎている過去データは再送対象外にする
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const clinic = String(row[idxClinic] || '').trim();
    const rawName = String(row[idxName] || '').trim();
    const dateObj = parseDateToSafeDateObj(row[idxDate]);
    if (!clinic || !rawName || !dateObj) continue;

    // すでに勤務日が過ぎている過去データはスキップ（送信履歴削除後の再送防止）
    if (dateObj < today) continue;

    const dateKey = fastFormatDate(dateObj);
    const cleanShortName = rawName.split(/[（(]/)[0].replace(/\s+/g, '').replace(/先生$/, '');
    const matchedQueueRec = syncedOhisaMap.get(`${clinic}_${dateKey}_${cleanShortName}`);

    const uniqueKey = matchedQueueRec
      ? `OHISA_INITIAL_${matchedQueueRec.uniqueKey}`
      : `OHISA_INITIAL_${clinic}_${dateKey}_${rawName.replace(/\s+/g, '')}`;

    // 永続管理シートで「初回通知済」または送信履歴シートに記録があればスキップ
    if ((matchedQueueRec && matchedQueueRec.initialSent) || hasSentHistory('お久しぶり_初回', uniqueKey)) continue;

    const rawTime = String((idxTime2 !== null && row[idxTime2]) ? row[idxTime2] : row[idxTime1] || '').trim();
    const isOrienRequired = String(row[idxOrien] || '').trim() === '必須';
    const adjustedTime = shiftStartTimeBack30Min(rawTime, isOrienRequired);

    const dept = idxDept !== null ? String(row[idxDept] || '').trim() : '';
    const clinicDisplay = (CONFIG.DEPT_LABEL_CLINICS.includes(clinic) && dept) ? `【${clinic}/${dept}】` : `【${clinic}】`;
    const doctorName = rawName.endsWith('先生') ? rawName : `${rawName}先生`;
    const w = ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
    const displayDate = `${dateKey}(${w})`;

    // 診療体制（〇〇先生と２診 / １診体制）の取得
    let shinTaisei = (matchedQueueRec && matchedQueueRec.shinTaisei) ? matchedQueueRec.shinTaisei : '';
    if (!shinTaisei && typeof buildCurrentAndNextMonthShiftLookup_ === 'function' && typeof buildShinTaiseiText_ === 'function') {
      if (!shiftLookupForFallback) {
        shiftLookupForFallback = buildCurrentAndNextMonthShiftLookup_();
      }
      const iki = idxIki !== null ? String(row[idxIki] || '').replace(/[^\d]/g, '').trim() : '';
      shinTaisei = buildShinTaiseiText_({
        clinic: clinic,
        dept: dept,
        dateKey: dateKey,
        doctorName: doctorName,
        iki: iki,
        rawTimeStr: rawTime,
        timeStr: adjustedTime
      }, shiftLookupForFallback.dailyClinicShiftsMap);
    }
    if (!shinTaisei) shinTaisei = '１診体制';

    const mentionHeader = buildClinicMentionsText([clinic], false);
    const message = `${mentionHeader}[info][title]お久しぶり勤務&オリエンテーション依頼[/title]
対象拠点：${clinicDisplay}
勤務日：${displayDate}
対象時間：${adjustedTime}
医師名：${doctorName}
診療体制：${shinTaisei}
[code]お久しぶり勤務のため、院内オリエンテーションをお願いいたします。[/code]
[/info]`;

    sendChatworkMessage(message);
    recordSentHistory('お久しぶり_初回', uniqueKey, doctorName, clinic, dateKey);

    if (matchedQueueRec && typeof markTargetQueueFlag_ === 'function') {
      markTargetQueueFlag_(matchedQueueRec.uniqueKey, '初回通知済');
    }
    Utilities.sleep(1500);
  }
}