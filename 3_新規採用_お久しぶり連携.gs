/**
 * =================================================================
 * 【3_新規採用_お久しぶり連携.gs】（要件4・5）
 * ・機能②：DS部に〈採用報告〉が投稿されたタイミングで個人情報カット版を拠点へ投稿
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

  // 〈採用報告〉シートから該当医師データを動的に取得（列ハードコードなし）
  const saiyoOpened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.SAIYO_REPORT, 1);
  const { values, headerMap } = saiyoOpened;
  if (values.length <= 1) return;

  const idxClinic = requireColumnAny(headerMap, ['拠点'], '〈採用報告〉');
  const idxDept   = requireColumnAny(headerMap, ['診療科'], '〈採用報告〉');
  const idxDate   = requireColumnAny(headerMap, ['勤務日'], '〈採用報告〉');
  const idxTime1  = requireColumnAny(headerMap, ['勤務時間'], '〈採用報告〉');
  const idxTime2  = optionalColumnAny(headerMap, ['勤務時間が上記に当てはまらい場合']);
  const idxName   = requireColumnAny(headerMap, ['氏名'], '〈採用報告〉');
  const idxGender = requireColumnAny(headerMap, ['性別'], '〈採用報告〉');
  const idxSpec   = requireColumnAny(headerMap, ['専門科目'], '〈採用報告〉');

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

  const clinic = String(targetRow[idxClinic] || '').trim();
  const dept = String(targetRow[idxDept] || '').trim();
  const dateObj = parseDateToSafeDateObj(targetRow[idxDate]);
  const timeStr = String((idxTime2 !== null && targetRow[idxTime2]) ? targetRow[idxTime2] : targetRow[idxTime1] || '').trim();
  const rawName = String(targetRow[idxName] || '').trim();
  const doctorName = rawName.endsWith('先生') ? rawName : `${rawName}先生`;
  const gender = String(targetRow[idxGender] || '').trim();
  const specialty = String(targetRow[idxSpec] || '').trim();

  const w = dateObj ? ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()] : '';
  const dateFormatted = dateObj ? `${fastFormatDate(dateObj)}(${w})` : String(targetRow[idxDate] || '').trim();
  const uniqueKey = `NEW_INITIAL_${clinic}_${fastFormatDate(dateObj || new Date())}_${doctorName}`;

  if (hasSentHistory('新規採用_初回', uniqueKey)) {
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
[/info]`;

  sendChatworkMessage(message);
  recordSentHistory('新規採用_初回', uniqueKey, doctorName, clinic, fastFormatDate(dateObj || new Date()));
}

/**
 * 機能③：「お久しぶり勤務」シートの未送信行をチェックし、
 * オリエンテーション必須なら30分前補正して投稿する
 */
function runOhisashiburiInitialCheck() {
  const opened = openSheetWithHeaders(CONFIG.SS_ID.SAIYO, CONFIG.SHEET.OHISASHIBURI, 1);
  const { values, headerMap } = opened;
  if (values.length <= 1) return;

  const idxClinic = requireColumnAny(headerMap, ['拠点'], 'お久しぶり勤務');
  const idxDate   = requireColumnAny(headerMap, ['勤務日'], 'お久しぶり勤務');
  const idxTime1  = requireColumnAny(headerMap, ['勤務時間（自動で30分前に補正します。）', '勤務時間'], 'お久しぶり勤務');
  const idxTime2  = optionalColumnAny(headerMap, ['勤務時間が上記に当てはまらい場合']);
  const idxName   = requireColumnAny(headerMap, ['氏名'], 'お久しぶり勤務');
  const idxDept   = optionalColumnAny(headerMap, ['診療科']);
  const idxOrien  = requireColumnAny(headerMap, ['30分前オリエンテーションの有無'], 'お久しぶり勤務');

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const clinic = String(row[idxClinic] || '').trim();
    const rawName = String(row[idxName] || '').trim();
    const dateObj = parseDateToSafeDateObj(row[idxDate]);
    if (!clinic || !rawName || !dateObj) continue;

    const dateKey = fastFormatDate(dateObj);
    const uniqueKey = `OHISA_INITIAL_${clinic}_${dateKey}_${rawName.replace(/\s+/g, '')}`;
    if (hasSentHistory('お久しぶり_初回', uniqueKey)) continue;

    const rawTime = String((idxTime2 !== null && row[idxTime2]) ? row[idxTime2] : row[idxTime1] || '').trim();
    const isOrienRequired = String(row[idxOrien] || '').trim() === '必須';
    const adjustedTime = shiftStartTimeBack30Min(rawTime, isOrienRequired);

    const dept = idxDept !== null ? String(row[idxDept] || '').trim() : '';
    const clinicDisplay = (CONFIG.DEPT_LABEL_CLINICS.includes(clinic) && dept) ? `【${clinic}/${dept}】` : `【${clinic}】`;
    const doctorName = rawName.endsWith('先生') ? rawName : `${rawName}先生`;
    const w = ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
    const displayDate = `${dateKey}(${w})`;

    const mentionHeader = buildClinicMentionsText([clinic], false);
    const message = `${mentionHeader}[info][title]お久しぶり勤務&オリエンテーション依頼[/title]
対象拠点：${clinicDisplay}
勤務日：${displayDate}
対象時間：${adjustedTime}
医師名：${doctorName}
[code]お久しぶり勤務のため、院内オリエンテーションをお願いいたします。[/code]
[/info]`;

    sendChatworkMessage(message);
    recordSentHistory('お久しぶり_初回', uniqueKey, doctorName, clinic, dateKey);
    Utilities.sleep(1500);
  }
}