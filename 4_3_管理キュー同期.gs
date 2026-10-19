/**
 * =================================================================
 * 【4_3_管理キュー同期.gs】
 * ・永続管理シート「⚙️_新規お久しぶり管理」の読み書き・マージ同期（syncTargetDoctorsQueue）
 * ・各種送信済みフラグの記録ヘルパー（markTargetQueueFlag_）
 * =================================================================
 */

/**
 * 永続管理シート「⚙️_新規お久しぶり管理」を〈採用報告〉・「お久しぶり勤務」・「貼付用」と同期し、
 * 全レコードの最新配列を返すメイン同期関数
 */
function syncTargetDoctorsQueue() {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);
  if (!ss.getSheetByName(CONFIG.SHEET.TARGET_QUEUE)) {
    initialSetup();
  }

  const qOpened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.TARGET_QUEUE, 1);
  const qSheet = qOpened.sheet;
  const qVals = qOpened.values;
  const qMap = qOpened.headerMap;

  const colIdx = {
    key:         requireColumnAny(qMap, ['ユニークキー'], CONFIG.SHEET.TARGET_QUEUE),
    type:        requireColumnAny(qMap, ['種別'], CONFIG.SHEET.TARGET_QUEUE),
    ts:          requireColumnAny(qMap, ['タイムスタンプ'], CONFIG.SHEET.TARGET_QUEUE),
    iki:         requireColumnAny(qMap, ['医籍番号'], CONFIG.SHEET.TARGET_QUEUE),
    name:        requireColumnAny(qMap, ['医師名'], CONFIG.SHEET.TARGET_QUEUE),
    gender:      requireColumnAny(qMap, ['性別'], CONFIG.SHEET.TARGET_QUEUE),
    spec:        requireColumnAny(qMap, ['専門科目'], CONFIG.SHEET.TARGET_QUEUE),
    fClinic:     requireColumnAny(qMap, ['フォーム拠点'], CONFIG.SHEET.TARGET_QUEUE),
    fDept:       requireColumnAny(qMap, ['フォーム診療科'], CONFIG.SHEET.TARGET_QUEUE),
    fDate:       requireColumnAny(qMap, ['フォーム勤務日'], CONFIG.SHEET.TARGET_QUEUE),
    fTime:       requireColumnAny(qMap, ['フォーム時間'], CONFIG.SHEET.TARGET_QUEUE),
    tClinic:     requireColumnAny(qMap, ['確定拠点'], CONFIG.SHEET.TARGET_QUEUE),
    tDept:       requireColumnAny(qMap, ['確定診療科'], CONFIG.SHEET.TARGET_QUEUE),
    tDate:       requireColumnAny(qMap, ['確定勤務日'], CONFIG.SHEET.TARGET_QUEUE),
    tTime:       requireColumnAny(qMap, ['確定時間'], CONFIG.SHEET.TARGET_QUEUE),
    status:      requireColumnAny(qMap, ['補正状態'], CONFIG.SHEET.TARGET_QUEUE),
    initSent:    requireColumnAny(qMap, ['初回通知済'], CONFIG.SHEET.TARGET_QUEUE),
    remind7Sent: requireColumnAny(qMap, ['1週間前リマインド済'], CONFIG.SHEET.TARGET_QUEUE),
    weeklySent:  requireColumnAny(qMap, ['週次リマインド済'], CONFIG.SHEET.TARGET_QUEUE),
    syncedAt:    requireColumnAny(qMap, ['最終同期日時'], CONFIG.SHEET.TARGET_QUEUE)
  };

  // 既存キューの読み込み
  const existingMap = new Map();
  for (let i = 1; i < qVals.length; i++) {
    const row = qVals[i];
    const k = String(row[colIdx.key] || '').trim();
    if (!k) continue;
    existingMap.set(k, {
      sheetRowIndex: i + 1,
      rowValues: row
    });
  }

  const shiftLookup = buildCurrentAndNextMonthShiftLookup_();
  const sourceEntries = fetchSourceEntriesFromSaiyoAndOhisa_();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 永続管理シート（⚙️_新規お久しぶり管理）とのマージ＆貼付用補正の反映
  const rowsToAppend = [];
  const finalRecords = [];

  sourceEntries.forEach(src => {
    const existing = existingMap.get(src.uniqueKey);

    let trueClinic = src.formClinic;
    let trueDept = src.formDept;
    let trueDateObj = src.formDateObj;
    let trueDateKey = src.formDateKey;
    let trueTime = src.formTime;
    let statusText = src.iki ? 'フォーム値（貼付用に前倒しなし）' : 'フォーム値（医籍番号なし）';
    let initSent = '';
    let remind7Sent = '';
    let weeklySent = '';

    if (existing) {
      const r = existing.rowValues;
      const savedTrueDate = parseDateToSafeDateObj(r[colIdx.tDate]);
      if (savedTrueDate) {
        trueDateObj = savedTrueDate;
        trueDateKey = fastFormatDate(savedTrueDate);
      }
      trueClinic  = String(r[colIdx.tClinic] || src.formClinic).trim();
      trueDept    = String(r[colIdx.tDept] || src.formDept).trim();
      trueTime    = String(r[colIdx.tTime] || src.formTime).trim();
      statusText  = String(r[colIdx.status] || statusText).trim();
      initSent    = String(r[colIdx.initSent] || '').trim();
      remind7Sent = String(r[colIdx.remind7Sent] || '').trim();
      weeklySent  = String(r[colIdx.weeklySent] || '').trim();
    }

    // 新規採用について貼付用シートで前倒しチェック
    if (src.type === '新規') {
      const resolved = resolveShiftWithPaste_(
        shiftLookup, today,
        src.iki, src.rawName, src.tsDateObj, src.formDateObj, trueDateObj,
        src.formClinic, src.formDept, src.formTime, statusText
      );
      if (resolved) {
        trueDateObj = resolved.dateObj;
        trueDateKey = resolved.dateKey;
        trueClinic  = resolved.clinic;
        trueDept    = resolved.dept;
        trueTime    = resolved.timeStr;
        statusText  = resolved.statusText;
      }
    }

    const newRowData = [];
    newRowData[colIdx.key]         = src.uniqueKey;
    newRowData[colIdx.type]        = src.type;
    newRowData[colIdx.ts]          = src.tsStr;
    newRowData[colIdx.iki]         = src.iki;
    newRowData[colIdx.name]        = src.rawName;
    newRowData[colIdx.gender]      = src.gender;
    newRowData[colIdx.spec]        = src.specialty;
    newRowData[colIdx.fClinic]     = src.formClinic;
    newRowData[colIdx.fDept]       = src.formDept;
    newRowData[colIdx.fDate]       = src.formDateKey;
    newRowData[colIdx.fTime]       = src.formTime;
    newRowData[colIdx.tClinic]     = trueClinic;
    newRowData[colIdx.tDept]       = trueDept;
    newRowData[colIdx.tDate]       = trueDateKey;
    newRowData[colIdx.tTime]       = trueTime;
    newRowData[colIdx.status]      = statusText;
    newRowData[colIdx.initSent]    = initSent;
    newRowData[colIdx.remind7Sent] = remind7Sent;
    newRowData[colIdx.weeklySent]  = weeklySent;
    newRowData[colIdx.syncedAt]    = nowStr;

    if (existing) {
      // 確定勤務日・確定拠点・確定時間・補正状態に変化があれば行を更新
      const prevTDate = fastFormatDate(parseDateToSafeDateObj(existing.rowValues[colIdx.tDate]));
      const prevTClinic = String(existing.rowValues[colIdx.tClinic] || '').trim();
      const prevTTime = String(existing.rowValues[colIdx.tTime] || '').trim();
      const prevStatus = String(existing.rowValues[colIdx.status] || '').trim();

      if (prevTDate !== trueDateKey || prevTClinic !== trueClinic || prevTTime !== trueTime || prevStatus !== statusText) {
        qSheet.getRange(existing.sheetRowIndex, 1, 1, newRowData.length).setValues([newRowData]);
      }
    } else {
      // 初回シート生成時に過去の古いデータまで一斉送信されないよう、すでに過去日のものは「過去日スキップ」を入れる
      if (trueDateObj < today) {
        newRowData[colIdx.initSent]    = '過去日済';
        newRowData[colIdx.remind7Sent] = '過去日済';
        newRowData[colIdx.weeklySent]  = '過去日済';
        initSent = '過去日済';
        remind7Sent = '過去日済';
        weeklySent = '過去日済';
      }
      rowsToAppend.push(newRowData);
    }

    const docWithTitle = src.rawName.endsWith('先生') ? src.rawName : `${src.rawName}先生`;
    const shortDoc = src.rawName.split(/[（(]/)[0].replace(/\s+/g, '').replace(/先生$/, '') + '先生';

    const recordObj = {
      uniqueKey: src.uniqueKey,
      type: src.type,
      iki: src.iki,
      doctorName: docWithTitle,
      shortDoctorName: shortDoc,
      gender: src.gender,
      specialty: src.specialty,
      dateObj: trueDateObj,
      dateKey: trueDateKey,
      clinic: trueClinic,
      dept: trueDept,
      rawTimeStr: src.type === 'お久しぶり' ? src.rawFormTime : trueTime,
      timeStr: trueTime,
      statusText: statusText,
      initialSent: Boolean(initSent),
      remind7Sent: Boolean(remind7Sent),
      weeklySent: Boolean(weeklySent),
      shinTaisei: '１診体制'
    };

    recordObj.shinTaisei = buildShinTaiseiText_(recordObj, shiftLookup.dailyClinicShiftsMap);
    finalRecords.push(recordObj);
  });

  if (rowsToAppend.length > 0) {
    qSheet.getRange(qSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return finalRecords;
}

/**
 * 永続管理シート（⚙️_新規お久しぶり管理）の特定フラグ列に送信日時を記録するヘルパー
 */
function markTargetQueueFlag_(uniqueKey, columnHeaderName, customValue) {
  const opened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.TARGET_QUEUE, 1);
  const { sheet, values, headerMap } = opened;
  const idxKey = requireColumnAny(headerMap, ['ユニークキー'], CONFIG.SHEET.TARGET_QUEUE);
  const idxFlag = requireColumnAny(headerMap, [columnHeaderName], CONFIG.SHEET.TARGET_QUEUE);
  const valToSet = customValue || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxKey] || '').trim() === uniqueKey) {
      sheet.getRange(i + 1, idxFlag + 1).setValue(valToSet);
      break;
    }
  }
}