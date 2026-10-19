/**
 * =================================================================
 * 【99_デバッグ確認.gs】
 * 「貼付用」シートの「スタッフコメント」列（スタッフコメント1〜5）のみを動的に対象とし、
 * 「ナイズ初」「オリエン30分」「初回勤務」等のコメントがあるシフトを検出して、
 * 今週分の対象一覧とリマインド文案をログ出力するデバッグスクリプト
 * ★列番号のハードコード一切なし ／ シート書き込み・Chatwork送信なし
 * =================================================================
 */

function debugCheckDirectApplicantsThisWeek() {
  Logger.log('================ 直接応募（スタッフコメント列のみ検知）デバッグ開始 ================');

  // 列インデックス(0始まり)をA, B, C...に変換するヘルパー
  const idxToColLetter = (idx) => {
    let letter = '';
    let temp = idx + 1;
    while (temp > 0) {
      let mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  };

  // 1. 期間設定（今週：月曜〜日曜）
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dayOfWeek = today.getDay(); // 0:日, 1:月 ... 6:土
  const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  const startOfThisWeek = new Date(today);
  startOfThisWeek.setDate(today.getDate() + diffToMonday);
  const endOfThisWeek = new Date(startOfThisWeek);
  endOfThisWeek.setDate(startOfThisWeek.getDate() + 6);

  const thisWeekStartKey = fastFormatDate(startOfThisWeek);
  const thisWeekEndKey   = fastFormatDate(endOfThisWeek);

  Logger.log(`・基準日（今日） : ${fastFormatDate(today)}`);
  Logger.log(`・今週の対象期間 : ${thisWeekStartKey} 〜 ${thisWeekEndKey}`);

  // 2. 既存のシフトインデックス＆〈採用報告〉・「お久しぶり勤務」データを取得
  const shiftLookup = buildCurrentAndNextMonthShiftLookup_();
  const formSourceEntries = fetchSourceEntriesFromSaiyoAndOhisa_(shiftLookup);

  // すでに〈採用報告〉または「お久しぶり勤務」に存在する医師のセット（医籍番号 ＆ 正規化氏名）
  const existingFormIkiSet = new Set();
  const existingFormNameSet = new Set();
  formSourceEntries.forEach(e => {
    if (e.iki) existingFormIkiSet.add(e.iki);
    const nKey = normalizeDoctorName_(e.rawName);
    if (nKey) existingFormNameSet.add(nKey);
  });

  // -----------------------------------------------------------------
  // 3. 「貼付用」シートから「スタッフコメント」列のみを動的に全取得してスキャン
  // -----------------------------------------------------------------
  const pasteOpened = openSheetWithHeaders(CONFIG.SS_ID.MIJYUSOKU, CONFIG.SHEET.PASTE, 1);
  const pVals = pasteOpened.values;
  const pMap = pasteOpened.headerMap;
  const rawHeaderRow = pVals[0] || [];

  const idxName   = requireColumnAny(pMap, ['名前', '氏名'], '貼付用');
  const idxIki    = requireColumnAny(pMap, ['医籍番号'], '貼付用');
  const idxClinic = requireColumnAny(pMap, ['クリニック名', '拠点名'], '貼付用');
  const idxDept   = optionalColumnAny(pMap, ['診療科']);
  const idxDate   = requireColumnAny(pMap, ['勤務日'], '貼付用');
  const idxStart  = requireColumnAny(pMap, ['勤務開始時間'], '貼付用');
  const idxEnd    = requireColumnAny(pMap, ['勤務終了時間'], '貼付用');
  const idxKeii   = optionalColumnAny(pMap, ['入職経緯']);
  const idxShokai = optionalColumnAny(pMap, ['シフト紹介経緯']);

  // ★見出しに「スタッフコメント」を含む列（スタッフコメント1〜5）のみを動的抽出
  const staffCommentColumns = [];
  rawHeaderRow.forEach((h, idx) => {
    const hName = String(h || '').replace(/\r?\n/g, '').trim();
    if (hName.includes('スタッフコメント')) {
      staffCommentColumns.push({ idx: idx, colLetter: idxToColLetter(idx), name: hName });
    }
  });

  Logger.log(`\n▼ 動的検出したスタッフコメント列（計 ${staffCommentColumns.length} 列）:`);
  Logger.log('  ' + staffCommentColumns.map(c => `${c.colLetter}列("${c.name}")`).join(', '));

  // 初回勤務と判定するキーワード群
  const FIRST_WORK_KEYWORDS = ['ナイズ初', 'オリエン30分', 'オリエン３０分', 'オリエン', '初回勤務', '初勤務', '初回'];

  // 同一医師が複数日ヒットした場合に「最も早いシフト（初回）」に絞るためのマップ
  const commentHitsByDoctor = new Map();
  const allThisWeekCommentRawHits = [];

  for (let i = 1; i < pVals.length; i++) {
    const row = pVals[i];
    const dObj = parseDateToSafeDateObj(row[idxDate]);
    if (!dObj) continue;

    const docName = String(row[idxName] || '').trim();
    const clinic  = String(row[idxClinic] || '').trim();
    if (!docName || !clinic || docName.includes('休館') || clinic.includes('バックアップ') || clinic.includes('有給') || clinic.includes('欠勤')) continue;

    // スタッフコメント列のみをチェック
    const hitDetails = [];
    staffCommentColumns.forEach(col => {
      const cellText = String(row[col.idx] || '').trim();
      if (!cellText) return;
      const matchedKw = FIRST_WORK_KEYWORDS.find(kw => cellText.includes(kw));
      if (matchedKw) {
        hitDetails.push(`${col.colLetter}列[${col.name}]: "${cellText}" (キーワード:${matchedKw})`);
      }
    });

    if (hitDetails.length === 0) continue;

    const dKey = fastFormatDate(dObj);
    const iki = normalizeIkiNumber_(row[idxIki]);
    const nKey = normalizeDoctorName_(docName);
    const dept = idxDept !== null ? String(row[idxDept] || '').trim() : '';
    const sMin = safeParseTime(row[idxStart]);
    const eMin = safeParseTime(row[idxEnd]);
    const timeStr = (!isNaN(sMin) && !isNaN(eMin)) ? `${safeFormatTime(sMin)}-${safeFormatTime(eMin)}` : '';
    const keiiStr = idxKeii !== null ? String(row[idxKeii] || '').trim() : '';
    const shokaiStr = idxShokai !== null ? String(row[idxShokai] || '').trim() : '';

    const isAlreadyInForm = (iki && existingFormIkiSet.has(iki)) || (nKey && existingFormNameSet.has(nKey));

    const hitRecord = {
      rowNumber: i + 1,
      dateObj: dObj,
      dateKey: dKey,
      docName: docName,
      shortDoctorName: docName.split(/[（(]/)[0].replace(/先生|医師|[\s ]/g, '') + '先生',
      normName: nKey,
      iki: iki,
      clinic: clinic,
      dept: dept,
      timeStr: timeStr,
      keiiStr: keiiStr,
      shokaiStr: shokaiStr,
      hitDetailsText: hitDetails.join(' / '),
      isAlreadyInForm: isAlreadyInForm
    };

    if (dKey >= thisWeekStartKey && dKey <= thisWeekEndKey) {
      allThisWeekCommentRawHits.push(hitRecord);
    }

    // 医師ごとに最も早い日付のシフトを保持
    const docKey = iki || nKey;
    if (!commentHitsByDoctor.has(docKey) || dObj < commentHitsByDoctor.get(docKey).dateObj) {
      commentHitsByDoctor.set(docKey, hitRecord);
    }
  }

  // -----------------------------------------------------------------
  // 4. 今週分（10/19〜10/25）でスタッフコメントにヒットした全シフトの詳細ログ出力
  // -----------------------------------------------------------------
  Logger.log(`\n==================== ① 今週（${thisWeekStartKey}〜${thisWeekEndKey}）のスタッフコメント検知一覧（計 ${allThisWeekCommentRawHits.length} 件） ====================`);
  if (allThisWeekCommentRawHits.length === 0) {
    Logger.log('今週のシフトで該当キーワード（ナイズ初 / オリエン30分 / 初回勤務 等）を含むスタッフコメントはありませんでした。');
  } else {
    allThisWeekCommentRawHits.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());
    allThisWeekCommentRawHits.forEach((h, idx) => {
      const w = ['日', '月', '火', '水', '木', '金', '土'][h.dateObj.getDay()];
      const formStatus = h.isAlreadyInForm
        ? '（既に〈採用報告〉または「お久しぶり」シートにも記載あり）'
        : '★【新規検出！】〈採用報告〉シート未記載（スタッフコメントのみで検知）';
      Logger.log(
        `[${idx + 1}] 貼付用${h.rowNumber}行目: ${h.dateKey}(${w}) ${h.timeStr} ＠ ${h.clinic}${h.dept ? '/' + h.dept : ''}\n` +
        `    医師名 : ${h.docName} (医籍番号: ${h.iki || 'なし'}) | 入職経緯: "${h.keiiStr}" / 紹介経緯: "${h.shokaiStr}"\n` +
        `    検知内容: ${h.hitDetailsText}\n` +
        `    判定  : ${formStatus}`
      );
    });
  }

  // -----------------------------------------------------------------
  // 5. 既存の〈採用報告〉・「お久しぶり」＋ 今回の「スタッフコメント検知」を統合した今週のリマインド文案プレビュー
  // -----------------------------------------------------------------
  const syncedFormRecords = syncTargetDoctorsQueue();
  const integratedWeeklyList = syncedFormRecords
    .filter(e => e.dateKey >= thisWeekStartKey && e.dateKey <= thisWeekEndKey);

  // 〈採用報告〉シートに未掲載だったスタッフコメント検知医師（今週分）を追加
  commentHitsByDoctor.forEach(hit => {
    if (hit.isAlreadyInForm) return;
    if (hit.dateKey < thisWeekStartKey || hit.dateKey > thisWeekEndKey) return;

    const itemObj = {
      type: '新規',
      iki: hit.iki,
      doctorName: hit.docName.endsWith('先生') ? hit.docName : `${hit.docName}先生`,
      shortDoctorName: hit.shortDoctorName,
      dateObj: hit.dateObj,
      dateKey: hit.dateKey,
      clinic: hit.clinic,
      dept: hit.dept,
      rawTimeStr: hit.timeStr,
      timeStr: hit.timeStr,
      shinTaisei: '１診体制',
      isDirectFromComment: true,
      commentSummary: hit.hitDetailsText
    };
    itemObj.shinTaisei = buildShinTaiseiText_(itemObj, shiftLookup.dailyClinicShiftsMap);
    integratedWeeklyList.push(itemObj);
  });

  integratedWeeklyList.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

  Logger.log(`\n==================== ② 統合後の「今週の新規勤務医師・お久しぶり勤務医師」プレビュー（計 ${integratedWeeklyList.length} 名） ====================`);
  if (integratedWeeklyList.length === 0) {
    Logger.log('今週の対象者はいません。');
  } else {
    const byDate = new Map();
    const clinicsForMention = new Set();

    integratedWeeklyList.forEach(item => {
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
    const previewMessage = `${mentionHeader}[info][title]今週の新規勤務医師・お久しぶり勤務医師[/title]
${bodyLines.join('\n')}
オリエンテーションの設定のほどお願い申し上げます。
上記はあくまで予定になりますので、変更の場合は速やかにご連絡させていただきます。
[/info]`;

    Logger.log(`\n▼ 【統合後の週次リマインド文案】\n${previewMessage}`);
  }

  Logger.log('\n================ 直接応募（スタッフコメント列のみ検知）デバッグ終了 ================');
}