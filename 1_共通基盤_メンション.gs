/**
 * =================================================================
 * 【1_共通基盤_メンション.gs】
 * ・動的ヘッダー検索（列番号ハードコード完全禁止）
 * ・拠点マスタからのメンション（共用アカウント・拠点長ID）自動解決（横並び対応）
 * ・内閣府公式ページ（gaiyou.html / syukujitsu.csv）からの祝日動的取得
 * ・Chatwork送信＆送信履歴記録
 * =================================================================
 */

/**
 * 1行目の配列から「見出し名 → 列インデックス(0始まり)」のマップを生成
 */
function buildHeaderMap(headerRow) {
  const map = {};
  if (!headerRow) return map;
  for (let i = 0; i < headerRow.length; i++) {
    const key = String(headerRow[i]).replace(/\r?\n/g, '').trim();
    if (key !== '' && !(key in map)) {
      map[key] = i;
    }
  }
  return map;
}

/**
 * 複数の見出し候補から列インデックス(0始まり)を取得（必須列用）
 */
function requireColumnAny(headerMap, candidates, sheetLabel) {
  const list = Array.isArray(candidates) ? candidates : [candidates];
  for (const name of list) {
    if (name in headerMap) return headerMap[name];
    for (const key in headerMap) {
      if (key.includes(name)) return headerMap[key];
    }
  }
  throw new Error(`シート「${sheetLabel}」に見出し「${list.join(' / ')}」が見つかりません。`);
}

/**
 * 複数の見出し候補から列インデックス(0始まり)を取得（任意列用：なければnull）
 */
function optionalColumnAny(headerMap, candidates) {
  const list = Array.isArray(candidates) ? candidates : [candidates];
  for (const name of list) {
    if (name in headerMap) return headerMap[name];
    for (const key in headerMap) {
      if (key.includes(name)) return headerMap[key];
    }
  }
  return null;
}

/**
 * 指定スプレッドシート・シートを開き、valuesとheaderMapを返す
 */
function openSheetWithHeaders(ssId, sheetName, headerRowNumber = 1) {
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません (SS_ID: ${ssId})`);
  const values = sheet.getDataRange().getValues();
  const headerRow = values.length >= headerRowNumber ? values[headerRowNumber - 1] : [];
  return {
    ss: ss,
    sheet: sheet,
    values: values,
    headerMap: buildHeaderMap(headerRow)
  };
}

/**
 * 拠点マスタ（拠点名シート）を読み込み、どんな表記揺れからでも
 * 「正規記載」「共用アカウント（拠点To）」「拠点長ID（拠点長To）」「開院日」を引ける辞書を返す
 */
function getClinicMentionDictionary() {
  if (getClinicMentionDictionary.cache) return getClinicMentionDictionary.cache;

  const opened = openSheetWithHeaders(CONFIG.SS_ID.MASTER, CONFIG.SHEET.MASTER_CLINIC, 1);
  const { values, headerMap, sheet } = opened;

  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nameColIndices = [];
  rawHeaders.forEach((h, idx) => {
    const hStr = String(h).trim();
    if (hStr === '正規記載' || hStr.includes('表記揺れ')) {
      nameColIndices.push(idx);
    }
  });

  const idxSeiki = requireColumnAny(headerMap, ['正規記載'], '拠点名');
  const idxShared = requireColumnAny(headerMap, ['共用アカウント'], '拠点名');
  const idxLeader = requireColumnAny(headerMap, ['拠点長ID'], '拠点名');
  const idxOpenDate = optionalColumnAny(headerMap, ['開院日']);

  const dict = new Map();
  const openDateMap = new Map();

  const normalizeKey = (str) => {
    return String(str || '')
      .replace(/キャップスクリニック/g, '')
      .replace(/[【】\[\]（）() \s]/g, '')
      .replace(/小児科|内科/g, '')
      .replace(/^\d+\./, '')
      .replace(/D$/i, '')
      .trim();
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const seiki = String(row[idxSeiki] || '').trim();
    if (!seiki || seiki === 'MQC') continue;

    const sharedTag = String(row[idxShared] || '').trim();
    const leaderTag = String(row[idxLeader] || '').trim();
    const openDate = (idxOpenDate !== null && row[idxOpenDate] instanceof Date) ? row[idxOpenDate] : null;

    const info = {
      seiki: seiki,
      sharedAccount: sharedTag,
      leaderId: leaderTag,
      openDate: openDate
    };

    if (openDate) openDateMap.set(seiki, openDate);

    nameColIndices.forEach(cIdx => {
      const val = String(row[cIdx] || '').trim();
      if (val) {
        dict.set(val, info);
        const norm = normalizeKey(val);
        if (norm) dict.set(norm, info);
      }
    });
  }

  const result = { dict, openDateMap, normalizeKey };
  getClinicMentionDictionary.cache = result;
  return result;
}

/**
 * 拠点名の配列を受け取り、重複なしのメンション文字列（共用アカウント＋拠点長ID）を生成する
 * ★複数メンションを使う場合は1名ずつ改行せず、横並び（スペース区切り）で出力する
 */
function buildClinicMentionsText(clinicNames, includeToAll = false) {
  const { dict, normalizeKey } = getClinicMentionDictionary();
  const sharedSet = new Set();
  const leaderSet = new Set();

  (clinicNames || []).forEach(rawName => {
    if (!rawName) return;
    const clean = String(rawName).trim();
    const norm = normalizeKey(clean);
    const info = dict.get(clean) || dict.get(norm);
    if (info) {
      if (info.sharedAccount) sharedSet.add(info.sharedAccount);
      if (info.leaderId) leaderSet.add(info.leaderId);
    }
  });

  const lines = [];
  if (includeToAll) lines.push('[toall]');
  if (sharedSet.size > 0) lines.push(Array.from(sharedSet).join(' '));
  if (leaderSet.size > 0) lines.push(Array.from(leaderSet).join(' '));

  return lines.length > 0 ? lines.join('\n') + '\n' : '';
}

/**
 * Chatworkへメッセージを送信する共通関数
 */
function sendChatworkMessage(message) {
  const roomId = CONFIG.getTargetRoomId();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('CHATWORK_API_KEY') || props.getProperty('CHATWORK_API_TOKEN');

  if (!apiKey) {
    throw new Error('スクリプトプロパティに CHATWORK_API_KEY が設定されていません。');
  }

  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': apiKey },
    payload: { body: String(message) },
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error(`Chatwork送信失敗 (Room: ${roomId}, Code: ${code}): ${res.getContentText()}`);
  }
  Logger.log(`✅ Chatwork送信成功 (Room: ${roomId})`);
  return true;
}

/**
 * 送信履歴シート（⚙️_送信履歴）のチェック＆記録関数
 */
function hasSentHistory(category, uniqueKey) {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);
  if (!ss.getSheetByName(CONFIG.SHEET.LOG_HISTORY)) {
    initialSetup();
  }
  const opened = openSheetWithHeaders(CONFIG.SS_ID.SELF, CONFIG.SHEET.LOG_HISTORY, 1);
  const { values, headerMap } = opened;
  const idxCat = requireColumnAny(headerMap, ['種別'], '⚙️_送信履歴');
  const idxKey = requireColumnAny(headerMap, ['ユニークキー'], '⚙️_送信履歴');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxCat]).trim() === category && String(values[i][idxKey]).trim() === uniqueKey) {
      return true;
    }
  }
  return false;
}

function recordSentHistory(category, uniqueKey, doctorName, clinicName, workDateStr) {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);
  let sheet = ss.getSheetByName(CONFIG.SHEET.LOG_HISTORY);
  if (!sheet) {
    initialSetup();
    sheet = ss.getSheetByName(CONFIG.SHEET.LOG_HISTORY);
  }
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([nowStr, category, uniqueKey, doctorName || '', clinicName || '', workDateStr || '', CONFIG.getTargetRoomId()]);
}

/**
 * 日付・時間ユーティリティ
 */
function fastFormatDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

function formatJPDateWithDay(dateObj, includeYear = true) {
  const w = ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
  const m = dateObj.getMonth() + 1;
  const d = dateObj.getDate();
  return includeYear ? `${dateObj.getFullYear()}年${m}月${d}日(${w})` : `${m}月${d}日(${w})`;
}

function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) {
    if (isNaN(dateInput.getTime())) return null;
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }
  let dateStr = String(dateInput).trim();
  const jpMatch = dateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jpMatch) {
    return new Date(parseInt(jpMatch[1], 10), parseInt(jpMatch[2], 10) - 1, parseInt(jpMatch[3], 10));
  }
  dateStr = dateStr.split(/[（(]/)[0].trim().replace(/-/g, '/');
  const parts = dateStr.split('/');
  if (parts.length >= 3) {
    const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, d = parseInt(parts[2], 10);
    if (!isNaN(y) && !isNaN(m) && !isNaN(d)) return new Date(y, m, d);
  }
  const fb = new Date(dateStr);
  return !isNaN(fb.getTime()) ? new Date(fb.getFullYear(), fb.getMonth(), fb.getDate()) : null;
}

function safeParseTime(t) {
  if (t === null || t === undefined || t === '') return NaN;
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  if (typeof t === 'string') {
    const p = t.trim().split(':');
    if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  }
  if (typeof t === 'number' && t >= 0 && t < 1) return Math.round(t * 24 * 60);
  return NaN;
}

function safeFormatTime(min) {
  return `${String(Math.floor(min / 60)).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;
}

/**
 * 開始時刻を30分前倒しに補正するヘルパー（例: "09:00~13:00" -> "08:30-13:00"）
 */
function shiftStartTimeBack30Min(timeRangeStr, shouldShift) {
  if (!timeRangeStr) return '';
  const clean = String(timeRangeStr).trim();
  const match = clean.match(/(\d{1,2}:\d{2})\s*[~〜～\-]\s*(\d{1,2}:\d{2})/);
  if (!match) return clean;

  let startMin = safeParseTime(match[1]);
  const endMin = safeParseTime(match[2]);
  if (isNaN(startMin) || isNaN(endMin)) return clean;

  if (shouldShift) {
    startMin = Math.max(0, startMin - 30);
  }
  return `${safeFormatTime(startMin)}-${safeFormatTime(endMin)}`;
}

/**
 * ★内閣府の公式ページ（https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html）から
 * 最新の祝日CSVリンクを動的に検出して国民の祝日データを取得する関数（Googleカレンダー不使用）
 */
function getOfficialHolidayMap() {
  if (getOfficialHolidayMap.cache) return getOfficialHolidayMap.cache;
  const holidayMap = new Map();
  const GAIYOU_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html';
  const BASE_DIR_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/';
  let csvUrl = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';

  try {
    // 1. 内閣府の gaiyou.html にアクセスし、ページ内の .csv リンクを動的に抽出
    const htmlRes = UrlFetchApp.fetch(GAIYOU_URL, { muteHttpExceptions: true });
    if (htmlRes.getResponseCode() === 200) {
      const htmlText = htmlRes.getContentText('Shift_JIS') || htmlRes.getContentText('UTF-8');
      const match = htmlText.match(/href=["']([^"']*?\.csv)["']/i);
      if (match && match[1]) {
        const href = match[1].trim();
        if (href.startsWith('http')) {
          csvUrl = href;
        } else if (href.startsWith('/')) {
          csvUrl = 'https://www8.cao.go.jp' + href;
        } else {
          csvUrl = BASE_DIR_URL + href.replace(/^\.\//, '');
        }
      }
    }

    // 2. 抽出した公式CSVを取得してパース
    const csvRes = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });
    if (csvRes.getResponseCode() === 200) {
      const lines = csvRes.getContentText('Shift_JIS').split(/\r?\n/);
      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        const parts = line.split(',');
        if (parts.length >= 2) {
          const d = parseDateToSafeDateObj(parts[0]);
          const name = String(parts[1] || '').trim();
          if (d) holidayMap.set(fastFormatDate(d), name);
        }
      }
    }
  } catch (e) {
    Logger.log('内閣府祝日データ取得エラー: ' + e.message);
  }

  getOfficialHolidayMap.cache = holidayMap;
  return holidayMap;
}

function isJapaneseHoliday(dateObj) {
  return getOfficialHolidayMap().has(fastFormatDate(dateObj));
}