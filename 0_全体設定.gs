/**
 * =================================================================
 * 【0_全体設定.gs】
 * 新スクリプト（拠点連絡・自動通知システム）の全設定一元管理ファイル
 * 原則：列のハードコードは一切禁止（すべて見出し名から動的に取得）
 * =================================================================
 */

const CONFIG = {
  // 本番切り替えスイッチ（false = 壁打ち用: 410981446 / true = 本番用: 449648673）
  IS_PRODUCTION: true,
  ROOM_ID_PROD: '449648673',
  ROOM_ID_TEST: '410981446',

  // チャットワークルームID取得ヘルパー
  getTargetRoomId: function() {
    return this.IS_PRODUCTION ? this.ROOM_ID_PROD : this.ROOM_ID_TEST;
  },

  // 関連スプレッドシートID
  SS_ID: {
    SELF: '1jmRdwUEUNhC2iTZyqQC3IqdBIJGZiCdR9RR2An_yvw8',     // この管理用スプシ
    MASTER: '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs',   // 1. 拠点マスタ
    MIJYUSOKU: '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA', // 2. 未充足スプシ
    SHIFT: '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co',     // 3. シフト元スプシ
    SAIYO: '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ'      // 4. 採用・お久しぶりスプシ
  },

  // シート名定義
  SHEET: {
    MASTER_CLINIC: '拠点名',
    PASTE: '貼付用',
    KAKUNIN: '確認用',
    FUZAI: '医師不在拠点',
    AUTO_TEXT: '文章自動作成',
    CLOSED: '休館日',
    IRREGULAR: '変則営業',
    KAKUTEI_SHIFT: '確定シフト',
    SAIYO_REPORT: '〈採用報告〉',
    OHISASHIBURI: 'お久しぶり勤務',
    DIRECT_FIRST: 'DS直接応募初回勤務',
    URL_SHEET: 'URL',
    // 本スプシ内に自動生成する管理用シート
    LOG_HISTORY: '⚙️_送信履歴',
    EVENT_QUEUE: '⚙️_連携キュー',
    TARGET_QUEUE: '⚙️_新規お久しぶり管理'
  },

  // 特殊拠点（小児科・内科を分けて表記する拠点）
  DEPT_LABEL_CLINICS: ['北葛西', '亀有']
};

/**
 * 初回セットアップ用関数
 * 1. Chatwork APIキーの保存（未設定の場合）
 * 2. 管理用シート（⚙️_送信履歴 / ⚙️_連携キュー / ⚙️_新規お久しぶり管理）の自動作成
 * 3. 不要な初期シート「シート1」の自動削除
 */
function initialSetup() {
  const apiKey = '';
  if (apiKey !== '') {
    PropertiesService.getScriptProperties().setProperty('CHATWORK_API_KEY', apiKey);
    Logger.log('Chatwork APIキーをスクリプトプロパティに保存しました。');
  }

  const ss = SpreadsheetApp.openById(CONFIG.SS_ID.SELF);

  // 1. 送信履歴シートの自動生成（48時間で自動掃除）
  let logSheet = ss.getSheetByName(CONFIG.SHEET.LOG_HISTORY);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEET.LOG_HISTORY);
    logSheet.appendRow(['送信日時', '種別', 'ユニークキー', '対象医師', '対象拠点', '勤務日', '送信先RoomID']);
    logSheet.setFrozenRows(1);
    Logger.log('シート「⚙️_送信履歴」を作成しました。');
  }

  // 2. 連携キューシートの自動生成（48時間で自動掃除）
  let queueSheet = ss.getSheetByName(CONFIG.SHEET.EVENT_QUEUE);
  if (!queueSheet) {
    queueSheet = ss.insertSheet(CONFIG.SHEET.EVENT_QUEUE);
    queueSheet.appendRow(['リクエスト日時', 'イベント種別', 'ペイロードJSON', '処理ステータス', '処理完了日時']);
    queueSheet.setFrozenRows(1);
    Logger.log('シート「⚙️_連携キュー」を作成しました。');
  }

  // 3. 新規・お久しぶり永続管理シートの自動生成（削除しない台帳）
  let targetSheet = ss.getSheetByName(CONFIG.SHEET.TARGET_QUEUE);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(CONFIG.SHEET.TARGET_QUEUE);
    targetSheet.appendRow([
      'ユニークキー',
      '種別',
      'タイムスタンプ',
      '医籍番号',
      '医師名',
      '性別',
      '専門科目',
      'フォーム拠点',
      'フォーム診療科',
      'フォーム勤務日',
      'フォーム時間',
      '確定拠点',
      '確定診療科',
      '確定勤務日',
      '確定時間',
      '補正状態',
      '初回通知済',
      '1週間前リマインド済',
      '週次リマインド済',
      '最終同期日時'
    ]);
    targetSheet.setFrozenRows(1);
    Logger.log('シート「⚙️_新規お久しぶり管理」を作成しました。');
  }

  // 4. 不要な「シート1」が残っていれば自動削除
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
    Logger.log('不要な初期シート「シート1」を自動削除しました。');
  }

  Logger.log('初期セットアップが完了しました。');
}

/**
 * スプレッドシートを開いたときのカスタムメニュー
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('拠点連絡システム')
    .addItem('1. 1週間分の未充足・不在拠点を今すぐ投稿', 'runWeeklyAbsencePost')
    .addItem('2. 新規採用（URLシートC2）を今すぐ拠点へ投稿', 'runNewHirePostFromUrlSheet')
    .addItem('3. お久しぶり勤務の未送信分を今すぐチェック＆投稿', 'runOhisashiburiInitialCheck')
    .addSeparator()
    .addItem('新規・お久しぶり管理キューを貼付用と今すぐ同期', 'syncTargetDoctorsQueue')
    .addItem('4. 1週間前リマインド（新規・お久しぶり）を今すぐ実行', 'runOneWeekBeforeReminder')
    .addItem('5. 今週の新規・お久しぶり週次リマインド（月曜用）を今すぐ実行', 'runMondayWeeklyReminder')
    .addItem('6. 明日の新規・お久しぶり前日リマインドを今すぐ実行', 'runTomorrowNewAndOhisaReminder')
    .addItem('7. 拠点間移動発生医師の事前通知を今すぐ実行', 'runClinicTransferPost')
    .addSeparator()
    .addItem('管理用シートの初期作成（初回のみ）', 'initialSetup')
    .addItem('定期トリガーを一括設定する（初回のみ）', 'setupAllTriggers')
    .addToUi();
}