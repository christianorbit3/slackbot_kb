// TaskReminderController関連の定数
const CONTROLLER_SPREADSHEET_ID = '1VzSP6Ab61nlcYKbNU_5iMVtUAr1zn9vuL7WcaKL6oEw';
const CONTROLLER_SHEET_NAME = 'TaskReminderController';
const CONTROLLER_STATUS_ACTIVE = 'Active';

// TaskReminderControllerシートの列インデックス
const CONTROLLER_COL_PROJECT_NAME = 0;  // 案件名 (A列)
const CONTROLLER_COL_SHEET_ID = 1;      // sheetID (B列)
const CONTROLLER_COL_CHANNEL_ID = 2;    // channelID (C列)
const CONTROLLER_COL_STATUS = 3;        // ステータス (D列)

// --- 定数定義 ---
const SHEET_NAME_TASKS = 'Tasks';
const STATUS_COMPLETED = '完了';

// Tasksシートのヘッダー名の定数（列位置は動的に取得）
const HEADER_TASK_TYPE = 'タスク種別';
const HEADER_DUE_DATE = '期日';
const HEADER_STATUS = 'ステータス';
const HEADER_SUMMARY = '概要';
const HEADER_SLACK_ID = 'SlackID';
const HEADER_ASSIGNEE = 'アサイン';
const HEADER_EFFORT = '想定工数';
const HEADER_DETAILS = '詳細';

/**
 * エラーをログに記録
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラーが発生したコンテキスト
 * @param {Object} [additionalInfo] - 追加の情報（オプション）
 */
function logError(error, context, additionalInfo = {}) {
  const timestamp = new Date().toISOString();
  const errorInfo = {
    timestamp,
    context,
    message: error.message,
    stack: error.stack,
    ...additionalInfo
  };
  
  console.error(JSON.stringify(errorInfo, null, 2));
  
  // エラーログをスプレッドシートに記録
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("error_logs");
    
    if (sheet) {
      sheet.appendRow([
        timestamp,
        context,
        error.message,
        error.stack,
        JSON.stringify(additionalInfo)
      ]);
    }
  } catch (logError) {
    console.error("Failed to log error to sheet:", logError);
  }
}

function logMessageToSheet(threadId, ts, userId, text) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("messages");
  
  if (!sheet) {
    throw new Error("messages sheet not found");
  }

  const timestamp = new Date();
  sheet.appendRow([
    threadId,
    ts,
    userId,
    text,
    timestamp
  ]);
}

function getThreadMessageLogs(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("messages");
  
  if (!sheet) {
    throw new Error("messages sheet not found");
  }

  const data = sheet.getDataRange().getValues();
  
  // ヘッダー行をスキップして、該当するthread_tsのメッセージを抽出
  return data.slice(1) // ヘッダー行をスキップ
    .filter(row => row[0] === threadId)
    .map(row => ({
      threadId: row[0],
      ts: row[1],
      userId: row[2],
      text: row[3],
      timestamp: row[4]
    }))
    .sort((a, b) => {
      // タイムスタンプで昇順ソート（古い順）
      const tsA = parseFloat(a.ts.replace("'", ""));
      const tsB = parseFloat(b.ts.replace("'", ""));
      return tsA - tsB;
    });
}

/**
 * Geminiプロンプトをスプレッドシートに記録
 * @param {string} prompt - 送信したプロンプト
 * @param {string} response - Geminiからのレスポンス
 * @param {boolean} outputJson - JSON出力モードかどうか
 * @param {string} functionName - 呼び出し元の関数名
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function logGeminiPromptToSheet(prompt, response, outputJson = false, functionName = "") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("gemini_logs");
  
  if (!sheet) {
    // シートが存在しない場合は作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = spreadsheet.insertSheet("gemini_logs");
    
    // ヘッダー行を追加
    newSheet.getRange(1, 1, 1, 6).setValues([[
      "タイムスタンプ", "関数名", "JSON出力モード", "プロンプト", "レスポンス", "文字数"
    ]]);
  }

  const timestamp = new Date();
  const promptLength = prompt ? prompt.length : 0;
  const responseLength = response ? response.length : 0;
  
  // 新しい行を追加
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("gemini_logs");
  targetSheet.appendRow([
    timestamp,
    functionName,
    outputJson ? "JSON" : "テキスト",
    prompt,
    response,
    `プロンプト: ${promptLength}文字, レスポンス: ${responseLength}文字`
  ]);
}

/**
 * 商談JSONをスプレッドシートに保存
 * @param {string} threadId - スレッドID
 * @param {Object} json - 商談JSON
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveTaskJson(threadId, json) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("taskJson");
  
  if (!sheet) {
    throw new Error("tasks sheet not found");
  }

  const timestamp = new Date();
  const existingRow = findTaskRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(JSON.stringify(json));
    sheet.getRange(existingRow, 4).setValue(timestamp);
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId,
      JSON.stringify(json),
      timestamp,
      timestamp
    ]);
  }
}

/**
 * 商談JSONをスプレッドシートから取得
 * @param {string} threadId - スレッドID
 * @return {Object|null} 商談JSON（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getTaskJson(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("taskJson");
  
  if (!sheet) {
    throw new Error("tasks sheet not found");
  }

  const row = findTaskRow(threadId);

  if (!row) {
    return null;
  }

  const jsonStr = sheet.getRange(row, 2).getValue();

  try {
    json = JSON.parse(jsonStr);
    return json;

  } catch (error) {
 
    throw new Error(`Invalid JSON format in tasks sheet: ${error.message}`);
  }
}

/**
 * 商談JSONの行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findTaskRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("taskJson");
  
  if (!sheet) {
    throw new Error("tasks sheet not found");
  }

  const data = sheet.getDataRange().getValues();

  // 後ろから検索すれば、最初に見つかったものが最新
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === threadId) {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}

/**
 * 設定値を取得
 * @param {string} key - 設定キー
 * @return {string|null} 設定値（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getConfig(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("config");
  
  if (!sheet) {
    throw new Error("config sheet not found");
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}


/**
 * 商談登録の確認状態を保存
 * @param {string} threadId - スレッドID
 * @param {Object} confirmation - 確認状態
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveTaskConfirmation(threadId, confirmation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_confirmations");
  
  if (!sheet) {
    // シートが存在しない場合は作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = spreadsheet.insertSheet("task_confirmations");
    
    // ヘッダー行を追加
    newSheet.getRange(1, 1, 1, 5).setValues([[
      "threadId", "status", "json"
    ]]);
    
    // ヘッダー行のスタイル設定
    newSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    newSheet.getRange(1, 1, 1, 5).setBackground("#f0f0f0");
  }

  const existingRow = findTaskConfirmationRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(confirmation.status);
    sheet.getRange(existingRow, 3).setValue(JSON.stringify(confirmation.json));
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId, // 明示的に文字列に変換
      confirmation.status,
      JSON.stringify(confirmation.json)
    ]);
  }
}

/**
 * タスク登録の確認状態を取得
 * @param {string} threadId - スレッドID
 * @return {Object|null} 確認状態（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getPendingTaskConfirmation(threadId) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_confirmations");
  if (!sheet) {
    return null;
  }
  const row = findTaskConfirmationRow(threadId);
  if (!row) {
    return null;
  }
  const status = sheet.getRange(row, 2).getValue();
  const jsonStr = sheet.getRange(row, 3).getValue();

  try {
    const json = JSON.parse(jsonStr);
    return {
      status: status,
      json: json,
    };
  } catch (error) {
    throw new Error(`Invalid JSON format in task confirmations sheet: ${error.message}`);
  }
}

/**
 * 商談登録の確認状態の行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findTaskConfirmationRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_confirmations");
  
  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const storedThreadId = data[i][0];

    if (storedThreadId == threadId && data[i][1]=="pending") {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}

/**
 * スレッドの種類を保存
 * @param {string} threadId - スレッドID
 * @param {string} processType - プロセスの種類（'create', 'soql', 'update', 'communication'）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveThreadProcessType(threadId, processType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("thread_processes");

  const existingRow = findThreadProcessRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(processType);
    sheet.getRange(existingRow, 3).setValue(new Date());
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId,
      processType,
      new Date()
    ]);
  }
}

/**
 * スレッドの種類を取得
 * @param {string} threadId - スレッドID
 * @return {string|null} プロセスの種類（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getThreadProcessType(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("thread_processes");
  
  if (!sheet) {
    return null;
  }

  const row = findThreadProcessRow(threadId);
  if (!row) {
    return null;
  }

  return sheet.getRange(row, 2).getValue();
}

/**
 * スレッドのプロセス行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findThreadProcessRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("thread_processes");
  
  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(threadId)) {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}

/**
 * スレッドのプロセス種類をリセット
 * @param {string} threadId - スレッドID
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function resetThreadProcessType(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("thread_processes");
  
  if (!sheet) {
    return;
  }

  const row = findThreadProcessRow(threadId);
  if (row) {
    // 既存の行を削除
    sheet.deleteRow(row);
  }
}

function getCsvAsBase64(fileId, gid) {
  // URL から fileId と gid を抽出
  //const fileId = (csvUrl.match(/\/d\/([-\w]+)/) || [])[1];
  //const gid    = (csvUrl.match(/gid=([0-9]+)/) || [])[1] || '0';
  if (!fileId) throw new Error('fileId 抽出失敗: ' + csvUrl);

  // Drive API ではなく Spreadsheet の export エンドポイントを直叩き
  const exportUrl =
    `https://docs.google.com/spreadsheets/d/${fileId}/export?format=csv&gid=${gid}`;

  const res = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`CSV エクスポート失敗 (${res.getResponseCode()}) → ${exportUrl}\n`
                    + res.getContentText());
  }

  return Utilities.base64Encode(res.getBlob().getBytes());
}

// --- Task Reminder Related Functions ---
// These constants are also defined in taskReminder.js, ensure they are consistent or managed globally.
// const SHEET_NAME_TASKS_FOR_REMINDER = 'Tasks'; // ... (これらの定数定義を削除)
// const STATUS_COMPLETED_FOR_REMINDER = '完了';
// ... (他の関連定数も削除)

/**
 * スプレッドシートから未完了のタスクを取得します。
 * @param {string} spreadsheetId - スプレッドシートのID。
 * @return {Array<Object>|null} 未完了タスクの配列。エラー時やタスクがない場合はnullまたは空配列。
 * Taskオブジェクト: { ... }
 */
// function getPendingTasks(spreadsheetId) { ... } // この関数全体を削除

/**
 * (テスト用)指定されたスプレッドシートにサンプルタスクデータを生成します。
 * シート 'Tasks' が存在する必要があります。
 * @param {string} spreadsheetId - データを生成するスプレッドシートのID。
 */
function generateSampleTaskData(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(SHEET_NAME_TASKS_FOR_REMINDER); 

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_TASKS_FOR_REMINDER);
    Logger.log('シート「%s」を作成しました。', SHEET_NAME_TASKS_FOR_REMINDER);
  } else {
    sheet.clearContents(); // 既存のデータをクリア
    Logger.log('シート「%s」の既存データをクリアしました。', SHEET_NAME_TASKS_FOR_REMINDER);
  }

  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const dayAfterTomorrow = new Date(today);
  dayAfterTomorrow.setDate(today.getDate() + 2);

  const sampleData = [
    ['タスク種別', '期日', 'ステータス', '概要', 'SlackID', 'アサイン', '想定工数', '詳細'],
    ['開発', Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '未着手', '過去タスク1', 'U01234567', '山田太郎', '2h', 'これは過去のタスクです'],
    ['設計', Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '着手中', '今日タスク1', 'U98765432', '鈴木花子', '3h', 'これは今日のタスクです'],
    ['開発', Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '未着手', '今日タスク2', 'UABCDEFGH', '佐藤一郎', '1d', 'これも今日のタスク'],
    ['テスト', Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '未着手', '未来タスク1', 'UXYZ12345', '田中次郎', '4h', 'これは未来のタスクです'],
    ['資料作成', Utilities.formatDate(dayAfterTomorrow, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '完了', '完了タスク1', 'UFOOBARBAZ', '伊藤三郎', '1h', 'これは完了済みのタスクなので通知されないはず'],
    ['設計', Utilities.formatDate(dayAfterTomorrow, Session.getScriptTimeZone(), 'yyyy/MM/dd'), '未着手', '未来タスク2', 'UFOOBARBAZ', '伊藤三郎', '5h', 'これも未来のタスク'],
    ['その他', '2024/01/01', '未着手', 'かなり過去のタスク', '', '匿名希望', 'N/A', 'Slack IDなし'], // Slack IDなし
    ['開発', 45321, '未着手', 'シリアル値期日タスク', 'U SERIAL ', '田中', '1d', 'Excelシリアル値の期日'], // Excelシリアル値 (2024/01/31)
  ];

  sheet.getRange(1, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  Logger.log('サンプルデータをシート「%s」に書き込みました。', SHEET_NAME_TASKS_FOR_REMINDER);
  SpreadsheetApp.getUi().alert('サンプルデータをシート「' + SHEET_NAME_TASKS_FOR_REMINDER + '」に書き込みました。');
}

/**
 * 最初の空白行を探す
 * @param {Sheet} sheet - スプレッドシートのシート
 * @param {Object} columnMap - 列インデックスマップ
 * @param {number} startRow - 検索開始行（デフォルト: 2）
 * @return {number} 空白行の行番号（見つからない場合は最後の行+1）
 */
function findFirstEmptyRow(sheet, columnMap, startRow = 2) {
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(startRow, columnMap.summary + 1, lastRow - startRow + 1, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) { // 概要列が空の場合
      return startRow + i;
    }
  }
  
  return lastRow + 1; // 空白行が見つからない場合は最後の行+1を返す
}

/**
 * タスクをスプレッドシートに登録
 * @param {Object} taskJson - タスク情報
 * @return {Object} 登録結果
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function createTaskToSheet(taskJson) {
  const sheet = SpreadsheetApp.openById(taskJson.SheetId).getSheetByName("Tasks");
  
  if (!sheet) {
    throw new Error("Tasks sheet not found");
  }

  // ヘッダー行を取得して列インデックスマップを作成
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnMap = createColumnIndexMap(headerRow);
  
  // 必要な列が存在するか確認
  const requiredColumns = {
    summary: HEADER_SUMMARY,
    dueDate: HEADER_DUE_DATE,
    status: HEADER_STATUS,
    assignee: HEADER_ASSIGNEE
  };
  
  const missingColumns = Object.entries(requiredColumns)
    .filter(([key]) => columnMap[key] === undefined)
    .map(([_, header]) => header);
  
  if (missingColumns.length > 0) {
    throw new Error(`必要な列が見つかりません: ${missingColumns.join(", ")}`);
  }

  // 最初の空白行を探す
  const targetRow = findFirstEmptyRow(sheet, columnMap);
  
  // 各列に値を設定
  const values = {
    summary: taskJson.概要,
    dueDate: taskJson.期日,
    status: "", // デフォルトのステータス
    assignee: taskJson.アサイン
  };

  // 各列に値を設定
  Object.entries(values).forEach(([key, value]) => {
    const columnIndex = columnMap[key] + 1; // 1-based index
    sheet.getRange(targetRow, columnIndex).setValue(value);
  });

  return {
    success: true,
    message: "タスクを登録しました",
    row: targetRow
  };
}

/**
 * アクティブなタスクシートのIDをTaskReminderControllerシートから取得します。
 * @return {Array<string>} アクティブなタスクシートのIDの配列。
 */
function getActiveTaskSheetIdsFromController() {

  let controllerSheet;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // まずアクティブなスプレッドシートから試みる
    controllerSheet = ss.getSheetByName(CONTROLLER_SHEET_NAME);

    if (!controllerSheet && CONTROLLER_SPREADSHEET_ID) {
      // 見つからなければIDで開く
      const controllerSs = SpreadsheetApp.openById(CONTROLLER_SPREADSHEET_ID);
      controllerSheet = controllerSs.getSheetByName(CONTROLLER_SHEET_NAME);
    }

    if (!controllerSheet) {
      throw new Error(`Sheet "${CONTROLLER_SHEET_NAME}" not found.`);
    }
  } catch (e) {
    logError(e, "getActiveTaskSheetIdsFromController: Failed to access TaskReminderController sheet");
    return [];
  }

  const data = controllerSheet.getDataRange().getValues();
  const activeSheetIds = [];
  // ヘッダー行をスキップ (インデックス1から開始)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // ステータスがActiveで、SheetIDが存在する場合
    if (row[CONTROLLER_COL_STATUS] === CONTROLLER_STATUS_ACTIVE && row[CONTROLLER_COL_SHEET_ID]) {
      activeSheetIds.push(String(row[CONTROLLER_COL_SHEET_ID]).trim());
    }
  }
  return activeSheetIds;
}

/**
 * 指定されたSlackUserIDの未完了タスクを特定のシートから取得します。
 * taskReminder.jsのgetPendingTasksを参考にし、ユーザーフィルターを追加しています。
 * @param {string} slackUserId - SlackユーザーID。
 * @param {string} sheetId - スプレッドシートID。
 * @return {Array<Object>} タスク情報の配列。
 */
function getPendingTasksForUserFromSheet(slackUserId, sheetId) {
  const SHEET_NAME_TASKS = 'Tasks'; // taskReminder.jsの定数
  const STATUS_COMPLETED = '完了';   // taskReminder.jsの定数

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    logError(e, `getPendingTasksForUserFromSheet: Failed to open spreadsheet ID: ${sheetId}`);
    return [];
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
  if (!sheet) {
    Logger.log(`getPendingTasksForUserFromSheet: Sheet "${SHEET_NAME_TASKS}" not found in spreadsheet ID: ${sheetId}`);
    return [];
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) return []; // データなし、またはヘッダーのみ

  const headerRow = values[0];
  const columnMap = createColumnIndexMap(headerRow);

  // validateRequiredColumns は taskReminder.js から利用可能と仮定
  const missingRequired = validateRequiredColumns(columnMap);
  if (missingRequired.length > 0) {
    Logger.log(`getPendingTasksForUserFromSheet: Missing required columns (${missingRequired.join(', ')}) in sheet ID ${sheetId}`);
    return [];
  }

  if (columnMap.slackId === undefined) {
    Logger.log(`getPendingTasksForUserFromSheet: SlackID column (HEADER_SLACK_ID) not found in sheet ID ${sheetId} using map: ${JSON.stringify(columnMap)}`);
    return [];
  }
  if (columnMap.status === undefined) {
    Logger.log(`getPendingTasksForUserFromSheet: Status column (HEADER_STATUS) not found in sheet ID ${sheetId}`);
    return [];
  }
   if (columnMap.summary === undefined) {
    Logger.log(`getPendingTasksForUserFromSheet: Summary column (HEADER_SUMMARY) not found in sheet ID ${sheetId}`);
    return [];
  }
  if (columnMap.dueDate === undefined) {
    Logger.log(`getPendingTasksForUserFromSheet: DueDate column (HEADER_DUE_DATE) not found in sheet ID ${sheetId}`);
    return [];
  }

  const tasks = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const taskStatus = row[columnMap.status] ? String(row[columnMap.status]).trim() : '';
    const taskSlackId = row[columnMap.slackId] ? String(row[columnMap.slackId]).trim() : '';

    if (taskSlackId === slackUserId && taskStatus !== STATUS_COMPLETED) {
      const summary = row[columnMap.summary] ? String(row[columnMap.summary]).trim() : '概要なし';
      if (!summary || summary === '概要なし') {
        // Logger.log(`Skipping task with empty summary at row ${i+1} in sheet ID ${sheetId}`);
        continue;
      }
      
      const dueDateValue = row[columnMap.dueDate];
      let dueDateObj;
      // Try to parse the date value
      if (dueDateValue instanceof Date && !isNaN(dueDateValue)) {
        dueDateObj = dueDateValue;
      } else if (typeof dueDateValue === 'string' || typeof dueDateValue === 'number') {
        dueDateObj = new Date(dueDateValue);
        if (isNaN(dueDateObj.getTime())) { // Check if date is invalid
          dueDateObj = null; 
        }
      } else {
        dueDateObj = null;
      }

      let formattedDueDate = '期日不明';
      if (dueDateObj) {
        formattedDueDate = Utilities.formatDate(dueDateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      } else {
         // taskReminder.jsでは未来の日付を設定しているが、ここでは「期日不正」として扱う
        formattedDueDate = '期日不正';
      }
      
      tasks.push({
        summary: summary,
        dueDate: formattedDueDate,
        status: taskStatus,
        taskType: columnMap.taskType !== undefined && row[columnMap.taskType] ? String(row[columnMap.taskType]) : '',
        assignee: columnMap.assignee !== undefined && row[columnMap.assignee] ? String(row[columnMap.assignee]) : '',
        effort: columnMap.effort !== undefined && row[columnMap.effort] ? String(row[columnMap.effort]) : '',
        details: columnMap.details !== undefined && row[columnMap.details] ? String(row[columnMap.details]) : '',
        sheetId: sheetId
      });
    }
  }
  return tasks;
}

/**
 * タスク完了の確認状態を保存
 * @param {string} threadId - スレッドID
 * @param {Object} confirmation - 確認状態
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveTaskCompleteConfirmation(threadId, confirmation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_complete_confirmations");
  
  if (!sheet) {
    // シートが存在しない場合は作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = spreadsheet.insertSheet("task_complete_confirmations");
    
    // ヘッダー行を追加
    newSheet.getRange(1, 1, 1, 5).setValues([[
      "threadId", "status", "json"
    ]]);
    
    // ヘッダー行のスタイル設定
    newSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    newSheet.getRange(1, 1, 1, 5).setBackground("#f0f0f0");
  }

  const existingRow = findTaskCompleteConfirmationRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(confirmation.status);
    sheet.getRange(existingRow, 3).setValue(JSON.stringify(confirmation.json));
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId, // 明示的に文字列に変換
      confirmation.status,
      JSON.stringify(confirmation.json)
    ]);
  }
}

/**
 * タスク完了の確認状態を取得
 * @param {string} threadId - スレッドID
 * @return {Object|null} 確認状態（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getPendingTaskCompleteConfirmation(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_complete_confirmations");
  if (!sheet) {
    return null;
  }
  const row = findTaskCompleteConfirmationRow(threadId);
  if (!row) {
    return null;
  }
  const status = sheet.getRange(row, 2).getValue();
  const jsonStr = sheet.getRange(row, 3).getValue();

  try {
    const json = JSON.parse(jsonStr);
    return {
      status: status,
      json: json,
    };
  } catch (error) {
    throw new Error(`Invalid JSON format in task complete confirmations sheet: ${error.message}`);
  }
}

/**
 * タスク完了の確認状態の行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findTaskCompleteConfirmationRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("task_complete_confirmations");
  
  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const storedThreadId = data[i][0];

    if (storedThreadId == threadId && data[i][1]=="pending") {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}

/**
 * チャンネルIDに紐づくタスクシートIDを取得
 * @param {string} channelId - SlackチャンネルID
 * @return {string|null} タスクシートID（見つからない場合はnull）
 */
function getSheetIdFromChannelId(channelId) {
  try {
    const controllerSheet = SpreadsheetApp.openById(CONTROLLER_SPREADSHEET_ID)
      .getSheetByName(CONTROLLER_SHEET_NAME);
    
    if (!controllerSheet) {
      return null;
    }

    const data = controllerSheet.getDataRange().getValues();
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetId = String(row[CONTROLLER_COL_SHEET_ID]).trim();
      const status = String(row[CONTROLLER_COL_STATUS]).trim();
      const channelIdFromController = String(row[CONTROLLER_COL_CHANNEL_ID]).trim();
      
      if (status === CONTROLLER_STATUS_ACTIVE && channelIdFromController === channelId) {
        return sheetId;
      }
    }
    return null;
  } catch (error) {
    Logger.log('タスクシートの検索中にエラーが発生しました: ' + error.message);
    return null;
  }
}

/**
 * タスクを完了状態に更新
 * @param {string} sheetId - スプレッドシートID
 * @param {string} taskSummary - タスク概要
 * @return {boolean} 更新が成功したかどうか
 */
function completeTaskInSheet(sheetId, taskSummary) {
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
    
    if (!sheet) {
      throw new Error("タスクシートが見つかりませんでした。");
    }

    // ヘッダー行を取得して列インデックスマップを作成
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMap = createColumnIndexMap(headerRow);

    // データを取得
    const data = sheet.getDataRange().getValues();
    let foundTask = false;

    // 検索文字列をトリミング
    const normalizedTaskSummary = taskSummary
      .replace(/[「」:：]/g, '') // 特殊文字を削除
      .trim(); // 前後の空白を削除

    // 2行目から検索（1行目はヘッダー）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const summary = String(row[columnMap.summary] || '')
        .replace(/[「」:：]/g, '') // 特殊文字を削除
        .trim(); // 前後の空白を削除
      const status = String(row[columnMap.status] || '').trim();

      // 概要が一致し、かつ未完了のタスクを探す
      if (summary === normalizedTaskSummary && status !== STATUS_COMPLETED) {
        // ステータスを「完了」に更新
        sheet.getRange(i + 1, columnMap.status + 1).setValue(STATUS_COMPLETED);
        foundTask = true;
        break;
      }
    }

    return foundTask;
  } catch (error) {
    Logger.log('タスクの完了処理中にエラーが発生しました: ' + error.message);
    return false;
  }
}

/**
 * カレンダーイベントJSONをスプレッドシートに保存
 * @param {string} threadId - スレッドID
 * @param {Object} json - イベントJSON
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveEventJson(threadId, json) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("eventJson");
  
  if (!sheet) {
    // シートが存在しない場合は作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = spreadsheet.insertSheet("eventJson");
    
    // ヘッダー行を追加
    newSheet.getRange(1, 1, 1, 4).setValues([[
      "threadId", "json", "created", "updated"
    ]]);
  }

  const timestamp = new Date();
  const existingRow = findEventRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(JSON.stringify(json));
    sheet.getRange(existingRow, 4).setValue(timestamp);
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId,
      JSON.stringify(json),
      timestamp,
      timestamp
    ]);
  }
}

/**
 * カレンダーイベントJSONをスプレッドシートから取得
 * @param {string} threadId - スレッドID
 * @return {Object|null} イベントJSON（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getEventJson(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("eventJson");
  
  if (!sheet) {
    return null;
  }

  const row = findEventRow(threadId);

  if (!row) {
    return null;
  }

  const jsonStr = sheet.getRange(row, 2).getValue();

  try {
    return JSON.parse(jsonStr);
  } catch (error) {
    throw new Error(`Invalid JSON format in eventJson sheet: ${error.message}`);
  }
}

/**
 * カレンダーイベントJSONの行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findEventRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("eventJson");
  
  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  // 後ろから検索すれば、最初に見つかったものが最新
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === threadId) {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}

/**
 * カレンダーイベント登録の確認状態を保存
 * @param {string} threadId - スレッドID
 * @param {Object} confirmation - 確認状態
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function saveEventConfirmation(threadId, confirmation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("event_confirmations");
  
  if (!sheet) {
    // シートが存在しない場合は作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = spreadsheet.insertSheet("event_confirmations");
    
    // ヘッダー行を追加
    newSheet.getRange(1, 1, 1, 3).setValues([[
      "threadId", "status", "json"
    ]]);
    
    // ヘッダー行のスタイル設定
    newSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    newSheet.getRange(1, 1, 1, 3).setBackground("#f0f0f0");
  }

  const existingRow = findEventConfirmationRow(threadId);
  
  if (existingRow) {
    // 既存の行を更新
    sheet.getRange(existingRow, 2).setValue(confirmation.status);
    sheet.getRange(existingRow, 3).setValue(JSON.stringify(confirmation.json));
  } else {
    // 新しい行を追加
    sheet.appendRow([
      "'" + threadId, // 明示的に文字列に変換
      confirmation.status,
      JSON.stringify(confirmation.json)
    ]);
  }
}

/**
 * カレンダーイベント登録の確認状態を取得
 * @param {string} threadId - スレッドID
 * @return {Object|null} 確認状態（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function getPendingEventConfirmation(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("event_confirmations");
  if (!sheet) {
    return null;
  }
  const row = findEventConfirmationRow(threadId);
  if (!row) {
    return null;
  }
  const status = sheet.getRange(row, 2).getValue();
  const jsonStr = sheet.getRange(row, 3).getValue();

  try {
    const json = JSON.parse(jsonStr);
    return {
      status: status,
      json: json,
    };
  } catch (error) {
    throw new Error(`Invalid JSON format in event confirmations sheet: ${error.message}`);
  }
}

/**
 * カレンダーイベント登録の確認状態の行を検索
 * @param {string} threadId - スレッドID
 * @return {number|null} 行番号（見つからない場合はnull）
 * @throws {Error} スプレッドシート操作に失敗した場合
 */
function findEventConfirmationRow(threadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("event_confirmations");
  
  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const storedThreadId = data[i][0];

    if (storedThreadId == threadId && data[i][1]=="pending") {
      return i + 1; // 行番号なので+1
    }
  }
  
  return null;
}
