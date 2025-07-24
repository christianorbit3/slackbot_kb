/**
 * ヘッダー行から列のインデックスマップを作成します。
 * @param {Array<string>} headerRow - ヘッダー行の配列
 * @return {Object} 列名をキーとし、インデックスを値とするオブジェクト
 */
function createColumnIndexMap(headerRow) {
  const columnMap = {};
  
  headerRow.forEach((header, index) => {
    const trimmedHeader = String(header).trim();
    switch (trimmedHeader) {
      case HEADER_TASK_TYPE:
        columnMap.taskType = index;
        break;
      case HEADER_DUE_DATE:
        columnMap.dueDate = index;
        break;
      case HEADER_STATUS:
        columnMap.status = index;
        break;
      case HEADER_SUMMARY:
        columnMap.summary = index;
        break;
      case HEADER_SLACK_ID:
        columnMap.slackId = index;
        break;
      case HEADER_ASSIGNEE:
        columnMap.assignee = index;
        break;
      case HEADER_EFFORT:
        columnMap.effort = index;
        break;
      case HEADER_DETAILS:
        columnMap.details = index;
        break;
    }
  });
  
  return columnMap;
}

/**
 * 必須列が存在するかチェックします。
 * @param {Object} columnMap - 列インデックスマップ
 * @return {Array<string>} 不足している必須列名の配列
 */
function validateRequiredColumns(columnMap) {
  const requiredColumns = ['dueDate', 'status', 'summary'];
  const missingColumns = [];
  
  requiredColumns.forEach(col => {
    if (columnMap[col] === undefined) {
      missingColumns.push(col);
    }
  });
  
  return missingColumns;
}

/**
 * 指定スプレッドシートの未完了タスクを Slack に通知します。
 * @param {string} spreadsheetId - タスク管理シートのスプレッドシート ID。
 * @param {string} channelId     - 投稿先チャンネル ID。
 * @param {boolean} [disableMention=false] - trueの場合、SlackIDの代わりにAssigneeを表示（メンションなし）
 */
function projectTaskReport(spreadsheetId, channelId, disableMention = false) {
  try {
    const tasks = getPendingTasks(spreadsheetId);
    if (!tasks || tasks.length === 0) {
      Logger.log('未完了のタスクはありません。spreadsheetId: %s', spreadsheetId);
      return;
    }

    const { tasksTodayOrPast, tasksFuture } = classifyTasksByDueDate(tasks);

    // 各ブロック内で期日昇順にソート
    tasksTodayOrPast.sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime());
    tasksFuture.sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime());
    
    const message = buildTaskSlackMessage(tasksTodayOrPast, tasksFuture, disableMention, spreadsheetId);

    if (message.trim() === '') {
      Logger.log('投稿するタスクメッセージがありません。spreadsheetId: %s', spreadsheetId);
      return;
    }
    
    // slack.js の postToSlack を呼び出す想定
    // この関数がグローバルスコープに存在する必要があります。
    if (typeof postToSlack !== 'function') {
        throw new Error('postToSlack関数が見つかりません。slack.jsがロードされているか確認してください。');
    }
    postToSlack(channelId, message);
    
    const mentionMode = disableMention ? 'メンション無効（Assignee表示）' : 'メンション有効（SlackID使用）';
    Logger.log('タスクリマインダーを Slack に投稿しました。channelId: %s, モード: %s', channelId, mentionMode);

  } catch (e) {
    Logger.log('projectTaskReport でエラーが発生しました: %s\\nStack: %s', e.message, e.stack);
    // 必要に応じて、エラーを呼び出し元に再スローするか、特定のチャンネルにエラー通知を送る
  }
}

/**
 * スプレッドシートから未完了のタスクを取得します。
 * @param {string} spreadsheetId - スプレッドシートのID。
 * @return {Array<Object>|null} 未完了タスクの配列。エラー時やタスクがない場合はnullまたは空配列。
 * Taskオブジェクト: {
 *   taskType: string,
 *   dueDate: Date,
 *   status: string,
 *   summary: string,
 *   slackId: string,
 *   assignee: string,
 *   effort: string,
 *   details: string,
 *   rawRowData: Array<any> // 元の行データ (デバッグ用など)
 * }
 */
function getPendingTasks(spreadsheetId) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log('スプレッドシートが開けませんでした。ID: %s, エラー: %s', spreadsheetId, e.message);
    return null;
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
  if (!sheet) {
    Logger.log('シート "%s" が見つかりませんでした。スプレッドシートID: %s', SHEET_NAME_TASKS, spreadsheetId);
    return null;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) {
    Logger.log('タスクシートにデータがありません（ヘッダーのみ、または空）。シート名: %s', SHEET_NAME_TASKS);
    return [];
  }

  // ヘッダー行から列インデックスマップを作成
  const headerRow = values[0];
  const columnMap = createColumnIndexMap(headerRow);
  
  // 必須列の存在チェック
  const missingColumns = validateRequiredColumns(columnMap);
  if (missingColumns.length > 0) {
    throw new Error(`必須列が見つかりません: ${missingColumns.join(', ')}。スプレッドシートID: ${spreadsheetId}`);
  }

  Logger.log('列マッピング: %s', JSON.stringify(columnMap));

  // util.js に parseDateFromSheetValue が定義されている想定
  if (typeof parseDateFromSheetValue !== 'function') {
      throw new Error('parseDateFromSheetValue関数が見つかりません。util.jsがロードされているか確認してください。');
  }

  const tasks = [];
  // 1行目はヘッダーなので、2行目 (インデックス 1) から処理
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[columnMap.status] ? String(row[columnMap.status]).trim() : '';
    
    if (status !== STATUS_COMPLETED) {
      // 概要列が空白の場合は処理をスキップ
      const summary = row[columnMap.summary] ? String(row[columnMap.summary]).trim() : '';
      if (!summary) {
        //Logger.log('概要が空白のため処理をスキップします。行: %s, スプレッドシートID: %s', i + 1, spreadsheetId);
        continue;
      }

      const dueDateValue = row[columnMap.dueDate];
      const dueDate = parseDateFromSheetValue(dueDateValue); // util.jsの関数を利用

      let finalDueDate;
      if (!dueDate) {
        Logger.log('無効な期日形式です。進行中タスクとして扱います。行: %s, 値: %s, スプレッドシートID: %s', i + 1, dueDateValue, spreadsheetId);
        // 期日が不正な場合は、未来の日付（例：2099年12月31日）を設定して進行中タスクとして扱う
        finalDueDate = new Date(2099, 11, 31); // 2099年12月31日
      } else {
        finalDueDate = dueDate;
      }

      // SlackID が空の場合はメンションしない
      const slackId = columnMap.slackId !== undefined && row[columnMap.slackId] ? 
                      String(row[columnMap.slackId]).trim() : '';
      
      tasks.push({
        taskType: columnMap.taskType !== undefined ? String(row[columnMap.taskType] || '') : '',
        dueDate: finalDueDate,
        status: status,
        summary: summary.slice(0, 100), // 30文字以内
        slackId: slackId,
        assignee: columnMap.assignee !== undefined ? String(row[columnMap.assignee] || '') : '',
        effort: columnMap.effort !== undefined ? String(row[columnMap.effort] || '') : '',
        details: columnMap.details !== undefined ? String(row[columnMap.details] || '') : '',
        rawRowData: row,
        isInvalidDate: !dueDate // 期日が不正だったかどうかのフラグ
      });
    }
  }
  return tasks;
}

/**
 * SlackユーザーIDの形式を検証します。
 * @param {string} slackId - 検証するSlackユーザーID
 * @return {boolean} 有効な形式の場合true
 */
function isValidSlackUserId(slackId) {
  if (!slackId || typeof slackId !== 'string') {
    return false;
  }
  // SlackユーザーIDは通常 U で始まる英数字
  // 例: U1234567890, USLACKBOT など
  return /^U.*$/.test(slackId.trim());
}

/**
 * タスクを期日に基づいて「今日または過去」と「未来」の2つのグループに分類します。
 * @param {Array<Object>} tasks - タスクオブジェクトの配列。
 * @return {{tasksTodayOrPast: Array<Object>, tasksFuture: Array<Object>}} 分類されたタスク。
 */
function classifyTasksByDueDate(tasks) {
  const tasksTodayOrPast = [];
  const tasksFuture = [];

  const today = new Date();
  today.setHours(0, 0, 0, 0); // JSTの今日0時0分0秒

  tasks.forEach(task => {
    const taskDueDate = new Date(task.dueDate);
    taskDueDate.setHours(0, 0, 0, 0); // タスクの期日も0時0分0秒に正規化して比較

    if (taskDueDate.getTime() <= today.getTime()) {
      tasksTodayOrPast.push(task);
    } else {
      tasksFuture.push(task);
    }
  });
  return { tasksTodayOrPast, tasksFuture };
}

/**
 * タスクを担当者別にグループ化します
 * @param {Array<Object>} tasks - タスクの配列
 * @return {Object} 担当者をキーとしたタスクグループ
 */
function groupTasksByAssignee(tasks) {
  const groups = {};
  
  tasks.forEach(task => {
    let assigneeKey;
    
    // SlackIDまたはAssigneeを使用してグループ化
    if (task.slackId && task.slackId.trim()) {
      const trimmedSlackId = task.slackId.trim();
      if (isValidSlackUserId(trimmedSlackId)) {
        assigneeKey = `slack:${trimmedSlackId}`;
      } else {
        assigneeKey = `invalid:${trimmedSlackId}`;
      }
    } else if (task.assignee && task.assignee.trim()) {
      assigneeKey = `assignee:${task.assignee.trim()}`;
    } else {
      assigneeKey = 'unknown:未割り当て';
    }
    
    if (!groups[assigneeKey]) {
      groups[assigneeKey] = [];
    }
    groups[assigneeKey].push(task);
  });
  
  return groups;
}

function buildTaskSlackMessage(tasksTodayOrPast, tasksFuture, disableMention = false, spreadsheetId = '') {
  let message = '';
  const scriptTimeZone = Session.getScriptTimeZone(); // "Asia/Tokyo" が期待される

  // Helper to format a single task line (without mention)
  const formatTaskLine = (task, isIndented = true) => {
    // 期日の表示処理
    let dueDateFormatted;
    if (task.isInvalidDate) {
      dueDateFormatted = '期日不正';
    } else {
      dueDateFormatted = Utilities.formatDate(task.dueDate, scriptTimeZone, 'yyyy/MM/dd');
    }
    
    const indent = isIndented ? '  • ' : '- ';
    return `${indent}${task.summary}  期日: ${dueDateFormatted}`;
  };

  // Helper to format assignee header
  const formatAssigneeHeader = (assigneeKey, disableMention) => {
    const [type, value] = assigneeKey.split(':');
    
    if (disableMention) {
      // メンション無効モード
      switch (type) {
        case 'slack':
        case 'invalid':
          return `👤 ${value}`;
        case 'assignee':
          return `👤 ${value}`;
        case 'unknown':
          return `👤 ${value}`;
        default:
          return `👤 ${value}`;
      }
    } else {
      // 通常モード（メンション有効）
      switch (type) {
        case 'slack':
          return `👤 <@${value}>`;
        case 'invalid':
          Logger.log('無効なSlackID形式です。SlackID: %s', value);
          return `👤 [無効ID: ${value}]`;
        case 'assignee':
          return `👤 ${value}`;
        case 'unknown':
          return `👤 ${value}`;
        default:
          return `👤 ${value}`;
      }
    }
  };

  if (tasksTodayOrPast.length > 0) {
    message += '🚨 ■期日が本日中または過去のタスク\n';
    message += '下記、早急に対応して下さい\n\n';
    
    const urgentGroups = groupTasksByAssignee(tasksTodayOrPast);
    
    // 担当者別にソート（SlackIDを持つ人を優先）
    const sortedAssignees = Object.keys(urgentGroups).sort((a, b) => {
      const aIsSlack = a.startsWith('slack:');
      const bIsSlack = b.startsWith('slack:');
      if (aIsSlack && !bIsSlack) return -1;
      if (!aIsSlack && bIsSlack) return 1;
      return a.localeCompare(b);
    });
    
    sortedAssignees.forEach(assigneeKey => {
      const assigneeTasks = urgentGroups[assigneeKey];
      // 担当者内でも期日昇順にソート
      assigneeTasks.sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime());
      
      message += `${formatAssigneeHeader(assigneeKey, disableMention)}\n`;
      assigneeTasks.forEach(task => {
        message += `${formatTaskLine(task, true)}\n`;
      });
      message += '\n'; // 担当者間のスペース
    });
  }

  if (tasksFuture.length > 0) {
    message += '■案件の進行中タスクのリマインドです\n';
    
    const futureGroups = groupTasksByAssignee(tasksFuture);
    
    // 担当者別にソート（SlackIDを持つ人を優先）
    const sortedAssignees = Object.keys(futureGroups).sort((a, b) => {
      const aIsSlack = a.startsWith('slack:');
      const bIsSlack = b.startsWith('slack:');
      if (aIsSlack && !bIsSlack) return -1;
      if (!aIsSlack && bIsSlack) return 1;
      return a.localeCompare(b);
    });
    
    sortedAssignees.forEach(assigneeKey => {
      const assigneeTasks = futureGroups[assigneeKey];
      // 担当者内でも期日昇順にソート
      assigneeTasks.sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime());
      
      message += `${formatAssigneeHeader(assigneeKey, disableMention)}\n`;
      assigneeTasks.forEach(task => {
        message += `${formatTaskLine(task, true)}\n`;
      });
      message += '\n'; // 担当者間のスペース
    });
  }
  
  // スプレッドシートへのリンクを追加
  if (spreadsheetId) {
    const sheetUrl = generateTasksSheetUrl(spreadsheetId);
    message += `📋 <${sheetUrl}|タスクシートを開く>`;
    message += `📋 <https://www.notion.so/TaskReminder-200b821ba4fa80218d6fd41d37e74624?pvs=4|マニュアルを開く>`;
  }
  
  return message.trim();
}

/**
 * 指定されたスプレッドシートのTasksシートのgidを取得します。
 * @param {string} spreadsheetId - スプレッドシートのID
 * @return {string|null} TasksシートのgidまたはNull（見つからない場合）
 */
function getTasksSheetGid(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
    
    if (!sheet) {
      Logger.log('Tasksシートが見つかりませんでした。スプレッドシートID: %s', spreadsheetId);
      return null;
    }
    
    return sheet.getSheetId().toString();
  } catch (e) {
    Logger.log('Tasksシートのgid取得でエラーが発生しました。スプレッドシートID: %s, エラー: %s', spreadsheetId, e.message);
    return null;
  }
}

/**
 * スプレッドシートのTasksシートへの直接リンクを生成します。
 * @param {string} spreadsheetId - スプレッドシートのID
 * @return {string} TasksシートへのURL
 */
function generateTasksSheetUrl(spreadsheetId) {
  const gid = getTasksSheetGid(spreadsheetId);
  
  if (gid) {
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${gid}`;
  } else {
    // gidが取得できない場合はデフォルトのURL
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
  }
}

/**
 * TaskReminderControllerシートからActiveな案件情報を取得します。
 * @return {Array<Object>} Activeな案件の配列
 * 案件オブジェクト: {
 *   projectName: string,
 *   sheetId: string,
 *   channelId: string,
 *   status: string
 * }
 */
function getActiveProjects() {
  let controllerSpreadsheet;
  try {
    controllerSpreadsheet = SpreadsheetApp.openById(CONTROLLER_SPREADSHEET_ID);
  } catch (e) {
    Logger.log('TaskReminderControllerスプレッドシートが開けませんでした。ID: %s, エラー: %s', CONTROLLER_SPREADSHEET_ID, e.message);
    return [];
  }

  const controllerSheet = controllerSpreadsheet.getSheetByName(CONTROLLER_SHEET_NAME);
  if (!controllerSheet) {
    Logger.log('シート "%s" が見つかりませんでした。スプレッドシートID: %s', CONTROLLER_SHEET_NAME, CONTROLLER_SPREADSHEET_ID);
    return [];
  }

  const dataRange = controllerSheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) {
    Logger.log('TaskReminderControllerシートにデータがありません（ヘッダーのみ、または空）。');
    return [];
  }

  const activeProjects = [];
  // 1行目はヘッダーなので、2行目 (インデックス 1) から処理
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[CONTROLLER_COL_STATUS] ? String(row[CONTROLLER_COL_STATUS]).trim() : '';
    
    if (status === CONTROLLER_STATUS_ACTIVE) {
      const projectName = row[CONTROLLER_COL_PROJECT_NAME] ? String(row[CONTROLLER_COL_PROJECT_NAME]).trim() : '';
      const sheetId = row[CONTROLLER_COL_SHEET_ID] ? String(row[CONTROLLER_COL_SHEET_ID]).trim() : '';
      const channelId = row[CONTROLLER_COL_CHANNEL_ID] ? String(row[CONTROLLER_COL_CHANNEL_ID]).trim() : '';
      
      // 必須項目のチェック
      if (!projectName || !sheetId || !channelId) {
        Logger.log('必須項目が不足しています。行: %s, 案件名: %s, sheetID: %s, channelID: %s', 
                   i + 1, projectName, sheetId, channelId);
        continue;
      }
      
      activeProjects.push({
        projectName: projectName,
        sheetId: sheetId,
        channelId: channelId,
        status: status
      });
    }
  }
  
  return activeProjects;
}




/**
 * TaskReminderControllerシートのActiveな案件に対してタスクリマインダーを実行します。
 * @param {boolean} [disableMention=false] - trueの場合、SlackIDの代わりにAssigneeを表示（メンションなし）
 */
function executeTaskRemindersForActiveProjects(disableMention = false) {
  try {
    Logger.log('=== TaskReminderController実行開始 ===');
    
    const activeProjects = getActiveProjects();
    
    if (activeProjects.length === 0) {
      Logger.log('Activeな案件が見つかりませんでした。');
      return;
    }
    
    Logger.log('Activeな案件数: %s', activeProjects.length);
    
    let successCount = 0;
    let errorCount = 0;
    
    activeProjects.forEach((project, index) => {
      try {
        Logger.log('案件 %s/%s: %s (sheetID: %s, channelID: %s)', 
                   index + 1, activeProjects.length, project.projectName, project.sheetId, project.channelId);
        disableMention = false
        projectTaskReport(project.sheetId, project.channelId, disableMention);
        successCount++;
        
        Logger.log('案件 "%s" のタスクリマインダー実行完了', project.projectName);
        
      } catch (e) {
        errorCount++;
        Logger.log('案件 "%s" でエラーが発生しました: %s', project.projectName, e.message);
      }
    });
    
    Logger.log('=== TaskReminderController実行完了 ===');
    Logger.log('成功: %s件, エラー: %s件', successCount, errorCount);
    
  } catch (e) {
    Logger.log('executeTaskRemindersForActiveProjects でエラーが発生しました: %s\\nStack: %s', e.message, e.stack);
  }
}

// --- テスト用の関数 ---
// この関数は手動実行やテストフレームワークから呼び出すことを想定しています。
// トリガーからは projectTaskReport を直接呼び出すか、
// spreadsheetId と channelId を固定したラッパー関数をトリガー設定します。
function testProjectTaskReport() {
  // これらのIDは実際の環境に合わせてください
  const TEST_SPREADSHEET_ID = '163HYrgm2uCwop2wrQ9FXP_KvCoKU5-BD0SSBcuF59F8'; //AI_Ingage
  const TEST_SLACK_CHANNEL_ID = `C08FNUD2Q1Y`;
  
  Logger.log('テスト開始: projectTaskReport (%s, %s)', TEST_SPREADSHEET_ID, TEST_SLACK_CHANNEL_ID);
  projectTaskReport(TEST_SPREADSHEET_ID, TEST_SLACK_CHANNEL_ID,false);
  Logger.log('テスト終了');
}


/**
 * TaskReminderController実行のテスト関数（メンション無効モード）
 */
function testExecuteTaskRemindersNoMention() {
  Logger.log('=== TaskReminderController実行テスト開始（メンション無効） ===');
  executeTaskRemindersForActiveProjects(true); // メンション無効
  Logger.log('=== TaskReminderController実行テスト終了 ===');
}

/**
 * メンション無効モードのテスト関数
 */
function testProjectTaskReportNoMention() {
  const TEST_SPREADSHEET_ID = '18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw'; //AI_Ingage
  const TEST_SLACK_CHANNEL_ID = `C07RPAX70JD`;
  
  Logger.log('テスト開始: projectTaskReport (メンション無効モード) (%s, %s)', TEST_SPREADSHEET_ID, TEST_SLACK_CHANNEL_ID);
  projectTaskReport(TEST_SPREADSHEET_ID, TEST_SLACK_CHANNEL_ID, true); // disableMention = true
  Logger.log('テスト終了');
}

/**
 * TaskReminderControllerシートの内容を確認するテスト関数
 */
function testTaskReminderController() {
  try {
    Logger.log('=== TaskReminderController テスト開始 ===');
    
    const activeProjects = getActiveProjects();
    
    Logger.log('取得したActiveな案件数: %s', activeProjects.length);
    
    activeProjects.forEach((project, index) => {
      Logger.log('案件 %s: %s', index + 1, JSON.stringify(project));
    });
    
    if (activeProjects.length === 0) {
      Logger.log('Activeな案件がありません。TaskReminderControllerシートの内容を確認してください。');
    }
    
  } catch (e) {
    Logger.log('testTaskReminderController でエラーが発生しました: %s', e.message);
  }
}
