/**
 * メインエントリーポイントモジュール
 * 
 * このモジュールは、SlackからのWebhookリクエストを受け取り、
 * 以下の処理を実行します：
 * 1. Slackイベントの検証と処理
 * 2. タスク情報の抽出と検証
 * 3. Salesforceへのタスク登録
 * 4. Notionへのタスク情報の記録
 */

/**
 * SlackからのWebhookリクエストを処理
 * @param {Object} e - リクエストオブジェクト
 * @return {Object} レスポンス
 * @throws {Error} リクエスト処理に失敗した場合
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    // URL Verification
    if (data.type === 'url_verification') {
      // challenge 値を "そのまま" plain-text で返す
      return ContentService
        .createTextOutput(data.challenge)
        .setMimeType(ContentService.MimeType.TEXT);
    }

    // ここから3秒ルール対策
    // イベントID取得（リトライ検出用）
    const eventId = data.event_id;
    if (!eventId || !data.event) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false,
        error: "No event data"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3秒ルール対策（リトライ検出）
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventLog");
    const eventIds = sheet.getRange("A:A").getValues();
    if (eventIds.some(row => row[0] === eventId)) {
      // 2回目以降は即時応答
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        message: "Event already processed"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 初回はIDを記録
    sheet.appendRow([eventId, new Date(), "processing"]);
    // ここまで3秒ルール対策

    // 署名の検証
    //if (!verifySlackSignature(e)) {
      //throw new Error("Invalid Slack signature");
    //}
    
    // イベントの処理（非同期）
    processSlackMessage(data.event);
    
    // 即時レスポンス
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      message: "Event received"
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    logError(error, "Webhookリクエスト処理中にエラーが発生");
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Slackメッセージの処理
 * @param {Object} event - Slackイベント
 */
function processSlackMessage(event) {
  const threadId = event.thread_ts || event.ts;
  const pastMessages = getThreadMessageLogs(threadId);
  
  // スレッドのプロセス種類を取得
  const currentProcessType = getThreadProcessType(threadId);
  
  if (currentProcessType) {
    // 既存のスレッドでプロセス種類が確定している場合
    switch (currentProcessType) {
      case 'getTasks':
        getTasksSlackProcess(event);
        break;
      case 'completeTask':
        completeTaskSlackProcess(event);
        break;
      case 'createTask':
        createTaskSlackProcess(event);
        break;
      case 'communication':
        communicationSlackProcess(event);
        break;
      case 'getCalendar':
        getCalendarAvailabilitySlackProcess(event);
        break;
      case 'createEvent':
        createEventSlackProcess(event);
        break;
    }
  } else {
    // 新しいスレッドまたはプロセス種類が未確定の場合
    const processResult = determineProcessType(event.text, pastMessages);
    
    if (processResult.processType && processResult.confidence > 0.7) {
      // 高い確信度でプロセス種類を判定できた場合
      saveThreadProcessType(threadId, processResult.processType);
      
      // 判定結果を通知
      postMessage(event.channel, `このスレッドは「${getProcessTypeName(processResult.processType)}」として処理します。`, null, threadId);
      
      // 対応するプロセスを実行
      switch (processResult.processType) {
        case 'getTasks':
          getTasksSlackProcess(event);
          break;
        case 'completeTask':
          completeTaskSlackProcess(event);
          break;
        case 'createTask':
          createTaskSlackProcess(event);
          break;
        case 'communication':
          communicationSlackProcess(event);
          break;
        case 'getCalendar':
          getCalendarAvailabilitySlackProcess(event);
          break;
        case 'createEvent':
          createEventSlackProcess(event);
          break;
      }
    } else {
      // プロセス種類を判定できない場合、ユーザーに選択を促す
      const blocks = [
        {
          type: "divider"
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "このスレッドで行いたいことはなんですか？？"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "*以下のいずれかの操作を実行できます：*"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "📝 *持っているタスクの一覧の確認*\n`mytask` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "✏️ *タスクを完了に*\n`done タスク概要` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "✏️ *タスクの新規作成*\n`create タスク概要 期日` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "📅 *カレンダーの空き枠確認*\n`calendar メールアドレス1 メールアドレス2 ...` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "🗓️ *カレンダー予約の作成*\n`event 予定名 開始時間 時間 招待アドレス` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "💬 *通常の会話*\n`会話` などと入力してください"
          }
        },
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "または、具体的に何をしたいかをお知らせください。"
          }
        },
        {
          type: "divider"
        }
      ];
      
      postMessage(event.channel, "処理の種類を選択してください", blocks, threadId);
    }
  }
}

/**
 * プロセス種類の表示名を取得
 * @param {string} processType - プロセス種類
 * @return {string} 表示名
 */
function getProcessTypeName(processType) {
  const names = {
    'getTasks': 'タスク一覧の取得',
    'completeTask': 'タスクの完了',
    'createTask': 'タスクの作成',
    'communication': '通常の会話',
    'getCalendar': 'カレンダーの空き枠確認',
    'createEvent': 'カレンダー予約の作成'
  };
  return names[processType] || processType;
}

/**
 * メッセージイベントを処理
 * @param {Object} event - メッセージイベント
 * @throws {Error} メッセージ処理に失敗した場合
 */
function createTaskSlackProcess(event) {
  // メッセージをログに記録
  logMessageToSheet(
    "'" + (event.thread_ts || event.ts),
    "'" + event.ts,
    event.user,
    event.text
  );

  // そのthread_tsに紐づく過去のメッセージログを取得
  const threadId = event.thread_ts || event.ts;
  var pastMessages = getThreadMessageLogs(threadId);
  
  // タスク登録の確認状態をチェック
  const confirmation = getPendingTaskConfirmation(threadId);

  if (confirmation && confirmation.status === "pending") {
    // 肯定の返事かどうかを判定
    const isPositive = isPositiveResponse(event.text, pastMessages);
    
    if (isPositive) {
      // タスクを登録
      try {
        const result = createTaskToSheet(confirmation.json);
        
        // タスク情報を整形して表示
        const message = `タスクを登録しました`;
        
        postMessage(event.channel, message, null, threadId);
        
        // 確認状態を更新
        saveTaskConfirmation(threadId, {
          ...confirmation,
          status: "completed"
        });

        // スレッドのプロセス種類をリセット
        resetThreadProcessType(threadId);
        
        // 次のアクションを促すメッセージを送信
        const blocks = [
          {
            type: "divider"
          },
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: "🎉 *タスクの登録が完了しました！*"
            }
          },
          {
            type: "divider"
          }
        ];
        
        postMessage(event.channel, "ご希望がございましたら、次のアクションを指示して下さい", blocks, threadId);
        
        return; // タスク登録が完了したら処理を終了
      } catch (error) {
        postMessage(event.channel, `タスクの登録中にエラーが発生しました: ${error.message}`, null, threadId);
        logError(error, "タスク登録中にエラーが発生", { json: confirmation.json });
        
        // 確認状態を更新
        saveTaskConfirmation(threadId, {
          ...confirmation,
          status: "error"
        });
        
        return; // エラーが発生したら処理を終了
      }
    } else {
      // 否定の返事の場合、確認状態をリセット
      saveTaskConfirmation(threadId, {
        ...confirmation,
        status: "cancelled"
      });
      
      postMessage(event.channel, "タスクの登録をキャンセルしました。", null, threadId);
      return; // キャンセルしたら処理を終了
    }
  }

  // チャンネルIDに紐づくタスクシートを検索
  const targetSheetId = getSheetIdFromChannelId(event.channel);
  if (!targetSheetId) {
    postMessage(event.channel, "このチャンネルに紐づくタスクシートが見つかりませんでした。", null, threadId);
    return;
  }

  // 既存のタスクJSONを取得
  const existingJson = getTaskJson(threadId);

  // もう一度更新
  pastMessages = getThreadMessageLogs(threadId);
  
  // 現在の日付を取得
  const currentDate = new Date();
  
  // タスク情報を抽出
  const json = callGeminiForJson(event.text, existingJson, pastMessages, currentDate);

  // シートIDを設定
  json.SheetId = targetSheetId;

  // タスク情報を保存
  saveTaskJson(threadId, json);

  // 不足項目をチェック
  checkMissingFields(event.channel, threadId, json);
}

/**
 * 通常の会話を処理する関数
 * @param {Object} event - Slackイベント
 */
function communicationSlackProcess(event) {
  // メッセージをログに記録
  logMessageToSheet(
    "'" + (event.thread_ts || event.ts),
    "'" + event.ts,
    event.user,
    event.text
  );
  const threadId = event.thread_ts || event.ts;
  const pastMessages = getThreadMessageLogs(threadId);

  // 会話履歴が30件を超えている場合
  if (pastMessages.length > 30) {
    const blocks = [
      {
        type: "divider"
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "⚠️ *会話が長くなりすぎています*"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "現在のスレッドでの会話を終了し、新しいスレッドで会話を続けることをお勧めします。\n\n新しいスレッドでは、より効率的に会話を進めることができます。"
        }
      },
      {
        type: "divider"
      }
    ];

    postMessage(event.channel, "会話を終了します", blocks, threadId);
    return;
  }

  // 過去のメッセージ履歴を会話形式に整形
  let conversationHistory = "";
  if (pastMessages && pastMessages.length > 0) {
    conversationHistory = pastMessages.map((msg, index) => {
      const speaker = msg.userId === "BOT" ? "システム" : `ユーザー${msg.userId}`;
      return `${speaker}: ${msg.text}`;
    }).join('\n');
  }

  // LLMに会話を生成させる
  const prompt = `
以下の会話履歴を踏まえて、自然な会話を続けてください。

# 会話履歴
${conversationHistory}

# 最新のメッセージ
${event.text}

# 注意事項
- 簡潔で分かりやすい回答を心がけてください
- 必要に応じて質問をしてください
- 会話の文脈を理解し、適切な応答をしてください
- システムとしての立場を保ち、プロフェッショナルな対応をしてください

# 出力形式
{
  "response": "応答メッセージ",
  "shouldAskQuestion": true/false,
  "question": "質問内容（shouldAskQuestionがtrueの場合のみ）"
}
`;

  try {
    const result = callGeminiApi(prompt, true);
    const parsedResult = extractAndParseJson(result);

    // 応答を送信
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: parsedResult.response
        }
      }
    ];

    // 質問がある場合は追加
    if (parsedResult.shouldAskQuestion && parsedResult.question) {
      blocks.push({
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*質問：* ${parsedResult.question}`
        }
      });
    }

    postMessage(event.channel, "会話の応答", blocks, threadId);

  } catch (error) {
    // エラー時の処理
    const errorBlocks = [
      {
        type: "divider"
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "❌ *エラーが発生しました*"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "申し訳ありません。会話の処理中にエラーが発生しました。\nもう一度お試しください。"
        }
      },
      {
        type: "divider"
      }
    ];

    postMessage(event.channel, "エラーが発生しました", errorBlocks, threadId);
    logError(error, "会話処理中にエラーが発生", { threadId, pastMessages });
  }
}

/**
 * タスク一覧取得のSlackプロセス
 * @param {Object} event - Slackイベント
 */
function getTasksSlackProcess(event) {
  const threadId = event.thread_ts || event.ts;
  const channelId = event.channel;
  let userMessageText = event.text;

  try {
    const pastMessages = getThreadMessageLogs(threadId);
    if (pastMessages.length > 0) {
      userMessageText = pastMessages.map(m => m.text).join('\n') + '\n' + userMessageText;
    }

    // callgemini.js の関数を呼び出し
    const extractionResult = extractTargetSlackUserId(userMessageText);

    let targetSlackUserId = null;
    if (extractionResult && extractionResult.found && extractionResult.slackUserId) {
      targetSlackUserId = extractionResult.slackUserId;
    } else {
      const blocks = [
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "どなたのタスク一覧を取得しますか？ 名前で指定してください。"
          }
        }
      ];
      postMessage(channelId, "タスク取得対象のユーザーを指定してください。", blocks, threadId);
      return;
    }

    // アクティブなタスクシートIDを取得 (spreadsheet.jsの関数を呼び出し)
    const activeSheetIds = getActiveTaskSheetIdsFromController();

    if (activeSheetIds.length === 0) {
      postMessage(channelId, "現在参照可能なアクティブなタスクシートがありません。", null, threadId);
      return;
    }

    // TaskReminderControllerから案件名のマッピングを作成
    const sheetIdToProjectName = new Map();
    try {
      const controllerSheet = SpreadsheetApp.openById(CONTROLLER_SPREADSHEET_ID)
        .getSheetByName(CONTROLLER_SHEET_NAME);
      
      if (controllerSheet) {
        const data = controllerSheet.getDataRange().getValues();
        // ヘッダー行をスキップして処理
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const sheetId = String(row[CONTROLLER_COL_SHEET_ID]).trim();
          const status = String(row[CONTROLLER_COL_STATUS]).trim();
          const projectName = String(row[CONTROLLER_COL_PROJECT_NAME]).trim();
          
          if (status === CONTROLLER_STATUS_ACTIVE && activeSheetIds.includes(sheetId)) {
            sheetIdToProjectName.set(sheetId, projectName);
          }
        }
      }
    } catch (error) {
      Logger.log('案件名の取得中にエラーが発生しました: ' + error.message);
    }

    let allUserTasks = [];
    activeSheetIds.forEach(sheetId => {
      const tasksFromSheet = getPendingTasksForUserFromSheet(targetSlackUserId, sheetId);
      allUserTasks = allUserTasks.concat(tasksFromSheet);
    });

    if (allUserTasks.length === 0) {
      postMessage(channelId, `<@${targetSlackUserId}>さんの未完了タスクは見つかりませんでした。`, null, threadId);
      return;
    }

    // 重複を排除
    const uniqueTasks = Array.from(new Set(allUserTasks.map(task => JSON.stringify(task))))
      .map(taskStr => JSON.parse(taskStr));

    // タスク一覧を表示するブロックを作成
    const blocks = [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: "未完了タスク一覧",
          emoji: true
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `<@${targetSlackUserId}> さんのタスク`
        }
      }
    ];

    // タスクを期日でソート
    uniqueTasks.sort((a, b) => {
      const dateA = new Date(a.dueDate);
      const dateB = new Date(b.dueDate);
      return dateA - dateB;
    });

    // 最大50個のタスクまで表示
    const maxTasks = Math.min(uniqueTasks.length, 50);
    const TASKS_PER_MESSAGE = 10; // 1メッセージあたりのタスク数

    // 最初のメッセージを送信
    postMessage(channelId, `<@${targetSlackUserId}>さんの未完了タスク一覧`, blocks, threadId);

    // タスクを複数のメッセージに分割して送信
    for (let i = 0; i < maxTasks; i += TASKS_PER_MESSAGE) {
      const messageBlocks = [];
      const endIndex = Math.min(i + TASKS_PER_MESSAGE, maxTasks);
      
      for (let j = i; j < endIndex; j++) {
        const task = uniqueTasks[j];
        messageBlocks.push({ type: "divider" });
        
        let taskDetail = "";
        const projectName = sheetIdToProjectName.get(task.sheetId);

        if (projectName) taskDetail += `*案件: ${projectName}*`;
        taskDetail += `\n*概要:* ${task.summary}\n*期日:* ${task.dueDate}`;
        if (task.status) taskDetail += `\n*ステータス:* ${task.status}`;
        
        if (task.sheetId) {
          const sheetUrl = `https://docs.google.com/spreadsheets/d/${task.sheetId}/edit`;
          taskDetail += `\n*シート:* <${sheetUrl}|開く>`;
        }

        messageBlocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: taskDetail
          }
        });
      }

      // 各メッセージにページ情報を追加
      const currentPage = Math.floor(i / TASKS_PER_MESSAGE) + 1;
      const totalPages = Math.ceil(maxTasks / TASKS_PER_MESSAGE);
      messageBlocks.push({
        type: "context",
        elements: [
          {
            type: "mrkdwn",
            text: `*ページ ${currentPage}/${totalPages}*`
          }
        ]
      });

      // 分割したメッセージを送信
      postMessage(channelId, "", messageBlocks, threadId);
    }

    // タスク数が多い場合は注意書きを追加
    if (uniqueTasks.length > maxTasks) {
      const footerBlocks = [{
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*注意:* 表示できるタスクは最大${maxTasks}件までです。残りの${uniqueTasks.length - maxTasks}件のタスクは省略されています。`
        }
      }];
      postMessage(channelId, "", footerBlocks, threadId);
    }

  } catch (error) {
    logError(error, "getTasksSlackProcess: Error processing request", { event });
    postMessage(channelId, "タスク一覧の取得中にエラーが発生しました。もう一度お試しください。", null, threadId);
  }
}

function completeTaskSlackProcess(event) {
  const threadId = event.thread_ts || event.ts;
  const channelId = event.channel;
  const userMessageText = event.text;

  try {
    // タスク完了の確認状態をチェック
    const confirmation = getPendingTaskCompleteConfirmation(threadId);

    if (confirmation && confirmation.status === "pending") {
      // 肯定の返事かどうかを判定
      const pastMessages = getThreadMessageLogs(threadId);
      const isPositive = isPositiveResponseToCompleteTask(userMessageText, pastMessages);
      
      if (isPositive) {
        // タスクを完了に更新
        try {
          const { targetSheetId, taskSummary } = confirmation.json;
          const success = completeTaskInSheet(targetSheetId, taskSummary);
          
          if (!success) {
            throw new Error(`タスク「${taskSummary}」が見つからないか、すでに完了しています。`);
          }

          // 確認状態を更新
          saveTaskCompleteConfirmation(threadId, {
            ...confirmation,
            status: "completed"
          });

          // スレッドのプロセス種類をリセット
          resetThreadProcessType(threadId);

          const message = `タスク「${taskSummary}」を完了にしました。`;
          postMessage(channelId, message, null, threadId);

        } catch (error) {
          postMessage(channelId, `タスクの完了処理中にエラーが発生しました: ${error.message}`, null, threadId);
          logError(error, "タスク完了処理中にエラーが発生", { confirmation });
          
          // 確認状態を更新
          saveTaskCompleteConfirmation(threadId, {
            ...confirmation,
            status: "error"
          });
        }
        return;
      } else {
        // 否定の返事の場合、確認状態をリセット
        saveTaskCompleteConfirmation(threadId, {
          ...confirmation,
          status: "cancelled"
        });
        
        postMessage(channelId, "タスクの完了をキャンセルしました。", null, threadId);
        return;
      }
    }

    // チャンネルIDに紐づくタスクシートを検索
    const targetSheetId = getSheetIdFromChannelId(channelId);
    if (!targetSheetId) {
      postMessage(channelId, "このチャンネルに紐づくタスクシートが見つかりませんでした。", null, threadId);
      return;
    }

    // Geminiを使用してタスク概要を抽出
    const extractionResult = extractTaskSummary(userMessageText);

    if (!extractionResult.found || !extractionResult.taskSummary) {
      postMessage(channelId, "タスクの概要を特定できませんでした。具体的なタスク概要を指定してください。", null, threadId);
      return;
    }

    const taskSummary = extractionResult.taskSummary;

    // タスクシートを開いてタスクの存在を確認
    const spreadsheet = SpreadsheetApp.openById(targetSheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
    
    if (!sheet) {
      postMessage(channelId, "タスクシートが見つかりませんでした。", null, threadId);
      return;
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
        foundTask = true;
        break;
      }
    }

    if (!foundTask) {
      postMessage(channelId, `タスク「${taskSummary}」が見つからないか、すでに完了しています。`, null, threadId);
      return;
    }

    // 確認状態を保存
    saveTaskCompleteConfirmation(threadId, {
      status: "pending",
      json: {
        targetSheetId: targetSheetId,
        taskSummary: taskSummary
      }
    });

    // 確認メッセージを送信
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `タスク「${taskSummary}」を完了にしますか？`
        }
      }
    ];
    postMessage(channelId, "タスク完了の確認", blocks, threadId);

  } catch (error) {
    logError(error, "completeTaskSlackProcess: Error processing request", { event });
    postMessage(channelId, "タスクの完了処理中にエラーが発生しました。", null, threadId);
  }
}

/**
 * カレンダーの空き枠を取得するSlackプロセス
 * @param {Object} event - Slackイベント
 * @return {boolean} 処理結果
 */
function getCalendarAvailabilitySlackProcess(event) {
  try {
    const threadId = event.thread_ts || event.ts;
    const channelId = event.channel;

    // 現在の日付を取得
    const currentDate = new Date();

    // カレンダー情報を抽出
    const calendarInfo = extractCalendarAvailability(event.text, currentDate);
    
    // 開始日を決定
    let startDate = null;
    if (calendarInfo.startDate && calendarInfo.startDate.trim() !== '') {
      startDate = calendarInfo.startDate;
    }
    
    // カレンダーの空き枠を解析
    const availability = analyzeCalendarAvailability(
      ["kiko.bandai@tsuide-inc.com","kiko.bandai@buki-ya.com","bandai@inbound.llc"],
      calendarInfo.days,
      calendarInfo.startTime,
      calendarInfo.endTime,
      startDate
    );

    if (availability.error) {
      postMessage(channelId, "カレンダーの空き枠の取得中にエラーが発生しました。", null,threadId);
      return false;
    }

    // 結果をSlackメッセージとして整形
    const blocks = [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: "📅 萬代のカレンダーの空き枠は下記です。カレンダーの調整をお願いします。",
          emoji: true
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*確認期間:* ${availability.startDate}から${calendarInfo.days}日間${calendarInfo.startDateDescription ? `（${calendarInfo.startDateDescription}から）` : ''}\n*時間範囲:* ${calendarInfo.startTime} - ${calendarInfo.endTime}`
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `恐れ入りますが、上記の時間枠の中で、kiko.bandai@buki-ya.comにカレンダーを招待して下さい！後で実施する場合タスクに登録し漏れないように、何卒よろしくお願いいたします。`
        }
      }
    ];

    // 各日の空き枠を追加
    availability.availability.forEach(day => {
      if (day.slots.length > 0) {
        const slotsText = day.slots.map(slot => 
          `• ${slot.start} - ${slot.end} (${slot.duration})`
        ).join('\n');

        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: `*${day.date} (${day.dayOfWeek}) - ${day.fullDate}*\n${slotsText}`
          }
        });
      }
    });

    // 空き枠がない場合のメッセージ
    if (availability.availability.every(day => day.slots.length === 0)) {
      blocks.push({
        type: "section",
        text: {
          type: "mrkdwn",
          text: "指定された期間内に空き枠が見つかりませんでした。"
        }
      });
    }

    // Slackにメッセージを送信
    postMessage(channelId, "カレンダーの空き枠確認結果", blocks, threadId);

    // カレンダー確認完了後、スレッドのプロセス種類をcreateEventに切り替え
    saveThreadProcessType(threadId, 'createEvent');
    
    // 次のアクション（イベント作成）を促すメッセージを送信
    const nextActionBlocks = [
      {
        type: "divider"
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "📅 *カレンダー予約を作成しますか？*"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "上記の空き時間でカレンダー予約を作成したい場合は、以下の形式で入力してください：\n\n`event 予定名 開始日時 時間 招待者メールアドレス`\n\n例：`event 打ち合わせ 2025-01-20 14:00 60分 example@company.com`"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "または、具体的な予定の詳細をお知らせください。"
        }
      },
      {
        type: "divider"
      }
    ];
    postMessage(channelId, "次のアクション", nextActionBlocks, threadId);
    
    return true;

  } catch (error) {
    logError(error, "getCalendarAvailabilitySlackProcess: Error processing request", { event });
    postMessage(event.channel, "カレンダーの空き枠の取得中にエラーが発生しました。", null, event.thread_ts || event.ts);
    return false;
  }
}

/**
 * 不足項目をチェック
 * @param {string} channelId - チャンネルID
 * @param {string} threadId - スレッドID
 * @param {Object} json - タスク情報
 */
function checkMissingFields(channelId, threadId, json) {
  try {
    // タスク情報を検証
    const validationResult = callGeminiForValidation(json, null, json.SheetId);

    if (!validationResult.isValid) {
      // エラーメッセージを送信
      const blocks = [
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "❌ *タスク情報に問題があります*"
          }
        }
      ];

      // エラーメッセージを追加
      if (validationResult.errors && validationResult.errors.length > 0) {
        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: validationResult.errors.map(error => `• ${error}`).join('\n')
          }
        });
      }

      // 改善提案を追加
      if (validationResult.suggestions && validationResult.suggestions.length > 0) {
        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: "*改善提案:*\n" + validationResult.suggestions.map(suggestion => `• ${suggestion}`).join('\n')
          }
        });
      }

      postMessage(channelId, "タスク情報の修正が必要です", blocks, threadId);
      return;
    }

    // タスク情報を確認状態として保存
    saveTaskConfirmation(threadId, {
      status: "pending",
      json: json
    });

    // 確認メッセージを送信
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "以下のタスクを作成します："
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*概要:* ${json.概要}\n*期日:* ${json.期日}\n*アサイン:* ${json.アサイン}\n*ステータス:* ${json.ステータス}`
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "この内容でタスクを作成しますか？"
        }
      }
    ];

    postMessage(channelId, "タスク作成の確認", blocks, threadId);

  } catch (error) {
    logError(error, "不足項目チェック中にエラーが発生", { json });
    postMessage(channelId, "タスク情報の検証中にエラーが発生しました。", null, threadId);
  }
}

/**
 * カレンダー予約作成のSlackプロセス
 * @param {Object} event - Slackイベント
 */
function createEventSlackProcess(event) {
  // メッセージをログに記録
  logMessageToSheet(
    "'" + (event.thread_ts || event.ts),
    "'" + event.ts,
    event.user,
    event.text
  );

  const threadId = event.thread_ts || event.ts;
  var pastMessages = getThreadMessageLogs(threadId);
  
  // カレンダー予約の確認状態をチェック
  const confirmation = getPendingEventConfirmation(threadId);

  if (confirmation && confirmation.status === "pending") {
    // 肯定の返事かどうかを判定
    const isPositive = isPositiveResponse(event.text, pastMessages);
    
    if (isPositive) {
      // カレンダー予約を作成
      try {
        const result = createCalendarEvent(confirmation.json);
        
        // 予約情報を整形して表示
        const message = `カレンダー予約を作成しました`;
        
        postMessage(event.channel, message, null, threadId);
        
        // 確認状態を更新
        saveEventConfirmation(threadId, {
          ...confirmation,
          status: "completed"
        });

        // スレッドのプロセス種類をリセット
        resetThreadProcessType(threadId);
        
        // 次のアクションを促すメッセージを送信
        const blocks = [
          {
            type: "divider"
          },
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: "🎉 *カレンダー予約の作成が完了しました！*"
            }
          },
          {
            type: "divider"
          }
        ];
        
        postMessage(event.channel, "ご希望がございましたら、次のアクションを指示して下さい", blocks, threadId);
        
        return; // 予約作成が完了したら処理を終了
      } catch (error) {
        postMessage(event.channel, `カレンダー予約の作成中にエラーが発生しました: ${error.message}`, null, threadId);
        logError(error, "カレンダー予約作成中にエラーが発生", { json: confirmation.json });
        
        // 確認状態を更新
        saveEventConfirmation(threadId, {
          ...confirmation,
          status: "error"
        });
        
        return; // エラーが発生したら処理を終了
      }
    } else {
      // 否定の返事の場合、確認状態をリセット
      saveEventConfirmation(threadId, {
        ...confirmation,
        status: "cancelled"
      });
      
      postMessage(event.channel, "カレンダー予約の作成をキャンセルしました。", null, threadId);
      return; // キャンセルしたら処理を終了
    }
  }

  // 既存のイベントJSONを取得
  const existingJson = getEventJson(threadId);

  // もう一度更新
  pastMessages = getThreadMessageLogs(threadId);
  
  // 現在の日付を取得
  const currentDate = new Date();
  
  // イベント情報を抽出
  const json = callGeminiForEventJson(event.text, existingJson, pastMessages, currentDate);

  // イベント情報を保存
  saveEventJson(threadId, json);

  // 不足項目をチェック
  checkMissingEventFields(event.channel, threadId, json);
}

/**
 * カレンダーイベントの不足項目をチェック
 * @param {string} channelId - チャンネルID
 * @param {string} threadId - スレッドID
 * @param {Object} json - イベント情報
 */
function checkMissingEventFields(channelId, threadId, json) {
  try {
    // イベント情報を検証
    const validationResult = callGeminiForEventValidation(json, null);

    if (!validationResult.isValid) {
      // エラーメッセージを送信
      const blocks = [
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "❌ *カレンダーイベント情報に問題があります*"
          }
        }
      ];

      // エラーメッセージを追加
      if (validationResult.errors && validationResult.errors.length > 0) {
        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: validationResult.errors.map(error => `• ${error}`).join('\n')
          }
        });
      }

      // 改善提案を追加
      if (validationResult.suggestions && validationResult.suggestions.length > 0) {
        blocks.push({
          type: "section",
          text: {
            type: "mrkdwn",
            text: "*改善提案:*\n" + validationResult.suggestions.map(suggestion => `• ${suggestion}`).join('\n')
          }
        });
      }

      postMessage(channelId, "カレンダーイベント情報の修正が必要です", blocks, threadId);
      return;
    }

    // イベント情報を確認状態として保存
    saveEventConfirmation(threadId, {
      status: "pending",
      json: json
    });

    // 確認メッセージを送信
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "以下のカレンダー予約を作成します："
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*タイトル:* ${json.title}\n*開始日時:* ${json.startDateTime}\n*時間:* ${json.duration || 30}分${json.guestEmail ? `\n*招待者:* ${json.guestEmail}` : ''}`
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "この内容でカレンダー予約を作成しますか？"
        }
      }
    ];

    postMessage(channelId, "カレンダー予約作成の確認", blocks, threadId);

  } catch (error) {
    logError(error, "カレンダーイベント不足項目チェック中にエラーが発生", { json });
    postMessage(channelId, "カレンダーイベント情報の検証中にエラーが発生しました。", null, threadId);
  }
}

/**
 * すべてのモジュールのテストを実行
 * @throws {Error} テストが失敗した場合
 */
function test_all() {
  try {
    // Slack APIのテスト
    //test_slack();
    
    // Salesforce APIのテスト
    //test_salesforce();
    
    // Notion APIのテスト
    test_notion();
    
    // Gemini APIのテスト
    test_gemini();
    
    console.log("All Tests Completed Successfully");
  } catch (error) {
    console.error("Test Failed:", error);
    throw error;
  }
} 
/**
 * createTaskSlackProcessのテスト関数
 * @throws {Error} テストが失敗した場合
 */
function test_createTaskSlackProcess() {
  try {
    // テスト用のモックイベントを作成
    const mockEvent = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "create テストタスク 2024-04-01 萬代"
    };

    const mockEvent2 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "OK"
    };

    // テスト用のスプレッドシートIDを設定
    const testSheetId = "1VzSP6Ab61nlcYKbNU_5iMVtUAr1zn9vuL7WcaKL6oEw"; // 実際のテスト用スプレッドシートIDに置き換えてください

    // テスト用のタスクJSONを作成
    const testTaskJson = {
      "概要": "テストタスク",
      "期日": "2025-06-01",
      "アサイン": "萬代 貴昂",
      "SheetId": testSheetId
    };

    // Json作製実行
    createTaskSlackProcess(mockEvent);

    //
    saveTaskConfirmation(mockEvent.thread_ts, {
      status: "pending",
      json: testTaskJson
    });

    // OK → Task作製
    createTaskSlackProcess(mockEvent2);
    return true;

  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}

/**
 * getTasksSlackProcessのテスト関数
 * @throws {Error} テストが失敗した場合
 */
function test_getTasksSlackProcess() {
  try {
    // テストケース1: メンション付きのメッセージ
    const mockEvent1 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "萬代のタスク一覧を表示して"
    };

    // テストケース2: メンションなしのメッセージ
    const mockEvent2 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "タスク一覧を表示して"
    };

    // テストケース3: 複数のメンションがある場合
    const mockEvent3 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "<@U01234567> と <@U76543210> のタスク一覧を表示して"
    };

    Logger.log("=== getTasksSlackProcess テスト開始 ===");

    // テストケース1の実行
    Logger.log("テストケース1: メンション付きのメッセージ");
    getTasksSlackProcess(mockEvent1);

    // テストケース2の実行
    //Logger.log("テストケース2: メンションなしのメッセージ");
    //getTasksSlackProcess(mockEvent2);

    // テストケース3の実行
    //Logger.log("テストケース3: 複数のメンションがある場合");
    //getTasksSlackProcess(mockEvent3);

    Logger.log("=== getTasksSlackProcess テスト終了 ===");
    return true;

  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}

/**
 * completeTaskSlackProcessのテスト関数
 * @throws {Error} テストが失敗した場合
 */
function test_completeTaskSlackProcess() {
  try {
    Logger.log("=== completeTaskSlackProcess テスト開始 ===");

    // テストケース1: タスク完了の確認状態がある場合
    const mockEvent1 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "はい"
    };

    // テストケース2: タスク概要の抽出と完了処理
    const mockEvent2 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "done テストタスクを完了"
    };

    // テストケース3: エラーケース（存在しないタスク）
    const mockEvent3 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "done 存在しないタスク"
    };

    // テストケース1の実行
    Logger.log("テストケース1: タスク完了の確認状態がある場合");
    // 確認状態を設定
    saveTaskCompleteConfirmation(mockEvent1.thread_ts, {
      status: "pending",
      json: {
        targetSheetId: "1VzSP6Ab61nlcYKbNU_5iMVtUAr1zn9vuL7WcaKL6oEw",
        taskSummary: "テストタスク"
      }
    });
    completeTaskSlackProcess(mockEvent1);

    // テストケース2の実行
    Logger.log("テストケース2: タスク概要の抽出と完了処理");
    completeTaskSlackProcess(mockEvent2);

    // テストケース3の実行
    Logger.log("テストケース3: エラーケース（存在しないタスク）");
    completeTaskSlackProcess(mockEvent3);

    Logger.log("=== completeTaskSlackProcess テスト終了 ===");
    return true;

  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}

/**
 * getCalendarAvailabilitySlackProcessのテスト関数
 * @throws {Error} テストが失敗した場合
 */
function test_getCalendarAvailabilitySlackProcess() {
  try {
    Logger.log("=== getCalendarAvailabilitySlackProcess テスト開始 ===");

    // テストケース1: 基本的なカレンダー確認
    const mockEvent1 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "calendar 直近3日 20時まで"
    };


    // テストケース1の実行
    Logger.log("テストケース1: 基本的なカレンダー確認");
    getCalendarAvailabilitySlackProcess(mockEvent1);

    Logger.log("=== getCalendarAvailabilitySlackProcess テスト終了 ===");
    return true;

  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}

/**
 * createEventSlackProcessのテスト関数
 * @throws {Error} テストが失敗した場合
 */
function test_createEventSlackProcess() {
  try {
    Logger.log("=== createEventSlackProcess テスト開始 ===");

    // テストケース1: 基本的なイベント作成
    const mockEvent1 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "event テストミーティング 2025-01-20 14:00 60分 test@example.com"
    };

    // テストケース2: 確認の返事
    const mockEvent2 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "はい"
    };

    // テストケース1の実行
    Logger.log("テストケース1: 基本的なイベント作成");
    createEventSlackProcess(mockEvent1);

    // テスト用のイベントJSONを作成
    const testEventJson = {
      "title": "テストミーティング",
      "startDateTime": "2025-01-20 14:00",
      "duration": 60,
      "guestEmail": "test@example.com"
    };

    // 確認状態を設定
    saveEventConfirmation(mockEvent1.thread_ts, {
      status: "pending",
      json: testEventJson
    });

    // テストケース2の実行
    Logger.log("テストケース2: 確認の返事");
    createEventSlackProcess(mockEvent2);

    Logger.log("=== createEventSlackProcess テスト終了 ===");
    return true;

  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}