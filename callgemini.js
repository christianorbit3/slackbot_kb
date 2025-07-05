/**
 * Gemini 2.5 Pro (I/O edition) で月次レポートを生成する
 *
 * @param {string} monthlySummaryCsvBinB64  base64 エンコード済み CSV
 * @param {string} weeklySummaryCsvBin2B64  base64 エンコード済み CSV
 * @param {string} dailySummaryCsvBinB64    base64 エンコード済み CSV
 * @param {string} mdTemplate               報告書テンプレート（Markdown）
 * @param {Date}   today                    処理基準日
 * @return {string}                          Gemini から返ってきた Markdown
 */
function callGeminiInternalReport(
  monthlySummaryCsvBinB64,
  weeklySummaryCsvBin2B64,
  dailySummaryCsvBinB64,
  mdTemplate,
  today
) {
  /* ---------- 事前計算 ---------- */
  const todayString        = getTodayJSTString(today);
  const bizDaysThisMonth   = getBizDaysThisMonth(today);
  const remainingBizDays   = getRemainingBizDays(today);
  const usedBusinessDayProrated = (bizDaysThisMonth - remainingBizDays) / bizDaysThisMonth;
  const thisMonth  = thisMonthStartJST(today);
  const oneMonthAgo = lastMonthStartJST(today);
  const twoMonthAgo = twoMonthsAgoStartJST(today);

  /* ---------- API キー ---------- */
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定');

  /* ---------- Gemini API へ渡す contents を構築 ---------- */
  const contents = [];

  // Gemini には system ロールが無いので、先頭に「指示文」を user ロールで置く
  contents.push({
    role: 'user',
    parts: [
      {
        text: [
          'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい',
          '---'
        ].join('\n')
      }
    ]
  });

  // メイン指示
  contents.push({
    role: 'user',
    parts: [
      {
        text: [
          '私は広告代理店のプロジェクトマネージャーです。',
          'クライアントへの月次報告資料を作成しており、いまから共有する月次サマリーデータを、提示するテンプレートに沿った形式でまとめてください。',
          '途中で質問することなく、テンプレートの指示に従ってマークダウン形式のテキストで出力してください。マークダウンとして正しく出力することに注意して下さい。ただし「```markdown」や [####] 見出しは使わないで下さい。',
          'できるだけ深い考察をお願いします。また、分析を開始する前に、これまでの同様の分析指示の結果については混同しないようなるべく忘れてください。',
          '',
          'テンプレート内で使える変数: ',
          `- {{today}}=${todayString}`,
          `- {{usedBusinessDayProrated}}=${usedBusinessDayProrated}`,
          `- {{remainingBizDays}}=${remainingBizDays}`,
          `- {{bizDaysThisMonth}}=${bizDaysThisMonth}`,
          `- {{thisMonth}}=${thisMonth}`,
          `- {{oneMonthAgo}}=${oneMonthAgo}`,
          `- {{twoMonthAgo}}=${twoMonthAgo}`,
          '',
          '```',
          mdTemplate.trim(),
          '```',
          '',
          '報告書のテンプレートは以上です。提供データは monthlySummary（CSV）、weeklySummary（CSV）、dailySummary（CSV）の３つ。テンプレートの中で適切に使い分けて下さい。'
        ].join('\n')
      }
    ]
  });

  /* ---------- 大きな CSV は 20 000 文字ごとに分割 ---------- */
  const pushCsvChunks = (label, b64) => {
    const chunks = chunkString(base64ToUtf8(b64), 20000);
    for (let i = 0; i < chunks.length; ++i) {
      contents.push({
        role: 'user',
        parts: [
          {
            text: [
              `### ${label}.csv — part ${i + 1}`,
              '```',
              chunks[i],
              '```'
            ].join('\n')
          }
        ]
      });
    }
  };

  pushCsvChunks('monthlySummary', monthlySummaryCsvBinB64);
  pushCsvChunks('weeklySummary',  weeklySummaryCsvBin2B64);
  pushCsvChunks('dailySummary',   dailySummaryCsvBinB64);

  /* ---------- リクエスト本体 ---------- */
  const body = {
    model: 'gemini-2.5-pro-preview-06-05', // 2025-05 現在の最新 Preview ID
    contents: contents,
    generationConfig: {
      maxOutputTokens: 100000          // 30k トークン上限（必要に応じて調整）
    }
  };

  /* ---------- HTTP 送信 ---------- */
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${body.model}:generateContent?key=${apiKey}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    // Apps Script では「followRedirects: true」がデフォルトのため省略
    timeout: 180000 // 180 s
  };

  const resp     = UrlFetchApp.fetch(url, options);
  const respJson = JSON.parse(resp.getContentText());

  /* ---------- 返り値を抽出 ---------- */
  try {
    // Gemini のレスポンス構造:
    //  candidates[0].content.parts[0].text
    return respJson.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Gemini API error → ' + resp.getContentText());
  }
}

/**
 * Gemini 2.5 Pro (I/O edition) で
 * 「## Slackポスト内容」節だけを抜き出す関数（Apps Script）
 *
 * @param {string} report  完成済み Markdown レポート全文
 * @return {string}        Gemini が返した抜粋テキスト
 */
function callGeminiInternalReportSlackSummary(report) {
  /* -------------------- 認証キー -------------------- */
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定');

  /* -------------------- プロンプト -------------------- */
  // Gemini には system ロールが無いので、指示文を user ロールに含める
  const contents = [
    {
      role: 'user',
      parts: [
        {
          text: [
            'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい。'
          ].join('\n')
        }
      ]
    },
    {
      role: 'user',
      parts: [
        {
          text: [
            '下記のレポートから一部分のテキストを **改行・表記そのまま** 正確に抜き出して下さい。アウトプットはその内容のみをテキストで返して下さい。その他のコメントは不要です。',
            '',
            '### 抜き出してほしい箇所',
            '「Slackポスト内容」という見出しから、次の見出しまでのすべてのテキストを抽出して下さい。',
            '',
            '以下、抽出対象のレポート全文です。',
            '```',
            report.trim(),
            '```'
          ].join('\n')
        }
      ]
    }
  ];

  /* -------------------- リクエスト -------------------- */
  const body = {
    model: 'gemini-2.5-pro-preview-06-05',   // 2025-05 時点の最新 Preview ID
    contents: contents,
    generationConfig: {
      maxOutputTokens: 100000                 // 元と同等の上限
    }
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${body.model}:generateContent?key=${apiKey}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    timeout: 180000                           // 180 s
  };

  /* -------------------- 呼び出し -------------------- */
  const respText = UrlFetchApp.fetch(url, options).getContentText();

  try {
    const respJson = JSON.parse(respText);
    // Gemini の戻り値は candidates[0].content.parts[0].text
    return respJson.candidates[0].content.parts[0].text;
  } catch (e) {
    return 'slackに投稿する内容の抽出に失敗しました';
  }
}

function callGeminiSearchKeywordReport(
  csvBinB64,
  mdTemplate,
) {

  /* ---------- API キー ---------- */
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定');

  /* ---------- Gemini API へ渡す contents を構築 ---------- */
  const contents = [];

  // Gemini には system ロールが無いので、先頭に「指示文」を user ロールで置く
  contents.push({
    role: 'user',
    parts: [
      {
        text: [
          'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい',
          '---'
        ].join('\n')
      }
    ]
  });

  // メイン指示
  contents.push({
    role: 'user',
    parts: [
      {
        text: [
          '私は広告代理店のプロジェクトマネージャーです。',
          'あるアカウントの検索キーワードを作成しており、いまから共有する検索キーワードのデータを、提示するテンプレートに沿った形式でまとめてください。',
          '途中で質問することなく、テンプレートの指示に従ってマークダウン形式のテキストで出力してください。マークダウンとして正しく出力することに注意して下さい。ただし「```markdown」や [####] 見出しは使わないで下さい。',
          'できるだけ深い考察をお願いします。また、分析を開始する前に、これまでの同様の分析指示の結果については混同しないようなるべく忘れてください。',
          '',
          '',
          '```',
          mdTemplate.trim(),
          '```',
          '',
          '報告書のテンプレートは以上です。提供データは テンプレートの中で適切に使い分けて下さい。'
        ].join('\n')
      }
    ]
  });

  /* ---------- 大きな CSV は 20 000 文字ごとに分割 ---------- */
  const pushCsvChunks = (label, b64) => {
    const chunks = chunkString(base64ToUtf8(b64), 20000);
    for (let i = 0; i < chunks.length; ++i) {
      contents.push({
        role: 'user',
        parts: [
          {
            text: [
              `### ${label}.csv — part ${i + 1}`,
              '```',
              chunks[i],
              '```'
            ].join('\n')
          }
        ]
      });
    }
  };

  pushCsvChunks('searchKeywordSummary', csvBinB64);

  /* ---------- リクエスト本体 ---------- */
  const body = {
    model: 'gemini-2.5-pro-preview-06-05', // 2025-05 現在の最新 Preview ID
    contents: contents,
    generationConfig: {
      maxOutputTokens: 100000          // 30k トークン上限（必要に応じて調整）
    }
  };

  /* ---------- HTTP 送信 ---------- */
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${body.model}:generateContent?key=${apiKey}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    // Apps Script では「followRedirects: true」がデフォルトのため省略
    timeout: 180000 // 180 s
  };

  const resp     = UrlFetchApp.fetch(url, options);
  const respJson = JSON.parse(resp.getContentText());

  /* ---------- 返り値を抽出 ---------- */
  try {
    // Gemini のレスポンス構造:
    //  candidates[0].content.parts[0].text
    return respJson.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Gemini API error → ' + resp.getContentText());
  }
}

/**
 * Gemini APIとの連携モジュール
 * 
 * このモジュールは、Google Gemini APIを使用してタスク情報の抽出と検証を行います。
 * 2段階の処理を実装：
 * 1. 会話テキストからタスク情報を抽出（JSON形式）
 * 2. 抽出された情報の必須項目チェック
 */

/**
 * Gemini APIの設定を取得
 * @return {Object} API設定
 * @throws {Error} 設定の取得に失敗した場合
 */
function getGeminiConfig() {
  const apiKey = getScriptProperty("GEMINI_API_KEY");
  const model = getConfig("GEMINI_MODEL") || "gemini-2.5-pro-preview-06-05";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  
  return { apiKey, model, url };
}

/**
 * Gemini APIを呼び出す
 * @param {string} prompt - プロンプト
 * @return {Object} APIレスポンス
 * @throws {Error} API呼び出しに失敗した場合
 */
function callGeminiApi(prompt, outputJson=false) {
  const { apiKey, url } = getGeminiConfig();

  const payloadObject = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };
  if (outputJson) {
    payloadObject.generationConfig = {
      responseMimeType: "application/json"
    };
  }
  
  let response, result, responseText;
  
  try {
    response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey
      },
      payload: JSON.stringify(payloadObject)
    });

    result = JSON.parse(response.getContentText());
    
    if (!result.candidates || !result.candidates[0]?.content?.parts?.[0]?.text) {
      throw new Error("Invalid response from Gemini API");
    }
    
    responseText = result.candidates[0].content.parts[0].text;
    
    // 成功時のログ記録
    logGeminiPromptToSheet(prompt, responseText, outputJson);
    
    return responseText;
    
  } catch (error) {
    // エラー時もログ記録（レスポンスはエラーメッセージ）
    const errorMessage = `エラー: ${error.message}`;
    logGeminiPromptToSheet(prompt, errorMessage, outputJson);
    
    throw error;
  }
}

/**
 * JSON文字列を抽出してパース
 * @param {string} text - テキスト
 * @return {Object} パースされたJSON
 * @throws {Error} JSONの抽出またはパースに失敗した場合
 */
function extractAndParseJson(text) {
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error("Failed to extract JSON from text");
  }
  
  try {
    return JSON.parse(jsonMatch[0]);
  } catch (error) {
    throw new Error(`Failed to parse JSON: ${error.message}`);
  }
}

/**
 * Geminiを使用してタスク情報を抽出
 * @param {string} text - ユーザーメッセージ
 * @param {Object} existingJson - 既存のタスクJSON
 * @param {Array} pastMessages - 過去のメッセージ履歴
 * @param {Date} currentDate - 現在の日付
 * @return {Object} 抽出されたタスク情報
 */
function callGeminiForJson(text, existingJson = null, pastMessages, currentDate = new Date()) {
  // 現在の日付を YYYY-MM-DD 形式で取得
  const todayString = currentDate.toLocaleDateString('ja-JP', { 
    year: 'numeric', 
    month: '2-digit', 
    day: '2-digit' 
  }).replace(/\//g, '-');

  // 過去のメッセージ履歴を会話形式に整形
  let conversationHistory = "";
  if (pastMessages && pastMessages.length > 0) {
    conversationHistory = pastMessages.map((msg, index) => {
      const speaker = msg.userId === "BOT" ? "システム" : `ユーザー${msg.userId}`;
      return `${speaker}: ${msg.text}`;
    }).join('\n');
  }

  // [Users]シートから担当者のフルネーム一覧を取得
  let assigneeList = [];
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = spreadsheet.getSheetByName("User");
    if (usersSheet) {
      const lastRow = usersSheet.getLastRow();
      if (lastRow > 0) {
        const assigneeRange = usersSheet.getRange(1, 1, lastRow, 1);
        assigneeList = assigneeRange.getValues()
          .flat()
          .filter(name => name !== ""); // 空のセルを除外
      }
    }
  } catch (error) {
    logError(error, "担当者一覧の取得中にエラーが発生");
  }

  const prompt = `
以下の会話履歴と最新のメッセージから、タスク情報を抽出してください。

# 会話履歴
${conversationHistory}

# 最新のメッセージ
${text}

# 抽出すべき情報
- タスクの概要
- 期日（YYYY-MM-DD形式）
- 担当者（フルネーム）

# 出力形式
以下のJSON形式で出力してください：
{
  "概要": "タスクの概要",
  "期日": "期日（YYYY-MM-DD形式）",
  "アサイン": "担当者のフルネーム",
  "ステータス": "未着手"
}

# アサイン（担当者）のフルネーム一覧
${assigneeList.join('\n')}

# 注意事項
- アサインが指定されていない場合は、メッセージから担当者を特定してください
- タスクの概要が不明確な場合は、質問を返してください
- 担当者は必ず上記の一覧から選択してください
- 今日の日付は${todayString}です
- 現在の年は${currentDate.getFullYear()}年です
- 現在の月は${currentDate.getMonth() + 1}月です
- 現在の日は${currentDate.getDate()}日です
- 現在の曜日は${['日', '月', '火', '水', '木', '金', '土'][currentDate.getDay()]}曜日です

`;

  try {
    const result = callGeminiApi(prompt, true);
    const parsedResult = extractAndParseJson(result);
    
    // 既存のJSONとマージ
    const mergedJson = {
      ...existingJson,
      ...parsedResult
    };
    
    return mergedJson;
  } catch (error) {
    logError(error, "タスク情報の抽出中にエラーが発生", { text, existingJson, pastMessages });
    throw error;
  }
}

/**
 * Geminiを使用してタスク情報を検証
 * @param {Object} json - 検証するタスク情報
 * @param {string} slackId - SlackユーザーID
 * @param {string} sheetId - タスクシートのID
 * @return {Object} 検証結果
 */
function callGeminiForValidation(json, slackId, sheetId) {
  // [Users]シートから担当者のフルネーム一覧を取得
  let assigneeList = [];
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const usersSheet = spreadsheet.getSheetByName("User");
    if (usersSheet) {
      const lastRow = usersSheet.getLastRow();
      if (lastRow > 0) {
        const assigneeRange = usersSheet.getRange(1, 1, lastRow, 1);
        assigneeList = assigneeRange.getValues()
          .flat()
          .filter(name => name !== ""); // 空のセルを除外
      }
    }
  } catch (error) {
    logError(error, "担当者一覧の取得中にエラーが発生", { sheetId });
  }

  const prompt = `
以下のタスク情報を検証してください。

# タスク情報
${JSON.stringify(json, null, 2)}

アサイン（担当者）のフルネーム一覧
${assigneeList.join('\n')}

# 検証項目
1. タスクの概要が明確か
2. 期日が適切な形式（YYYY-MM-DD）で指定されているか
3. アサインが上記の一覧から正しく選択されているか
4. ステータスが適切か

# 出力形式
以下のJSON形式で出力してください：
{
  "isValid": true/false,
  "errors": ["エラーメッセージ1", "エラーメッセージ2", ...],
  "suggestions": ["改善提案1", "改善提案2", ...]
}

# 注意事項
- エラーがある場合は、具体的な改善方法を提案してください
- 担当者は必ず上記の一覧から選択してください
- アサイン先は、必ず次の候補の中から、「表記揺れがないように」設定して下さい。
`;

  try {
    const result = callGeminiApi(prompt, true);
    return extractAndParseJson(result);
  } catch (error) {
    logError(error, "タスク情報の検証中にエラーが発生", { json, slackId, sheetId });
    throw error;
  }
}

/**
 * メッセージが肯定の返事かどうかを判定
 * @param {string} message - 判定するメッセージ
 * @param {Array} pastMessages - 過去のメッセージ履歴
 * @return {boolean} 肯定の返事かどうか
 * @throws {Error} 判定処理に失敗した場合
 */
function isPositiveResponse(message, pastMessages = []) {
  // 過去のメッセージ履歴を会話形式に整形
  let conversationHistory = "";
  if (pastMessages && pastMessages.length > 0) {
    conversationHistory = pastMessages.map((msg, index) => {
      const speaker = msg.userId === "BOT" ? "システム" : `ユーザー${msg.userId}`;
      return `${speaker}: ${msg.text}`;
    }).join('\n');
  }

  const prompt = `
以下の会話履歴と最新のメッセージを確認し、最新のメッセージが「タスクを登録しますか？」という質問に対する肯定の返事かどうかを判定してください。

# 会話履歴
${conversationHistory}

# 最新のメッセージ
${message}

# 判定基準
- 「はい」「OK」「了解」「承知」「お願いします」などの肯定表現を含む場合は「true」
- 「いいえ」「やめて」「キャンセル」などの否定表現を含む場合は「false」
- 明確な肯定・否定表現がない場合は「false」
- 質問や不明確な表現の場合は「false」

# 出力形式
{
  "isPositive": true/false,
  "reason": "判定理由"
}
`;

  try {
    const result = callGeminiApi(prompt, true);
    const parsedResult = extractAndParseJson(result);
    return parsedResult.isPositive;
  } catch (error) {
    logError(error, "肯定返事の判定中にエラーが発生", { message, pastMessages });
    return false; // エラー時は安全のためfalseを返す
  }
}

/**
 * メッセージからプロセス種類を判定
 * @param {string} message - 判定するメッセージ
 * @param {Array} pastMessages - 過去のメッセージ履歴
 * @return {Object} 判定結果
 * @throws {Error} 判定処理に失敗した場合
 */
function determineProcessType(message, pastMessages = []) {
  const prompt = `
以下のメッセージと会話履歴から、ユーザーが実行したい処理の種類を判定してください。

# 入力メッセージ
${message}

# 会話履歴
${pastMessages.map(m => `${m.userId}: ${m.text}`).join('\n')}

# 判定対象の処理種類
- getTasks: タスク一覧の取得（例：mytask、タスク一覧を表示して）
- completeTask: タスクの完了（例：done、完了、タスクを完了に）
- createTask: タスクの作成（例：create、タスクを作成、新規タスク）
- getCalendar: カレンダーの空き枠確認（例：calendar、カレンダー、空き時間、予定）
- createEvent: カレンダー予約の作成（例：event、予約、予定を作成、ミーティング、会議）
- communication: 通常の会話（上記以外の会話）

# 出力形式
以下のJSON形式で出力してください：
{
  "processType": "判定された処理種類",
  "confidence": 0.0-1.0の確信度
}

# 注意事項
- 確信度は0.0から1.0の間で、1.0が最も確実
- 確信度が0.7未満の場合は、ユーザーに確認が必要
- カレンダー関連のメッセージは、メールアドレスや予定、空き時間などのキーワードを含む場合にgetCalendarとして判定
- 予約作成関連のメッセージは、予定名、時間、会議などのキーワードを含む場合にcreateEventとして判定
`;

  try {
    const result = callGeminiApi(prompt, true);
    return extractAndParseJson(result);
  } catch (error) {
    logError(error, "determineProcessType: Error processing request", { message, pastMessages });
    return {
      processType: "communication",
      confidence: 0.0
    };
  }
}

/**
 * Slackメッセージからタスク参照対象のSlackUserIDを抽出します。
 * @param {string} userMessage - ユーザーからのメッセージ（過去のメッセージも含む可能性あり）。
 * @return {Object} 抽出結果。形式: { slackUserId: "UXXXXXXX" | null, found: true | false }
 * @throws {Error} API呼び出しまたはJSON処理に失敗した場合
 */
function extractTargetSlackUserId(userMessage) {
  const prompt = `
以下のメッセージから、タスク一覧を取得したい人物のSlackユーザーIDを抽出してください。
SlackユーザーIDは通常"U"で始まり、その後に英数字が続く形式です（例: U0123ABCDE）。
メンション形式（例: <@U0123ABCDE>）で記述されている場合は、ID部分のみを抽出してください。
メッセージ内に複数のユーザーIDが含まれる場合は、最も妥当と思われるものを1つだけ抽出してください。
U08QLADMSUFのタスクを検索することはありません。検索対象から除外して下さい。

#人物名とSlackUserIDとの対応表
萬代 貴昂	UNZ5061JM
村上 幸平	U08L19046DR
菊地 奏	U0871U28E5U
越村 大樹	U078F8Y6YE8
岡藤 隆平	U05N84AQ6R2
畑野 敏宏	U076R5UAMHT
KB 総務	U08QYSCN9HV
遠藤 未来彦	U073FNWBYV7
石神 沙織	U08RVAXE5U7
清水 幹郎	U0820MVEGNB
佐藤守	U08FZ6ACLRF
菰田 滉平	U088WF58DUG
倉林 真吾	U07FQDYARED
大森 一樹	D077GSCQZNJ
山下諒	D08GCE4RCK1
前田旭	D06G2DZQS3W
神田蒼空	U08KS9122QN
矢代千明	U079Z75UQVC
松浦侑人	D08BJ232L3X

#メッセージ
"""
${userMessage}
"""

抽出したSlackユーザーIDをJSON形式で返してください。見つからない場合は "found" を false にしてください。
{
  "slackUserId": "抽出されたSlackID" または null,
  "found": true または false
}
`;

  try {
    const geminiResponse = callGeminiApi(prompt, true); // outputJson = true
    const parsedResult = extractAndParseJson(geminiResponse);
    
    // <@UXXXXX> のような形式からIDを抽出する処理を追加
    if (parsedResult && parsedResult.found && parsedResult.slackUserId) {
      const match = parsedResult.slackUserId.match(/<@(.*?)>/);
      if (match && match[1]) {
        parsedResult.slackUserId = match[1];
      }
    }
    return parsedResult;

  } catch (error) {
    logError(error, "extractTargetSlackUserId: Error during Slack User ID extraction", { userMessage });
    // エラー時は見つからなかったとして返す
    return { slackUserId: null, found: false };
  }
}

/**
 * メッセージからタスク概要を抽出
 * @param {string} message - ユーザーメッセージ
 * @return {Object} 抽出結果
 */
function extractTaskSummary(message) {
  const prompt = `
以下のメッセージから、完了させたいタスクの概要を抽出してください。
メッセージは「done」や「完了」「yes」などのコマンドを含む可能性がありますが、それらは除外してタスクの概要のみを抽出してください。

# メッセージ
${message}

# 出力形式
{
  "taskSummary": "抽出されたタスク概要",
  "found": true/false
}
`;

  try {
    const result = callGeminiApi(prompt, true);
    return extractAndParseJson(result);
  } catch (error) {
    Logger.log('タスク概要の抽出中にエラーが発生しました: ' + error.message);
    return { found: false, taskSummary: null };
  }
}

function isPositiveResponseToCompleteTask(message, pastMessages = []) {
  // 過去のメッセージ履歴を会話形式に整形
  let conversationHistory = "";
  if (pastMessages && pastMessages.length > 0) {
    conversationHistory = pastMessages.map((msg, index) => {
      const speaker = msg.userId === "BOT" ? "システム" : `ユーザー${msg.userId}`;
      return `${speaker}: ${msg.text}`;
    }).join('\n');
  }

  const prompt = `
以下の会話履歴と最新のメッセージを確認し、最新のメッセージが「タスクを完了にしますか？」という質問に対する肯定の返事かどうかを判定してください。

# 会話履歴
${conversationHistory}

# 最新のメッセージ
${message}

# 判定基準
- 「はい」「OK」「了解」「承知」「お願いします」などの肯定表現を含む場合は「true」
- 「いいえ」「やめて」「キャンセル」などの否定表現を含む場合は「false」
- 明確な肯定・否定表現がない場合は「false」
- 質問や不明確な表現の場合は「false」

# 出力形式
{
  "isPositive": true/false,
  "reason": "判定理由"
}
`;

  try {
    const result = callGeminiApi(prompt, true);
    const parsedResult = extractAndParseJson(result);
    return parsedResult.isPositive;
  } catch (error) {
    logError(error, "肯定返事の判定中にエラーが発生", { message, pastMessages });
    return false; // エラー時は安全のためfalseを返す
  }
}

/**
 * メッセージからカレンダー情報を抽出
 * @param {string} message - ユーザーメッセージ
 * @param {Date} currentDate - 現在の日付
 * @return {Object} 抽出されたカレンダー情報
 */
function extractCalendarAvailability(message, currentDate = new Date()) {
  // 現在の日付を YYYY-MM-DD 形式で取得
  const todayString = currentDate.toLocaleDateString('ja-JP', { 
    year: 'numeric', 
    month: '2-digit', 
    day: '2-digit' 
  }).replace(/\//g, '-');

  const prompt = `
以下のメッセージから、カレンダーの空き枠を確認するために必要な情報を抽出してください。

# 入力メッセージ
${message}

# 抽出すべき情報
- メールアドレス（複数の場合も含む）
- 確認する日数（指定がない場合は3日）
- 時間範囲（指定がない場合は9:00-18:00）
- 開始日（「来週の火曜日」「明日」「1月20日」など、指定がない場合は今日から）

# 出力形式
以下のJSON形式で出力してください：
{
  "found": true/false,
  "emails": ["メールアドレス1", "メールアドレス2", ...],
  "days": 日数（数値）,
  "startTime": "開始時間（HH:MM形式）",
  "endTime": "終了時間（HH:MM形式）",
  "startDate": "開始日（YYYY-MM-DD形式、指定がない場合は空文字列）",
  "startDateDescription": "開始日の説明（「来週の火曜日」「明日」など、指定がない場合は空文字列）"
}

# 注意事項
- 日数や時間範囲が指定されていない場合は、デフォルト値を設定してください
- 開始日が相対的な表現（「来週の火曜日」「明日」など）の場合は、startDateDescriptionに記録し、startDateは可能であれば具体的な日付（YYYY-MM-DD形式）に変換してください
- 開始日が指定されていない場合は、startDateとstartDateDescriptionは空文字列にしてください
- 今日の日付は${todayString}です
- 現在の年は${currentDate.getFullYear()}年です
- 現在の月は${currentDate.getMonth() + 1}月です
- 現在の日は${currentDate.getDate()}日です
- 現在の曜日は${['日', '月', '火', '水', '木', '金', '土'][currentDate.getDay()]}曜日です
`;

  try {
    const result = callGeminiApi(prompt, true);
    return extractAndParseJson(result);
  } catch (error) {
    logError(error, 'extractCalendarAvailability: Error processing request', { message });
    return {
      found: false,
      error: error.message
    };
  }
}

/**
 * Geminiを使用してカレンダーイベント情報を抽出
 * @param {string} text - ユーザーメッセージ
 * @param {Object} existingJson - 既存のイベントJSON
 * @param {Array} pastMessages - 過去のメッセージ履歴
 * @param {Date} currentDate - 現在の日付
 * @return {Object} 抽出されたイベント情報
 */
function callGeminiForEventJson(text, existingJson = null, pastMessages, currentDate = new Date()) {
  // 現在の日付を YYYY-MM-DD 形式で取得
  const todayString = currentDate.toLocaleDateString('ja-JP', { 
    year: 'numeric', 
    month: '2-digit', 
    day: '2-digit' 
  }).replace(/\//g, '-');

  // 過去のメッセージ履歴を会話形式に整形
  let conversationHistory = "";
  if (pastMessages && pastMessages.length > 0) {
    conversationHistory = pastMessages.map((msg, index) => {
      const speaker = msg.userId === "BOT" ? "システム" : `ユーザー${msg.userId}`;
      return `${speaker}: ${msg.text}`;
    }).join('\n');
  }

  const prompt = `
以下の会話履歴と最新のメッセージから、カレンダーイベント情報を抽出してください。

# 会話履歴
${conversationHistory}

# 最新のメッセージ
${text}

# 抽出すべき情報
- イベントのタイトル（予定名）
- 開始日時（YYYY-MM-DD HH:MM形式）
- 時間（分単位、デフォルト30分）
- 招待するメールアドレス（任意）

# 出力形式
以下のJSON形式で出力してください：
{
  "title": "イベントのタイトル",
  "startDateTime": "開始日時（YYYY-MM-DD HH:MM形式）",
  "duration": 時間（分単位、数値）,
  "guestEmail": "招待するメールアドレス（任意）"
}

# 注意事項
- 開始日時が不明確な場合は、質問を返してください
- 時間が指定されていない場合は30分をデフォルトとしてください
- 1時間の指定がある場合は60分としてください
- 招待するメールアドレスが指定されていない場合は空文字列としてください
- 日付が指定されていない場合は今日の日付を使用してください
- 今日の日付は${todayString}です
- 現在の年は${currentDate.getFullYear()}年です
- 現在の月は${currentDate.getMonth() + 1}月です
- 現在の日は${currentDate.getDate()}日です
- 現在の曜日は${['日', '月', '火', '水', '木', '金', '土'][currentDate.getDay()]}曜日です
`;

  try {
    const result = callGeminiApi(prompt, true);
    const parsedResult = extractAndParseJson(result);
    
    // 既存のJSONとマージ
    const mergedJson = {
      ...existingJson,
      ...parsedResult
    };
    
    return mergedJson;
  } catch (error) {
    logError(error, "カレンダーイベント情報の抽出中にエラーが発生", { text, existingJson, pastMessages });
    throw error;
  }
}

/**
 * Geminiを使用してカレンダーイベント情報を検証
 * @param {Object} json - 検証するイベント情報
 * @param {string} slackId - SlackユーザーID
 * @return {Object} 検証結果
 */
function callGeminiForEventValidation(json, slackId) {
  const prompt = `
以下のカレンダーイベント情報を検証してください。

# イベント情報
${JSON.stringify(json, null, 2)}

# 検証項目
1. イベントのタイトルが明確か
2. 開始日時が適切な形式（YYYY-MM-DD HH:MM）で指定されているか
3. 時間（duration）が適切な値（分単位）で指定されているか
4. 招待するメールアドレスが有効な形式か（指定されている場合）

# 出力形式
以下のJSON形式で出力してください：
{
  "isValid": true/false,
  "errors": ["エラーメッセージ1", "エラーメッセージ2", ...],
  "suggestions": ["改善提案1", "改善提案2", ...]
}

# 注意事項
- エラーがある場合は、具体的な改善方法を提案してください
- 開始日時は現在時刻より未来である必要があります
- 時間は15分以上240分以下の範囲で指定してください
`;

  try {
    const result = callGeminiApi(prompt, true);
    return extractAndParseJson(result);
  } catch (error) {
    logError(error, "カレンダーイベント情報の検証中にエラーが発生", { json, slackId });
    throw error;
  }
}