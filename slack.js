/**
 * 渡された URL（通常は Notion ページ URL）を
 * 指定チャンネルにポストするだけの関数。
 *
 * @param {string} slackChannelId  Slack チャンネル ID (例: 'C0123456789')
 * @param {string} message         投稿メッセージ（URL 含む）
 */
function postToSlack(slackChannelId, message) {
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
  if (!slackToken) throw new Error('SLACK_TOKEN が未設定');

  const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method      : 'post',
    contentType : 'application/json; charset=utf-8',
    payload     : JSON.stringify({
      channel : slackChannelId,
      text    : message
    }),
    headers: { Authorization: `Bearer ${slackToken}` }
  });

  const ok = JSON.parse(res.getContentText()).ok;
  if (!ok) throw new Error(`Slack 送信失敗 → ${res.getContentText()}`);
}

/**
 * Slack APIとの連携モジュール
 * 
 * このモジュールは、Slack APIを使用して以下の機能を提供します：
 * - メッセージの送信
 * - 署名の検証
 * - タスク情報の処理
 * - スレッドメッセージの管理
 */

/**
 * Slack APIの設定を取得
 * @return {Object} API設定
 * @throws {Error} 設定の取得に失敗した場合
 */
function getSlackConfig() {
  const token = getScriptProperty("SLACK_TOKEN");
  const verifyToken = getScriptProperty("SLACK_VERIFY_TOKEN");
  
  if (!token || !verifyToken) {
    throw new Error("Slack configuration is incomplete");
  }
  
  return { token, verifyToken };
}

/**
 * Slackの署名を検証
 * @param {Object} e - リクエストオブジェクト
 * @return {boolean} 検証結果
 * @throws {Error} 検証処理に失敗した場合
 */
function verifySlackSignature(e) {
  const { verifyToken } = getSlackConfig();
  const timestamp = e.parameter.timestamp;
  const signature = e.parameter.signature;
  
  if (!timestamp || !signature) {
    throw new Error("Missing required parameters for signature verification");
  }
  
  // タイムスタンプの検証（5分以内）
  const now = Math.floor(Date.now() / 1000);
  if (Math.abs(now - timestamp) > 300) {
    throw new Error("Request timestamp is too old");
  }
  
  // 署名の検証
  const payload = `${timestamp}.${JSON.stringify(e.postData.contents)}`;
  const hmac = Utilities.computeHmacSha256Signature(payload, verifyToken);
  const computedSignature = Utilities.base64Encode(hmac);
  
  return signature === computedSignature;
}

/**
 * Slackにメッセージを送信
 * @param {string} channel - チャンネルID
 * @param {string} text - メッセージテキスト
 * @param {Object} [blocks] - ブロックキット（オプション）
 * @param {string} [threadTs] - スレッドのタイムスタンプ（オプション）
 * @throws {Error} メッセージ送信に失敗した場合
 */
function postMessage(channel, text, blocks = null, threadTs = null) {
  const { token } = getSlackConfig();
  const url = "https://slack.com/api/chat.postMessage";
  
  const payload = {
    channel: channel,
    text: text,
    thread_ts: threadTs
  };
  
  if (blocks) {
    payload.blocks = blocks;
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload)
  });
  
  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    throw new Error(`Slack API Error: ${result.error}`);
  }
  // ボットのメッセージもログに記録
  logMessageToSheet(
    "'" + (threadTs || result.ts),
    "'" + result.ts,
    "BOT", // ボットのユーザーIDとして"BOT"を使用
    text
  );
  return result;
}

/**
 * タスク情報の不足項目を確認し、質問メッセージを送信
 * @param {string} channel - チャンネルID
 * @param {string} threadTs - スレッドのタイムスタンプ
 * @param {Object} json - タスクJSON
 * @throws {Error} 不足項目チェックに失敗した場合
 */
function checkMissingFields(channel, threadTs, json) {
  const validation = callGeminiForValidation(json);

  if (validation.missing.length > 0 || validation.questions.length > 0) {
    const blocks = [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "以下の情報が不足しています："
        }
      }
    ];
    
    if (validation.missing.length > 0) {
      blocks.push({
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*不足項目：*\n${validation.missing.map(item => `• ${item}`).join("\n")}`
        }
      });
    }
    
    if (validation.questions.length > 0) {
      blocks.push({
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*追加情報が必要：*\n${validation.questions.map(q => `• ${q}`).join("\n")}`
        }
      });
    }
    
    postMessage(channel, "タスク情報の追加が必要です", blocks, threadTs);
    const formattedJson = "```json\n" + JSON.stringify(json, null, 2) + "\n```";
    postMessage(channel, formattedJson, null, threadTs);
  } else {
    postMessage(channel, "タスク情報に必要な情報が揃いました。", null, threadTs);
      
    // タスク情報を整形して表示
    const formattedJson = `以下のタスク情報を登録しますか？はい、OKなどで回答して下さい。
\`\`\`json
${JSON.stringify(json, null, 2)}
\`\`\`
`;
    postMessage(channel, formattedJson, null, threadTs);
      
    // タスク登録の確認状態を保存
    saveTaskConfirmation(threadTs, {
      status: "pending",
      json: json
    });
  }
  
}