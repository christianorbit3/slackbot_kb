/**
 * OpenAI o3-2025-04-16 に投げて Markdown を受け取る
 */
function callO3MonthlyReport(csvBinB64, mdTemplate) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey || !apiKey.startsWith('sk-')) throw new Error('OPENAI_API_KEY が未設定 or 不正');

  const messages = [
    {
      role: 'system',
      content: 'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい'
    },
    {
      role: 'user',
      content: [
        '私は広告代理店のプロジェクトマネージャーです。クライアントへの月次報告資料を作成しており、いまから共有する月次サマリーデータを、提示するテンプレートに沿った形式でまとめてください。表はcategoryごとに行出力するようにお願いします。途中で質問することなく、テンプレートの指示に従ってマークダウンファイルで出力してください。できるだけ深い考察をお願いします。また、分析を開始する前に、これまでの同様の分析指示の結果については、混同しないようになるべく忘れてください。',
        '```',
        mdTemplate.trim(),
        '```',
        '報告書のテンプレートは以上です。以下が、あなたが推論すべきデータです。月次推移データとなります。'
      ].join('\n')
    }
  ];

  // CSV が長い場合は 20 000 文字ごとにチャンクして別メッセージで渡す
  for (let i = 0, chunks = chunkString(base64ToUtf8(csvBinB64), 20000); i < chunks.length; ++i) {
    messages.push({
      role: 'user',
      content: [
        `### monthlySummary.csv — part ${i + 1}`,
        '```',
        chunks[i],
        '```'
      ].join('\n')
    });
  }
  Logger.log(messages)

  const payload = {
    model:    'o3-pro-2025-06-10',
    max_completion_tokens:30000,
    messages: messages
  };

  const options = {
    method : 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 180000       // 180 s
  };

  const resp     = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const respJson = JSON.parse(resp.getContentText());

  try {
    return respJson.choices[0].message.content;
  } catch (e) {
    throw new Error('OpenAI API error → ' + resp.getContentText());
  }
}
function callO3InternalReport(monthlySummaryCsvBinB64,weeklySummaryCsvBin2B64, dailySummaryCsvBinB64, mdTemplate, today) {
  const todayString = getTodayJSTString(today);
  const bizDaysThisMonth = getBizDaysThisMonth(today)
  const remainingBizDays = getRemainingBizDays(today);
  const usedBusinessDayProrated = (bizDaysThisMonth-remainingBizDays) / bizDaysThisMonth;
  const thisMonth = thisMonthStartJST(today);
  const oneMonthAgo = lastMonthStartJST(today);
  const twoMonthAgo = twoMonthsAgoStartJST(today);
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey || !apiKey.startsWith('sk-')) throw new Error('OPENAI_API_KEY が未設定 or 不正');

  const messages = [
    {
      role: 'system',
      content: 'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい'
    },
    {
      role: 'user',
      content: [
        '私は広告代理店のプロジェクトマネージャーです。',
        'クライアントへの月次報告資料を作成しており、いまから共有する月次サマリーデータを、提示するテンプレートに沿った形式でまとめてください。',
        '表はcategoryごとに行出力するようにお願いします。途中で質問することなく、テンプレートの指示に従ってマークダウン形式のテキストで出力してください。マークダウンとして正しく出力することに注意して下さい。ただし、「```markdown」[####(見出し4)]は使わないで下さい。',
        'できるだけ深い考察をお願いします。また、分析を開始する前に、これまでの同様の分析指示の結果については、混同しないようになるべく忘れてください。',
        `また、テンプレートの中の変数として、下記変数を利用して下さい`,
        `- {{today}}=${todayString}`,
        `- {{usedBusinessDayProrated}}=${usedBusinessDayProrated}`,
        `- {{remainingBizDays}}=${remainingBizDays}`,
        `- {{bizDaysThisMonth}}=${bizDaysThisMonth}`,
        `- {{thisMonth}}=${thisMonth}`,
        `- {{oneMonthAgo}}=${oneMonthAgo}`,
        `- {{twoMonthAgo}}=${twoMonthAgo}`,
        '```',
        mdTemplate.trim(),
        '```',
        '報告書のテンプレートは以上です。',
        'あなたに提供するデータは2つで、montlySummaryは月次の集計データ、weeklySummaryは週次の集計データとなります。テンプレートの中で適切に使い分けて下さい。'
      ].join('\n')
    }
  ];

  // CSV が長い場合は 20 000 文字ごとにチャンクして別メッセージで渡す
  for (let i = 0, chunks = chunkString(base64ToUtf8(monthlySummaryCsvBinB64), 20000); i < chunks.length; ++i) {
    messages.push({
      role: 'user',
      content: [
        `### monthlySummary.csv — part ${i + 1}`,
        '```',
        chunks[i],
        '```'
      ].join('\n')
    });
  }
  for (let i = 0, chunks = chunkString(base64ToUtf8(weeklySummaryCsvBin2B64), 20000); i < chunks.length; ++i) {
    messages.push({
      role: 'user',
      content: [
        `### weeklySummary.csv — part ${i + 1}`,
        '```',
        chunks[i],
        '```'
      ].join('\n')
    });
  }
  for (let i = 0, chunks = chunkString(base64ToUtf8(dailySummaryCsvBinB64), 20000); i < chunks.length; ++i) {
    messages.push({
      role: 'user',
      content: [
        `### dailySummary.csv — part ${i + 1}`,
        '```',
        chunks[i],
        '```'
      ].join('\n')
    });
  }

  Logger.log(messages)

  const payload = {
    model:    'o3-pro-2025-06-10',
    max_output_tokens:100000,
    input: messages
  };

  const options = {
    method : 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 180000       // 180 s
  };

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/responses', options);
  const respJson = JSON.parse(resp.getContentText());
  Logger.log(respJson); 

  try {
    return respJson.output[0].content[0].text;
  } catch (e) {
    throw new Error('OpenAI API error → ' + resp.getContentText());
  }
}

function callo3InternalReportSlackSummary(report) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey || !apiKey.startsWith('sk-')) throw new Error('OPENAI_API_KEY が未設定 or 不正');

  const messages = [
    {
      role: 'system',
      content: 'あなたは広告代理店で働くデータアナリストです。何よりもデータを正確に処理することを優先し、プロジェクトマネージャーの意見にも従いながら統計的に正しい推論を述べて下さい'
    },
    {
      role: 'user',
      content: [
        '下記のレポートから一部分のテキストを正確に抜き出して下さい。改行などはそのままで忠実に抜き出して下さい',
        '# 抜き出してほしい箇所',
        '「Slackポスト内容」という見出しから、次の見出しまでのすべてのテキストを抽出して下さい。',
        '以下、抽出対象のレポートです。',
        '```',
        report.trim(),
        '```'
      ].join('\n')
    }
  ];

  Logger.log(messages)

  const payload = {
    model:    'o3-pro-2025-06-10',
    max_completion_tokens:30000,
    messages: messages
  };

  const options = {
    method : 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 180000       // 180 s
  };

  const resp     = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const body = resp.getContentText();
  
  
  try {
    const respJson = JSON.parse(body);
    return respJson.choices[0].message.content;
  } catch (e) {
    return "slackに投稿する内容の抽出に失敗しました"
    //throw new Error('OpenAI API error → ' + resp.getContentText());
  }
}