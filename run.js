function runExternalReport(today,  mention, slackChannelId, monthlySheetId, monthlyGid, weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl) {
  const monthlySummarycsvB64  = getCsvAsBase64(monthlySheetId, monthlyGid);
  const weeklySummarycsvB64  = getCsvAsBase64(weeklySheetId, weeklyGid);
  const dailySummarycsvB64  = getCsvAsBase64(dailySheetId, dailyGid);
  const mdText  = fetchNotionMarkdown(templateUrl);

  const report  = callGeminiInternalReport(monthlySummarycsvB64, weeklySummarycsvB64, dailySummarycsvB64, mdText, today);  
  Logger.log(report);
  const newPageId = writeMarkdownToNotion(outputUrl, report, "ExternalReport");
  const slackSummary = callo3InternalReportSlackSummary(report);
  const message = `${mention} 本日の定例用レポートです。 \nhttps://www.notion.so/${newPageId} をご確認下さい。\n■テンプレURL\n${templateUrl}`;
  //postToSlack(slackChannelId, message);
  return report;
}

function runInternalReport(today,  mention, slackChannelId, monthlySheetId, monthlyGid, weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl) {
  const monthlySummarycsvB64  = getCsvAsBase64(monthlySheetId, monthlyGid);
  const weeklySummarycsvB64  = getCsvAsBase64(weeklySheetId, weeklyGid);
  const dailySummarycsvB64  = getCsvAsBase64(dailySheetId, dailyGid);
  const mdText  = fetchNotionMarkdown(templateUrl);

  const report  = callGeminiInternalReport(monthlySummarycsvB64, weeklySummarycsvB64, dailySummarycsvB64, mdText, today);  
  Logger.log(report);
  const newPageId = writeMarkdownToNotion(outputUrl, report, "InternalReport");
  const slackSummary =callGeminiInternalReportSlackSummary(report);
  Logger.log(slackSummary);
  const message = `${mention} 本日のインターナルレポートです。 \nhttps://www.notion.so/${newPageId} をご確認下さい。\n${slackSummary}\n■テンプレURL\n${templateUrl}`;
  postToSlack(slackChannelId, message);
  return report;
}

function runSerchWordReport(slackChannelId, sheetId, gid, templateUrl, outputUrl) {
  const csvB64  = getCsvAsBase64(sheetId, gid);
  const mdText  = fetchNotionMarkdown(templateUrl);
  const report  = callGeminiSearchKeywordReport(csvB64, mdText);

  const newPageId = writeMarkdownToNotion(outputUrl, report, "SearchKeyword");
  postToSlack(slackChannelId, `本日の検索キーワードレポートを更新しました\nhttps://www.notion.so/${newPageId}`);


}
