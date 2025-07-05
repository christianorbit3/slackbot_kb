/********************************************************************
 * Tasks
 ********************************************************************/
function taskRemindTrigger_Bandai(){
  try {
    const mockEvent1 = {
      thread_ts: "1234567890.123456",
      ts: "1234567890.123456",
      channel: "C08U08RM6UU",
      user: "UNZ5061JM",
      text: "萬代のタスク一覧を表示して"
    };
    getTasksSlackProcess(mockEvent1);
    return true;
  } catch (error) {
    Logger.log("テスト失敗: " + error.message);
    throw error;
  }
}

/********************************************************************
 * TEMPLATE
 ********************************************************************/
function templateProjectReport() {
  const sheetId="1fDttXcNIM4ItBiyRG4mAY_9oqxlfSkPDu5LpbF9o3TQ";
  const gid="402384898";
  const templateUrl="https://www.notion.so/template-YYYYMMDD_-MTG-1d8b821ba4fa8057aa7df3a1f5b13526";
  const outputUrl = "";
  runMonthlyReport(sheetId, gid, templateUrl, outputUrl);
}

function ingageProjectInternal() {
  const monthlySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const monthlyGid = "0";
  const weeklySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const weeklyGid = "1641990680";
  const dailySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const dailyGid = "1721929597";
  const templateUrl="https://www.notion.so/ingageInternalAdReportTemplate-1e6b821ba4fa80588edcdb602f9a4e65";
  const outputUrl = "https://www.notion.so/ingageMonthlyAdReportOutput-1e6b821ba4fa80bc9961fb903d644017";
  const today =  makeJstDate(2025, 4, 30);

  runInternalReport(today, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

/********************************************************************
 * ASUENE
 ********************************************************************/
function projectInternalReport_Asuene() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C07JQ3MRNBB";
  const monthlySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const monthlyGid = "0";
  const weeklySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const weeklyGid = "765352429";
  const dailySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const dailyGid = "154818473";
  const templateUrl="https://www.notion.so/Asuene-AdReport_InternalTemplate-1e7b821ba4fa80eda72ff86b58327aa3";
  const outputUrl = "https://www.notion.so/Asuene-AdReport_Output-1e7b821ba4fa807f862ff8fb6712e0d1";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

function projectExternalReport_Asuene() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C07JQ3MRNBB";
  const monthlySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const monthlyGid = "0";
  const weeklySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const weeklyGid = "765352429";
  const dailySheetId = "1Axn886sHKOBm9Tf93B5eUElw2PYF01UxUQwI4CuIYkU";
  const dailyGid = "154818473";
  const templateUrl="https://www.notion.so/Asuene-AdReport_ExternalTemplate-1fbb821ba4fa80d789d9e9b4d28e2ebb";
  const outputUrl = "https://www.notion.so/Asuene-AdReport_Output-1e7b821ba4fa807f862ff8fb6712e0d1";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runExternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

/********************************************************************
 * INGAGE
 ********************************************************************/
function projectInternalReport_Ingage() {
  const mention = "<@U05N84AQ6R2> <@UNZ5061JM>";
  const slackChannelId = "C07RPAX70JD";
  const monthlySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const monthlyGid = "0";
  const weeklySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const weeklyGid = "1641990680";
  const dailySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const dailyGid = "1721929597";
  const templateUrl="https://www.notion.so/Ingage-AdReport_InternalTemplate-1e6b821ba4fa80588edcdb602f9a4e65";
  const outputUrl = "https://www.notion.so/Ingage-AdReport_Output-1e6b821ba4fa80bc9961fb903d644017";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

function projectExternalReport_Ingage() {
  const mention = "<@U05N84AQ6R2> <@UNZ5061JM>";
  const slackChannelId = "C07RPAX70JD";
  const monthlySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const monthlyGid = "0";
  const weeklySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const weeklyGid = "1641990680";
  const dailySheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const dailyGid = "1721929597";
  const templateUrl="https://www.notion.so/Ingage-AdReport_ExternalTemplate-1e6b821ba4fa80ceb021cc47c391e58a";
  const outputUrl = "https://www.notion.so/Ingage-AdReport_Output-1e6b821ba4fa80bc9961fb903d644017";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runExternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

function searchKeywordReport_Ingage() {
  const slackChannelId = "C07RPAX70JD";
  const sheetId = "18A1w2fhWw0CQwBmQtVeIg-5QWpMIL07Gap9Ddbu3jHw";
  const gid = "1229389282";
  const templateUrl="https://www.notion.so/template_ingage_search-md-1e7b821ba4fa805ca468ce608e087e26";
  const outputUrl = "https://www.notion.so/Ingage-AdReport_Output-1e6b821ba4fa80bc9961fb903d644017";
  runSerchWordReport(slackChannelId, sheetId, gid, templateUrl, outputUrl)
}


/********************************************************************
 * CAM
 ********************************************************************/
 function projectInternalReport_Cam() {
  const mention = "<@U08L19046DR> <@UNZ5061JM>";
  const slackChannelId = "C07MRV4LMBM";
  const monthlySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const monthlyGid = "1404504805";
  const weeklySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const weeklyGid = "2053942090";
  const dailySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const dailyGid = "2019164879";
  const templateUrl="https://www.notion.so/Cam-AdReport_InternalTemplate-1f3b821ba4fa8080bef8f52c30608a17";
  const outputUrl = "https://www.notion.so/Cam-AdReport_Output-1f3b821ba4fa80688777cc7f54f251bd";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

function projectExternalReport_Cam() {
  const mention = "<> <@UNZ5061JM>";
  const slackChannelId = "C07MRV4LMBM";
  const monthlySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const monthlyGid = "1404504805";
  const weeklySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const weeklyGid = "2053942090";
  const dailySheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74";
  const dailyGid = "2019164879";
  const templateUrl= "https://www.notion.so/Cam-AdReport_ExternalTemplate-1f3b821ba4fa8099ad15ca0d2c654bf6";
  const outputUrl = "https://www.notion.so/Cam-AdReport_Output-1f3b821ba4fa80688777cc7f54f251bd";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runExternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

/********************************************************************
 * TONE
 ********************************************************************/
 function toneProjectReport() {
  const sheetId = "1GOG4QLEVGhx7ycFlKchFIi23U-Y8uHA9DcwB3QJhkmw";
  const gid = "1128847627";
  const sheetId2 = "1GOG4QLEVGhx7ycFlKchFIi23U-Y8uHA9DcwB3QJhkmw";
  const gid2 = "693866427";
  const templateUrl="https://www.notion.so/toneWeeklyAdReportTemplate-1e3b821ba4fa80e49e0ec626ec73d625";
  const outputUrl = "https://www.notion.so/toneAdReportOutput-1e3b821ba4fa802d911fcfbbb4d3b9aa";
  runInternalReport(sheetId, gid,sheetId2, gid2, templateUrl, outputUrl);
}

function projectInternalReport_Tone() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C08F7E19CRE";
  const monthlySheetId = "1GOG4QLEVGhx7ycFlKchFIi23U-Y8uHA9DcwB3QJhkmw";
  const monthlyGid = "1128847627";
  const weeklySheetId = "1GOG4QLEVGhx7ycFlKchFIi23U-Y8uHA9DcwB3QJhkmw";
  const weeklyGid = "693866427";
  const dailySheetId = "1GOG4QLEVGhx7ycFlKchFIi23U-Y8uHA9DcwB3QJhkmw";
  const dailyGid = "736515481";
  const templateUrl="https://www.notion.so/Tone-AdReportTemplate_Internal-1e3b821ba4fa80e49e0ec626ec73d625";
  const outputUrl = "https://www.notion.so/Tone-AdReportTemplate_Output-1e3b821ba4fa802d911fcfbbb4d3b9aa";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

/********************************************************************
 * MANNEN
 ********************************************************************/
 function projectInternalReport_Mannen() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C0810P2NDQE";
  const monthlySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const monthlyGid = "0";
  const weeklySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const weeklyGid = "765352429";
  const dailySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const dailyGid = "154818473";
  const templateUrl="https://www.notion.so/AdReportTemplate_Internal-1f4b821ba4fa8075ab4ecbaecf094667";
  const outputUrl = "https://www.notion.so/AdReportTemplate_Output-1f4b821ba4fa803b9e47d4da0b73595e";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

function projectExternalReport_Mannen() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C0810P2NDQE";
  const monthlySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const monthlyGid = "0";
  const weeklySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const weeklyGid = "765352429";
  const dailySheetId = "1cb4iK4wnKYAfCOrRjfES7KoquVabvP50IkzJQf45iCs";
  const dailyGid = "154818473";
  const templateUrl="https://www.notion.so/AdReportTemplate_Internal-1f4b821ba4fa8075ab4ecbaecf094667";
  const outputUrl = "hhttps://www.notion.so/AdReportTemplate_Output-1f4b821ba4fa803b9e47d4da0b73595e";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runExternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}

/********************************************************************
 * BGRASS
 ********************************************************************/
 function projectInternalReport_Bgrass() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C07LBLDAJNR";
  const monthlySheetId = "14Qt0cC23cVB9iCzq_nMDJu_64bqUlk27C1lUCgVN4V0";
  const monthlyGid = "1772394036";
  const weeklySheetId = "14Qt0cC23cVB9iCzq_nMDJu_64bqUlk27C1lUCgVN4V0";
  const weeklyGid = "1355736120";
  const dailySheetId = "14Qt0cC23cVB9iCzq_nMDJu_64bqUlk27C1lUCgVN4V0";
  const dailyGid = "710302977";
  const templateUrl="https://www.notion.so/WAKECareer-AdReportTemplate_Internal-21ab821ba4fa80d3b6e3c4e3c86211be";
  const outputUrl = "https://www.notion.so/WAKECareer-AdReport_Output-21ab821ba4fa80d6b0c4d63e9f8eb40d";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}


/********************************************************************
 * XTalent
 ********************************************************************/
 function projectInternalReport_XTalent() {
  const mention = "<@U0871U28E5U> <@UNZ5061JM>";
  const slackChannelId = "C08EWLG50V7";
  const monthlySheetId = "1y7WLk1XsrSnmv5YJLcmTcP-28ccT5P7g9FOVeKYMDfI";
  const monthlyGid = "1052678223";
  const weeklySheetId = "1y7WLk1XsrSnmv5YJLcmTcP-28ccT5P7g9FOVeKYMDfI";
  const weeklyGid = "717535607";
  const dailySheetId = "1y7WLk1XsrSnmv5YJLcmTcP-28ccT5P7g9FOVeKYMDfI";
  const dailyGid = "1639556954";
  const templateUrl="https://www.notion.so/Xtalent-AdReportTemplate_Internal-200b821ba4fa800a88ecf645f321bd69";
  const outputUrl = "https://www.notion.so/XTalent-AdReport_Output-21ab821ba4fa80428b33d863bc146219";
  //const today =  makeJstDate(2025, 4, 30);
  const today = new Date();
  runInternalReport(today, mention, slackChannelId, monthlySheetId, monthlyGid,weeklySheetId, weeklyGid, dailySheetId, dailyGid, templateUrl, outputUrl);
}