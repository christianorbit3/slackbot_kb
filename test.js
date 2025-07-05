function testFunction() {
const today = new Date('2025-04-01T00:00:00+09:00'); // 時間を 00:00:00 に丸める
 Logger.log(getRemainingBizDays(new Date()));
}

function testSpackPost(){
  const message = "<@UNZ5061JM> <@U0871U28E5U> test test test";
  const slackChannelId = "C08F7E19CRE";
  postToSlack(slackChannelId, message);
}