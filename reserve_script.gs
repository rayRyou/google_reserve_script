var calId = "bu6ddoalk3ghqj7dmmbr288jno@group.calendar.google.com";

function getSchedulesByCalendar(e) {
  try {
    var form = FormApp.getActiveForm();
    var newItem = e.response.getItemResponses();
    var canReserve = false;
    var reserveDate = null;
    var mailAddress = e.response.getRespondentEmail();
    var name = null;
    var sumMinute = 0;
    var menus = ""
    for (var cnt = 0; cnt < newItem.length; cnt++ ) {
      var question = newItem[cnt].getItem();
      var title = question.getTitle();
      if (title == "予約日") { // 予約日の質問タイトルと同じにする
        reserveDate = newItem[cnt].getResponse();
      }else if (title == "メニュー") { // メニューの質問タイトルと同じにする
        var resAry = newItem[cnt].getResponse();
        for (var resCnt = 0; resCnt < resAry.length; resCnt++) {
          var res = resAry[resCnt];
          let texts = res.split(" (所要時間")
          var minStr = texts[1];
          minStr = minStr.split("分")[0]
          let min = Number(minStr);
          sumMinute += min
          if (menus.length != 0) {
            menus += " + "
          }
          menus += texts[0]
        }
        Logger.log(sumMinute.toString())
      }else if (title == "名前"){
        name = newItem[cnt].getResponse();
      }
    }
    if (sumMinute != 0 && reserveDate != null) {
      var startDate = new Date(reserveDate)
      var endDate = new Date(reserveDate)
      endDate.setMinutes(endDate.getMinutes() + sumMinute)
      canReserve = checkScheduleCheck(startDate, endDate)
    }

    // add to calendar
    // mail
    var subject = null
    var body = null
    if (canReserve){
      let startDate = new Date(reserveDate)
      var endDate = new Date(reserveDate)
      endDate.setMinutes(endDate.getMinutes() + sumMinute)

      addNewEventToCalendar(name, menus, mailAddress, startDate, endDate)

      // 成功メール
      subject = "【予約完了】申し込みされた時間で、予約できました"
      body = `${name}様\n\nご登録いただいた内容で承りました。\n予約日:${reserveDate}`
    }else{
      // 失敗メール
      subject = "【予約失敗】申し込みされた時間では、予約できませんでした"
      body = `${name}様\n申し訳ございません。\n登録していただいたお時間では予約枠がありませんでした。\nカレンダーに空きがあった場合でも他の方が先に予約してしまうことがございます。\n\nお手数ですが、再度カレンダーを確認して、ご登録をお願いいたします。\n\nカレンダー : https://calendar.google.com/calendar/u/0?cid=YnU2ZGRvYWxrM2docWo3ZG1tYnIyODhqbm9AZ3JvdXAuY2FsZW5kYXIuZ29vZ2xlLmNvbQ`
    }
    if (mailAddress != null) {
      Logger.log(mailAddress)
      const options = {
        name: "自動予約bot"
      }
      GmailApp.sendEmail(mailAddress, subject, body, options)
    }else{
      Logger.log(mailAddress)
    }
  }catch(e){
    Logger.log(e)
  }
}

function checkScheduleCheck(startDate, endDate) {
  var result = false;
  try {
    const calendar = CalendarApp.getCalendarById(calId);
    const start = new Date(startDate)
    const end = new Date(endDate)
    const events = calendar.getEvents(start, end)
    if (events.length > 0) {
      result = false
    }else{
      result = true
    }
  }catch (e) {
    console.log(e);
  }
  return result;
}
function addNewEventToCalendar(name, menu, address, startDate, endDate){

  try {
    var title = `${name}様 ${menu}`
    const start = new Date(startDate)
    const end = new Date(endDate)
    const option = {
      description: `${name}様\n\n申し込みメニュー:${menu}`,
      location: '〒277-0011 千葉県柏市東上町2-28 第一水戸屋ビル 店舗棟 3F "R2"',
    }
    const calendar = CalendarApp.getCalendarById(calId);
    event = calendar.createEvent(title, start, end, option)
    event.setVisibility(CalendarApp.Visibility.PRIVATE)
  }catch (e) {
    console.log(e);
  }
  

} 