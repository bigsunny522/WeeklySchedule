function showDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('dialog')
      .setWidth(430)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '登録フォーム');
}

function registerData(dateString, startTimeString, endTimeString, content, tag) {
  // 日付と時間の文字列を解析して日付オブジェクトに変換
  var dateParts = dateString.split('-');
  var startTimeParts = startTimeString.split(':');
  var endTimeParts = endTimeString.split(':');

  // Dateオブジェクトの月は0から始まるため、月の値を1減らす
  var date = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  var startTime = new Date('1970-01-01T' + startTimeParts[0] + ':' + startTimeParts[1]);
  var endTime = new Date('1970-01-01T' + endTimeParts[0] + ':' + endTimeParts[1]);

  // 新しいデータを追加
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATABASE");
  var newRow = [date, startTime, endTime, content, tag];
  sheet.appendRow(newRow);
  Logger.log(date+ startTime+ endTime+ content+ tag)

  // 日付を基準にして A 列を昇順に並び替え
  sheet.getRange("A2:E").sort({column: 1, ascending: true});
  copyDataToHomeSheet();
  generateWeeklySchedule();
}

