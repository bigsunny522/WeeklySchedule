function resetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var databaseSheet = ss.getSheetByName("DATABASE");
  var homeSheet = ss.getSheetByName("HOME");

  // HOMEシートのセルを空にして、セルの色を白にする
  homeSheet.getRange('C4:O39').clear({ contentsOnly: true, formatOnly: true });
  homeSheet.getRangeList(['Q4:S100','U3:V100','X3:Y100']).clear({ contentsOnly: true, formatOnly: true });
  homeSheet.getRangeList(['B2:O39','C4:O39','B2:B3']).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  //homeSheet.getRangeList(['C2:C39','E2:E39','G2:G39','I2:I39','K2:K39','M2:M39']).setBorder(true, true, true, true, false, false, "glay", SpreadsheetApp.BorderStyle.SOLID);

  // DATABASEシートのF列の中身を空にする
  databaseSheet.getRange('F:F').clear({ contentsOnly: true, formatOnly: true });

  // DATABASEシートのA列の日付を確認して、今日よりも1週間以上前のものはAからF列まで削除して全体を上に上げる
  var today = new Date();
  var lastRow = databaseSheet.getLastRow();
  var data = databaseSheet.getRange('A2:A' + lastRow).getValues();
  for (var i = 0; i < data.length; i++) {
    var date = new Date(data[i][0]);
    if (date < today - 7 * 24 * 60 * 60 * 1000) { // 1週間以上前の日付かどうかをチェック
      databaseSheet.deleteRow(i + 2); // 行番号は1-indexedなので+2する
    }
  }

  copyDataToHomeSheet();
  generateSchedule();
  generateWeeklySchedule()
}

