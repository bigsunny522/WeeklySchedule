function copyDataToHomeSheet() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATABASE");
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HOME");
  var today = new Date();
  var nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // 来週の日付
  
  var data = sourceSheet.getRange("A:F").getValues(); // F列を含めてデータを取得
  
  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][0]);
    var tag = data[i][4];
    var content = data[i][3];
    var done = data[i][5]; // Doneの状態を確認

    // Doneが入力されている行はスキップ
    if (done === "Done") {
      continue;
    }

    // タグに応じて色を決定
    var color;
    if(tag == "予定"){
      color = "#ffcccc"; // 予定用の色
    }else if(tag == "課題"){
      color = "#ccffcc"; // 課題用の色
    }else if(tag == "趣味"){
      color = "#a6beed" //趣味用の色
    }else{
      color = "#ccccff"; // その他の場合の色
    }
    
    // 今日の予定の場合
    if (date.toDateString() === today.toDateString()) {
      var targetColumn;
      if (tag === "予定") {
        targetColumn = 17; // Q列
      } else if (tag === "課題") {
        targetColumn = 18; // R列
      } else {
        targetColumn = 19; // S列
      }
      var targetRow = findEmptyRow(targetSheet, targetColumn);
      targetSheet.getRange(targetRow+1, targetColumn).setValue(content);
      sourceSheet.getRange(i + 1, 6).setValue("Done"); // F列に特定の値を入れる
    }
    // 今週の予定の場合
    else if (date >= today && date < nextWeek) {
      var targetRow = findEmptyRow(targetSheet, 21,0); // U列
      targetSheet.getRange(targetRow, 21).setValue(date);
      targetSheet.getRange(targetRow, 22).setValue(content);
      sourceSheet.getRange(i + 1, 6).setValue("Done"); // F列に特定の値を入れる
      targetSheet.getRange(targetRow, 22).setBackground(color); // 背景色を変更する
      targetSheet.getRange("U3:V").sort({column: 21, ascending: true});
      if (tag === "課題") {
        var targetRow = findEmptyRow(targetSheet, 17,20); // Q列
        targetSheet.getRange(targetRow, 17).setValue(date);
        targetSheet.getRange(targetRow, 18).setValue(content);
      }
    }
    // 来週以降の予定の場合
    else if (date >= nextWeek) {
      var targetRow = findEmptyRow(targetSheet, 24,0); // X列
      targetSheet.getRange(targetRow, 24).setValue(date);
      targetSheet.getRange(targetRow, 25).setValue(content);
      sourceSheet.getRange(i + 1, 6).setValue("Done"); // F列に特定の値を入れる
      targetSheet.getRange(targetRow, 25).setBackground(color); // 背景色を変更する
      targetSheet.getRange("X3:Y").sort({column: 24, ascending: true});
      if (tag === "課題") {
        var targetRow = findEmptyRow(targetSheet, 17,20); // Q列
        targetSheet.getRange(targetRow, 17).setValue(date);
        targetSheet.getRange(targetRow, 18).setValue(content);
      }

    }
  }
}

// 指定した列から空白行を探す関数
function findEmptyRow(sheet, column,startfindrow) {
  var values = sheet.getRange(4, column, sheet.getLastRow() - 3).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[startfindrow+i][0] === "") {
      return startfindrow+i + 3;
    }
  }
  return sheet.getLastRow() + 1;
}