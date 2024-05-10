function generateSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var databaseSheet = ss.getSheetByName("DATABASE");
  var homeSheet = ss.getSheetByName("HOME");

  // 今日の日付を取得し、7日後の日付を取得
  var today = new Date();
  today.setDate(today.getDate() -1);
  var nextWeek = new Date();
  nextWeek.setDate(today.getDate() + 7);

  // データベースから日付、開始時間、終了時間、タグの列を取得
  var dataRange = databaseSheet.getDataRange();
  var dataValues = dataRange.getValues();

  for (var i = 1; i < dataValues.length; i++) { // ヘッダー行をスキップする
    var rowData = dataValues[i];
    var dateText = rowData[0];
    
    // A列の日付のセルが空白の場合、処理をスキップ
    if (!dateText) {
      continue;
    }

    var startDate = new Date(dateText);
    // 今日から7日以内のデータのみ処理
    if (startDate >= today && startDate <= nextWeek){
      var startTime = rowData[1];
      var endTime = rowData[2];
      var content = rowData[3];
      var tag = rowData[4];

      // 開始時間と終了時間が設定されている場合のみ処理を続行
      if (startTime && endTime) {
          var startRow = getRowForTime(startTime);
          if(endTime.getMinutes()==0){
            var endRow = getRowForTime(endTime)-1;
          }else{
            var endRow = getRowForTime(endTime);
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
            homeSheet.getRange(startRow,getColumnForDate(startDate)).setValue(content);

          // 開始時間と終了時間で特定されたセルの背景色を変更し、タグをコピーする
          for (var j = startRow; j <= endRow; j++) {
              var cell = homeSheet.getRange(j, getColumnForDate(startDate));
              cell.setBackground(color);
          }
      } else {
          Logger.log("Error: Start Time or End Time is missing."); // 開始時間または終了時間がない場合のエラーログ
      }
    }
  }
}

// 日付を指定のフォーマットに変換する関数
function getRowForTime(time) {
  // 時刻が文字列でない場合、文字列に変換する
  if (typeof time !== 'string') {
    time = time.toString();
  }

  var hours = parseInt(time.split(":")[0]); // 時間部分を取得
  var minutes = parseInt(time.split(":")[1]); // 分部分を取得
  var row = 4 + (hours - 6) * 2; // 6:00からの経過時間に応じて行を計算

  // 分が30分以上なら1行下に移動
  if (minutes >= 30) {
    row++;
  }
  return row;
}

// 時間に対応する行を取得する関数
function getRowForTime(time) {
  var hours = time.getHours();
  var minutes = time.getMinutes();
  var row = 4 + (hours - 6) * 2; // 6:00からの経過時間に応じて行を計算
  // 分が30分以上なら1行下に移動
  if (minutes >= 30) {
    row++;
  }
  return row;
}

// 日付から列を取得する関数
function getColumnForDate(date) {
  var today = new Date();
  var differenceInDays = Math.ceil((date - today) / (1000 * 60 * 60 * 24));
  var column = differenceInDays * 2 + 3; // 今日の列は3列目から始まるため

  // 負の値や非整数の場合は3列目（今日の列）に設定する
  if (column < 3 || !Number.isInteger(column)) {
    column = 3;
  }
  return column;
}

