function generateWeeklySchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var databaseSheet = ss.getSheetByName("DATABASE");
  var homeSheet = ss.getSheetByName("HOME");

  var currentDate = new Date(); // 今日の日付を取得

  var daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"]; // 曜日のリスト

  var currentDayOfWeek = daysOfWeek[currentDate.getDay()]; // 今日の曜日を取得

  var homeDaysOfWeek = homeSheet.getRange(2, 3, 1, 13).getValues()[0]; // HOMEシートの2行目のCからO列までの曜日のリストを取得  

  if (startColumnIndex !== -1) { // "月"が見つかった場合の処理
    for (var i=0; i<=5 ; i++) { // 連続した5つの列（K列からO列）を処理する
      var dayofweek = databaseSheet.getRange(1,11+i).getValue(); 
      console.log(dayofweek)
      var startColumnIndex = homeDaysOfWeek.indexOf(dayofweek); // homeDaysOfWeekリストから"月"のインデックスを取得
      
      for (var j=2; j<=15; j++) {
        var column = startColumnIndex + 3; // HOMEシートの3列目から曜日の列が開始するため、インデックスに+3する

        var schedule = databaseSheet.getRange(j, 11+i).getValue(); // DATABASEシートのK列から順に予定を取得

        if (schedule) { // 予定が存在する場合の処理
          var startTime = databaseSheet.getRange(j, 9).getValue(); // DATABASEシートのI列から順に開始時間を取得

          var endTime = databaseSheet.getRange(j, 10).getValue(); // DATABASEシートのJ列から順に終了時間を取得
          
          var startRow = getRowForTime(startTime); // 開始時間に対応する行を取得
          var endRow = getRowForTime(endTime); // 終了時間に対応する行を取得（開始時間と同じ行）

          // セルの背景色を設定（内容が書かれているセルと同じ色、白色の場合は灰色に設定）
          var cellColor = databaseSheet.getRange(j, 11+i).getBackground();
          Logger.log(cellColor)
          if (cellColor === "#ffffff") {
            var color="#cccccc";
          } else {
            var color=cellColor;
          }

          // データを書き込む
          homeSheet.getRange(startRow, column).setValue(schedule);
          for (var k = startRow; k <= endRow; k++) {
                var cell = homeSheet.getRange(k, column);
                cell.setBackground(color);
          }
        }
      }
    }
  }
}