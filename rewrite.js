//マクロ実行用ボタン作成
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: '実行する', functionName: 'fetchAndWrite'}
  ];
  spreadsheet.addMenu('回答を取得', menuItems);
}


function fetchAndWrite() {
  //書き込み先のシートを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  //アンケート回答をanswersに格納
  var sheetAns = SpreadsheetApp.getActive().getSheetByName('フォームの回答 1');
  var answers = sheetAns.getDataRange().getValues();
  
  //回答を格納
  var date = new Array();
  var name = new Array();
  var price = new Array();
  var w_t = new Array();
  var who = new Array(); 
  for(var i = 0 ; i < answers.length; i++) {
    if(answers[i][0] != '' && i > 0) {
      //空白除去などの成形する場合はここで。
      date.push(answers[i][0]);
      name.push(answers[i][1]);
      price.push(answers[i][2]);
      w_t.push(answers[i][3]);
      who.push(answers[i][4]);
    }
  }
  
  //書き込み先のセルを取得 lastRow = 書き込み先の最終行番号
  var lastRow = sheet.getRange("A1:A").getValues().filter(String).length + 1;
  for(var j = 0 ; j < name.length ; j++) {
    sheet.getRange(lastRow + j, 1).setValue(name[0 + j])  //A列の最後のセルに書き込み
    sheet.getRange(lastRow + j, 2).setValue(price[0 + j])  //B列の最後のセルに書き込み
    sheet.getRange(lastRow + j, 3).setValue(w_t[0 + j])  //C列の最後のセルに書き込み
    sheet.getRange(lastRow + j, 4).setValue(who[0 + j])  //D列の最後のセルに書き込み
    sheet.getRange(lastRow + j, 7).setValue(date[0 + j])  //G列の最後のセルに書き込み
  }
  
  //アンケート回答をクリア
  sheetAns.getDataRange().setValue('');
  
}
