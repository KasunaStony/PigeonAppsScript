function addProjectToTranslator(interpretor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var EngTrans = ss.getSheetByName("英翻");
  var arrangement = ss.getSheetByName("时间安排");
  var trans = interpretor.split(",");
  var row = 1;  
  var lastCol = EngTrans.getLastColumn();  
  var findTrans = false;
  for(var i = 0; i < trans.length; i++){
    findTrans = false;
    while(!EngTrans.getRange(row, 1).getValue().equals("")){
      if(EngTrans.getRange(row, 1).getValue().equals(trans[i])){ 
        EngTrans.getRange(row, lastCol).setValue(
          EngTrans.getRange(row, lastCol).getValue().equals("")?
          arrangement.getRange(5, 8).getValue():
          EngTrans.getRange(row, lastCol).getValue() + " " +arrangement.getRange(5, 8).getValue()); 
        findTrans = true;
        break;
      }
      row++;
    }  
    if(!findTrans){
      Browser.msgBox(trans[i] + ": 没有找到你哎_:(´ཀ`」 ∠):_ \\n\\n拜托看到这条消息去跟群管理讲一下~(｡・ω・｡)ﾉ♥\\n\\n如果不是英翻的话请无视这条消息~");
    }
  }
}

function addProjectToTimeLine(timeLine) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var time = ss.getSheetByName("时间轴");
  var arrangement = ss.getSheetByName("时间安排");
  var row = 1;
  var lastCol = time.getLastColumn();  
  var findTime = false;
    while(!time.getRange(row, 1).getValue().equals("")){
      if(time.getRange(row, 1).getValue().equals(timeLine)){ 
        time.getRange(row, lastCol).setValue(
          time.getRange(row, lastCol).getValue().equals("")?
          arrangement.getRange(5, 8).getValue():
          time.getRange(row, lastCol).getValue() + " " +arrangement.getRange(5, 8).getValue()); 
        findTime = true;
        break;
      }
      row++;
    }  
    if(!findTime){
      Browser.msgBox(timeLine + ": 没有找到你哎_:(´ཀ`」 ∠):_ \\n\\n拜托看到这条消息去跟群管理讲一下~(｡・ω・｡)ﾉ♥\\n\\n如果不是英翻的话请无视这条消息~");
    }
  }

function findPersonById(s, id) {
  lastRow = s.getLastRow();
  for(i = 1; i <= lastRow; ++i)
    if(s.getRange(i, 1).equals(id))
      return i;
  return 0;
}