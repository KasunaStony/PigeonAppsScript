function checkWeek(){
  if(!isInSameWeek()){
    assignNewTasks();
  }
}
function isInSameWeek() {
  var arrangement = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间安排");
  return arrangement.getRange(5, 7).getValue().equals(getCurrentWeek());
}

function assignNewTasks(){
  var taskPool = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
    lastRow=taskPool.getLastRow();
    noTask=true;
  for(i=1;i<=lastRow;++i){
    if(taskPool.getRange(i,1).getBackground().equals("#ffffff") && !taskPool.getRange(i,1).isBlank()){
      projectAutoAdd(i);
    }
    
  }
}
