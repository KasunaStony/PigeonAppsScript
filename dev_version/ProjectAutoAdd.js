function projectAutoAdd(row) {
  s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  //检查是不是所有人员齐全
  function check(x) {
    return !s.getRange(row, x).getValue().equals("");
  }
  if(check(6) && check(7) && check(9) && check(11) && check(13)) {
    //如果人员齐全，添加到时间安排中
    addNewWork(row);
    //修改接坑处颜色
    setRowColor(s, row);
    //隐藏接坑处项目
    s.hideRows(row);
 }
}

function getTimeDuration(str) {
  //从接坑处中获取时间安排信息
  time = str.split("-")[1];
  return time.slice(0,2) + ":" + time.slice(2,4);
}

function getNum(str) {
  return 'E' + (parseInt(str.slice(1,3), 10) + 1);
  //得到上一个坑的编号后自动加一并返回
}

function addNewWork(row) {
  row = row -1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var arrangement = ss.getSheetByName("时间安排"); //获取时间安排表格
  var taskPool = ss.getSheetByName("接坑处"); //获取接坑处表格
  var data = taskPool.getDataRange().getValues(); //将接坑处表格所有信息存入data数组
  arrangement.insertRowBefore(5); //在第五行之上插入新的一行
  var range = arrangement.getRange('H5:O5'); //选中新的一行的相应位置
  var interpretor = data[row][8]+(data[row][10].equals("/")?"":(","+data[row][10]))+(data[row][12].equals("/")?"":(","+data[row][12])); //获取两个翻译 生成字符串
  var taskNum = getNum(arrangement.getRange("H6").getValue()); //调用getNum 获取项目编号
  var values = [[ taskNum, data[row][0],data[row][1],data[row][13],getTimeDuration(data[row][7]),interpretor,data[row][5],data[row][6] ]];
  //将所有接坑处数据按顺序存入values
  range.setValues(values); //将values赋值到选中区域
  if(data[row][10]=='/'){
    arrangement.getRange('A5:B5').merge();
    arrangement.getRange('D5:E5').merge();
    arrangement.getRange("L5").setValue("/");
  }//检测是否是单人翻译项目 若是则合并单元格
  arrangement.getRange(5, 1, 1, 34).setBackgroundColor("#f4cccc");
  arrangement.getRange(5, 7).setBackgroundColor("#ffffff");//设置该行背景颜色
  //如果是在本周
  if(arrangement.getRange(6, 7).getValue().equals(getCurrentWeek())){
    //合并周次单元格
    mergeWeekInfo(arrangement);
  }else{//不是则切换下一周
    //当前时间
    arrangement.getRange(5, 7).setValue(getCurrentWeek());
    //加粗底部边缘
    arrangement.getRange(5, 1, 1, 34).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    addNewWeekInTranslatorPage("英翻");
    addNewWeekInTranslatorPage("时间轴");
  }
  addProjectToTranslator(interpretor);
  addProjectToTranslator(data[row][5]);
  addProjectToTimeLine(data[row][6]);
}

function setRowColor(s, row) {
  var range = s.getRange(row, 1, 1, 34);
  range.setBackgroundColor("#b6d7a8");
}

function mergeWeekInfo(sheet){
  var mergedCellNumber = 0;
  var rowNum = 7;
  while((sheet.getRange(rowNum, 7).getValue()).equals("")){
    mergedCellNumber++;
    rowNum++;
  }
  sheet.getRange('G5:G' + (mergedCellNumber + 6).toString()).merge();
}

function addNewWeekInTranslatorPage(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var lastCol = sheet.getLastColumn();
  if(sheet.getRange(1, lastCol).getValue().equals("")){
    sheet.getRange(1, lastCol).setValue(getCurrentWeekNoDates());
  }else{
    sheet.insertColumnsAfter(lastCol, 10);
    sheet.getRange(1, lastCol + 1).setValue(getCurrentWeekNoDates());
  }
  
}
