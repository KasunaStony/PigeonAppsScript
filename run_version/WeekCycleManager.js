/*
用于获得周次的字符串
格式为 【年份后两位 月份缩写 周次\n(日期)】【19 Mar. W2\n(8-15)】
*/
function getCurrentWeek() {
  //获得当前时间
  var current = new Date();
  //年份
  var year = current.getUTCFullYear().toString().slice(2,4);
  //月份缩写
  var month = getMouthAbbreviation(current.getUTCMonth());
  //周次与日期范围
  var dateRange = getDateRange(current.getUTCFullYear(), current.getUTCMonth(), current.getUTCDate());
  
  return year + " " + month + " " + dateRange;
}
function getCurrentWeekNoDates(){
  //获得当前时间
  var current = new Date();
  //年份
  var year = current.getUTCFullYear().toString().slice(2,4);
  //月份缩写
  var month = getMouthAbbreviation(current.getUTCMonth());
  //周次与日期范围
  var dateRange = getDateRange(current.getUTCFullYear(), current.getUTCMonth(), current.getUTCDate());
  return year.toString() + month + dateRange.toString().slice(0,2);
}
/*
月份缩写
*/
function getMouthAbbreviation(month){
  
  switch(month){
    case 0:
      return "Jan.";
    case 1:
      return "Feb.";
    case 2:
      return "Mar.";
    case 3:
      return "Apr.";
    case 4:
      return "May";
    case 5:
      return "Jun.";
    case 6:
      return "Jul.";
    case 7:
      return "Aug.";
    case 8:
      return "Sep.";
    case 9:
      return "Oct.";
    case 10:
      return "Nov.";
    case 11:
      return "Dec.";          
  }
}

/*
平年闰年
*/
function isLeapYear(year){
  return ((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0);
}

function getLastDay(){
  var current = new Date();
  var year = current.getUTCFullYear();
  var month = current.getUTCMonth();
  var date = current.getUTCDate();
  var i;
  //日期数据
  var data;
  //起始日期
  var firstDate;
  //结束日期
  var lastDate;
  //周次
  var weekNumber = "W";
  //区分平闰年
  if(isLeapYear(year)){
    data = getLeapYear();
  }else{
    data = getNormalYear();
  }
  //循环查现在日期的区间
  for(i = 0; i < 4; i++){
    //如果在第一个区间
    if(i == 0 && date <= data[month][i]){
      //s.getRange(3, i + 1).setNote("in first if ");
      firstDate = 1;
      lastDate = data[month][i];
      weekNumber += (i+1).toString();
      break;
    }
    //另外的区间
    else if(date <= data[month][i]){
      firstDate = data[month][i-1] + 1;
      lastDate = data[month][i];
      weekNumber += (i+1).toString();
      break;
    }
  }
  
  return lastDate;
}
/*
周次与日期范围
*/
function getDateRange(year, month, date){
  var i;
  //日期数据
  var data;
  //起始日期
  var firstDate;
  //结束日期
  var lastDate;
  //周次
  var weekNumber = "W";
  //区分平闰年
  if(isLeapYear(year)){
    data = getLeapYear();
  }else{
    data = getNormalYear();
  }
  //循环查现在日期的区间
  for(i = 0; i < 4; i++){
    //如果在第一个区间
    if(i == 0 && date <= data[month][i]){
      //s.getRange(3, i + 1).setNote("in first if ");
      firstDate = 1;
      lastDate = data[month][i];
      weekNumber += (i+1).toString();
      break;
    }
    //另外的区间
    else if(date <= data[month][i]){
      firstDate = data[month][i-1] + 1;
      lastDate = data[month][i];
      weekNumber += (i+1).toString();
      break;
    }
  }
  
  return weekNumber + String.fromCharCode(10) +"(" +  (firstDate + "-" + lastDate) + ")";
}

/*
平年
*/
function getNormalYear(){
  var normalYear = [
    [7, 15, 23, 31],
    [7, 14, 21, 28],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31]
  ];
  return normalYear;
}
/*
闰年
*/
function getLeapYear(){
  var leapYear = [
    [7, 15, 23, 31],
    [7, 14, 21, 29],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31],
    [7, 14, 22, 30],
    [7, 15, 23, 31]
  ];
  return leapYear;
}

