function memberWorkDone(e) {
  var range = e.range;
  var ss = range.getSheet(), row = range.getRow();
  function get(x) {
    return ss.getRange(row, x).getValue();
  }
  function check(x) {
    return ss.getRange(row, x).getValue().equals("C");
  }
  if(ss.getSheetName().equals("时间安排")) {
    if((check(1) && check(2) && get(13).split(",").length == 2)||
      (check(1) && get(13).split(",").length == 1) || 
      (check(1) && check(2) && check(3) && get(13).split(",").length == 2)) {
        if((check(4) && check(5)) || (check(4) && get(13).split(",").length == 1)) {
          if(check(6))
            ss.getRange(row, 1, 1, 34).setBackgroundColor("#d9ead3");
          else
            ss.getRange(row, 1, 1, 34).setBackgroundColor("#fff2cc");
        }
        else
          ss.getRange(row, 1, 1, 34).setBackgroundColor("#fce5cd");
      }
    else
      ss.getRange(row, 1, 1, 34).setBackgroundColor("#f4cccc");
    ss.getRange(row, 7).setBackgroundColor("#ffffff");
  }
}


