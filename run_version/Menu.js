function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('字幕组')
    .addItem('登记', 'register')
    .addItem("挖坑", "newTask")
    .addItem("接坑", "takeTask")
    .addItem("手动排坑", "manuallySet")
    .addItem("神话语录", "mythQuote")
    .addToUi();
}

function register() {
  email = Session.getActiveUser().getEmail();
  PropertiesService.getUserProperties().setProperty("e-mail", email);
  Browser.msgBox("Welcome, " + email);
}

function newTask() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('NewTask'), "在？为什么挖坑？");
}
function manuallySet() {
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处"), false);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile('ManuallyPage').evaluate().setTitle("在？安排一下？"));
}
function mythQuote() {
  //SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("神话语录"), false);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('MythQuote').evaluate(), "在？看看语录？");
}

function addNewTask(o) {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  var row;
  if (o.cata.equals("game")) {
    row = 41;
  } else {
    for (var i = 1; i <= ss.getLastRow(); i++) {
      if (ss.getRange(i, 1).getValue().equals("动漫相关内容稿件")) {
        row = i + 1;
        break;
      }
    }
  }
  ss.insertRowBefore(row);
  ss.getRange("A" + row + ":H" + row).setValues([[o.name, o.lang, o.diff, o.type, o.abst, o.pfrd, "", o.frg1]]);
  ss.getRange(row, 14).setValue(o.adss);
  if (!o.frg2.equals(""))
    ss.getRange("J" + row).setValue(o.frg2);
  else
    ss.getRange("J" + row + ":K" + row).setValues([["/", "/"]]);
  if (!o.frg3.equals(""))
    ss.getRange("L" + row).setValue(o.frg3);
  else
    ss.getRange("L" + row + ":M" + row).setValues([["/", "/"]]);
  ss.getRange("A" + row + ":AH" + row).setBackgroundColor("#f4cccc");
  Browser.msgBox("挖坑成功！记得填哦");
}

function takeTask() {
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处"), false);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile('TakeTask').evaluate().setTitle("在？接接坑？"));
}

function newTakeTask(id) {
  email = PropertiesService.getUserProperties().getProperty("e-mail");
  if (email == null)
    return Browser.msgBox("请先登记！");
  row = parseInt(id.slice(0, -4), 10);
  task = id.slice(-4);
  switch (task) {
    case "pfrd": column = 6; break;
    case "tmln": column = 7; break;
    case "frg1": column = 9; break;
    case "frg2": column = 11; break;
    case "frg3": column = 13; break;
  }
  nickname = findNickname(email);
  if (nickname == undefined)
    return Browser.msgBox("接坑失败。");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处").getRange(row, column).setValue(findNickname(email));
  Browser.msgBox("接坑成功！\\n自己接的坑，跪着也要做完");
  s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  function check(x) {
    return !s.getRange(row, x).getValue().equals("");
  }
  if (check(6) && check(7) && check(9) && check(11) && check(13)) {
    changeRowColorToWhite(row);
  }
}

function findNickname(email) {
  english = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("英翻");
  timeline = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间轴");
  lastRow = english.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (english.getRange(i, 2).getValue().equals(email))
      return english.getRange(i, 1).getValue();
  lastRow = timeline.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (timeline.getRange(i, 2).getValue().equals(email))
      return timeline.getRange(i, 1).getValue();
  Browser.msgBox("未能找到你的信息！\\n请确认邮箱是否填写正确");
}

function changeRowColorToWhite(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  ss.getRange("A" + row + ":AH" + row).setBackgroundColor("#ffffff");
}
function mythQuoteSheet() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("神话语录"));
}
function nextQuote(randomRow) {
  var ran = parseInt(randomRow, 10);
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("神话语录数据库");
  var lastRow = ss.getLastRow();
  return ss.getRange(ran, 2).getValue().toString();

}