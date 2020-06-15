
function addMenu() {
  SpreadsheetApp.getUi()
  .createMenu('字幕组')
  .addItem('登记', 'register')
  .addItem("挖坑", "newTask")
  .addItem("接坑", "takeTask")
  .addItem("手动排坑", "manuallySet")
  .addItem("神话语录", "mythQuote")
  .addItem("Ui框架开发", "UiFrameworkDev")
  .addToUi();

}

function register() {
  var email_input = updateSessionEmail();
  var nickname_input = updateNickName();

  //for test
  //Browser.msgBox('email range: ' + findEmailRange(email_input));
  //Browser.msgBox('nickname range: ' + findNicknameRange(nickname_input));

  if(!findEmailRange(email_input) && !findNicknameRange(nickname_input)){
    //没有对应的邮箱和昵称的记录，询问是否是新人
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('记录中没有对应的邮箱和昵称的，也许是新人？', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      creatNewPigeon(email_input, nickname_input);
    } else {
      Browser.msgBox('请确认您的邮箱和昵称是否正确\\n' +
                     '邮箱： ' + email_input + '\\n' +
                     '昵称： ' + nickname_input + '\\n' +
                      '如果遇到bug，请尽快联系我们');
      Logger.log('注册邮箱： ' + email_input + '\\n' +
                '注册昵称： ' + nickname_input + '\\n');
    }
  } else if (String(findNickname(email_input)).toLowerCase() === String(nickname_input).toLowerCase()){
    Browser.msgBox(findNickname(email_input) + ' 登记成功！');
  }
  /*
   * old version
  var nick = findNickname(email);
  if(nick === ''){
    Browser.msgBox('登记失败：无法找到用户名!');
  }else{
    Browser.msgBox('登记成功，欢迎' + nick);

    //findNicknameRange(nick).setFontColor('green');
  }
  */
}


//注册新用户
//TODO: add real implementation
function creatNewPigeon(email, nickname) {
    Browser.msgBox('注册功能还未实现呜呜呜')
    return true
}


//将用户email更新到session的Properties中
function updateSessionEmail(){
  email = Session.getActiveUser().getEmail();
  if(validateEmail(email)){
    PropertiesService.getUserProperties().setProperty("e-mail",email);
  } else {
    //邮箱获取错误
    Logger.log('登记失败：邮箱获取错误!')
  }
  return email;
}


//根据用户输入储存nick name，
function updateNickName(){
  var nickname = Browser.inputBox('请输入昵称，不需要区分大小写喔~');
  while(String(nickname) ===''){
    Browser.msgBox('昵称不能为空！');
    nickname = Browser.inputBox('请输入昵称，不用区分大小写~');
  }
  PropertiesService.getUserProperties().setProperty("nickname",nickname);
  return nickname;
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
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('MythQuote').evaluate(),"在？看看语录？");
}

function addNewTask(o) {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  var row;
  if(o.cata.equals("game")){
  row = 41;
  }else{
  for(var i =1; i<=ss.getLastRow(); i++){
    if(ss.getRange(i, 1).getValue().equals("动漫相关内容稿件")){
      row = i+1;
      break;
    }
  }
}
ss.insertRowBefore(row);
  ss.getRange("A" + row +":H" + row).setValues([[o.name,o.lang,o.diff,o.type,o.abst,o.pfrd,"",o.frg1]]);
  ss.getRange(row, 14).setValue(o.adss);
  if(!o.frg2.equals("")) 
    ss.getRange("J"+row).setValue(o.frg2);
  else
    ss.getRange("J"+row+":K"+row).setValues([["/","/"]]);
  if(!o.frg3.equals("")) 
    ss.getRange("L"+row).setValue(o.frg3);
  else
    ss.getRange("L"+row+":M"+row).setValues([["/","/"]]);
  ss.getRange("A"+row+":AH"+row).setBackgroundColor("#f4cccc");
  Browser.msgBox("挖坑成功！记得填哦");
}


function takeTask() {
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处"), false);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile('TakeTask').evaluate().setTitle("在？接接坑？"));
}


function newTakeTask(id) {
  email = PropertiesService.getUserProperties().getProperty("e-mail");
  if(email==null)
    return Browser.msgBox("请先登记！");
  row = parseInt(id.slice(0, -4), 10);
  task = id.slice(-4);
  switch(task) {
    case "pfrd":column=6;break;
    case "tmln":column=7;break;
    case "frg1":column=9;break;
    case "frg2":column=11;break;
    case "frg3":column=13;break;
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处").getRange(row, column).setValue(findNickname(email));
  Browser.msgBox("接坑成功！\\n自己接的坑，跪着也要做完");
  s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  function check(x) {
    return !s.getRange(row, x).getValue().equals("");
  }
  if(check(6) && check(7) && check(9) && check(11) && check(13)) { 
  changeRowColorToWhite(row);
  }
}


function findNickname(email) {
  english = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("英翻");
  timeline = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间轴");
  lastRow = english.getLastRow();
  for(i=1;i<=lastRow;++i)
    if(english.getRange(i, 2).getValue().equals(email))
      return english.getRange(i,1).getValue();
  lastRow = timeline.getLastRow();
    for(i=1;i<=lastRow;++i)
    if(timeline.getRange(i, 2).getValue().equals(email))
      return timeline.getRange(i,1).getValue();

  //未找到则返回空
  //TODO: 设置错误flag
  //Browser.msgBox("未能找到你的信息！\\n请确认邮箱是否填写正确");
  return '';
}


//不区分大小写的搜索
function findNicknameRange(nickname) {
  english = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("英翻");
  timeline = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间轴");
  lastRow = english.getLastRow();
  for(i=1;i<=lastRow;++i)
    if(String(english.getRange(i, 1).getValue()).toLowerCase() === String(nickname).toLowerCase())
      return english.getRange(i,1);
  lastRow = timeline.getLastRow();
  for(i=1;i<=lastRow;++i)
    if(String(timeline.getRange(i, 1).getValue()).toLowerCase() === String(nickname).toLowerCase())
      return timeline.getRange(i,1);
  //Browser.msgBox('未能找到对应的昵称！\\n');
  //未找到则返回null
  //TODO: 设置错误flag
  return null;
}


function findEmail(nickname) {
  english = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("英翻");
  timeline = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间轴");
  lastRow = english.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (english.getRange(i, 1).getValue().equals(nickname))
      return english.getRange(i, 2).getValue();
  lastRow = timeline.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (timeline.getRange(i, 1).getValue().equals(nickname))
      return timeline.getRange(i, 2).getValue();
  //未找到则返回空
  //TODO: 设置错误flag
  //Browser.msgBox("未能找到你的信息！\\n请确认昵称是否填写正确");
  return '';
}


function findEmailRange(email) {
  english = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("英翻");
  timeline = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("时间轴");
  lastRow = english.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (english.getRange(i, 2).getValue().equals(email))
      return english.getRange(i, 2);
  lastRow = timeline.getLastRow();
  for (i = 1; i <= lastRow; ++i)
    if (timeline.getRange(i, 2).getValue().equals(email))
      return timeline.getRange(i, 2);
  //未找到则返回null
  //TODO: 设置错误flag
  return null;
}


function changeRowColorToWhite(row){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("接坑处");
  ss.getRange("A" + row + ":AH" + row).setBackgroundColor("#ffffff");
}


function mythQuoteSheet(){
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("神话语录"));
}


function nextQuote(randomRow){
  var ran = parseInt(randomRow, 10);
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("神话语录数据库");
  var lastRow = ss.getLastRow();
  return ss.getRange(ran, 2).getValue().toString();
  
}


function validateEmail(email) {
  //for test
  Logger.log('validate: '+ email);

  var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}

function UiFrameworkDev() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile('UiDevPage').evaluate().setTitle("试验新功能"));

}