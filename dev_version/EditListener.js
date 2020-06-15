function onEdit(e) {

  if(e.range.getSheet().getSheetName().equals("时间安排"))
    memberWorkDone(e);
  if(e.range.getSheet().getSheetName().equals("接坑处"))
    //projectAutoAdd(e.range.getRow());
  if(e.range.getSheet().getSheetName().equals("test")){
    //only for testing and debugging purpose
    e.range.setNote(getUrl());
  }

  //TODO: find a decent trigger or event for checkSession
  checkSessionData();
}