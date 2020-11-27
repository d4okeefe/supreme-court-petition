function onInstall(){
  onOpen();
}

function onOpen() {
  DocumentApp.getUi() 
      .createMenu('Supreme Court Guidance')
      .addItem('Show cover guidance', 'showCoverNotes')
      .addItem('Show index guidance', 'showIndexNotes')
      .addItem('Show brief guidance', 'showBriefNotes')
      .addToUi();
}

function showCoverNotes(){
  var ui = HtmlService.createTemplateFromFile('sidebar_cover').evaluate();
  ui.setTitle('Guidance for a Supreme Court Cover');
  DocumentApp.getUi().showSidebar(ui);  
}

function showIndexNotes(){
  var ui = HtmlService.createTemplateFromFile('sidebar_index').evaluate();
  ui.setTitle('Guidance for a Supreme Court Index');
  DocumentApp.getUi().showSidebar(ui);
}

function showBriefNotes(){
  var ui = HtmlService.createTemplateFromFile('sidebar_brief').evaluate();
  ui.setTitle('Guidance for a Supreme Court Brief');
  DocumentApp.getUi().showSidebar(ui);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function setPageSizeToLetter(){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  if(body.getPageHeight() !== 792 && body.getPageWidth() !== 612){
    body.setPageHeight(792);
    body.setPageWidth(612);
  }
  return "Letter: 8-1/2\" x 11\"";
}

function setPageSizeToBooklet(){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  if(body.getPageHeight() !== 666 && body.getPageWidth() !== 441){
    body.setPageHeight(666);
    body.setPageWidth(441);
  }
  return "Booklet: 6-1/8\" x 9-1/4\"";
}

function increaseSpaceAroundAllParagraphs(){
  var document = DocumentApp.getActiveDocument();
  var paras = document.getBody().getParagraphs();
  for(var i = 0; i < paras.length; i++){
    
    var currSpaceBefore = paras[i].getSpacingBefore();
    var currSpaceAfter = paras[i].getSpacingAfter();
    
    if(i === 0){
      paras[i].setSpacingBefore(0);
      paras[i].setSpacingAfter(currSpaceAfter + 2);      
    }else if(i === paras.length - 1){
      paras[i].setSpacingBefore(currSpaceBefore + 2);
      paras[i].setSpacingAfter(0);      
    }else{
      paras[i].setSpacingBefore(currSpaceBefore + 2);
      paras[i].setSpacingAfter(currSpaceAfter + 2);      
    }
  }
  return "Text spacing expanded";
}

function decreaseSpaceAroundAllParagraphs(){
  var document = DocumentApp.getActiveDocument();
  var paras = document.getBody().getParagraphs();
  for(var i = 0; i < paras.length; i++){
    
    var currSpaceBefore = paras[i].getSpacingBefore();
    var currSpaceAfter = paras[i].getSpacingAfter();
    
    if(i === 0){
      paras[i].setSpacingBefore(0);
      paras[i].setSpacingAfter(currSpaceAfter - 2);      
    }else if(i === paras.length - 1){
      paras[i].setSpacingBefore(currSpaceBefore - 2);
      paras[i].setSpacingAfter(0);      
    }else{
      paras[i].setSpacingBefore(currSpaceBefore - 2);
      paras[i].setSpacingAfter(currSpaceAfter - 2);      
    }
  }
  return "Text spacing reduced";
}

function replacePetitioners(petits){
  Logger.log(typeof petits);
  Logger.log(petits);
  var doc = DocumentApp.getActiveDocument();
  var f = "{{PETS}}";
  var body = doc.getBody();
  if(petits){
    body.replaceText(f, petits);
  } else {
    Logger.log("no data");
  }
}
function replaceRespondents(resps){
  Logger.log(typeof resps);
  Logger.log(resps);
  var doc = DocumentApp.getActiveDocument();
  var f = "{{RESPS}}";
  var body = doc.getBody();
  if(resps){
    body.replaceText(f, resps);
  } else {
    Logger.log("no data");
  }
}
function replaceLowerCourt(lowerct){
  Logger.log(typeof lowerct);
  Logger.log(lowerct);
  var doc = DocumentApp.getActiveDocument();
  var f = "{{LOWER_COURT}}";
  var body = doc.getBody();
  if(lowerct){
    body.replaceText(f, lowerct);
  } else {
    Logger.log("no data");
  }
}