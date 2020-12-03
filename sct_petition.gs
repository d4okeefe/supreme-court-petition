function onInstall(){
  onOpen();
}

function onOpen() {
  DocumentApp.getUi() 
      .createMenu('Supreme Court Guidance')
      .addItem('Show petition guidance', 'showMainNotes')
//      .addItem('Show cover guidance', 'showCoverNotes')
//      .addItem('Show index guidance', 'showIndexNotes')
//      .addItem('Show brief guidance', 'showBriefNotes')
      .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


function showMainNotes(){
  var ui = HtmlService.createTemplateFromFile('sidebar_main').evaluate();
  ui.setTitle('Guidance for a Supreme Court Petition');
  DocumentApp.getUi().showSidebar(ui);  
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



/*
 * Inserts heading into brief. Called from buttons on sidebar.
 * @param heading String sent by html page to Google script.
 */
function insertHeading(heading){
  var d = DocumentApp.getActiveDocument();
  var s = d.getSelection();
  if(s){
    s.getRangeElements()[0].getElement().asParagraph().insertText(0, heading.toUpperCase());
  }
  var c = d.getCursor();
  if(c){
    c.getElement().asParagraph().insertText(0, heading.toUpperCase());
    c.getElement().asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING1);
  }
}



function check_setParagraphStyle(){
  setParagraphStyle("normal");
}

function setParagraphStyle(style){
  var d = DocumentApp.getActiveDocument();
  
  if(style === "normal"){
    style = DocumentApp.ParagraphHeading.NORMAL;
  }else if(style === "h1"){
    style = DocumentApp.ParagraphHeading.HEADING1;
  }else if(style === "h2"){
    style = DocumentApp.ParagraphHeading.HEADING2;
  }else if(style === "h3"){
    style = DocumentApp.ParagraphHeading.HEADING3;
  }else if(style === "h4"){
    style = DocumentApp.ParagraphHeading.HEADING4;
  }else if(style === "h5"){
    style = DocumentApp.ParagraphHeading.HEADING5;
  }else if(style === "h6"){
    style = DocumentApp.ParagraphHeading.HEADING6;
  }
  
  var selection = d.getSelection();
  if(selection){
    var rElements = selection.getRangeElements();
    for(var i = 0; i < rElements.length; i++){
      var r = rElements[i];
      var e = r.getElement();
      if(e.getType() === DocumentApp.ElementType.PARAGRAPH){
        e.asParagraph().setHeading(style);
      }
    }
  }else{
    Logger.log("No selection, checking cursor...");
    var cursor = d.getCursor();
    if(cursor){
      var cursor_element = cursor.getElement();
      var rangeBuilder = doc.newRange();
      rangeBuilder.addElement(cursor_element);
      var new_r = rangeBuilder.build();
      new_r.getRangeElements();
      for(var i = 0; i < rElements.length; i++){
        var r = rElements[i];
        var e = r.getElement();
        if(e.getType() === DocumentApp.ElementType.PARAGRAPH){
          e.asParagraph().setHeading(style);
        }
      }
    }
  }
  Logger.log("No selection or cursor");
  // set paragraph style under cursor or selection
  
  return "style change";
}









function setPageSizeToLetter(){
  var d = DocumentApp.getActiveDocument();
  var b = d.getBody();
  if(b.getPageHeight() !== 792 && b.getPageWidth() !== 612){
    b.setPageHeight(792);
    b.setPageWidth(612);
  }
  return "Letter: 8-1/2\" x 11\"";
}

function setPageSizeToBooklet(){
  var d = DocumentApp.getActiveDocument();
  var b = d.getBody();
  if(b.getPageHeight() !== 666 && b.getPageWidth() !== 441){
    b.setPageHeight(666);
    b.setPageWidth(441);
  }
  return "Booklet: 6-1/8\" x 9-1/4\"";
}

function increaseSpaceAroundAllParagraphs2(){
  var d = DocumentApp.getActiveDocument();
  var r = getCoverRange(d);
  
  for(var i = 0; i < r.length; i++){
    var e = r[i].getElement();
    var t = e.getType();
    Logger.log("orig_type: " + t);
    
    var count = 0;
    while(t !== DocumentApp.ElementType.PARAGRAPH){
      if(t === DocumentApp.ElementType.TABLE){break;}
      if(!e){break}
      
      var new_e = e.getParent();
      Logger.log("count: " + count++ + "; new_parent: " + new_e.getType());
      e = new_e;
      t = e.getType();
    }
    Logger.log("BREAK");
    
    if(t === DocumentApp.ElementType.PARAGRAPH){
      //Logger.log("para found");
      
      var p = e.asParagraph();
      var currSpaceBefore = p.getSpacingBefore();
      var currSpaceAfter = p.getSpacingAfter();
    
      if(i === 0){
        p.setSpacingBefore(0);
        p.setSpacingAfter(currSpaceAfter + 2);      
      }else if(i === r.length - 1){
        p.setSpacingBefore(currSpaceBefore + 2);
        p.setSpacingAfter(0);      
      }else{
        p.setSpacingBefore(currSpaceBefore + 2);
        p.setSpacingAfter(currSpaceAfter + 2);      
      }
    }      
  }
  return "Text spacing expanded";
}

function decreaseSpaceAroundAllParagraphs2(){
  var d = DocumentApp.getActiveDocument();
  var r = getCoverRange(d);
  
  for(var i = 0; i < r.length; i++){
    var e = r[i].getElement();
    var t = e.getType();
    Logger.log("orig_type: " + t);
    
    var count = 0;
    while(t !== DocumentApp.ElementType.PARAGRAPH){
      if(t === DocumentApp.ElementType.TABLE){break;}
      if(!e){break}
      
      var new_e = e.getParent();
      Logger.log("count: " + count++ + "; new_parent: " + new_e.getType());
      e = new_e;
      t = e.getType();
    }
    Logger.log("BREAK");
    
    if(t === DocumentApp.ElementType.PARAGRAPH){
      //Logger.log("para found");
      
      var p = e.asParagraph();
      var currSpaceBefore = p.getSpacingBefore();
      var currSpaceAfter = p.getSpacingAfter();
    
      if(i === 0){
        p.setSpacingBefore(0);
        p.setSpacingAfter(currSpaceAfter - 2);      
      }else if(i === r.length - 1){
        p.setSpacingBefore(currSpaceBefore - 2);
        p.setSpacingAfter(0);      
      }else{
        p.setSpacingBefore(currSpaceBefore - 2);
        p.setSpacingAfter(currSpaceAfter - 2);      
      }
    }      
  }
  return "Text spacing reduced";
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

