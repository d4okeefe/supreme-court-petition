function runModal_PageIndex() {
  var htmlOutput = HtmlService
    .createHtmlOutput('<div><iframe width="1120" height="630" src="https://www.youtube.com/embed/t8yNMHzrIFw?controls=0" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></div>')
    .setWidth(1200)
    .setHeight(800);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'How to manually page index');
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

/**
 * Get heading argument from TOC
 * Assumes that heading text is unique, which it should be with date stamp
 * @param {string} heading text
 * @return {string} heading argument for linking
 * @private
 */
function getHeadingArgFromTOC_(doc, headingText) {
  var bodyChild;
  var TOCChild;
  var TOCLinkURL;
  var headingArg;
  
  var body = doc.getBody();
  var numBodyChildren = body.getNumChildren();

  for (var i = 0; i < numBodyChildren; i++) {
    bodyChild = body.getChild(i);
    if (bodyChild.getType() == DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      var numTOCChildren = bodyChild.getNumChildren();
      for (var j = 0; j < numTOCChildren; j++) {
        TOCChild = bodyChild.getChild(j);
        if (TOCChild.getType() == DocumentApp.ElementType.PARAGRAPH ||
            TOCChild.getType() == DocumentApp.ElementType.LIST_ITEM) {
          if (TOCChild.getText() && TOCChild.getText() == headingText) {
            TOCLinkURL = TOCChild.getLinkUrl();
            if (TOCLinkURL) {
              headingArg = TOCLinkURL;
              break;
            }       
          }
        }
      }
    }
  }
  return headingArg;
}


function testRoman(){
  var result = romanize(10);
  Logger.log(typeof(result));
  result = result.toLowerCase();
  Logger.log(result);
}

function getIndexHeadings(){
  var d = DocumentApp.getActiveDocument();
  var ir = DocumentApp.getActiveDocument().getNamedRanges('index')[0].getRange();
  var elements = ir.getRangeElements();
  var heading_text = [];
  for(var i = 0; i < elements.length; i++){
    var e = elements[i].getElement();
    if(e.getType() === DocumentApp.ElementType.PARAGRAPH){
      var para = e.asParagraph();
      var h = para.getHeading();
      //Logger.log(h);
      for(var j = 0; j < headings.length; j++){
        if(h === headings[j]){
          heading_text.push(para.getText());
        }
      }
    }
  }
  for(var k = 0; k < heading_text.length; k++){
    Logger.log("heading_text: " + heading_text[k]);
  }
}




/*
 * Attempt to update TOC entry index page numbers
 */
function updateTocEntries(){
  var d = DocumentApp.getActiveDocument();
  var b = d.getBody();
  var toc = b.findElement(DocumentApp.ElementType.TABLE_OF_CONTENTS);
  var toc_element = toc.getElement().asTableOfContents();
  var cnt = toc_element.getNumChildren();
  for(var i = 0; i < cnt; i++){
    var child = toc_element.getChild(i);
    var t = child.getType();
    Logger.log(t);
    if(t === DocumentApp.ElementType.PARAGRAPH){
      var para = child.asParagraph();
      Logger.log(para.getText() + ": " + para.getLinkUrl());
    }
  }
}


/*
 * Can I create custom TOC?
 */
function createTOC(){
  var d = DocumentApp.getActiveDocument();
  var p = d.getBody().getParagraphs();
  var h = [];
  var toc = {}
  for(var i = 0; i < p.length; i++){
    if(p[i].getHeading() !== DocumentApp.ParagraphHeading.NORMAL){
      toc = {
        "text": p[i].getText(),
        "heading": p[i].getHeading()
      }
      h.push(toc);
    }
  }
  Logger.log(h);
  
//  var c = d.getSelection();
//  var new_toc = []
//  for(var i = 0; i < h.length; i++){
//    var new_p = Docs.newParagraph();
//    new_p.paragraphStyle = toc["heading"];
//   
//  }
}