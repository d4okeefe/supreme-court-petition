/*
 * Set index page numbers in Petition. This function works, but
 * often sets the page numbers in incorrect order. When it captures
 * headers, the program doesn't know the order the headers appear
 * in the document
 */
function setIndexNumbersInHeaders(){
  var d = DocumentApp.getActiveDocument();
  var h = d.getHeader();
  var par = h.getParent();
  var cnt = par.getNumChildren();
  Logger.log("Section count: " + cnt);
  
  var index_pages = [
    "i", "ii", "iii", "iv", "v",
    "vi", "vii", "viii", "ix", "x",
    "xi", "xii", "xiii", "xiv", "xv"
  ];
  
  // get index headers
  var header_count = 0;
  var header_list = [];
  for(var i = 0; i < cnt; i++){
    var chd = par.getChild(i);
    if(chd.getType() === DocumentApp.ElementType.HEADER_SECTION){
      header_count++;
      header_list.push(chd);
    }
  }
  Logger.log("header_count: " + header_count);
  Logger.log("header_list.length: " + header_list.length);
  
  var pg_num_itm = 0;
  
  for(var i = 0; i < header_count; i++){
    //var chd = par.getChild(i);
    //Logger.log(chd);
    //if(chd.getType() === DocumentApp.ElementType.HEADER_SECTION){
    Logger.log(header_list[i]);
//      if(i > 0 && i < header_count.length - 1){
    var paras = header_list[i].asHeaderSection().getParagraphs();
    if(i === 0){continue;}
    if(i === header_list.length - 2){continue;}
    
    Logger.log("num paragraphs: " + paras.length);
    
    if(paras.length === 0){
      header_list[i].insertParagraph(0, index_pages[pg_num_itm++]);
    } else if(paras.length === 1){
      var p = paras[0];
      p.clear();
      p.setText(index_pages[pg_num_itm++]);
    } else {
      Logger.log("Problem here: too many paragraphs");
    }
  }
}


/**
 * TOC OVERVIEW
 * Google docs provides an automated Table of Contents.
 * We plan to include this TOC in the templates that we
 * build for clients.
 * However, in the template, index page numbers are set
 * manually, so they will be incorrect on update.
 * LG wants to provide a method of updating those page
 * numbers with one click, using Apps Script.
 *
 * Step 1: Identify headings in index.
 * (Will a preset "namedrange" be stable enough?)
 * Use a preset NamedRange to identify index headings.
 *
 * Step 2: Use list of index headings to identify
 * matches in TOC.
 *
 * Step 3: Create dictionary of numbers, using this to
 * replace the numbers in the TOC.
 */
 function isDocHeading(h){
  var headings = [
    DocumentApp.ParagraphHeading.HEADING1,
    DocumentApp.ParagraphHeading.HEADING2,
    DocumentApp.ParagraphHeading.HEADING3,
    DocumentApp.ParagraphHeading.HEADING4,
    DocumentApp.ParagraphHeading.HEADING5,
    DocumentApp.ParagraphHeading.HEADING6
  ];
  
  for (var i = 0; i < headings.length; i++) {
    if(h === headings[i]){
      return true;
    }
  }
  return false;
}
 
/**
  * Capture headings from 'index', which is a NamedRange
  *
  */
function captureHeadingsFromIndex(){
  var d = DocumentApp.getActiveDocument();
  var toc = [];
  var index_r = d.getNamedRanges('index');
  if(index_r){
    var r = index_r[0].getRange();
    var elements = r.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var e = elements[i].getElement();
      if(e.getType() === DocumentApp.ElementType.PARAGRAPH){
        var p = e.asParagraph();
        var para_t = p.getHeading();
        if(isDocHeading(para_t)){
          toc.push(p.getText());
        }
      }
    }
  }
//  for(var itm in toc){
//    Logger.log(toc[itm]);
//  }
  return toc;
}

function romanize(num) {
  var lookup = {M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1},
      roman = '',
      i;
  for ( i in lookup ) {
    while ( num >= lookup[i] ) {
      roman += i;
      num -= lookup[i];
    }
  }
  return roman;
}

function correctTOCIndexPageNumbers(){
  var d = DocumentApp.getActiveDocument();
  var r = d.getNamedRanges('index')[0].getRange();
  
  var index_headings = captureHeadingsFromIndex();
  
  if(r){
    var elements = r.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i].getElement();
      var ele_type = element.getType();
      if(ele_type === DocumentApp.ElementType.TABLE_OF_CONTENTS){
        var toc = element.asTableOfContents();
        var toc_item_count = toc.getNumChildren();
        
        for(var j = 0; j < toc_item_count; j++){
          var toc_itm = toc.getChild(j);
          var toc_itm_type = toc_itm.getType();
          if(toc_itm_type === DocumentApp.ElementType.PARAGRAPH){
            var edit = toc_itm.editAsText();
            var txt = edit.getText();
            for(var k = 0; k < index_headings.length; k++){
              if(k === 0){ edit.setBold(false); }
              if(txt.includes(index_headings[k])){
                var num = /\d+/.exec(txt);
                var rn = romanize(num - 1);
                edit.replaceText("\\d+", rn.toLowerCase());
              }
            }
          }
        }
      }
    }
  }
}

