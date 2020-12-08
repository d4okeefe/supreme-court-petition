function insertSignatureTable(){

}

function normalStyle(){
  var style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  style[DocumentApp.Attribute.FONT_FAMILY] = "Century Schoolbook";
  //style[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  //style[DocumentApp.Attribute.SPACING_AFTER] = 6;
  //style[DocumentApp.Attribute.LINE_SPACING] = 1.3;
  return style;
}

/**
  * Set index page numbers in Petition. This function works, but
  * often sets the page numbers in incorrect order. When it captures
  * headers, the program doesn't know the order the headers appear
  * in the document.
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
 
 /**
  * Converts arabic numbers to roman numerals.
  * @param h ParagraphHeading value
  * @return Boolean
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
    //Logger.log(h);
    if(h === headings[i]){
      //Logger.log(true);
      return true;
    }
  }
  return false;
}
 
/**
  * Capture headings from 'index', which is a NamedRange
  * @return String array of headings in 'index' NamedRange
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
  for(var j = 0; j < toc.length; j++){
    //Logger.log(toc[j]);
  }
  return toc;
}

/**
  * Converts arabic numbers to roman numerals.
  * @param num An arabic number
  * @return String representation of roman numeral, in upper case.
  */
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

/**
  * Simple function to print array.
  */
function printArray(a){
  for (var i = 0; i < a.length; i++) {
    Logger.log(a[i]);
  }
}

/**
  * Corrects TOC entry page numbers.
  * @return String ui update
  */
function correctTOCIndexPageNumbers(){
  var d = DocumentApp.getActiveDocument();
  var r = d.getNamedRanges('index')[0].getRange();
  
  var index_headings = captureHeadingsFromIndex();
  //printArray(index_headings);
  
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
            edit.setBold(false);
            var toc_txt = edit.getText();
            var rx = /^(.*)\t(\d+)/;
            var toc_itm = toc_txt.match(rx);
            if(toc_itm){
//              Logger.log(toc_itm[1]);
//              Logger.log(toc_itm[2]);
              
              for(var k = 0; k < index_headings.length; k++){
                
                if(toc_itm[1] === index_headings[k]){
                  edit.setBold(false);
                  var rn = romanize(toc_itm[2] - 1);
                  edit.replaceText("\\d+", rn.toLowerCase());
                }
              }
              
            } else {
              // skip -- toc item bad structure
            }
          } // end if Paragraph
        } // end for toc_item
      } // end if ele_type TABLE_OF_CONTENTS
    } // end for elements
  } // end if r
  return "Page numbers updated. See Table of Contents."
} // end function