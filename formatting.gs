function getInfo(){
  var id = DocumentApp.getActiveDocument().getId();
  var d = Docs.Documents.get(id);
  
}
function styleInfo() {
  var id = DocumentApp.getActiveDocument().getId();
  var d = Docs.Documents.get(id);
  var s = d.namedStyles.styles;
  for(var i = 0; i < s.length; i++){
    Logger.log("type: " + s[i].namedStyleType);
    Logger.log("indentStart: " + s[i].paragraphStyle.indentStart);
    Logger.log("indentFirstLine: " + s[i].paragraphStyle.indentFirstLine);
    Logger.log("lineSpacing: " + s[i].paragraphStyle.lineSpacing);
    Logger.log("alignment: " + s[i].paragraphStyle.alignment);
  }
}
function rangeInfo(){
  var id = DocumentApp.getActiveDocument().getId();
  var d = Docs.Documents.get(id);
  var r = d.namedRanges;
  Logger.log(r);
}
function pageInfo(){
  var id = DocumentApp.getActiveDocument().getId();
  var d = Docs.Documents.get(id);

}
