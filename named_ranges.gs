/*
 * Gets 'cover' NamedRange in document.
 * @return Range of 'cover' in document
 */
 function getCoverRange(d){
  var rngs = d.getNamedRanges();
  for(var i = 0; i < rngs.length; i++){
    if(rngs[i].getName() === "cover"){
      return rngs[i].getRange().getRangeElements();
    }
  }
}
/*
 * Sets 'cover' NamedRange in template document.
 */
function createCoverNamedRange(){
  var d = DocumentApp.getActiveDocument();
  var rangeBuilder = d.newRange();
  var body = d.getBody();
  var startElem = body.findText("No\.").getElement();
  var endElem = body.findText("www.lookingglass.legal").getElement();
  rangeBuilder.addElementsBetween(startElem, endElem);
  d.addNamedRange("cover", rangeBuilder);
}
/*
 * Sets 'index' NamedRange in template document.
 */
function createIndexNamedRange(){
  var d = DocumentApp.getActiveDocument();
  var rangeBuilder = d.newRange();
  var body = d.getBody();
  var startElem = body.findText("QUESTION PRESENTED").getElement();
  var endElem = body.findText("End index").getElement();
  rangeBuilder.addElementsBetween(startElem, endElem);
  d.addNamedRange("index", rangeBuilder);
}
/*
 * Removes NamedRange in template document.
 */
function deleteNamedRange(){
  var d = DocumentApp.getActiveDocument();
  var r = d.getNamedRanges('cover');
  r[0].remove();
  Logger.log(r);
}
/*
 * Selects NamedRange in template document.
 */
function selectNamedRange(){
  var d = DocumentApp.getActiveDocument();
  var rng = d.getNamedRanges();
  for(var i = 0; i < rng.length; i++){
    Logger.log(rng[i].getId());
    Logger.log(rng[i].getName());
    if(rng[i].getName() === 'index'){
//      for(var j = 0; j < rng[i].length; j++){
//        var e = rng[i].getRange().getRangeElements()[j].getElement();
//      }
      d.setSelection(rng[i].getRange());
    }
  }
}