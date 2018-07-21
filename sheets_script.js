var DOCID="1gfLacBAD69Xh-TkvvOwzP5GvH9DFokrGeNM7y9vrZg8";
var questionWidthInPoints = 150;
var verticalPadding = 0;

var sections = {
  "Date of Birth": "Essentials",
  "Mother's first name": "Mother",
  "Father's first name": "Father",
  "Child's doctor's name": "Medical",
  "How would you describe your child's development to this point in time?": "Development",
  "Child's country of birth": "Cultural",
  "First sibling's name": "Siblings",
};

var swaps = {
  "Relationship to Child (contact2)": "Relationship to Child",
};

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getActive().addMenu('Printable', [
    {name: 'all', functionName: 'printAll_'},
    {name: 'selected', functionName: 'printSelected_'},
  ]);
}

function swap(headings, values, swaps) {
  for (var h1 in swaps) {
    // check if the property/key is defined in the object itself, not in parent
    if (swaps.hasOwnProperty(h1)) {
        pos1 = headings.indexOf(h1);
        pos2 = headings.indexOf(swaps[h1]);
        if (pos1 == -1 || pos2 == -1) {
          // log? alert?
          return;
        }
        tmp = headings[pos1];
        headings[pos1] = headings[pos2];
        headings[pos2] = tmp;

                                     tmp = values[pos1];
        values[pos1] = values[pos2];
        values[pos2] = tmp;

    }
  }
}
    

function printSelected_() {
  printRows_(true);    
}
    
function printAll_(){
  printRows_(false);
}

function addChild_(body, headings, values) {
  var lastName = lookup_(headings, values, "Child's last name");
  var firstName = lookup_(headings, values, "Child's first name");
  
  var p = body.appendParagraph(lastName+ ", " + firstName);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  swap(headings, values, swaps);
  
  var table = body.appendTable();
  for (var j=0; j<headings.length; j++) {
    var sectionHeading = sections[headings[j]];

    if (sectionHeading) {
      var p = body.appendParagraph(sectionHeading);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      table = body.appendTable();
      table.setBorderWidth(0);
      
    }
    if (values[j] == "") {
      continue;
    }
    var row = table.appendTableRow();
    var k = row.appendTableCell(headings[j]);
    k.setWidth(questionWidthInPoints);
    setPadding_(k);
    var v = row.appendTableCell(values[j]);
    setPadding_(v);
  }
  body.appendPageBreak();
}

function setPadding_(cell) {
  cell.setPaddingTop(verticalPadding);
  cell.setPaddingBottom(verticalPadding);
}

function lookup_(headings, values,  key){
  return values[headings.indexOf(key)];
}

function printRows_(chooseSelected) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastC = sheet.getLastColumn();
  var headingRange = sheet.getRange(1,1,1,lastC);
  var headings = headingRange.getValues()[0];
  var firstRow = 2;
  var lastRow = sheet.getLastRow();
  if (chooseSelected){
    var selected = sheet.getActiveRange();
    firstRow = selected.getRow();
    lastRow = selected.getLastRow();
  }
  var nSelectedRows = lastRow - firstRow + 1;
  var values = sheet.getRange(firstRow, 1, nSelectedRows, lastC).getValues();
  // var doc = DocumentApp.openById(DOCID);
  var doc = DocumentApp.create("spreadsheet report");
  var body = doc.getBody();

  for (var j=0; j<values.length; j++){
    addChild_(body, headings, values[j]);
  }
  doc.saveAndClose();
    
  var link = HtmlService.createHtmlOutput('<a href="'+doc.getUrl()+'">click to open doc</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Printable version'); 
}

