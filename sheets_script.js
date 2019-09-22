var DOCID="1gfLacBAD69Xh-TkvvOwzP5GvH9DFokrGeNM7y9vrZg8";

// Swap the order of columns with these headings.
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

function printSelected_() {
  printRows_(true);    
}
    
function printAll_(){
  printRows_(false);
}

/**
 * Print all (or selected) rows to newly-created doc. Show link to the doc. 
 */
function printRows_(chooseSelected) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastC = sheet.getLastColumn();
  var headingsRange = sheet.getRange(1,1,1,lastC);
  var headingsRichText = headingsRange.getRichTextValues()[0];
  var headingsValues = headingsRange.getDisplayValues()[0];
  var aliases = sheet.getRange(2,1,1, lastC).getDisplayValues()[0];
  
  var firstRow = 3;
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
    childValues = values[j];
    firstName = childValues[headingsValues.indexOf("Child's first name")];
    lastName = childValues[headingsValues.indexOf("Child's last name")];
    addChild_(body, headingsRichText, aliases, firstName, lastName, childValues);
  }
  doc.saveAndClose();
    
  var link = HtmlService.createHtmlOutput('<a href="'+doc.getUrl()+'">click to open doc</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Printable version'); 
}

/* Add child (values) to document (body).  Break up by headings. */
function addChild_(body, headings, aliases, firstName, lastName, values) {
  body.appendParagraph(lastName+ ", " + firstName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var p = body.appendParagraph("");

  // swap(headings, values, swaps);
  
  for (var j=0; j<headings.length; j++) {
    var h = headings[j];
    var hStyle = h.getTextStyle();
    var hContent = aliases[j];
    if (hContent == "") {
      hContent = h.getText();
    }
    var hType = "normal";
    if (hStyle.isBold() && hStyle.isItalic()) {
      hType = "heading";
    } else if (hStyle.isItalic()) {
      hType = "run on";
    } else if (hStyle.isBold()) {
      hType = "isolate";
    }
    
    if (hType == "heading") {
      body.appendParagraph(hContent).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      p = body.appendParagraph("");
      hContent = h.getText();
    }

    if (values[j] == "") {
      continue;
    }

    if (hType == "normal") {
      p = body.appendParagraph("");
    }
    if (hType == "isolate") {
      p = body.appendParagraph("");
      p = body.appendParagraph("");
    }
    
    p.appendText(hContent + ": ").setItalic(true);
    p.appendText(values[j] +"; ").setItalic(false);
    if (hType == "isolate") {
      p = body.appendParagraph("");
    }
  }
  body.appendPageBreak();
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
