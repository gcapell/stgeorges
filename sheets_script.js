var DOCID="1gfLacBAD69Xh-TkvvOwzP5GvH9DFokrGeNM7y9vrZg8";
var questionWidthInPoints = 150;
var verticalPadding = 0;

var sections = {
  "Date of Birth": "Essentials",
  "Mother's first name": "Mother",
  "Father's first name": "Father",
  "Contact Full Name": "Additional Contacts",
  "Child's doctor's name": "Medical",
  "How would you describe your child's development to this point in time?": "Development",
  "Child's country of birth": "Cultural",
  "First sibling's name": "Siblings",
};

// Swap the order of columns with these headings.
var swaps = {
  "Relationship to Child (contact2)": "Relationship to Child",
};
var short_heading_names = {
  "Timestamp": "",
  "Email Address": "",
  "Date of Birth": "DOB",
  "Child's last name": "",
  "Child's first name": "",
  "Is the child known by any other/former name?": "former",
  "Child's gender": "",
  "Address": "",
  "Suburb": "",
  "Post Code": "",
  "Home phone": "",
  
  "Mother's first name": "",
  "Mother's last name": "",
  "Mother's mobile phone": "mob",
  "Mother's email": "",
  "Any other names by which mother is known": "aka",
  "Mother's Occupation": "occ",
  "Mother's work phone": "wp",
  "Mother's work address": "work",
  "Mother's work hours": "work hours",
  "Mother's Employer's name": "employer",
  "Mother's work days": "work",
  "same address as child?": "same address",
  "Address": "address",
  "home phone": "ph",
  
  "Father's first name": "",
  "Father's last name": "",
  "Father's Mobile Phone": "",
  "Any other names by which Father is known": "aka",
  "Father's work phone": "work",
  "Father's email": "",
  "work days": "work",
  "Father's employer's name": "employer",
  "Father's occupation": "occ",
  "Father's Work Address": "work",
  "Father's work hours": "hours",
  "address same as child's?": "same address",
  "Street Address": "",
  "Suburb": "",
  "Postcode": "",
  "Home phone": "",
  "do both parents have access rights to the child?": "both access",
  "If you answered \"No\" above we will need to see court orders. Please advise if you have court orders.": "orders",
  "Would you like to nominate additional contacts?": "additional contacts",
  
  "Contact Full Name": 	"",
  "Address": "",
  "Relationship to Child": "rel",
  "Mobile phone and Home Phone" : "",
  "Work phone": "work",
  "Contact's authorisations": "auth",
  "Contact Full Name (contact2)": "contact2",
  "Address (contact2)": "",
  "Relationship to Child (contact2)": "rel",
  "Mobile phone and Home Phone": "",
  "Work phone": "",
  
  "Child's doctor's name": "doctor",
  "Doctor's phone": "",
  "Doctor's Address": "",
  "Medicare Number": "medicare",
  "Private Health Fund": "insurer",
  "Please describe your child's present state of health": "health",
  "Has your child been hospitalised?": "hospitalised",
  "Fever, Chicken Pox, Measles, Mumps,German Measles.Which of these apply to your child?": "common illnesses",
  "Does Your Child Take Any Regular Medications?": "meds?",
  "If \"yes\" Please Tell Us About The Medication Your Child Takes": "",
  "Has your child ever been diagnosed as being at risk of Anaphylaxis": "Anaphylaxis",
  "If yes, please provide details": "",
  "Has your child ever been diagnosed with asthma": "asthma",
  "If so, please provide details below": "",
  "Does your Child have any allergies": "allergies",
  "First Allergy": "",
  "Second Allergy": "",
  "Second Allergy": "",
  "What allergic reaction has your child had?": "reaction",
  
  "How would you describe your child's development to this point in time?": "",
  "how do you think your child will settle at preschool": "settle",
  "what would you like your child to gain from their time at preschool?": "gain",
  "is there anything you are concerned about that you would like us to monitor or anything we can help your child with?": "concerns",
  "what things does your child particularly like to do?": "favourite activity",
  "is your child toilet trained?": "toilet trained",
  "If \"no\" what is the current situation with your child's toilet training?": "",
  "Does your child sleep during the day?": "day sleep",
  "If so, at what time and for how long?": "",
  "Does your child have a security toy": "toy",
  "Describe toys (and names)": "",
  "How would you describe your child's general behaviour": "behaviour",
  "Where does your child prefer to play? e.g. outdoors": "play location",
  "Is your child used to playing with other children?": "used to other kids",
  "What is your child's preference when playing?": "play pref",
  "How talkative is your child when playing/interacting with others?": "talkative",
  "How does your child relate to adults and other children?": "relating",
  "Your child's reaction to unfamiliar adults?": "unfamiliar adults",
  "Your child's reaction in new situations?": "new situations",	
  "What is your child's response to animals?": "animals",
  "Does your child have siblings?": "siblings",
  "First sibling's name and age": "",
  "second sibling's name and age": "",
  "third sibling's name and age": "",
  "If your child has attended another preschool or child care centre, what is the name of the previous Centre?": "previous centre",
  "How many days per week did your child attend?": "days/wk",
  "When did they start?": "start",
  "Will they still be attending next year?": "next year",
  
  "Child's country of birth": "country",
  "Main language spoken by child": "language",
  "Child's cultural background": "child culture",
  "Mother's cultural background": "mother",
  "Father's cultural background": "father",
  "Would you like us to consider any cultural, religious or dietary requirements or additional needs?": "extra needs",
  "Is there any other food or drink that your child is not allowed to have other than for allergy or cultural reasons?": "food/drink",
  "Is MonTuesWed": "MonTueWed",					
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

/* Add child (values) to document (body).  Break up by headings. */
function addChild_(body, headings, values) {
  var lastName = lookup_(headings, values, "Child's last name");
  var firstName = lookup_(headings, values, "Child's first name");
  
  body.appendParagraph(lastName+ ", " + firstName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var p = body.appendParagraph("");

  swap(headings, values, swaps);
  
  for (var j=0; j<headings.length; j++) {
    var sectionHeading = sections[headings[j]];

    if (sectionHeading) {
      body.appendParagraph(sectionHeading).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      p = body.appendParagraph("");
    }
    if (values[j] == "") {
      continue;
    }
    var h = headings[j];
    if (h in short_heading_names) {
      h = short_heading_names[h];
    }
    if (h) {
      p.appendText(h+": ").setBold(true);
    }
    p.appendText(values[j] +"; ").setBold(false);
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

function lookup_(headings, values,  key){
  return values[headings.indexOf(key)];
}
