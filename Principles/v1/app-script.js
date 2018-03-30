// To reference this library from the script of another document:
// 1. Open the other document 
// 2. Tools > Script Editor
// 3. In the Script Editor (we'll call this the Target Script), comment out or delete any existing code.
// 4. In the Target Script, go to Resources > Libraries
// 5. Get the Project Key for *this* project by going File > Project properties and look for Project Key (for Principle Code Master, the project key is: M8_cBPX5rhK-GY7wGvwW3ZYJhnl668_xO)
// 6. Paste the Project Key into the Add a Project input on the Libraries dialog on the Target Script
// 7. In the Libraries dialog, set a Version (choose the latest) and set Development Mode to "On". Development Mode allows the Target Script to call the latest update of this project, regardless of if it has been versioned
// 8. Insert the following snippet as the only code in the Target Script
// function onOpen(ev) {
//   return PrincipleCodeMaster.onOpen(ev);
// }
// For reference this is the project's key identifier: M8_cBPX5rhK-GY7wGvwW3ZYJhnl668_xO


function onOpen(ev) {
  var ui = DocumentApp.getUi();
  var menu1 = ui.createMenu('Principles Utilities')
  .addSubMenu(ui.createMenu('Insert Data Classification Label')
                   .addItem('Public', 'PrincipleCodeMaster.class1')
                   .addItem('Staff Confidential', 'PrincipleCodeMaster.class2')
                   .addItem('Workgroup Confidential', 'PrincipleCodeMaster.class3')
                   .addItem('Individual Confidential', 'PrincipleCodeMaster.class4'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Insert Risk Label')
                   .addItem('LOW', 'PrincipleCodeMaster.risk1')
                   .addItem('MEDIUM', 'PrincipleCodeMaster.risk2')
                   .addItem('HIGH', 'PrincipleCodeMaster.risk3')
                   .addItem('MAXIMUM', 'PrincipleCodeMaster.risk4'))
  .addSeparator()

  .addSubMenu(ui.createMenu('Insert Assessment')
                   .addItem('Consistently Followed', 'PrincipleCodeMaster.gradeConsistently')
                   .addItem('Generally Followed', 'PrincipleCodeMaster.gradeGenerally')
                   .addItem('Occasionally Followed', 'PrincipleCodeMaster.gradeOccasionally')
                   .addItem('Rarely Followed', 'PrincipleCodeMaster.gradeRarely')
                   .addItem('N/A', 'PrincipleCodeMaster.gradeNA'))

  menu1.addToUi();
}


// Data classification menus
function insertClassification(title1, color1) {
  var doc = DocumentApp.getActiveDocument();
  var c = doc.getCursor();
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = color1;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Open Sans';
  style[DocumentApp.Attribute.FONT_SIZE] = 10;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.UNDERLINE] = false;

  c.insertText(title1);
  
  var txt = c.getSurroundingText();
  var cpos_start = c.getSurroundingTextOffset();
  var cpos_end = cpos_start + title1.length-1;
  
  txt.setLinkUrl(cpos_start, cpos_end, 'https://wiki.mozilla.org/Security/Data_Classification');
  txt.setAttributes(cpos_start, cpos_end, style);
}

function insertGrade(title1, color1) {
  var doc = DocumentApp.getActiveDocument();
  var c = doc.getCursor();
  if (c == null) {
    var selection = doc.getSelection();   //check for cases where the user selected text instead of had a blinking cursor
    if (selection) {
      var elements = selection.getRangeElements();
      for (var i = 0; i < elements.length; i++) {    //iterate through selected elements
        var tempElement = elements[i];
        // DocumentApp.getUi().alert('1' + tempElement.getElement().getType());

        if (tempElement.getElement().getType() == DocumentApp.ElementType.TEXT || tempElement.getElement().getType() == DocumentApp.ElementType.TABLE_CELL) {
          tempElement.getElement().setText('');    //clear selected text -- works for either type TEXT or type TABLE_CELL

          var position = doc.newPosition(tempElement.getElement(), 0);  //set the cursor at the beginning of the text or beginning of the table cell
          doc.setCursor(position);
          c = doc.getCursor();
        }
      }
    }
  }
  if (c) {
    var style = {};
    // style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ffffff';
    // style[DocumentApp.Attribute.BACKGROUND_COLOR] = color1;
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Open Sans';
    style[DocumentApp.Attribute.FONT_SIZE] = 10;
    style[DocumentApp.Attribute.BOLD] = true;
    style[DocumentApp.Attribute.UNDERLINE] = false;
    c.insertText(title1);
  
    var txt = c.getSurroundingText();
    var cpos_start = c.getSurroundingTextOffset();
    var cpos_end = cpos_start + title1.length-1;
  
    txt.setAttributes(cpos_start, cpos_end, style);
  
    var ocell = c.getElement().getParent().asTableCell();
    // ocell.setBackgroundColor(color1); 
  }
  else {
    DocumentApp.getUi().alert('Cannot find a cursor. Try clearing the selection and try again.');
  }
    
}

function gradeConsistently() {
  insertGrade('Principle is consistently followed (>90% of the time)', '#000000');
}
function gradeGenerally() {
  insertGrade('Principle is generally followed (60-90% of the time)', '#000000');
}
function gradeOccasionally() {
  insertGrade('Principle is occasionally followed (15-60% of the time)', '#000000');
}
function gradeRarely() {
  insertGrade('Principle is rarely followed (<15% of the time)', '#000000');
}
function gradeNA() {
  insertGrade('N/A', '#000000');

}

function class1() {
  insertClassification('Public', '#7a7a7a');
}
function class2() {
  insertClassification('Staff Confidential', '#4a6785');
}
function class3() {
  insertClassification('Workgroup Confidential', '#ebbd30'); // this is not the standard level color, but the standard color is far too flashy in gdocs.
                                                             // this color code is very close but darker and does not look as flashy in gdocs.
}
function class4() {
  insertClassification('Individual Confidential', '#d04437');
}

// Risk label menus
function insertRisk(title1, color1) {
  var doc = DocumentApp.getActiveDocument();
  var c = doc.getCursor();
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = color1;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Open Sans';
  style[DocumentApp.Attribute.FONT_SIZE] = 10;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.UNDERLINE] = false;

  
  var txt = c.getSurroundingText();
  //c.insertText(title1);
  //var cpos_start = c.getSurroundingTextOffset();
  //var cpos_end = cpos_start + title1.length-1;
  // Override position so that its always at the start of the string, so that we can parse it properly
  cpos_start = 0;
  cpos_end = title1.length-1;
  txt.insertText(cpos_start, title1+' '); //add extra spacing for convenience
  txt.setLinkUrl(cpos_start, cpos_end, 'https://wiki.mozilla.org/Security/Standard_Levels');
  txt.setAttributes(cpos_start, cpos_end, style);
}

function risk1() {
  insertRisk('LOW', '#7a7a7a');
}
function risk2() {
  insertRisk('MEDIUM', '#4a6785');
}
function risk3() {
  insertRisk('HIGH', '#ebbd30'); // this is not the standard level color, but the standard color is far too flashy in gdocs
}
function risk4() {
  insertRisk('MAXIMUM', '#d04437');
}
