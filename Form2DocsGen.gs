// this is an example code on Google Apps Script, I recommend to modify based your own
const spreadSheet = SpreadsheetApp.openById("(your form container Spredsheet ID)");
const sheet = spreadSheet.getSheetByName("Sheet Name of Your Form data");
const form = FormApp.openById("(Your Google Form ID)");
const folder = DriveApp.getFolderById("Your Drive Directory");

function onSubmit(e) {
  const questions = e.response.getGradableItemResponses();
  const timePart = questions[1].getResponse();
  const responseId = e.response.getId();
  const responses = form.getResponses();
  let row  = 1;
  let order = 0;
  for (const response of responses) {
    row++;
    if (response.getGradableItemResponses()[1].getResponse() === timePart) {
      order++;
    }
    if (response.getId() === responseId) {
      break;
    }
  }
  let message = "Form received\n\n";
  let Title = "";
  let DocID = "";
  let AuthorHandle = "";
  let Abstract = "";
  for (const question of questions) {
    const answer = question.getResponse();
    switch(question.getItem().getTitle()) {
      case "Title":
        Title = answer;
        break;
      case "AuthorHandle":
        AuthorHandle = answer;
        break;
      case "Abstract":
        Abstract = answer;
        break;
      default:
        //sheet.getRange(row, 4).setValue(answer);
        message += `【` +question.getItem().getTitle()+ `】\n ${answer}\n\n`;
        break;
    }
  }
  const timestamp = e.response.getTimestamp();
  sheet.getRange(row, 2).setValue(Utilities.formatDate(timestamp, "JST", "yyyyMMdd")); // ID is for the 2nd column
  DocID= CreateNewDoc(Title);
  sheet.getRange(row, 8).setValue("https://docs.google.com/document/d/"+DocID); // ID column (count your columns)
  ReplaceDoc(DocID,'{{Title}}', Title)
  ReplaceDoc(DocID,'{{AuthorHandle}}', AuthorHandle)
  ReplaceDoc(DocID,'{{Abstract}}', Abstract)
  const subject = "New Document is arrived";
  GmailApp.sendEmail("your mail address", subject, message);
}

function CreateNewDoc(title){
  var date = new Date();
  date.setDate(date.getDate()); // may be not unique if you submit several times at a day!
  var formattedDate = Utilities.formatDate(date, "JST", "yyyyMMdd");
  var fileName = title;
  var sourcefile = DriveApp.getFileById("(template file ID)"); // give your template doc file
  newfile = sourcefile.makeCopy( formattedDate + "-(temp)" + fileName);
  var doc = DocumentApp.openById(newfile.getId());
  Logger.log(doc.getName());
  return doc.getId();
}

function ReplaceDoc(_docID, _Keyword, _NewText) {
  let doc = DocumentApp.openById(_docID)
  let docbody = doc.getBody(); 
  docbody.replaceText(_Keyword,_NewText);  
  doc.saveAndClose(); 
}
