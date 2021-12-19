// Spreadsheet object of survey data
const SHEET = SpreadsheetApp.getActive().getSheetByName('Responses');

const SENT_EMAIL = "EMAIL_SENT";

// Confimation email, after forum is completed 
function sendEmail() {

  var data = getData();  // data = [ {} , {} , {} ]
  var i = 2;  // Initial start row index for 'Email_Stat' Column
  
  data.forEach(function(row){
    if(row.address && row.email_stat !== SENT_EMAIL) {
      // Send current row values to getMessage() -> returns an HTML body with current row values embedded in templateHTML
      var message = getMessage(row.name,row.gear, row.due);
      var recipientEmail = row.address
      var subject = "Gear Rental"
       MailApp.sendEmail(recipientEmail, subject, message, {htmlBody : message + "<BR/> <BR/>" });
       SHEET.getRange(i,9).setValue(SENT_EMAIL); // Set 'EMAIL_SENT' key to each address to ensure no duplicates 
    };
     i++;
  });
};
  
  //Overdue email reminder! On Seperate Event Trigger
  function overDueEmail () {
     var data = getData(); 
     
     data.forEach(function(row) {
      if(row.overdue_stat) {
       subject = "Overdue: VOCO Gear";
       message = getMessageOverdue(row.name,row.gear,row.due);
       MailApp.sendEmail(row.address, subject, message , {htmlBody : message});
     }
     });
     
  }


  function getMessage(name, gear, due_date ) {
    // Retrieve html template from 'email_templ.html' file
    var htmlOutput = HtmlService.createHtmlOutputFromFile("email_templ")
    var message = htmlOutput.getContent();
    
    // Replace message placeholders with values 
    message = message.replace("%name", name.split(" ")[0]).replace("%gear",gear).replace("%due_date",due_date);

    return message;

  }
  
  function getMessageOverdue (name,gear,due_date) {
    var htmlOutput = HtmlService.createHtmlOutputFromFile("overDueMessage");
    var message = htmlOutput.getContent();
    message = message.replace("%name", name).replace("%gear",gear).replace("%due_date", due_date);

    return message; 
  }

  // Function that reads and returns data from the 'Responses' Spreadsheet 
  function getData() {
    
    // Reads data from the Responses spreadsheet 
    var startRow = 2;  // Row 2
    var startCol = 1;  // Column A
  
    var values = SHEET.getRange(startRow, startCol, SHEET.getLastRow()-1 , SHEET.getLastColumn()).getValues(); 
    // values =  a list of lists, each list represents a row in the spreadsheet e.g. values = [ [colA, colB, . . . ] , [colA, colB, . . . ], ... ]
    

    // List will hold all rows as dictionary types 
    var data = [];

  // Iterate through each row, mapping rows values to correct key (gear, name, etc)
  values.forEach(function(value){

    // Intialize current dictionary 
    var row = {};

    row.name = value[1]; // Index for columns starts at 0 (0 : Timestamp, 1 : Name, 2: Gear, ...)
    row.gear = value[2];  // "" 
    row.checkout = Utilities.formatDate(new Date(value[3]), "GMT+1", "MM/dd/yyyy"); 
    row.due = Utilities.formatDate(new Date(value[4]), "GMT+1", "MM/dd/yyyy");    
    row.address = value[5]
    row.phone = value[6];
    row.email_stat = value[8]
    row.overdue_stat = value[10]

    data.push(row)
  });

  // data = [ {row1 values }, {row2 values} ... ]
  return data;
  }
  
