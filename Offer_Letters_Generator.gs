/// Functions on Google Apps Scripts to automatically generate offer letters based on a queue of requests from the company teams
/// These are also considered Cloud Functions since they are deployed on Time-Based Triggers on GCP

function createOfferDocs() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name of the Spreadsheet you started the Google Apps Script Project on');
  var startRow = 7; // First row of data to process
  var numRows = sheet.getLastRow(); // Number of rows to process
  var dataRange = sheet.getRange(startRow, 3, numRows, 28);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  var today = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy")
  
  for (var i = 0; i < data.length; ++i) {
  
      var row = data[i],
          status = row[17];
          
      if (status == 'Compiled') {
      
          // Normal Offer Letter Template grom Google Docs
          var templateID = '***********************************************';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              salary = row[9],
              firstPayDate = Utilities.formatDate(new Date(row[23]), "GMT+1", "MM/dd/yyyy"),
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
          
          var body = DocumentApp.openById(newDocID).getBody();
          body.replaceText('##TodaysDate##', today)
          body.replaceText('##FullName##', name)
          body.replaceText('##JobTitle##', title)
          body.replaceText('##ManagersName##', manager)
          body.replaceText('##Location##', location)
          body.replaceText('##Salary##', salary)
          body.replaceText('##FirstPayDate##', firstPayDate)
          body.replaceText('##StartDate##', startDate)
          body.replaceText('##EndDate##', endDate)
          
          DriveApp.getFileById(newDocID).setName(name + ' ' + title);
          
          sheet.getRange(startRow + i, 20).setValue('Drafted');
          sheet.getRange(startRow + i, 27).setValue(newDocID);
          
      }
      
      else if (status == 'ITP') {
      
          // Instructor Training Participant Offer Letter Template from Google Docs
          var templateID = '************************************************';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              salary = row[9],
              firstPayDate = Utilities.formatDate(new Date(row[23]), "GMT+1", "MM/dd/yyyy"),
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
          
          var body = DocumentApp.openById(newDocID).getBody();
          body.replaceText('##TodaysDate##', today)
          body.replaceText('##FullName##', name)
          body.replaceText('##JobTitle##', title)
          body.replaceText('##ManagersName##', manager)
          body.replaceText('##Location##', location)
          body.replaceText('##FirstPayDate##', firstPayDate)
          body.replaceText('##StartDate##', startDate)
          body.replaceText('##EndDate##', endDate)
          
          DriveApp.getFileById(newDocID).setName(name + ' ' + title);
          
          sheet.getRange(startRow + i, 20).setValue('Drafted');
          sheet.getRange(startRow + i, 27).setValue(newDocID);
      
    }
  }
}



function convertPDF() {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2020 Internships');
    var startRow = 7; // First row of data to process
    var numRows = sheet.getLastRow(); // Number of rows to process
    var dataRange = sheet.getRange(startRow, 3, numRows, 28);
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    
    
    for (var i = 0; i < data.length; ++i) {
        var row = data[i],
            status = row[17],
            newDocID = row[24];
        
        if (status == 'Drafted' && newDocID !== '') {
        
            var newDoc = DriveApp.getFileById(newDocID);
            var docFolder = DriveApp.getFileById(newDocID).getParents().next().getParents().next();
            var docBlob = newDoc.getBlob().getAs('application/pdf');
            var newPDFFile = docFolder.createFile(docBlob);
        
            newPDFFile.setName(newDoc.getName() + ".pdf");
            
            // docFolder.removeFile(newDoc);
            // sheet.getRange(startRow + i, 27).setValue(newPDFFile.getId()); // Not needed anymore since the Python Script downloads from Google Docs format and not PDF format
            
            
    }
  }
}




function notify_PaycomOnboarding() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2020 Internships');
  var startRow = 7; // First row of data to process
  var numRows = sheet.getLastRow(); // Number of rows to process
  var dataRange = sheet.getRange(startRow, 3, numRows, 28);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i],
        notification = row[24], // Empty column next to messages
        signed = row[17],
        paycom = row[19]; 
      
    
      
    if (notification !== 'NOTIFIED' && signed == 'Signed' && paycom == 'Yes') {  // Prevents sending duplicates
    
        var employeeName = row[0];
    
        var message = {
        to: row[1],
        subject: 'Minerva Summer Internship Onboarding Information',
        htmlBody: "Insert Onboarding/Welcome Message in HTML format. Excluded in this file for confidentiallity purposes",
        name: "Automatic Paycom Onboarding Email"
        };
        
        MailApp.sendEmail(message);
        sheet.getRange(startRow + i, 27).setValue('NOTIFIED');
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
    
  }     
 }
}
 


function moveFileId(fileId, toFolderId) {
   var file = DriveApp.getFileById(fileId);
   var source_folder = DriveApp.getFileById(fileId).getParents().next();
   var folder = DriveApp.getFolderById(toFolderId)
   folder.addFile(file);
   source_folder.removeFile(file);
}
