/// Functions on Google Apps Scripts to automatically generate offer letters based on a queue of requests from the company teams
/// These are also considered Cloud Functions since they are deployed on Time-Based Triggers on GCP

function mytest() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2020 Internships');
  var startRow = 7; // First row of data to process
  var numRows = sheet.getLastRow(); // Number of rows to process
  var dataRange = sheet.getRange(startRow, 3, numRows, 28);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  var today = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy");
  Logger.log(today);
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i],
        status = row[17],
        startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy");
        
        
    if(status != 'Signed' && today > startDate) {
      var name = row[0],
          title = row[3],
          manager = row[5],
          date = new Date(),
          newStartDate = new Date(date.setTime(date.getTime() + 7*86400000));
          // bestFirstPayDate = new Date(date.setTime(date.getTime() + 14*86400000));
          
      Logger.log(name);
      Logger.log(title);
      Logger.log(Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy"));
    }
  }
}


function createOfferDocs() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2020 Internships');
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
      
          // Normal Offer Letter Template
          var templateID = '13UKRBWA3GLhkxs7fOyKxcpC6COWIgnPfNcYoLfudh4A';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              salary = row[9],
              firstPayDate = Utilities.formatDate(new Date(row[23]), "GMT+1", "MM/dd/yyyy"),
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
              
          if(today > startDate) {
            var date = new Date(),
                newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
                startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy")
                // bestFirstPayDate = new Date(date.setTime(date.getTime() + 14*86400000));
                
            Logger.log(name);
            Logger.log(title);
            Logger.log(startDate);
          }
          
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
      
          // Instructor Training Participant Offer Letter Template
          var templateID = '11rCgnA-PSgnF7fN0H7p_G9EV4S_sZ5UDgqNBZJexSIc';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
              
          if(today > startDate) {
            var date = new Date(),
                newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
                startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy")
                // bestFirstPayDate = new Date(date.setTime(date.getTime() + 14*86400000));
                
            Logger.log(name);
            Logger.log(title);
            Logger.log(startDate);
          }
          
          var body = DocumentApp.openById(newDocID).getBody();
          body.replaceText('##TodaysDate##', today)
          body.replaceText('##FullName##', name)
          body.replaceText('##JobTitle##', title)
          body.replaceText('##ManagersName##', manager)
          body.replaceText('##Location##', location)
          body.replaceText('##StartDate##', startDate)
          body.replaceText('##EndDate##', endDate)
          
          DriveApp.getFileById(newDocID).setName(name + ' ' + title);
          
          sheet.getRange(startRow + i, 20).setValue('Drafted');
          sheet.getRange(startRow + i, 27).setValue(newDocID);
      
      }
      
      else if (status == 'Contractor') {
      
          // Contractor Offer Letter Template
          var templateID = '14xv8I2z5W1-EQ_JvuUt0z0OcIhStmQLxvmwBPF2SiZM';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              salary = row[9],
              firstPayDate = Utilities.formatDate(new Date(row[23]), "GMT+1", "MM/dd/yyyy"),
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
              
          if(today > startDate) {
            var date = new Date(),
                newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
                startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy")
                // bestFirstPayDate = new Date(date.setTime(date.getTime() + 14*86400000));
                
            Logger.log(name);
            Logger.log(title);
            Logger.log(startDate);
          }
          
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
          
          DriveApp.getFileById(newDocID).setName('CTR_' + name + ' ' + title);
          
          sheet.getRange(startRow + i, 20).setValue('Drafted');
          sheet.getRange(startRow + i, 27).setValue(newDocID);
          
               
      }
      
      else if (status == 'Contractor_ITP') {
      
          // Contractor Instructor Training Participant Offer Letter Template
          var templateID = '16LqPGaLgwtjhfE8BmQSDxog-_qxMVOoPuouEzyL1txI';
          var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();
      
          var name = row[0],
              title = row[3],
              manager = row[5],
              location = row[12],
              startDate = Utilities.formatDate(new Date(row[7]), "GMT+1", "MM/dd/yyyy"),
              endDate = Utilities.formatDate(new Date(row[8]), "GMT+1", "MM/dd/yyyy");
          
          if(today > startDate) {
            var date = new Date(),
                newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
                startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy")
                // bestFirstPayDate = new Date(date.setTime(date.getTime() + 14*86400000));
                
            Logger.log(name);
            Logger.log(title);
            Logger.log(startDate);
          }
          
          var body = DocumentApp.openById(newDocID).getBody();
          body.replaceText('##TodaysDate##', today)
          body.replaceText('##FullName##', name)
          body.replaceText('##JobTitle##', title)
          body.replaceText('##ManagersName##', manager)
          body.replaceText('##Location##', location)
          body.replaceText('##StartDate##', startDate)
          body.replaceText('##EndDate##', endDate)
          
          DriveApp.getFileById(newDocID).setName('CTR_' + name + ' ' + title);
          
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
        // posStatus = row[22].slice(0,8);
    
    if (notification !== 'NOTIFIED' && notification !== 'DEACTIVATED' && signed == 'Signed' && paycom == 'Yes') {  // Prevents sending duplicates
        
        var employeeName = row[0];
    
        var message = {
        to: row[1],
        subject: 'Minerva Summer Internship Onboarding Information',
        htmlBody: "<p class='c8'><span class='c4 c0 c2'>Dear " + employeeName + ",</span></p><p class='c1'><span class='c4 c0'></span></p><p class='c8 c2'><span class='c4 c0'>Congratulations on your Summer Internship Offer at Minerva Schools at KGI! </span></p><p class='c8 c2'><span class='c0'>As a member of the Finance Team, my co-worker Endrit (</span><span class='c9'><a class='c3' href='mailto:endrit@minerva.kgi.edu'>endrit@minerva.kgi.edu</a></span><span class='c4 c0'>) and I are here to help you in whatever you may need. To get started with your payroll onboarding, please read the following information carefully:</span></p><p class='c1'><span class='c4 c0'></span></p><ul class='c7 lst-kix_2eux5zgtgly2-0 start'><li class='c5 c2'><span class='c0'>Different than in previous Summer Internships and Work Study opportunities at Minerva, you will now be paid through a new payroll service called </span><span class='c9'><a class='c3' href='https://www.google.com/url?q=https://www.paycom.com/&amp;sa=D&amp;ust=1590098804675000'>Paycom</a></span><span class='c4 c0'>. In fact, Paycom will be replacing Zenefits and ADP going forward.</span></li><li class='c5 c2'><span class='c0'>You should be receiving (if you haven&#39;t yet) a username and password to log into your account in Paycom in the next few days. Follow the </span><span class='c9'><a class='c3' href='https://www.google.com/url?q=https://drive.google.com/file/d/18RTuMqZEI8CQvMPy9DSRnRk8FdwzQ5wQ/view?usp%3Dsharing&amp;sa=D&amp;ust=1590098804676000'>guidelines here</a></span><span class='c4 c0'>&nbsp;if needed.</span></li></ul><ul class='c7 lst-kix_2eux5zgtgly2-1 start'><li class='c8 c2 c13'><span class='c0'>You can also take a look at </span><span class='c9'><a class='c3' href='https://www.google.com/url?q=https://paycom.zoom.us/rec/share/4dJqJaj353pOTa-V1VHtWvQQJsPcX6a80yga-PcOnhnIAdBZ10djuBJte5Anq8rB&amp;sa=D&amp;ust=1590098804676000'>this video</a></span><span class='c0'>&nbsp;training. </span><span class='c6 c0'>Password:</span><span class='c0'>&nbsp;</span><span class='c0 c2'>1S*.4474</span></li></ul><ul class='c7 lst-kix_2eux5zgtgly2-0'><li class='c5 c2'><span class='c4 c0'>You may also receive a separate email from Paycom requesting you to complete the on-boarding checklist. Once you have received your temporary login credentials please complete the checklist and let us know whether there are mistakes in your pre-filled data (SSN, birthday, the spelling of your name, etc).</span></li></ul><p class='c8 c2 c10'><span class='c0 c15'>**Bank details must be entered 3 days before initial payday for payments to be processed.</span></p><ul class='c7 lst-kix_2eux5zgtgly2-0'><li class='c5 c2'><span class='c6 c0'>In the Employer I-9 Form section, please</span><span class='c6'>&nbsp;refer to </span><span class='c9 c6'><a class='c3' href='https://www.google.com/url?q=https://www.uscis.gov/i-9-central/acceptable-documents/list-documents/form-i-9-acceptable-documents&amp;sa=D&amp;ust=1590098804677000'>this resource</a></span><span class='c6'>&nbsp;for further information. List A documents are prefered. If multiple documents are necessary in your specific situation, you can merge them in a single PDF format document to upload </span><span class='c9 c6'><a class='c3' href='https://www.google.com/url?q=https://smallpdf.com/merge-pdf&amp;sa=D&amp;ust=1590098804678000'>here.</a></span></li><li class='c5 c2'><span class='c4 c0'>Please note you&#39;ll need to enter your bank details before payments can be processed.</span></li><li class='c5 c2'><span class='c0'>According to </span><span class='c9'><a class='c3' href='https://www.google.com/url?q=https://docs.google.com/spreadsheets/d/1ZMgC4QrAmFufr4RLwx7-gcmADy3kKfjqzHUAqEVZju0/edit?usp%3Dsharing&amp;sa=D&amp;ust=1590098804678000'>this calendar</a></span><span class='c0'>, you are to attest your weekly work every week no later than the Monday after the worked week. You can follow </span><span class='c9'><a class='c3' href='https://www.google.com/url?q=https://drive.google.com/file/d/16OT_k5S370IOkCNjhPVVql6OSPRzRpMG/view?usp%3Dsharing&amp;sa=D&amp;ust=1590098804679000'>these instructions</a></span><span class='c4 c0'>&nbsp;on how to attest your work on Paycom.</span></li></ul><p class='c1'><span class='c4 c0'></span></p><p class='c8 c2'><span class='c4 c0'>If you are working remotely or logging into Paycom from one of the following countries:</span></p><p class='c1'><span class='c4 c0'></span></p><ul class='c7 lst-kix_mrn8xge48l0s-0 start'><li class='c5 c2'><span class='c0 c4'>Russia</span></li><li class='c5 c2'><span class='c4 c0'>China</span></li><li class='c5 c2'><span class='c4 c0'>Taiwan</span></li><li class='c5 c2'><span class='c4 c0'>Nigeria</span></li><li class='c5 c2'><span class='c4 c0'>Iran</span></li><li class='c5 c2'><span class='c4 c0'>Syria</span></li><li class='c5 c2'><span class='c4 c0'>Turkey</span></li><li class='c5 c2'><span class='c4 c0'>Brazil</span></li><li class='c5 c2'><span class='c4 c0'>Hungary</span></li><li class='c2 c5'><span class='c4 c0'>Romania</span></li></ul><p class='c1'><span class='c4 c0'></span></p><p class='c8 c2'><span class='c0 c2'>Paycom&rsquo;s security feature blocks access to the above countries. Please </span><span>forward this email to Endrit Sylejmani (</span><span class='c9'><a class='c3' href='mailto:endrit@minerva.kgi.edu'>endrit@minerva.kgi.edu</a></span><span>)</span><span class='c0 c2'>&nbsp;with your Public IP address so we can request access for you. &nbsp;</span><span class='c9 c2'><a class='c3' href='https://www.google.com/url?q=https://www.whatismyip.com/what-is-my-public-ip-address/&amp;sa=D&amp;ust=1590098804680000'>This source</a></span><span class='c4 c0 c2'>&nbsp;may be useful for Windows, Mac, iOS and Android users. Note: you may experience issues trying to access Paycom from a different IP address (a location/device different than your usual workspace).</span></p><p class='c1'><span class='c4 c0 c2'></span></p><p class='c8 c2'><span class='c0 c2'>Feel free to take a look at our just created </span><span class='c9 c2'><a class='c3' href='https://www.google.com/url?q=https://docs.google.com/document/d/1xr0DPjpI0h8l4860m3gfEvtXTbWprNV4NnGYCjbaAOc/edit?usp%3Dsharing&amp;sa=D&amp;ust=1590098804681000'>FAQ</a></span><span class='c0 c2'>&nbsp;before asking directly to us. This document will keep growing and being updated. As a result, email questions that are already answered there are unlikely to be replied.</span></p><p class='c1'><span class='c4 c0'></span></p><p class='c1'><span class='c4 c0'></span></p><p class='c8 c2'><span class='c4 c0'>Do not hesitate in contacting us for any further questions or concerns.</span></p><p class='c1'><span class='c4 c0'></span></p><p class='c8 c2'><span class='c4 c0'>Best,</span></p><p class='c8 c11'><span class='c4 c16'></span></p>",
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
