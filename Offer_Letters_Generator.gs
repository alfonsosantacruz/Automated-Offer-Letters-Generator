/// Functions on Google Apps Scripts to automatically generate offer letters based on a queue of requests from the company teams
/// These are also considered Cloud Functions since they are deployed on Time-Based Triggers on GCP

var fieldPositions = {
  studentEmail: 2,
  studentFullName: 3,
  citizenship: 4,
  wsEligible: 6,
  onRotation: 7,
  company: 8,
  role: 9,
  manager: 10,
  salaryRate: 11,
  projectCode: 12,
  ctdSignOff: 13,
  paycomID: 14,
  paycomDeptCode: 17,
  managerEmail: 18,
  avgHours: 21,
  positionUpdateBool: 22
};

var actionsByColumn = {
  2: "StudentEmail",
  3: "StudentFullName",
  4: "Citizenship",
  6: "Eligibility",
  7: "Rotation",
  8: "Company",
  9: "Role",
  10: "Manager",
  11: "Salary",
  12: "Project Code",
  13: "CTD Sign Off",
  14: "PaycomID",
  17: "Dept Code",
  18: "Manager Email",
  21: "Average Hours",
  22: "Position Update Bool",
  101: "Doc_Unsigned_Offer_Link",
  102: "PDF_Unsigned_Offer_Link",
  103: "HelloSign_Offer_UUID",
  104: "Signed_Offer_Link",
  105: "Onboarding_Email_Sent"
};

var currentDashboardSheetName = "Spring Dashboard";
var previousDashboardSheetName = "Dashboard";
var updatesSheetName = "Updates";
var projectCodesSheetName = "PaycomProjectCodes";
var checklistSheetName = "Offers Checklist";
var managersSheetName = "PaycomManagersByEmail";
var paycomAnalysisSheetName = "PaycomYTDAnalysis";
var formConfigSheetName = "Form Config";

// Google Docs IDs for templates
var contractorOfferTemplateFileId = "";
var standardOfferTemplateFileId = "";

var salaryRate = 20;
var officialStartDate = "1/8/2022";
var normalEndDate = "4/22/2022";
var seniorEndDate = "5/27/2022";
var currentGraduationYear = 2022;

// Google Drive IDs to folders for offer letters
var docDraftsFolderId = "";
var pdfDraftsFolderId = "";
var pdfSignedFolderId = "";

function createOfferDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName(checklistSheetName);
  var startRow = 2;
  var dataRange = ss.getRange(startRow, 1, ss.getLastRow(), ss.getLastColumn());
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  var today = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy")
  
  for (var i = 0; i < data.length; ++i) {
    var studentData = data[i],
        eligibility = studentData[3],
        status = studentData[12],
        offerType = studentData[13],
        docUnsingedOfferLink = studentData[14];

    var hasNoOfferLetterYet = docUnsingedOfferLink == ""

    if (status == "Official" && eligibility == "Yes" && hasNoOfferLetterYet) {      
      if (offerType == 'Standard') {
        generateOfferLetter(sheet, standardOfferTemplateFileId, studentData, i + 2, today);
      } else if (offerType == 'Contractor') {
        generateOfferLetter(sheet, contractorOfferTemplateFileId, studentData, i + 2, today);
      }
    }
  }
}

function generateOfferLetter(sheet, templateID, studentData, rowIndex, today) {
  var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();

  var name = studentData[0],
      expectedGraduationYear = studentData[2].toString(),
      title = studentData[6],
      manager = studentData[7],
      studentId = studentData[9],
      startDate = Utilities.formatDate(new Date(officialStartDate), "GMT+1", "MM/dd/yyyy"),
      fileName = name + ' ' + title;

  var studentsEndDate = expectedGraduationYear == currentGraduationYear ? seniorEndDate : normalEndDate;
  var endDate = Utilities.formatDate(new Date(studentsEndDate), "GMT+1", "MM/dd/yyyy");

  if(today > startDate) {
    var date = new Date(),
        newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
        startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy") 
  }
  
  var body = DocumentApp.openById(newDocID).getBody();
  body.replaceText('##TodaysDate##', today)
  body.replaceText('##FullName##', name)
  body.replaceText('##GraduationYear##', expectedGraduationYear.slice(2))
  body.replaceText('##JobTitle##', title)
  body.replaceText('##ManagersName##', manager)
  body.replaceText('##StartDate##', startDate)
  body.replaceText('##EndDate##', endDate)
  body.replaceText('##StudentID##', studentId)
  
  DriveApp.getFileById(newDocID).setName(fileName);
  DocumentApp.openById(newDocID).saveAndClose();

  var newDocHyperLink = `https://docs.google.com/document/d/${newDocID}/edit`

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, rowIndex, 101, newDocHyperLink);

  DriveApp.getFileById(newDocID).moveTo(DriveApp.getFolderById(docDraftsFolderId));

  granularConvertPDF(sheet, newDocID, rowIndex);

  SpreadsheetApp.flush();
}

function granularConvertPDF(sheet, newDocID, rowIndex) {          
  var newDoc = DriveApp.getFileById(newDocID);
  var docFolder = DriveApp.getFolderById(pdfDraftsFolderId);
  var docBlob = newDoc.getBlob().getAs('application/pdf');
  var newPDFFile = docFolder.createFile(docBlob);

  newPDFFile.setName(newDoc.getName() + ".pdf");
  
  var newPDFFileID = newPDFFile.getId();
  var newPDFHyperLink = `https://drive.google.com/file/d/${newPDFFileID}/view`;

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, rowIndex, 102, newPDFHyperLink);
  SpreadsheetApp.flush();
}
