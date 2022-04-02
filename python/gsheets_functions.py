### Google Sheets API functions

# Posts update in updates sheet
# Inspired from postUpdatedStudentInfoAsUpdate function in the Apps Script implementation

# Code of action per column in the Google Sheet
actions_by_column = {
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
  101: "Doc_Unsigned_Offer_Link",
  102: "PDF_Unsigned_Offer_Link",
  103: "HelloSign_Offer_UUID",
  104: "Signed_Offer_Link",
  105: "Onboarding_Email_Sent"
}

def post_updated_info_as_update(updates_sheet,
                                last_row_index,
                                today, 
                                email, 
                                name, 
                                paycomID, 
                                action_num, 
                                value):
    if value:
        action_code = actions_by_column[action_num]
        
        updates_sheet.update_cell(last_row_index, 1, today)
        updates_sheet.update_cell(last_row_index, 2, email)
        updates_sheet.update_cell(last_row_index, 3, name)
        updates_sheet.update_cell(last_row_index, 4, paycomID)
        updates_sheet.update_cell(last_row_index, 5, action_code)
        updates_sheet.update_cell(last_row_index, 6, value)
        
