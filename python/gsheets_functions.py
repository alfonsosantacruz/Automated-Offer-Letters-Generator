### Google Sheets API functions

# Posts update in updates sheet
# Inspired from postUpdatedStudentInfoAsUpdate function in the Apps Script implementation

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
        
