### Main Function
import datetime
from gdrive_functions import download_file, upload_file
from hellosign_functions import send_sign_req, download_completed_offer, send_reminder

### Main Function

def main(sheet, drive_service, sign_client, folder_id, ctr_folder_id, drafted_path, completed_path):
    """
    Input: sheet (Obj). Google Sheets Client File Object. Supposed to open the offers Tracker Sheet
            drive_service (Obj). Google Drive Client. Bridges the Script with an authorized Project with Google Drive API enabled
            sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests
            folder_id (String). Google Drive Folder ID to upload the completed files
            draft_path (String): Path in local machine to store drafted documents templates from Google Drive
            completed_path (String): Path in local machine to store the completed documents from Hello Sign
    
    Ouput: Depending on the conditions and status on the HR Spreadsheet
    - Downloads drafted documents to be signed from Googe Drive to our local machine on a specified path
    - Sends signature requests with document template through Hello Sign
    - Downloads completed documents from Hello SIgn to local machine
    - Uploads completed file to a specific folder in Google Drive
    """
    hires = sheet.get_all_values()
    today = datetime.datetime.today().strftime('%A')
    print(len(hires), today)
    
    for i in range(7,len(hires)):
        status = hires[i - 1][19]
        
        if status == 'Drafted':
            hire_name = hires[i - 1][2]
            hire_email = hires[i - 1][3]
            hire_position = hires[i - 1][5]
            # file_name = hire_name + '-' + hire_position
            file_id = hires[i - 1][26] # Brings the DocID number for the download
            
            # Downloads the file using the Google Drive API
            file_name = download_file(drive_service, file_id, drafted_path)
                
            # Sends the Document through Hello Sign
            sign_req = send_sign_req(sign_client, file_name, hire_name, hire_email, drafted_path)
            
            # Updates Status of the Offer Letter in the Google Sheet
            sheet.update_cell(i, 20, 'Sent')
            sheet.update_cell(i,27, sign_req.signature_request_id)
            
            print('Updated Offer Letter status for ', hire_name, hire_email)
            
            print("")
            print("---------------------------------------------------------------------------------")
            print("")
            
        elif status == 'Sent':
            sign_req_id = hires[i - 1][26]
            sign_req = sign_client.get_signature_request(sign_req_id)
            file_title = sign_req.title
            if sign_req.is_complete == True:
                
                # CHecks whether the offer is for a Contractor
                if file_title[:3] == 'CTR':
                    # Downloads Signed Offer Letter to local machine
                    download_completed_offer(sign_client, sign_req, file_title, sign_req_id, completed_path)

                    # Uploads Offer Letter to Google Drive using the API (to folder with folder_id)
                    moved_file_id = upload_file(file_title, ctr_folder_id, completed_path)
                    
                else:
                    # Downloads Signed Offer Letter to local machine
                    download_completed_offer(sign_client, sign_req, file_title, sign_req_id, completed_path)

                    # Uploads Offer Letter to Google Drive using the API (to folder with folder_id)
                    moved_file_id = upload_file(file_title, folder_id, completed_path)

                sheet.update_cell(i, 20, 'Signed')
                sheet.update_cell(i, 27, moved_file_id)
                
                print('Updated Offer Letter status for ', file_title)
                
                print("")
                print("---------------------------------------------------------------------------------")
                print("")
            
            elif sign_req.is_complete == False and today == 'Wednesday':
                hire_email = hires[i - 1][3]
                
                # Sends reminder to hires whose signatures are pending
                send_reminder(sign_client, sign_req_id, hire_email)
                
                print('Sent reminder for ', file_title)
                
                print("")
                print("---------------------------------------------------------------------------------")
                print("")
                
                
                
