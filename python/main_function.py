### Main Function
import datetime
from gdrive_functions import download_file, upload_file
from hellosign_functions import send_sign_req, download_completed_offer, send_reminder

### Main Function

def main(updates_sheet, checklist_sheet, drive_service, sign_client, folder_id, drafted_path, completed_path):
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
    hires = checklist_sheet.get_all_values()
    num_rows_updates_sheet = len(updates_sheet.get_all_values())
    # day_of_week = datetime.date.today().strftime('%A')
    today = datetime.date.today().strftime("%m/%d/%Y")
    
    for i in range(1, len(hires)):
        hire_name = hires[i][0]
        hire_email = hires[i][1]
        paycomID = hires[i][8]
        
        # Brings the DocID number for the download
        file_url = hires[i][14]
        hellosign_offer_uuid = hires[i][16]
        signed_offer_url = hires[i][17]
        
        if file_url and not hellosign_offer_uuid:
            file_id = str(file_url.split("/")[5])
            
            print(hire_name, hire_email, paycomID)
            
            # Downloads the file using the Google Drive API
            file_name = download_file(drive_service, file_id, drafted_path)
                            
            # Sends the Document through Hello Sign
            sign_req = send_sign_req(sign_client, file_name, hire_name, hire_email, drafted_path)
            
            # Updates Status of the Offer Letter in the Google Sheet
            post_updated_info_as_update(updates_sheet,
                                num_rows_updates_sheet + 1,
                                today, 
                                hire_email, 
                                hire_name, 
                                paycomID, 
                                103, 
                                sign_req.signature_request_id)
            num_rows_updates_sheet += 1
            
            print('Updated Offer Letter status for', hire_name, hire_email)
            
            time.sleep(7)
            
            print("")
            print("---------------------------------------------------------------------------------")
            print("")
            
        elif file_url and hellosign_offer_uuid and not signed_offer_url:
            sign_req = sign_client.get_signature_request(hellosign_offer_uuid)
            sign_req_created_datetime = datetime.date.fromtimestamp(sign_req.created_at)
            file_title = sign_req.title
            if sign_req.is_complete == True:
                # Downloads Signed Offer Letter to local machine
                download_completed_offer(sign_client, 
                                         sign_req, 
                                         file_title, 
                                         hellosign_offer_uuid, 
                                         completed_path)

                # Uploads Offer Letter to Google Drive using the API (to folder with folder_id)
                signed_file_id = upload_file(drive_service, file_title, folder_id, completed_path)

                signed_file_url = f"https://drive.google.com/file/d/{signed_file_id}/view"
                
                post_updated_info_as_update(updates_sheet,
                                num_rows_updates_sheet + 1,
                                today, 
                                hire_email, 
                                hire_name, 
                                paycomID, 
                                104, 
                                signed_file_url)
                num_rows_updates_sheet += 1
                
                print('Updated Offer Letter status for ', file_title)
                
                time.sleep(7)
                
                print("")
                print("---------------------------------------------------------------------------------")
                print("")
            
            elif abs((sign_req_created_datetime - datetime.date.today()).days) >= 1:
                hire_email = hires[i][1]
                
                # print(sign_req_created_datetime)
                # Sends reminder to hires whose signatures are pending
                send_reminder(sign_client, hellosign_offer_uuid, hire_email)

                print('Sent reminder for ', file_title)

                print("")
                print("---------------------------------------------------------------------------------")
                print("")
