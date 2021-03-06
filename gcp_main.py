### Imports libraries and methods to initialize Google Drive API service
from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import io
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

### Imports Libraries and methods to Initialize Google Sheets API Service
import gspread
from oauth2client.service_account import ServiceAccountCredentials

### Imports hellosign SDK and Initializes Client from HelloSign API
from hellosign_sdk import HSClient

### Other libraries
import datetime
from google.cloud import storage

print("Imported libraries")


#######################################################################
##################### Google Drive API functions ######################
#######################################################################

# Downloads the file using the Google Drive API
def download_file(drive_service, file_id, drafted_path):
    """
    Downloads Google Doc file from Google Drive to an specified directory path in .pdf format.
    Input:
          drive_service (Obj): Google Drive API client service after verifying credentials
          file_id (String): Target Google Docs file ID.
          filename (String): File name with which the Google Doc will be saved
          drafted_path (String): local/cloud storage directory path where the downloaded file will be stored.
    Output: filename (String): The original name of the downloaded file.
    """
 
    # HTTP Request to download a Google Doc file as pdf format
    request = drive_service.files().export_media(fileId = file_id, mimeType='application/pdf')
    # HTTP GET Request to obtain the metadata of the downloaded file
    filename = drive_service.files().get(fileId = file_id).execute()['name']

    # Saves the file in our local machine to send for signing
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100), filename)
    with io.open(drafted_path + filename + '.pdf', 'wb') as f:
        fh.seek(0)
        f.write(fh.read())
        
    return filename
        
        
def upload_file(drive_service, file_name, folder_id, completed_path):
    """
    Uploads a .pdf file from local/cloud directory to an specified folder in Google Drive.
    Input:
          drive_service (Obj): Google Drive API client service after verifying credentials
          file_name (String): File name with which the .pdf file will be saved in Google Drive
          file_id (String): Target Google Drive folder ID.
          completed_path (String): Local/cloud storage directory path where the file will be retrieved to then be stored in Google Drive.
    Output:
          moved_file_id (String): The id of the uploaded file in the target directory
    """
    file_metadata = {'name': file_name + '.pdf'}
    media = MediaFileUpload(completed_path + file_name + '.pdf', 
                            mimetype = 'application/pdf')
    
    file = drive_service.files().create(body = file_metadata,
                                       media_body = media,
                                       fields = 'id, parents').execute()

    # Move the file to the new folder
    moved_file = drive_service.files().update(fileId = file.get('id'),
                                        addParents = folder_id,
                                        removeParents = file.get('parents')[0],
                                        fields = 'id, parents').execute()

    print('Uploaded Letter for ', file_name)
    
    return moved_file.get('id')



#######################################################################
##################### Hello Sign API Functions ########################
#######################################################################

def send_sign_req(sign_client, filename, name, email, drafted_path, cfo_email, cfo_name):
    """
    Input:
          sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests
          filename (String): The name of the file to be sent through Hello Sign
          name (String): The name of the new hire to sign the offer letter via Hello Sign
          email (String): The email address of the new hire to sign the offer letter via Hello Sign
          drafted_path (String): The local/cloud path where the to-be-sent file can be retrieved.
          cfo_email: The email address of the CFO (company authority) to sign the offer letter via Hello Sign
          cfo_name: The name of the CFO (company authority) to sign the offer letter via Hello Sign
    Output:
          sign_req (Obj): Hello Sign signature request object
    """
    # Sends the Document through Hello Sign
    sign_req = sign_client.send_signature_request(
        test_mode = False,
        title = filename + ' Summer Internship 2020 Offer Letter',
        subject = filename + ' Summer Internship 2020 Offer Letter',
        message = 'Hi, Please review and sign the offer letter through Hello Sign as part of your Summer 2020 Internship Position at Minerva. Best!',
        signers = [
            {'email_address': email, 'name': name},
            # CFO Email is a must in each letter.
            {'email_address': cfo_email, 'name': cfo_name} # ENV VAR
        ],
        use_text_tags = True,
        hide_text_tags = True,
        files = [drafted_path + filename + '.pdf'])

    print('Used HelloSign to send Letter to ', name, email)
    
    return sign_req


def download_completed_offer(sign_client, sign_req, file_title, sign_req_id, completed_path):
    """
    Downloads completed offer letters from Hello Sign
    Input: 
          sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests
          sign_req (Obj): Hello Sign signature request object of the offer letter to be downloaded
          file_title (String): The name that will be given to the downloaded file
          sign_req_id (String): Signature request ID of the offer letter to be downloaded
          completed_path (String): Path to the local/cloud storage directory to store the downloaded offer letter.
    """
    sign_client.get_signature_request_file(signature_request_id = sign_req_id,
                                                        filename = completed_path + file_title + '.pdf',
                                                        file_type ='pdf')

    
    
def send_reminder(sign_client, sign_req_id, hire_email):
    """
    Sends signature reminder to specific hires.
    Input: 
         sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests.
         sign_req_id (String): Signature request ID of the offer letter to be downloaded.
         hire_email (String): The email address of the hire to be reminded.
    """
    sign_client.remind_signature_request(signature_request_id = sign_req_id, email_address = hire_email)



#######################################################################
##################### Clients and Services Init #######################
#######################################################################

def services_init(hello_sign_api_key, sheet_name, bucket_name):
    """
    Initializes all the services from the used APIs
    Input: 
          hello_sign_api_key (String): API Key for the Hello Sign Client. Obtained from the websiste
          sheet_name (String): The name of the Google Sheet requesting access to
          bucket_name (String): The name of the Cloud Storage bucket where the APIs credentials files are stored.
    Output:
          sheet (Obj). Google Sheets Client File Object. Supposed to open the offers Tracker Sheet
          drive_service (Obj). Google Drive Client. Bridges the Script with an authorized Project with Google Drive API enabled
          sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests
    """      
    # Initializes a client using my corporate email API Key
    sign_client = HSClient(api_key = hello_sign_api_key) #ENV VAR

    print("Defined Hello Sign Client")

    # Initializes Cloud Storage client to access the Google Sheets API and Google Drive API clients credentials files
    # In the same GCP project, a storage bucket was created and these credentials files uploaded
    # The following lines of code access these files and downloads them to the temporary Cloud Function directory
    storage_client = storage.Client()
    bucket = storage_client.get_bucket(bucket_name)
    blobs = bucket.list_blobs(delimiter = '/')

    for blob in blobs:
        destination_uri = '/tmp/{}'.format(blob.name) 
        blob.download_to_filename(destination_uri)

    print("Downloaded Google Drive API and Google Sheets API .json and .pickle clients files")

    # Initializes Google Sheets API Service from client
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(opla_path, scope) #ENV VAR
    client = gspread.authorize(creds)
    
    # Obtains a Google Sheets service
    sheet = client.open(sheet_name).sheet1 #ENV VAR

    print("Defined Google Sheets Service")

    # Initializes Google Drive API Service from client
    SCOPES = ['https://www.googleapis.com/auth/drive']
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    #if os.path.exists('/tmp/tokendownload2.pickle'):
    with open(token_path, 'rb') as token:
        creds = pickle.load(token)

    print("Obtained token pickle file")

    # # Generates token file for the first time if not found in bucket blob
    # # If there are no (valid) credentials available, let the user log in.
    # if not creds or not creds.valid:
    #     if creds and creds.expired and creds.refresh_token:
    #         creds.refresh(Request())
    #     else:
    #         flow = InstalledAppFlow.from_client_secrets_file(
    #             '/tmp/client_id.json', SCOPES)
    #         creds = flow.run_local_server(port=0)
        
    #     # Save the credentials for the next run
    #     with open('/tmp/tokendownload2.pickle', 'wb') as token:
    #         pickle.dump(creds, token)

    # Obtains a Google Drive service
    drive_service = build('drive', 'v3', credentials = creds, cache_discovery = False)

    print("Defined Google Drive Service")

    return sheet, drive_service, sign_client


#######################################################################
###### Main function that performs the purpose of the pipeline ########
#######################################################################

def main(sheet, drive_service, sign_client, folder_id, ctr_folder_id, drafted_path, completed_path, cfo_email, cfo_name):
    """
    Input: sheet (Obj). Google Sheets Client File Object. Supposed to open the offers Tracker Sheet
            drive_service (Obj). Google Drive Client. Bridges the Script with an authorized Project with Google Drive API enabled
            sign_client (Obj). Hello Sign Client to make the documents completion verification and send signature requests
            folder_id (String). Google Drive Folder ID to upload the completed files
            folder_id (String). Google Drive Folder ID to upload the completed files for contractors
            draft_path (String): Path in local machine to store drafted documents templates from Google Drive
            completed_path (String): Path in local machine to store the completed documents from Hello Sign
            cfo_email (String): The email address of the company's CFO, who is the official signer of these offer letters
            cfo_name (String): The name of the company's CFO, who is the official signer of these offer letters
    
    Output: Depending on the conditions and status on the HR Spreadsheet
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
            filename = download_file(drive_service, file_id, drafted_path)
                
            # Sends the Document through Hello Sign
            sign_req = send_sign_req(sign_client, filename, name, email, drafted_path, cfo_email, cfo_name)
            
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
            if sign_req.is_complete == True:
                file_title = sign_req.title
                
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
            
            elif sign_req.is_complete == False and today == 'Monday':
                hire_email = hires[i - 1][3]
                
                # Sends reminder to hires whose signatures are pending
                send_reminder(sign_client, sign_req_id, hire_email)
                
                print('Sent reminder for ', sign_req.title)
                
                print("")
                print("---------------------------------------------------------------------------------")
                print("")



#######################################################################
####################### Main Handler Function #########################
#######################################################################


def handler(request):
    """
    Function to be triggered by the Cloud Function.
    Input: An HTTP Request. Preferably a POST request.
    Output: None. Proceeds with initializing the APIs services and unfold the pipeline process
    """
    if request.method == 'POST':
        print('Starting handler function')
        hello_sign_api_key = os.getenv('HELLO_SIGN_API_KEY')
        sheet_name = os.getenv('SHEET_FILE_NAME')

        # Folder used to store the uploaded signed offer letters
        folder_id = os.getenv('DRIVE_FOLDER_ID')
        ctr_folder_id = os.getenv('CTR_FOLDER_ID')
        
        # Bucket in Google Cloud Storage where the token.pkl and client.json files are
        bucket_name = os.getenv('BUCKET_NAME')

        # Paths with files names where the token.pkl and client.json are stored (Generally /tmp/ + filename)
        # Stored as ENV VAR to allow quick shift to direct Cloud Storage path if invocations exceed cloud function memory limit
        opla_path = os.getenv('OPLA_PATH')
        token_path = os.getenv('TOKEN_PATH')

        print('Successfully retrieved ENV VARs')

        sheet, drive_service, sign_client = services_init(hello_sign_api_key, sheet_name, bucket_name, opla_path, token_path)

        # Adds # to identify those are drafted letters
        # Might be reptitions for contractors with multiple positions
        drafted_path = '/tmp/#'
        completed_path = '/tmp/'

        # Adds COmpany signer as ENV VAR for privacy and quick switch if official signer is changed
        cfo_email, cfo_name = os.getenv('CFO_EMAIL'), os.getenv('CFO_NAME')

        main(sheet, drive_service, sign_client, folder_id, ctr_folder_id, drafted_path, completed_path, cfo_email, cfo_name) 
