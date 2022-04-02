### Imports libraries and methods to initialize Google Drive API service
from __future__ import print_function
import pickle
import os.path
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
from gdrive_functions import download_file, upload_file
from hellosign_functions import send_sign_req, download_completed_offer, send_reminder
from main_function import main
import datetime

print("Imported libraries")

# Initializes a client using my corporate email API Key
# Obtain API Key from your Hello Sign account
sign_client = HSClient(api_key='****************************************************')

print("Defined Hello Sign Client")


# Initializes Google Sheets API Service from client
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# Path to your JSON file that includes all the API credentials once enabled on GCP
creds = ServiceAccountCredentials.from_json_keyfile_name('********************************.json', scope)
client = gspread.authorize(creds)

# Obtains a Google Sheets service
sheet = client.open_by_key(sheet_file_id)
updates_sheet = sheet.worksheet("Updates")
checklist_sheet = sheet.worksheet("Offers Checklist")

print("Defined Google Sheets Service")


# Initializes Google Drive API Service from client
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = None

# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_id.json', SCOPES)
        creds = flow.run_local_server(port=0)
    
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# Obtains a Google Drive service
drive_service = build('drive', 'v3', credentials=creds)

print("Defined Google Drive Service")

# Folder used to store the uploaded signed offer letters
folder_id = '****************************' # The ID of the folder to store the signed offer letters
ctr_folder_id = '*********************************' # Where to send contractors signed offer letters
drafted_path = '../DraftedOfferLettersByScript/#' # Adds # to identify those are drafted letters
completed_path = '../SignedOfferLettersByScript/'

main(updates_sheet, checklist_sheet, drive_service, sign_client, folder_id, drafted_path, completed_path)
