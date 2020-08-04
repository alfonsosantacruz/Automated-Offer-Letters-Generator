### Google Drive API functions
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import io
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload


# Downloads the file using the Google Drive API
def download_file(drive_service, file_id, filename, drafted_path):
 
    request = drive_service.files().export_media(fileId = file_id, mimeType='application/pdf')

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
        
        
def upload_file(drive_service, file_name, folder_id, completed_path):
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
                                        fields='id, parents').execute()

    print('Uploaded Letter for ', file_name)
    
    return moved_file.get('id')
