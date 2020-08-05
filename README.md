# Automated-Offer-Letters-Generator
Deployed using GCP (Cloud Functions, Storage, Apps Scripts and Stackdriver), leverages Google Sheets API, Google Drive API, Gmail API and HelloSign API to automate a pipeline to generate, send and save (when signed) offer letters based on authorization conditions and data input from many company teams on an internal database.

Pending to explain every step in the process for replication by the reader. Do's and Don'ts. (With animation ~ 1 min)

### SQL Database Schema
This is the schema included in the Google Sheet. If the parameters of the following schema are added as columns to a Google Sheet and the correct credentials provided, the project can be replicated in your local machine. The schema is also shown to provide further understanding of the workflow, allowing the reader to adapt this project to other scenarios.

| Name | Description | 
|   :---:      |     :---:      | 
| Hire_Full_Name   | (String) Full name of the individual to hire  | 
| Corporate_Email     | (String) Corporate email of the hire. If not available, any contact email works | 
| Job_Title   | (String) Role title  | 
| Company     | (String) There are many companies in the payroll. Indicates which one the hire corresponds to.  | 
| Manager_Name   | (String) Name of the manager hiring this individual  | 
| Team_Name     | (String) Name of the team the hire will belong to | 
| Start_Date   | (Date) Could also be a string as long as it has the form mm/dd/yyyy. Start date of the offer | 
| End_Date     | (Date) Could also be a string as long as it has the form mm/dd/yyyy. End date of the offer  | 
| Weekly_Salary   | (Number) Weekly salary of the hire  | 
| Hours_Per_Week     | (Number) Number of hours it is expected for the hire to work each week | 
| Location   | (String) Location where the hire is to work. Set whether HQ, satellite office or remote.  | 
| Senior_Team_Approval     | (Boolean) Whether the Senior Team has approved this hire.  | 
| GND_Approval   | (Boolean) Whether this team has approved this hire  | 
| Offer_Letter_Status     | (Categorical) Drop-down menu indicating the type of hire. Could be Contractor, Full-Time, Foreign Contractor  | 
| Previously_Hired   | (Boolean) Whether the hire has been in the organization before. Could be being hired again or renewing its offer.  | 
| Already_In_Payroll     | (Boolean) Whether the hire is in the payroll system already  | 
| Note   | (String) Notes to indicate special cases for each hire | 
| First_Pay_Date     | (Date) First pay date to incldue in the offer for the hire  | 
| Misc_Data   | (String) Generally used to store IDs for offer letter docs or the IDs for HelloSign signature requests  | 


### References

##### Google Apps Script

- Google Apps Scripts Homepage (https://developers.google.com/apps-script)
- Link Google Apps Scripts project with an existing Google Cloud Platform Project (https://developers.google.com/apps-script/guides/cloud-platform-projects)
- Manage external APIs from Google Apps Scripts (https://developers.google.com/apps-script/guides/services/external)

##### Google Sheets API
- Google Sheets API and Python documentation (https://developers.google.com/sheets/api/quickstart/python)
- Google Sheets API and Python with gspread library (https://www.youtube.com/watch?v=vISRn5qFrkM)

##### Google Drive API
- Google Drive API Python Quickstart (https://developers.google.com/drive/api/v3/quickstart/python)
- Documentation to upload and Download file data using Google Drive API (https://developers.google.com/drive/api/v3/manage-uploads)
- Google Drive API Documentation to manage files and folders (https://developers.google.com/drive/api/v2/folder)
- Obtain metadata for a file object (https://developers.google.com/drive/api/v3/reference/files/get)
- Google Drive API Python Getting Started Upload, Download, Create Files Folder (https://www.youtube.com/watch?v=9OYYgJUAw-w)
- Resolve Erros from Google Drive API (https://developers.google.com/drive/api/v3/handle-errors#resolve_a_404_error_file_not_found_fileid)

##### Hello Sign API
- HelloSign Python SDK Repo end Dcocumentation (https://github.com/hellosign/hellosign-python-sdk)
- HelloSign API Documentation (https://app.hellosign.com/api/documentation#QuickStart)

##### Google Cloud PLatform

- Python Quickstart on GCP (https://cloud.google.com/functions/docs/quickstart-python)
- Google Cloud Functions Execution Environment (https://cloud.google.com/functions/docs/concepts/exec)
- Environment Variables in GCP (https://cloud.google.com/functions/docs/env-var#functions-deploy-command-python)
- Manage files from Cloud Storage in Cloud Functions as blobs (https://hackersandslackers.com/manage-files-in-google-cloud-storage-with-python/) and (https://medium.com/@Tim_Ebbers/import-a-file-to-gcp-cloud-storage-using-cloud-functions-9cf81db353dc)
- “ImportError: file_cache is unavailable” when using Python client for Google service account file_cache (https://stackoverflow.com/questions/40154672/importerror-file-cache-is-unavailable-when-using-python-client-for-google-ser)
- Collect Logs for the Google Cloud Functions App (https://help.sumologic.com/07Sumo-Logic-Apps/06Google/Google_Cloud_Functions/Collect_Logs_for_the_Google_Cloud_Functions_App)
- Cloud Functions tima based automation using Cloud Scheduler on GCP (https://rominirani.com/google-cloud-functions-tutorial-using-the-cloud-scheduler-to-trigger-your-functions-756160a95c43)
- Configuring cron job schedule in Cloud Scheduler on GCP (https://cloud.google.com/scheduler/docs/configuring/cron-job-schedules?&_ga=2.106010307.-1333731722.1582023546&_gac=1.128890238.1589515976.Cj0KCQjw2PP1BRCiARIsAEqv-pRCa92FrNZH9CIDwMKaDpV5PIYeR2pKjAd4jbb1TDfeurPI3YnqX1IaAnYJEALw_wcB#defining_the_job_schedule)

