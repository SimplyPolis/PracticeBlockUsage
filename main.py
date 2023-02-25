
from __future__ import print_function


import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import sys



argslen = len(sys.argv) #Handle args
Outfit=''
HowFar = 2
try:
    Outfit=sys.argv[1]
    if (argslen==3):
        HowFar=int(sys.argv[2])
except:
    print("You messed up the args")
    exit(1)

RANGE_NAME = "Sheet1!D7:D"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly','https://www.googleapis.com/auth/drive']



creds = None
if os.path.exists('token.json'): #Try to use a token
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # If it can not refresh grab from credentials file and web browser auth
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())


try:
        # create drive api client
        driveservice = build('drive', 'v3', credentials=creds) #services
        sheetsservice = build('sheets', 'v4', credentials=creds)
        files = []
        page_token = None

        date=(datetime.datetime.utcnow()-datetime.timedelta(weeks=4*HowFar)) #date calculations
        date=date.replace(microsecond=0).isoformat()
        while True:
            # pylint: disable=maybe-no-member
            response = driveservice.files().list(q="name contains '{}' and modifiedTime > '{}' and mimeType = 'application/vnd.google-apps.spreadsheet'".format(Outfit,date), #drive quarry
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(id, name, parents)',
                                            orderBy='modifiedTime desc',
                                            pageToken=page_token).execute()
            for file in response.get('files', []): #Go through files
                if(file.get("parents").count('14_D0GBku4KNJyDPbs4Zb2bL-EbymNRcR')==0 and file.get("parents").count('1cQx3hBZcHFqwE47CbHtciKu4ZWDzyWMw')==0): #Not a practice block
                    sheet = sheetsservice.spreadsheets()
                    result = sheet.values().get(spreadsheetId=file.get("id"),
                                                range=RANGE_NAME).execute()
                    values = result.get('values', []) #Get sheet
                    print(file.get("name"),"Uses:"+str(len(values)),"Link: https://docs.google.com/spreadsheets/d/"+file.get("id")) #print results
            files.extend(response.get('files', [])) #prevent loop
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

except HttpError as error: #something bad has happend
        print(F'An error occurred: {error}')
        files = None




