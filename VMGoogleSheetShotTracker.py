from __future__ import print_function
from ast import Global
from glob import glob
from operator import truediv

import os.path
from pkgutil import get_data
from urllib import response

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import re

from datetime import datetime


import unreal

def get_unreal_project_name():
    input_string = unreal.Paths.project_dir()

    match = re.search(r'/([^/]+)/$', input_string)

    if match:
        extracted_string = match.group(1)
        return extracted_string
    else:
        return False

def set_google_doc_link(link):

    global SPREADSHEET_ID
    match = re.search(r'd/(.*?)/edit', link)
    if match:
        SPREADSHEET_ID = match.group(1)
    else:
        print("Google Link Broken")

def check_credential_and_token():

    global CRED

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIAL_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())


    CRED = creds
    return creds

def get_google_services():
    
    global SHEET_SERVICE, CRED, DRIVE_SERVICE

    try:
        SHEET_SERVICE = build('sheets', 'v4', credentials=CRED).spreadsheets()
        DRIVE_SERVICE = build('drive', 'v3', credentials=CRED)

    except HttpError as err:
        print(err)
        return False


def get_empty_line():
    global SHEET_SERVICE, SPREADSHEET_ID, EMPTY_LINE

    try:
        result = SHEET_SERVICE.values().get(
            spreadsheetId = SPREADSHEET_ID,
            range=f"{DATE}!A:A"
        ).execute()

        EMPTY_LINE = len(result["values"]) + 1
    except HttpError as err:
        print(err)

    return False

def get_data_from_scope():
    global SHEET_SERVICE, SPREADSHEET_ID, SAMPLE_RANGE_NAME

    if SPREADSHEET_ID:
        try:

            result = SHEET_SERVICE.values().get(spreadsheetId=SPREADSHEET_ID,
                                        range=SAMPLE_RANGE_NAME).execute()
            values = result.get('values', [])

            if not values:
                print('No data found.')
                return

            for row in values:
                # Print columns A and E, which correspond to indices 0 and 4.
                print(row)
        except HttpError as err:
            print(err)
    else:
        pass
        #show_message("Google sheet link is not valid!")




def get_project_shot_sheet():
    global CRED, PROJECT_NAME, SHARE_DRIVE_ID, DRIVE_SERVICE, SHOT_SHEET_TEMPLATE, SPREADSHEET_ID, SPREADSHEET_FILE

    drive_service = None

    # Check if there is a folder for the current Unreal project
    #   if not create the folder with the project name and copy the template sheet under the folder
    try:
        files = []

        response = DRIVE_SERVICE.files().list(
            includeItemsFromAllDrives = True,
            supportsAllDrives = True,
            q = f"parents in '{SHARE_DRIVE_ID}' and mimeType='application/vnd.google-apps.spreadsheet' and name='{PROJECT_NAME}' and trashed=False"
        ).execute()

        currentProjectSheet = response.get('files',[])

        if (not currentProjectSheet):
            folder_metadata = {
                'name': PROJECT_NAME,
                'parents': [SHARE_DRIVE_ID],
                'mimeType': 'application/vnd.google-apps.spreadsheet'
            }
            try:
                file = DRIVE_SERVICE.files().copy(
                    supportsAllDrives = True,
                    fileId=SHOT_SHEET_TEMPLATE, 
                    body=folder_metadata).execute()
                
                currentProjectSheet.append(file)

                # Set permission for the file
                try:
                    permission = {'type': 'domain',
                                    'domain': 'versatile.media',
                                    'role': 'writer'}
                    DRIVE_SERVICE.permissions().create(
                        supportsAllDrives = True,
                        fileId=file['id'],
                        body=permission).execute()
                except HttpError as error:
                        print('Error while setting permission:', error)

            except HttpError as e:
                print(f"创建文件夹时出错：{str(e)}")
            
        SPREADSHEET_ID = currentProjectSheet[0]['id']
        SPREADSHEET_FILE = currentProjectSheet[0]

    except HttpError as error:
        print(F'Get file error: {error}')

# Check if there is a sheet named with current date, for example: 20231007
# if not, duplicate the template sheet and rename it with current date
def get_today_shot_sesssion_sheet():
    global DATE, SPREADSHEET_ID, TODAY_SHEET

    spreadsheet = SHEET_SERVICE.get(spreadsheetId = SPREADSHEET_ID).execute()
    sheets = spreadsheet['sheets']
    
    today_sheet_id = None
    template_sheet = None

    for sheet in sheets:
        if sheet['properties']['title'] == DATE:
            today_sheet_id = sheet['properties']['sheetId']
            break
        if sheet['properties']['title'] == "template":
            template_sheet = sheet
    

    if not today_sheet_id:
        try:
            today_sheet = SHEET_SERVICE.sheets().copyTo(
                spreadsheetId = SPREADSHEET_ID,
                sheetId = template_sheet['properties']['sheetId'],
                body = {"destinationSpreadsheetId": SPREADSHEET_ID}
            ).execute()

            today_sheet_id =  today_sheet['sheetId']

            print("New today sheet")
            print(today_sheet)

            body = {
                'requests' : [{
                    "updateSheetProperties": {                            
                        "properties": {
                            "sheetId": today_sheet_id,
                            "title": DATE
                            },
                        "fields": "title"
                    }
                }]
            }

            SHEET_SERVICE.batchUpdate(
                spreadsheetId = SPREADSHEET_ID,
                body = body
            ).execute()
        except HttpError as error:
            print("Duplicate Sheet Error:")
            print(error)
    
    TODAY_SHEET = today_sheet_id
  

def initial_tracker():
    global SPREADSHEET_FILE, SPREADSHEET_ID, TODAY_SHEET
    total_frames= 100
    text_lable = "Shot Tracker: Updating Google Sheet..."

    with unreal.ScopedSlowTask(total_frames, text_lable) as slow_task:
        slow_task.make_dialog(True)
        slow_task.enter_progress_frame(20)
        check_credential_and_token()

        slow_task.enter_progress_frame(20)
        get_google_services()

        slow_task.enter_progress_frame(20)
        get_project_shot_sheet()

        slow_task.enter_progress_frame(20)
        get_today_shot_sesssion_sheet()

        slow_task.enter_progress_frame(20)
        get_empty_line()

    data = {
        "spreadsheet" : SPREADSHEET_FILE,
        "sheet" : TODAY_SHEET
    }
    return data

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = None
SPREADSHEET_FILE = None
TODAY_SHEET = None

# Credential and token file path
current_directory = os.path.dirname(os.path.abspath(__file__))
print(current_directory)


CREDENTIAL_FILE = os.path.join(current_directory, 'credentials.json')
TOKEN_FILE = os.path.join(current_directory, 'token.json')

# The actual google sheet service and cred object
SHEET_SERVICE = None
DRIVE_SERVICE = None
CRED = None

# Project name
SHARE_DRIVE_ID = "0AJF9RCQtJivqUk9PVA"
PROJECT_NAME = get_unreal_project_name()
SHOT_SHEET_TEMPLATE = "17tHIjNhYN5VTSvO6h_ECp6_uBoNgqS7-WrQ5mIdL21w"

# Date
DATE =  datetime.now().strftime("%Y%m%d")

# Current Cursor, the empty line
EMPTY_LINE = 6

# Get the google doc link under 
#set_google_doc_link("https://docs.google.com/spreadsheets/d/17tHIjNhYN5VTSvO6h_ECp6_uBoNgqS7-WrQ5mIdL21w/edit#gid=0")
#initial_tracker()



def Record_Started_AddShot(slatename, slatenum, filepath, totalframe):

    initial_tracker()

    #Modify data
    global SHEET_SERVICE, SPREADSHEET_ID, DATE, EMPTY_LINE
    try:
        values = [
            #cell values
            [
                slatename, slatenum , filepath, totalframe
            ]
        ]
        body = {
            'values' : values
        }
        result = SHEET_SERVICE.values().update(
            spreadsheetId = SPREADSHEET_ID,
            range = f"{DATE}!A{EMPTY_LINE}:D{EMPTY_LINE}",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        return result

    except HttpError as error:
        print("Modify data error:", error)
        return error