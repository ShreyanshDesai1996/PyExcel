from __future__ import print_function
from cgi import test
from operator import truediv
import shutil
import xlrd
import re
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
import os.path


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/drive"]
creds = None
drive_service = None

uploadFilesToFolder = ""
uploadSheetToFolder = ""

inputFile = xlrd.open_workbook("C:/Users/Shrey/Downloads/nonsteel-cleaned-unsub.xls")
inputSheet = inputFile.sheet_by_index(0)

# specify which column contains the LAN directory links to the files
fileColumn = 4
copiedFilePath = ""


def doGoogleAuth():
    global creds
    global drive_service
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    drive_service = build("drive", "v3", credentials=creds)


def getFileFromLAN(url):  # maybe also pass file name based on po number from excel?
    source_path = r"\\mynetworkshare"
    dest_path = r"C:\TEMP"
    file_name = "\\myfile.txt"
    shutil.copyfile(source_path + file_name, dest_path + file_name)
    return True


def uploadFileToDrive(fileName, filePath, driveFolderId):
    global drive_service
    driveUrl = ""
    # do google work
    print("File uploaded")

    global creds
    try:

        file_metadata = {"name": fileName, "parents": [driveFolderId]}
        media = MediaFileUpload(filePath)
        file = (
            drive_service.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        print("File ID: " + file.get("id"))

    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f"An error occurred: {error}")


def addFileAndRowToExcel(rowNum, rowData):
    # row data is an array of what each row must contain

    # call uploadFileToDrive

    # after file is uploaded replace rowData array element at index fileColumn with the google drive file url

    print("Added PO")


def start():
    print("reading now")
    # do Google Auth
    doGoogleAuth()

    # test google auth
    # testGoogleAuth()

    global uploadFilesToFolder
    uploadFilesToFolder = "1B6JrT9z7sGv9cvRyc3wqInlJlkJoFDgm"
    uploadFileToDrive("google_test.py", "google_test.py", uploadFilesToFolder)
    # read row of local excel

    # getFileFromLAN

    # addFileAndRowToExcel


start()
