from __future__ import print_function
import shutil
import xlrd
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
import os.path

# Google Auth variables
# https://developers.google.com/drive/api/quickstart/python
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/drive"]
creds = None
drive_service = None
sheets_service = None

uploadFilesToFolder = "1B6JrT9z7sGv9cvRyc3wqInlJlkJoFDgm"
googleSheetId = ""

inputFile = xlrd.open_workbook("C:/Users/Shrey/Downloads/nonsteel-cleaned-unsub.xls")
inputSheet = inputFile.sheet_by_index(0)
firstDataRow = 1
fileUrlColIndex = 4
fileUrlColAlphabet = "E"

# specify which column contains the LAN directory links to the files
# assuming this is the last column
# copiedFilePath = ""


def doGoogleAuth():
    global creds, drive_service, sheets_service
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
    sheets_service = build("sheets", "v4", credentials=creds)


def getFileFromLAN(url):  # maybe also pass file name based on po number from excel?
    source_path = r"\\mynetworkshare"
    dest_path = r"C:\TEMP"
    file_name = "\\myfile.txt"
    shutil.copyfile(source_path + file_name, dest_path + file_name)
    return True


def uploadFileToDrive(fileName, filePath, driveFolderId):
    global drive_service
    global creds

    try:

        file_metadata = {"name": fileName, "parents": [driveFolderId]}
        media = MediaFileUpload(filePath)
        file = (
            drive_service.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        return file.get("id")

    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f"An error occurred: {error}")
        return False


def addFileAndRowToExcel(rowNum, rowData):
    global sheets_service, fileUrlColIndex, fileUrlColAlphabet, uploadFilesToFolder, googleSheetId
    # row data is a list of what each row must contain (array)(each element in an xlrd Cell object)
    rowAsListArr = []
    for i in range(len(rowData)):
        if i <= fileUrlColIndex:
            rowAsListArr.appent(rowData[i].value)
    # call uploadFileToDrive
    fileName = "PO" + str(rowNum) + ".pdf"
    filePath = str(rowData[fileUrlColIndex]).split("'")[1]
    fileId = uploadFileToDrive(fileName, filePath, uploadFilesToFolder)
    # after file is uploaded replace rowData array element at index fileColumn with the google drive file url
    rowData[fileUrlColIndex] = fileId

    # upload the row to the set google sheet
    sheetEditRange = "Sheet1!A" + rowNum + ":" + fileUrlColAlphabet + rowNum
    sheetValues = [rowAsListArr]
    reqBody = {"values": sheetValues}
    result = (
        sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=googleSheetId,
            range=sheetEditRange,
            valueInputOption="USER_ENTERED",
            body=reqBody,
        )
        .execute()
    )
    print("{0} cells updated.".format(result.get("updatedCells")))


def start():

    global uploadFilesToFolder, inputSheet, firstDataRow, fileColumn
    print("Starting...")
    # do Google Auth
    doGoogleAuth()

    # test google auth
    # testGoogleAuth()
    # upload a file to a folder
    # fileId = uploadFileToDrive("google_test.py", "google_test.py", uploadFilesToFolder)

    # if fileId:
    #    print(fileId)

    print("This excel has " + str(inputSheet.nrows) + " rows")
    # for cur_row in range(firstDataRow, inputSheet.nrows):
    #     for numcol in range(fileColumn):
    #         print("Row")

    currentOutputRow = 0
    for rowIndex in range(firstDataRow, inputSheet.nrows):
        rowObj = inputSheet.row(rowIndex)
        print(rowObj)
        # comment the below lines and test, make sure output is a list where each element is of the form: text:"+919741307999"
        addFileAndRowToExcel(currentOutputRow, rowObj)
        currentOutputRow = currentOutputRow + 1

    # addFileAndRowToExcel


start()
