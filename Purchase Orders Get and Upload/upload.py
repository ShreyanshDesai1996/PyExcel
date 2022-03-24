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


# Variables to set before running
uploadFilesToFolder = "1S55SU2XyfIRmfjYhfsW31vy7kHj5Ju4c"
googleSheetId = "1gMe0WxcE7H2d0bGEJmVc8LStP25HITmEzFRKIKbSUJE"

inputFile = xlrd.open_workbook("C:/Users/Shrey/Downloads/pos.xlsx")
inputSheet = inputFile.sheet_by_index(0)
firstDataRow = 1
fileUrlColIndex = 7  # specify which column contains the LAN directory links to the files, this program assumes it is the last column
fileUrlColAlphabet = "E"
filesDirectory = ""  # lan folder ending with slash


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
    global sheets_service, fileUrlColIndex, fileUrlColAlphabet, uploadFilesToFolder, googleSheetId, filesDirectory
    # row data is a list of what each row must contain (array)(each element in an xlrd Cell object)
    rowAsListArr = []
    for i in range(fileUrlColIndex):
        rowAsListArr.append(rowData[i].value)
    # call uploadFileToDrive
    fileName = rowData[fileUrlColIndex].value.split("\\")[1]
    filePath = filesDirectory + str(rowData[fileUrlColIndex].value)
    # https://stackoverflow.com/questions/7169845/using-python-how-can-i-access-a-shared-folder-on-windows-network
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
        # comment the below lines and test, make sure output is a list where each element is of the form: text:"+919741307999"
        addFileAndRowToExcel(currentOutputRow, rowObj)
        currentOutputRow = currentOutputRow + 1
    # addFileAndRowToExcel


start()
