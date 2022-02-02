import xlrd
import xlwt
from xlwt import Workbook
import re
import os

workbook = xlrd.open_workbook("C:/Users/Shrey/Downloads/slt2.xlsx")
sheet_1 = workbook.sheet_by_index(0)

outputWorkbook = Workbook()
outputSheet = outputWorkbook.add_sheet("Sheet 1")

currentOutputRow = 0

addedPlusCount = 0
addedPlusCodeCount = 0


def addToOutput(number, name, country):
    global currentOutputRow
    if len(number) > 9:
        number = cleanNumber(number, country)
        if len(number) > 9 and len(number) <= 13:
            number = addCountryCode(number, country)
            outputSheet.write(currentOutputRow, 0, name)
            outputSheet.write(currentOutputRow, 1, number)
            outputSheet.write(currentOutputRow, 2, country)
            currentOutputRow += 1


def addCountryCode(number, country):

    if len(number) == 10 and country == "India":
        global addedPlusCodeCount
        global addedPlusCount

        addedPlusCodeCount += 1
        print("Added +91 to:" + str(number))
        return "+91" + number

    elif len(number) == 12 and country == "India" and number[0:2] == "91":
        addedPlusCount += 1
        print("Added + to 12 dig number:" + str(number))
        return "+" + number

    elif len(number) == 13 and number[0] != "+":
        addedPlusCount += 1
        print("Added + to 13 dig number:" + str(number))
        return "+" + number
    else:
        return number


def cleanNumber(number, country):
    number = number.replace("-", "")
    number = number.replace(" ", "")
    number = number.replace("(", "")
    number = number.replace(")", "")
    number = number.replace(".", "")

    if number[0] == "+" and number[1] == "+":
        print("Found double + in " + number)
        number = number[1:]

    leadingZeroesRegex = "^0+(?!$)"  # removes all leading 0s
    number = re.sub(leadingZeroesRegex, "", number)

    return number


def splitCleanPutRow(number, name, country):
    if re.search("[a-zA-Z]", number):
        print("Number has an alphabet, ignoring " + number)
        return
    numbersArr = number.split("/")
    if len(numbersArr) > 1:
        for num in numbersArr:
            addToOutput(num, name, country)
        return
    else:
        numbersArr = number.split("\\")
        if len(numbersArr) > 1:
            for num in numbersArr:
                addToOutput(num, name, country)
            return
        else:
            numbersArr = number.split(",")
            if len(numbersArr) > 1:
                for num in numbersArr:
                    addToOutput(num, name, country)
                return
            else:
                numbersArr = number.split(";")
                if len(numbersArr) > 1:
                    for num in numbersArr:
                        addToOutput(num, name, country)
                    return
                else:
                    addToOutput(number, name, country)


def start():
    print("This excel has " + str(sheet_1.nrows) + " rows")
    numberColumns = [8, 9, 10, 11, 12]
    for cur_row in range(1, sheet_1.nrows):
        for numcol in numberColumns:
            if len(sheet_1.cell_value(cur_row, numcol)) > 9:
                splitCleanPutRow(
                    sheet_1.cell_value(cur_row, numcol),
                    sheet_1.cell_value(cur_row, 1),
                    sheet_1.cell_value(cur_row, 6),
                )


start()
outputWorkbook.save("output.xls")
