from operator import truediv
import xlrd
from xlwt import Workbook
import re
import os

# this application removes all phone numbers in the input file that are present in the reference file
# input file and reference file data must start from 0th row

inputFile = xlrd.open_workbook("C:/Users/Shrey/Downloads/input.xlsx")
inputSheet = inputFile.sheet_by_index(0)

referenceFile = xlrd.open_workbook("C:/Users/Shrey/Downloads/reference.xlsx")
referenceSheet = referenceFile.sheet_by_index(0)

outputFile = Workbook()
outputSheet = outputFile.add_sheet("Sheet 1")

currentOutputRow = 0


def isPresentInReferenceSheet(number):
    for curr_row in range(0, referenceSheet.nrows):
        if referenceSheet.cell_value(curr_row, 1) == number:
            return True
    return False


def start():
    global currentOutputRow
    duplicateCount = 0
    print("Input excel has " + str(inputSheet.nrows) + " rows")
    for cur_row in range(0, inputSheet.nrows):
        if not isPresentInReferenceSheet(inputSheet.cell_value(cur_row, 1)):
            outputSheet.write(currentOutputRow, 0, inputSheet.cell_value(cur_row, 0))
            outputSheet.write(currentOutputRow, 1, inputSheet.cell_value(cur_row, 1))
            currentOutputRow += 1
        else:
            print(
                "Duplicate found:"
                + str(inputSheet.cell_value(cur_row, 0))
                + " "
                + str(inputSheet.cell_value(cur_row, 1))
            )
            duplicateCount += 1
    print("Removed " + str(duplicateCount) + " duplicates")


start()
outputFile.save("output.xls")
