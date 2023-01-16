import sys
import os
import re
from openpyxl import load_workbook
from getExcelFileName import getExcelFileName

# Filenames
excelFile = getExcelFileName()
if excelFile == "None":
    print("No one *.XLSX files are find")
else:
    
    textToFind = "text_to_find.txt"
    resultFile = "result_file.txt"
    rowCounter = 0

    checkList = list()
    resultRowList = list()

    wb = load_workbook(excelFile)
    
    ## Take an active sheet from out workbook.
    ws = wb.active

    with open(textToFind) as rawValues:
        valuesListToSearch = rawValues.read().split(", ")

    for oneValue in valuesListToSearch:
        checkList.append(float(oneValue))
        checkList.append(0 - float(oneValue))

    for row in ws.values:
        rowCounter += 1
        for value in row:
            # If one or more rows will be as None
            try:
                if float(value) in checkList:
                    resultRowList.append(rowCounter)
            finally:
                continue

    with open(resultFile, 'w') as resultFilename:
        resultFilename.write(str(resultRowList))

    if len(resultRowList) != 0:
        print("Something is here!")
        input()
    else:
        print("Nothing to find..")
        input()
