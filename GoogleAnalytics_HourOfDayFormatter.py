# This program formats Google Analytics hour of day from 2022030115 to 2022/03/01 15
# for easy reading. Acceps either csv or xlsx, converts csv to xlsx, exports xlsx with _reformatted suffix.

#!/usr/local/bin/python3.10

import pandas as pd # import pandas module to convert csv to xlsx
from openpyxl import load_workbook # import openpyxl to parse and modify csv

print("This program formats Google Analytics hour of day from 2022030115 to 2022/03/01 15 for easy reading")


def reformat(file):
    # print("file", file)

    book = load_workbook(file)
    sheet = book.active  # iterable

    # check if a value is in a cell & removes row or column as applicable
    for row in sheet:
        for cell in row:
            try:
                if 'Hour of Day' in cell.value:
                    print("Hour of Day found")

                    start_column = cell.column
                    # start_point = cell.coordinate
                    start_point = "" + cell.coordinate[0] + str(cell.row+1) + ""
                    end_pointEstimate = "" + cell.coordinate[0] + str(cell.row+20) + ""

                    newEndPointArray = []
                    for row2 in sheet[f'{start_point}': f'{end_pointEstimate}']:
                        for cell2 in row2:
                            if cell2.value is None:
                                end_pointEstimate = "" + cell2.coordinate[0] + str(cell2.row) + ""
                                newEndPointArray.append(end_pointEstimate)
                                break
                    end_point = newEndPointArray[0]
                    end_point = end_point[0] + end_point[1] + str(int(end_point[2])-1)

                    for row in sheet[f'{start_point}': f'{end_point}']:
                        for cell in row:
                            # print("herez", cell.value)
                            x = str(cell.value)
                            y = " "
                            # print(x)
                            for index, char in enumerate(x):
                                if index == 3 or index == 5:
                                    # print(index)
                                    y += char + '/'
                                else:
                                    y += char
                                if index == 7:
                                    y += ' '
                            y = y.strip()
                            # print(repr(y))
                            sheet[cell.coordinate] = y
            except TypeError:
                continue

    if ".x" in file:
        docNameNew = file.split(".x")
        # print("docNameNew", docNameNew)
        fileUpdated = docNameNew[0] + "_reformatted" + ".x" + docNameNew[1]
        # print("fileUpdated", fileUpdated)
        book.save(filename=f'{fileUpdated}')
    
    print("Done")
    getFileName()

# function below changes csv to excel
def convertCSV2Excel(docName, excelFile):
    
    try:
        read_file = pd.read_csv (docName) # can handle filepath

        # convert to excel if no errors reading file
        try:
            read_file.to_excel (excelFile, index = None, header=False) # can handle filepath
            return True
        except UnboundLocalError:
            print("Error 2: Writing to file. Please make sure you have permission to write.")

    except pd.errors.ParserError:
        print("Error 1: reading file. Re-download/save & try again")
        getFileName()

# converter() function converts csv to xlsx then calls removeCurrencyColumn() to remove appropriate column
def converter(docName):

    docNameLastOnly = docName.split("/")

    if "csv" in docName: # checks if provided file is csv
        excelFile = docName.replace("csv", "xlsx")
        if convertCSV2Excel(docName, excelFile): # converts to xlsx & returns true
            reformat(excelFile)
        else:
            print("There was an issue with " + docNameLastOnly[len(docNameLastOnly)-1] + ", please try a different one.")
            getFileName()
    else:
        reformat(docName) # if not csv, removes column



def getFileName():
    docName = input("Enter csv or xlsx filename or path: ")
    converter(docName)


getFileName()