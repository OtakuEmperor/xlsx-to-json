import xlrd
import xlwt
import json
import os.path
import datetime
import os
import re

def getColNames(sheet):
    global startPointRow
    global startPointCol
    global colValuesPSValue
    rowSize = sheet.row_len(0)
    collen = len(sheet.col_values(0))
    for i in range(rowSize):
        colValue = sheet.col_values(i)
        for index, value in enumerate(colValue):
            if value == 'ID':
                trueStartPointRow = index
                startPointCol = i
                startPointRow = int(index)
                break
    
    colValues = sheet.row_values(startPointRow, startPointCol, rowSize )
    print(colValues)
    colValuesPSValue = sheet.row_values(trueStartPointRow + 2, startPointCol, rowSize )
    columnNames = []

    for index, value in enumerate( colValues ):
        if(colValuesPSValue[index] != 'PS' and colValuesPSValue[index] != 'S'):
            columnNames.append(value)
    return columnNames

def getRowData(row, columnNames):
    rowData = {}
    counter = 0
    dataIndex = 0

    for index, cell in enumerate(row):
        # check if it is of date type print in iso format
        if(colValuesPSValue[counter]!= 'PS' and colValuesPSValue[index] != 'S'):
            if cell.ctype==xlrd.XL_CELL_DATE:
                rowData[columnNames[dataIndex].replace(' ', '_')] = datetime.datetime(*xlrd.xldate_as_tuple(cell.value,0)).isoformat()
            else:
                rowData[columnNames[dataIndex].replace(' ', '_')] = cell.value
            dataIndex += 1
        counter +=1

    return rowData

def getSheetData(sheet, columnNames):
    nRows = sheet.nrows
    sheetData = []
    counter = 1

    for idx in range(startPointRow+3, nRows):
        row = sheet.row(idx)
        rowData = getRowData(row, columnNames)
        sheetData.append(rowData)

    return sheetData

def getWorkBookData(workbook):
    nsheets = workbook.nsheets
    counter = 0
    workbookdata = {}
    global colNames
    colNames = []

    for idx in range(0, nsheets):
        worksheet = workbook.sheet_by_index(idx)
        if(worksheet.name.startswith("$")):
            colNames.append(str(worksheet.name))
            columnNames = getColNames(worksheet)
            sheetdata = getSheetData(worksheet, columnNames)
            workbookdata[worksheet.name.lower().replace(' ', '_')] = sheetdata
            GeneJsonFile(worksheet.name.lower(), workbookdata)

def GeneJsonFile(name, data):
    output = open(name.replace("$","") +".json", "w+")
    tmpString = json.dumps(data, sort_keys=False, indent=2,  separators=(',', ": "))
    for name in colNames:
        replaceString = "\\"+str(name).lower();
        newstring = re.sub(replaceString, "singleSheet", tmpString)
    print(newstring)
    output.write(newstring)
    output.close()

def main(filename):
    if os.path.isfile(filename):
        workbook = xlrd.open_workbook(filename)
        workbookdata = getWorkBookData(workbook)

        print ("%s was created" %output.name)
    else:
        print ("Sorry, that was not a valid filename")

allfile = os.listdir(os.getcwd())

for excel in allfile:
    if os.path.isfile(excel) and excel.endswith(".xlsx") and not excel.startswith("~"):
        main(excel)