import xlrd, xlwt #For reading and writing excel's file 2003
import openpyxl

def read_table_xls_2003(filename):
    print("\n\tshowTable_xls_2003")
    rb = xlrd.open_workbook(filename, formatting_info=True)
    sheet = rb.sheet_by_index(0)

    maxRow = sheet.nrows
    maxCols = sheet.ncols
    # print(sheet.nrows)
    # print(sheet.ncols)

    strRow = sheet.row_values(0)
    print(f"|{strRow[0]}|{strRow[1]}| {strRow[2].ljust(12)}|")
    for numRow in range(1, sheet.nrows):
        strRow = sheet.row_values(numRow)
        strToSend = f"|{str(int(strRow[0])).ljust(6)}|{strRow[1].rjust(3).ljust(4)}| {strRow[2].ljust(12)}|"
        print(strToSend)

def write_table_xls_2003(filename, newFilename, sheetName):
    print("\n\tread_table_xls_2003")

    rb = xlrd.open_workbook(filename, formatting_info=True)
    sheet = rb.sheet_by_index(0)
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheetName)

    for numRow in range(0, sheet.nrows):
        strRow = sheet.row_values(numRow)
        for numCol in range(0, sheet.ncols):
            ws.write(numRow, numCol, strRow[numCol])

    wb.save(newFilename)
    print("New file save!")

def read_table_xlsx(filename):
    print("\n\tread_table_xlsx")
    wb = openpyxl.load_workbook(filename = filename)
    sheet = wb['Лист1']

    # strCell = sheet.cell(row=1, column=1).value #Begin at 1, not at 0
    # print(strCell)
    # print(sheet.max_column)
    # print(sheet.max_row)
    # # vals = [v[0].value for v in sheet['A1:C26']]
    # for numRow in vals:
    #     print(numRow)
    # print(sheet['C21'].value)
    for numRow in range(1, sheet.max_row + 1):
        bufferRow = list(range(sheet.max_column))
        for numColumn in range(1, sheet.max_column + 1):
            bufferRow[numColumn-1] = sheet.cell(row=numRow, column=numColumn).value
        print(bufferRow)



def write_table_xlsx(filename, newFilename, sheetName):
    print("\n\twrite_table_xlsx")
    wb = openpyxl.load_workbook(filename=filename)
    sheet = wb['Лист1']
    newWb = openpyxl.Workbook()
    newSheet = newWb.worksheets[0]
    newSheet.title = sheetName
    # del newWb["Sheet"]
    # newSheet = newWb.create_sheet("Test")
    # print(newWb.sheetnames)
    for numRow in range(1, sheet.max_row + 1):
        for numColumn in range(1, sheet.max_column + 1):
            newSheet.cell(row=numRow, column=numColumn).value = sheet.cell(row=numRow, column=numColumn).value

    newWb.save(newFilename)
    print("New file save!")

# read_table_xls_2003("testTable_2003.xls")
# write_table_xls_2003("testTable_2003.xls", "newTestTable_2003.xls", "Test_2003")
# read_table_xlsx("testTable.xlsx")
write_table_xlsx("testTable.xlsx", "newTestTable.xlsx", "Test")

print("\n---END---")
