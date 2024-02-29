import openpyxl
from SparksScheduleSearch import SparksScheduleSearch

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
    numRow = 1
    for numColumn in range(1, sheet.max_column + 1):
        if numColumn == 1:
            newSheet.column_dimensions[openpyxl.utils.get_column_letter(numColumn)].width = 10
        elif numColumn == 2:
            newSheet.column_dimensions[openpyxl.utils.get_column_letter(numColumn)].width = 7
        else:
            newSheet.column_dimensions[openpyxl.utils.get_column_letter(numColumn)].width = 20

        newSheet.cell(row=numRow, column=numColumn).font = openpyxl.styles.Font(bold=True)
        newSheet.cell(row=numRow, column=numColumn).alignment = openpyxl.styles.Alignment(horizontal='center')
        newSheet.cell(row=numRow, column=numColumn).value = sheet.cell(row=numRow, column=numColumn).value
    for numRow in range(2, sheet.max_row + 1):
        for numColumn in range(1, sheet.max_column + 1):
            newSheet.cell(row=numRow, column=numColumn).font = openpyxl.styles.Font(italic=True)
            if numColumn == 1:
                bufferAlignment = openpyxl.styles.Alignment(horizontal='right')
            elif numColumn == 2:
                bufferAlignment = openpyxl.styles.Alignment(horizontal='center')
            else:
                bufferAlignment = openpyxl.styles.Alignment(horizontal='left')
            newSheet.cell(row=numRow, column=numColumn).alignment = bufferAlignment
            newSheet.cell(row=numRow, column=numColumn).value = sheet.cell(row=numRow, column=numColumn).value

    newWb.save(newFilename)
    print("New file save!")

# read_table_xlsx("testTable.xlsx")
# write_table_xlsx("testTable.xlsx", "newTestTable.xlsx", "Test")

sparks = SparksScheduleSearch()
data = sparks.findFirstGhost()
for i in range(len(data)):
    print(i)
    for k in data[i].items():
        print(k[0].Name, k[1])
    print()

# for i in data:  # i dict
#     for k in i.items():
#         print(k[0].Name, k[1])
#         # print(k[0].Name)#EmployeeCard
#         # print(k[1])#list[tuple[DayType, ShiftLength, PlaceToWork]]
#     print()

print("\n---END---")
