import openpyxl
from SparksScheduleSearch import SparksScheduleSearch

def formatting_cell(sheet, row, column, value, fontSize, fontName, fontBold, fontItalic, alignmentHorizontal, alignmentVertical):
    sheet.cell(row=row, column=column).value = value
    sheet.cell(row=row, column=column).font = openpyxl.styles.Font(size=fontSize, name=fontName, bold=fontBold, italic=fontItalic)
    sheet.cell(row=row, column=column).alignment = openpyxl.styles.Alignment(horizontal=alignmentHorizontal, vertical=alignmentVertical)


filename = "Timetable.xlsx"
wb = openpyxl.Workbook()
sheet = wb.worksheets[0]
sheet.title = "sheetName"
#Todo Высчитывание числа и месяца
#--Init
sparks = SparksScheduleSearch()
timeTable = sparks.findFirstGhost()
mediumBorder = openpyxl.styles.Side(border_style="medium", color="000000")
thickBorder = openpyxl.styles.Side(border_style="thick", color="000000")
week = {1: 'ПН',
        2: 'ВТ',
        3: "СР",
        4: "ЧТ",
        5: "ПТ",
        6: "СБ",
        7: "ВС"}
staff = dict()
i = 0
for s in timeTable[0].items():
    staff[s[0].Name] = i
    i += 1

spaceBetweenTables = 1
numberOfTablesInLine = 3
tableWidth = 1 + len(week) + spaceBetweenTables  #+1 - begin from 1;+1 - space between tables
tableHeight = 1 + len(staff) + spaceBetweenTables
#--
#--Fill table all timeTables
for numOfTable in range(len(timeTable)):
    startingPointColumn = 1 + ((numOfTable % numberOfTablesInLine) * tableWidth)
    startingPointRow = 1 + ((numOfTable // numberOfTablesInLine) * tableHeight)
    sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn)].width = 10
    formatting_cell(sheet, startingPointRow, startingPointColumn, f"№ {numOfTable + 1}", 14, "Times New Roman", False, False, "center", "center")
    #--Formatting and x-filling the timetable
    for row in range(startingPointRow+1, len(staff)+startingPointRow+1):
        for column in range(startingPointColumn+1, len(week)+startingPointColumn+1):
            sheet.cell(row=row, column=column).border = openpyxl.styles.Border(right=mediumBorder,bottom=mediumBorder)
            formatting_cell(sheet, row, column, '✕', 14, "Times New Roman", False, False, "center", "center")
    #--
    #--Creating of header of timetable
    for day in range(1, len(week)+1):
        sheet.cell(row=startingPointRow, column=startingPointColumn+day).border = openpyxl.styles.Border(left=thickBorder, right=thickBorder, top=thickBorder, bottom=thickBorder)
        formatting_cell(sheet, startingPointRow, startingPointColumn+day, week[day], 14, "Times New Roman", True, False, "center", "center")
    staffNames = list(staff.keys())
    for memberOfStaff in range(len(staff)):
        sheet.cell(row=startingPointRow + memberOfStaff + 1, column=startingPointColumn).border = openpyxl.styles.Border(left=thickBorder, right=thickBorder, top=thickBorder, bottom=thickBorder)
        formatting_cell(sheet, startingPointRow + memberOfStaff + 1, startingPointColumn, staffNames[memberOfStaff], 14, "Times New Roman", True, True, "right", "center")
    #--
    #--Fill data to the timetable
    typeOfShift = '✕'#'✖'
    # print(numOfTable+1)
    for memberOfStaff in timeTable[numOfTable].items():  #dict[EmployeeCard, list[tuple[DayType, ShiftLength, PlaceToWork]]]
        # print(memberOfStaff[0].Name, memberOfStaff[1])
        for i in memberOfStaff[1]: #list[tuple[DayType, ShiftLength, PlaceToWork]]
            if i[2] == 'Hall':
                sheet.cell(row=startingPointRow + staff[memberOfStaff[0].Name] + 1, column=startingPointColumn + i[0]).fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor='F9E79F')
                if i[1] == 1:
                    typeOfShift = 'С'
                else:
                    typeOfShift = 'С2'
            else:
                sheet.cell(row=startingPointRow + staff[memberOfStaff[0].Name] + 1, column=startingPointColumn + i[0]).fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor='ABEBC6')
                typeOfShift = 'T'
            formatting_cell(sheet, startingPointRow + staff[memberOfStaff[0].Name] + 1, startingPointColumn + i[0], typeOfShift, 14, "Times New Roman", False, False, "center", "center")
    #--
    # print('\t----------')
#--
#--Close and save excel's file
wb.save(filename)
print("\n\t\tNew file save!")
#--
