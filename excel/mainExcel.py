import openpyxl
import datetime
import copy
from search.SparksScheduleSearch import SparksScheduleSearch

FILENAME_POOL_TIMETABLE = "../PoolTimetable.xlsx"
FILENAME_SCHEDULE_DATA_BASE = "../ScheduleDataBase.xlsx"
WEEK_LENGTH = 7
SPACE_BETWEEN_TABLES = 1
NUMBER_OF_TABLES_IN_LINE = 3
MEDIUM_BORDER = openpyxl.styles.Side(border_style="medium", color="000000")
THICK_BORDER = openpyxl.styles.Side(border_style="thick", color="000000")

def formatting_cell(sheet, row, column, value, fontSize, fontName, fontBold, fontItalic, alignmentHorizontal, alignmentVertical):
    sheet.cell(row=row, column=column).value = value
    sheet.cell(row=row, column=column).font = openpyxl.styles.Font(size=fontSize, name=fontName, bold=fontBold, italic=fontItalic)
    sheet.cell(row=row, column=column).alignment = openpyxl.styles.Alignment(horizontal=alignmentHorizontal, vertical=alignmentVertical)

def get_dated_week(datedWeek): #it gives dated week, that begin next monday
    date_1 = datetime.date(2024, 2, 29)
    date_2 = datetime.date(2024, 2, 28)
    delta_day = date_1 - date_2
    daysWeek = {
        1: 'ПН',
        2: 'ВТ',
        3: "СР",
        4: "ЧТ",
        5: "ПТ",
        6: "СБ",
        7: "ВС"}
    months = {
        1: 'Янв.',
        2: 'Февр.',
        3: 'Март',
        4: 'Апр.',
        5: 'Май',
        6: 'Июнь',
        7: 'Июль',
        8: 'Авг.',
        9: 'Сент.',
        10: 'Окт.',
        11: 'Нояб.',
        12: 'Дек.'}
    current_date = datetime.date.today()
    # current_date = datetime.date(2024, 11, 14)
    current_date += (WEEK_LENGTH-current_date.weekday()) * delta_day
    # print(f"{daysWeek[current_date.weekday() + 1]} | {current_date.strftime('%A| %d/%m/%y')}")
    for i in range(WEEK_LENGTH):
        # datedWeek.append(f"{daysWeek[current_date.weekday()+1]}, {current_date.day} {months[current_date.month]}")
        datedWeek.append(f"{current_date.day} {months[current_date.month]}")
        current_date += delta_day

def output_pool_of_schedule_to_excel(filename, searchMode):
    #--Init
    sparks = SparksScheduleSearch()
    timeTable = sparks.search(mode=searchMode) #'fast', 'part', 'full'
    wb = openpyxl.Workbook()
    sheet = wb.worksheets[0]
    sheet.title = "Выбор расписания"
    staff = dict()
    i = 0
    for s in timeTable[0].items():
        staff[s[0].Name] = i
        i += 1
    tableWidth = 1 + WEEK_LENGTH + SPACE_BETWEEN_TABLES  # +1 - begin from 1; +1 - space between tables
    tableHeight = 1 + len(staff) + SPACE_BETWEEN_TABLES
    #--
    #--Fill table all timeTables
    for numOfTable in range(len(timeTable)):
        startingPointColumn = 1 + ((numOfTable % NUMBER_OF_TABLES_IN_LINE) * tableWidth)
        startingPointRow = 1 + ((numOfTable // NUMBER_OF_TABLES_IN_LINE) * tableHeight)

        sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn)].width = 10
        for i in range(1, WEEK_LENGTH + 1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn+i)].width = 10#15 #if [day, dd mmmm]

        formatting_cell(sheet, startingPointRow, startingPointColumn, f"№ {numOfTable + 1}", 14, "Times New Roman", False, False, "center", "center")
        #--Formatting and x-filling the timetable
        for row in range(startingPointRow+1, len(staff)+startingPointRow+1):
            for column in range(startingPointColumn+1, WEEK_LENGTH + startingPointColumn + 1):
                sheet.cell(row=row, column=column).border = openpyxl.styles.Border(right=MEDIUM_BORDER,bottom=MEDIUM_BORDER)
                formatting_cell(sheet, row, column, '✕', 14, "Times New Roman", False, False, "center", "center")
        #--
        #--Creating of header of timetable
        datedWeek = list()
        get_dated_week(datedWeek)
        for day in range(1, WEEK_LENGTH + 1):
            sheet.cell(row=startingPointRow, column=startingPointColumn+day).border = openpyxl.styles.Border(left=THICK_BORDER, right=THICK_BORDER, top=THICK_BORDER, bottom=THICK_BORDER)
            formatting_cell(sheet, startingPointRow, startingPointColumn+day, datedWeek[day-1], 12, "Times New Roman", True, False, "center", "center")
        staffNames = list(staff.keys())
        for memberOfStaff in range(len(staff)):
            sheet.cell(row=startingPointRow + memberOfStaff + 1, column=startingPointColumn).border = openpyxl.styles.Border(left=THICK_BORDER, right=THICK_BORDER, top=THICK_BORDER, bottom=THICK_BORDER)
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
                        typeOfShift = 'С' #RUS С
                    else:
                        typeOfShift = 'С2' #RUS С
                else:
                    sheet.cell(row=startingPointRow + staff[memberOfStaff[0].Name] + 1, column=startingPointColumn + i[0]).fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor='ABEBC6')
                    typeOfShift = 'Т' #RUS Т
                formatting_cell(sheet, startingPointRow + staff[memberOfStaff[0].Name] + 1, startingPointColumn + i[0], typeOfShift, 14, "Times New Roman", False, False, "center", "center")
        #--
        # print('\t----------')
    #--
    #--Close and save excel's file
    wb.save(filename)
    print(f"\tFile '{filename}' was saved!")
    #--
    lengthOfPool = len(timeTable)
    return lengthOfPool

def update_schedule_data_base(filenamePoolTimetable, filenameSceduleDataBase, numChoosingTimetable):
    wbSceduleDataBase = openpyxl.load_workbook(filename = filenameSceduleDataBase)
    sheet = wbSceduleDataBase.worksheets[0]
    wbChooseTimetable = openpyxl.load_workbook(filename = filenamePoolTimetable)
    poolSheet = wbChooseTimetable.worksheets[0]
    # print(wb.worksheets)
    # sheet.cell(row=1, column=1).value #number of week schedule
    staffLength = 0
    r = 2
    while poolSheet.cell(row=r, column=1).value != None:
        staffLength += 1
        r += 1
    tableWidth = 1 + 7 + SPACE_BETWEEN_TABLES  # +1 - begin from 1; 7 - week length; +1 - space between tables
    tableHeight = 1 + staffLength + SPACE_BETWEEN_TABLES
    startingPointColumn = 1 + (sheet.cell(row=1, column=1).value * tableWidth)
    startingPointRow = 3
    poolPointColumn = 1 + (((numChoosingTimetable-1) % NUMBER_OF_TABLES_IN_LINE) * tableWidth)
    poolPointRow = 1 + (((numChoosingTimetable-1) // NUMBER_OF_TABLES_IN_LINE) * tableHeight)

    for i in range(tableHeight-SPACE_BETWEEN_TABLES):
        for j in range(tableWidth-SPACE_BETWEEN_TABLES):
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).value = poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).value
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).font = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).alignment = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).alignment)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).border = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).border)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).fill = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).fill)
            # formatting_cell(sheet, startingPointRow + i, startingPointColumn + j, poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).value,
            #                 poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font.sz, poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font.name,
            #                 poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font.b, poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font.i,
            #                 poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).alignment.horizontal, poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).alignment.vertical)

    sheet.cell(row=1, column=1).value = sheet.cell(row=1, column=1).value + 1 #number of week schedule + 1 after added
    wbSceduleDataBase.save(filenameSceduleDataBase)
    print(f"\nPool schedule number {numChoosingTimetable} was added to Schedule Data Base '{filenameSceduleDataBase}'")

if __name__ == "__main__":
    lengthOfPool = output_pool_of_schedule_to_excel(FILENAME_POOL_TIMETABLE, "part") #"fast", "part", "full"
    while True:
        numChoosingTimetable = input("\nEnter number of choosing schedule: ")
        try:
            numChoosingTimetable = int(numChoosingTimetable)
            if (numChoosingTimetable < 1 or numChoosingTimetable > lengthOfPool):
                print(f"\n\t**ERROR VALUE.\n\t\tPlease, enter value from 1 to {lengthOfPool}.")
            else:
                print(numChoosingTimetable)
                break
        except:
            print("\n\t**ERROR TYPE.\n\t\tPlease, enter integer value.")
    update_schedule_data_base(FILENAME_POOL_TIMETABLE, FILENAME_SCHEDULE_DATA_BASE, numChoosingTimetable)
