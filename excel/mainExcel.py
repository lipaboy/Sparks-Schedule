import openpyxl
import datetime
import copy
from search.SparksScheduleSearch import SparksScheduleSearch
from search.WeekScheduleExcelType import ShiftType
from search.WeekScheduleExcelType import EmployeeCard
from search.WeekScheduleExcelType import WeekScheduleExcelType

FILENAME_POOL_TIMETABLE = "../PoolTimetable.xlsx"
FILENAME_SCHEDULE_DATA_BASE = "../ScheduleDataBase.xlsx"
WEEK_LENGTH = 7
SPACE_BETWEEN_TABLES = 1
NUMBER_OF_TABLES_IN_LINE = 3
TABLE_0_COLUMN_WIDTH = 10 #For BD sheet min value 10
TABLE_1_COLUMN_WIDTH = 15 #For Staff and Tracks sheets
THIN_BORDER = openpyxl.styles.Side(border_style="thin", color="000000")
MEDIUM_BORDER = openpyxl.styles.Side(border_style="medium", color="000000")
THICK_BORDER = openpyxl.styles.Side(border_style="thick", color="000000")
CHAR_CROSS = '✕' #'✖'
CHAR_HALL = 'С' #RUS С
CHAR_HALF_HALL = 'С2' #RUS С
CHAR_TRACK = 'Т' #RUS Т
ERROR_STR_HEAD = "\n\t**ERROR"#! (numOfFunction)\n\t\t

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
    current_date = datetime.date.today()  # current_date = datetime.date(2024, 11, 14)
    current_date += (WEEK_LENGTH-current_date.weekday()) * delta_day
    for i in range(WEEK_LENGTH):
        datedWeek.append(f"{current_date.day} {months[current_date.month]}")
        current_date += delta_day

def output_pool_of_schedule_to_excel(filenamePoolTimetable, searchMode="fast"):# "fast", "part", "full"
    #--Init
    sparks = SparksScheduleSearch()
    timeTable = sparks.search(mode=searchMode) #'fast', 'part', 'full'
    wb = openpyxl.Workbook()
    sheet = wb.worksheets[0]
    sheet.title = "Выбор расписания"
    tableWidth = 1 + WEEK_LENGTH + SPACE_BETWEEN_TABLES  # +1 - begin from 1; +1 - space between tables
    tableHeight = 1 + len(timeTable[0]) + SPACE_BETWEEN_TABLES
    #--
    #--Fill table all timeTables
    for numOfTable in range(len(timeTable)):
        startingPointColumn = 1 + ((numOfTable % NUMBER_OF_TABLES_IN_LINE) * tableWidth)
        startingPointRow = 1 + ((numOfTable // NUMBER_OF_TABLES_IN_LINE) * tableHeight)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn)].width = TABLE_0_COLUMN_WIDTH
        for i in range(1, WEEK_LENGTH + 1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn+i)].width = TABLE_0_COLUMN_WIDTH#15 #if [day, dd mmmm]
        formatting_cell(sheet, startingPointRow, startingPointColumn, f"№ {numOfTable + 1}", 14, "Times New Roman", False, False, "center", "center")
        #--Formatting and x-filling the timetable
        for row in range(startingPointRow+1, len(timeTable[numOfTable])+startingPointRow+1):
            for column in range(startingPointColumn+1, WEEK_LENGTH + startingPointColumn + 1):
                sheet.cell(row=row, column=column).border = openpyxl.styles.Border(right=MEDIUM_BORDER,bottom=MEDIUM_BORDER)
                formatting_cell(sheet, row, column, CHAR_CROSS, 14, "Times New Roman", False, False, "center", "center")
        #--
        #--Creating of header of timetable
        datedWeek = list()
        get_dated_week(datedWeek)
        for day in range(1, WEEK_LENGTH + 1):
            sheet.cell(row=startingPointRow, column=startingPointColumn+day).border = openpyxl.styles.Border(left=THICK_BORDER, right=THICK_BORDER, top=THICK_BORDER, bottom=THICK_BORDER)
            formatting_cell(sheet, startingPointRow, startingPointColumn+day, datedWeek[day-1], 12, "Times New Roman", True, False, "center", "center")
        for numOfStaff in range(len(timeTable[numOfTable])):
            sheet.cell(row=startingPointRow + numOfStaff + 1, column=startingPointColumn).border = openpyxl.styles.Border(left=THICK_BORDER, right=THICK_BORDER, top=THICK_BORDER, bottom=THICK_BORDER)
            formatting_cell(sheet, startingPointRow + numOfStaff + 1, startingPointColumn, timeTable[numOfTable][numOfStaff].Name, 14, "Times New Roman", timeTable[numOfTable][numOfStaff].IsElder, True, "right", "center")
        #--
        #--Fill data to the timetable
        typeOfShift = CHAR_CROSS
        for numOfStaff in range(len(timeTable[numOfTable])):  #class EmployeeCard(name: str, isElder: bool, shifts: list[ShiftType])
            for shift in timeTable[numOfTable][numOfStaff].Shifts: #ShiftType = tuple[DayType, ShiftLength, PlaceToWork]
                if shift[2] == 'Hall':
                    sheet.cell(row=startingPointRow + numOfStaff + 1, column=startingPointColumn + shift[0]).fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor='F9E79F')
                    if shift[1] == 1:
                        typeOfShift = CHAR_HALL
                    else:
                        typeOfShift = CHAR_HALF_HALL
                else:
                    sheet.cell(row=startingPointRow + numOfStaff + 1, column=startingPointColumn + shift[0]).fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor='ABEBC6')
                    typeOfShift = CHAR_TRACK
                formatting_cell(sheet, startingPointRow + numOfStaff + 1, startingPointColumn + shift[0], typeOfShift, 14, "Times New Roman", False, False, "center", "center")
        #--
    #--
    #--Close and save excel's file
    wb.save(filenamePoolTimetable)
    print(f"\n\tFile '{filenamePoolTimetable}' was created!")
    #--
    lengthOfPool = len(timeTable)
    return lengthOfPool

def init_schedule_data_base(filenameSceduleDataBase):
    wbSceduleDataBase = openpyxl.Workbook()
    sheet = wbSceduleDataBase.worksheets[0]
    sheet.title = "Расписания"
    formatting_cell(sheet, 1, 1, 0, 14, "Times New Roman", False, False, "center", "center")
    wbSceduleDataBase.create_sheet("Персонал")
    sheet = wbSceduleDataBase.worksheets[1]
    columnCursor = 1
    sheet.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheet, 1, columnCursor, "Имя", 14, "Times New Roman", True, False, "center", "center")
    sheet.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=MEDIUM_BORDER, bottom=THICK_BORDER)
    columnCursor = 2
    sheet.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheet, 1, columnCursor, "Старший", 14, "Times New Roman", True, False, "center", "center")
    sheet.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=THICK_BORDER, bottom=THICK_BORDER)
    wbSceduleDataBase.create_sheet('Тачки')
    sheet = wbSceduleDataBase.worksheets[2]
    columnCursor = 1
    sheet.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheet, 1, columnCursor, "Имя", 14, "Times New Roman", True, False, "center", "center")
    sheet.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=MEDIUM_BORDER, bottom=THICK_BORDER)
    columnCursor = 2
    sheet.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheet, 1, columnCursor, "Кол-во Т", 14, "Times New Roman", True, False, "center", "center")
    sheet.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=THICK_BORDER, bottom=THICK_BORDER)
    wbSceduleDataBase.save(filenameSceduleDataBase)
    print(f"\n\tFile '{filenameSceduleDataBase}' was created!")

def update_schedule_data_base(filenameSceduleDataBase, filenamePoolTimetable, numChoosingTimetable):
    #Init
    wbSceduleDataBase = openpyxl.load_workbook(filename = filenameSceduleDataBase)
    sheet = wbSceduleDataBase.worksheets[0]
    wbChooseTimetable = openpyxl.load_workbook(filename = filenamePoolTimetable)
    poolSheet = wbChooseTimetable.worksheets[0]
    #--
    #Calc of staff's size
    staffLength = 0
    r = 2
    while poolSheet.cell(row=r, column=1).value != None:
        staffLength += 1
        r += 1
    #--
    #Init tables parameters
    tableWidth = 1 + 7 + SPACE_BETWEEN_TABLES  # +1 - begin from 1; 7 - week length; +1 - space between tables
    tableHeight = 1 + staffLength + SPACE_BETWEEN_TABLES
    startingPointColumn = 1 + (sheet.cell(row=1, column=1).value * tableWidth)
    startingPointRow = 3
    poolPointColumn = 1 + (((numChoosingTimetable-1) % NUMBER_OF_TABLES_IN_LINE) * tableWidth)
    poolPointRow = 1 + (((numChoosingTimetable-1) // NUMBER_OF_TABLES_IN_LINE) * tableHeight)
    #--
    #Update DB
    for j in range(tableWidth-SPACE_BETWEEN_TABLES):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(startingPointColumn+j)].width = TABLE_0_COLUMN_WIDTH
    for i in range(tableHeight-SPACE_BETWEEN_TABLES):
        for j in range(tableWidth-SPACE_BETWEEN_TABLES):
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).value = poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).value
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).font = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).font)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).alignment = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).alignment)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).border = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).border)
            sheet.cell(row=startingPointRow + i, column=startingPointColumn + j).fill = copy.copy(poolSheet.cell(row=poolPointRow + i, column=poolPointColumn + j).fill)

    sheet.cell(row=1, column=1).value = sheet.cell(row=1, column=1).value + 1 #number of week schedule + 1 after added
    sheet.cell(row=startingPointRow, column=startingPointColumn).value = "№ " + str(sheet.cell(row=1, column=1).value)
    #--
    wbSceduleDataBase.save(filenameSceduleDataBase)
    print(f"\nPool schedule number {numChoosingTimetable} was added to Schedule Data Base '{filenameSceduleDataBase}'")

def update_schedule_data_base_staff(filenameSceduleDataBase, poolOfNewStaff=None, numOfSelectedSchedule=-1):
    #Check the 2nd parameter
    if poolOfNewStaff != None:
        print(f"{ERROR_STR_HEAD}! (update_schedule_data_base_staff)\n\t\tType of data wasn't selected")#Todo Add update from array of data
        return None
    #--
    wbSceduleDataBase = openpyxl.load_workbook(filename=filenameSceduleDataBase)
    sheet = wbSceduleDataBase.worksheets[0]
    sheetStaff = wbSceduleDataBase.worksheets[1]
    #Clear sheet of staff
    sheetStaff.delete_cols(1)
    sheetStaff.delete_cols(1)
    columnCursor = 1
    sheetStaff.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheetStaff, 1, columnCursor, "Имя", 14, "Times New Roman", True, False, "center", "center")
    sheetStaff.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=MEDIUM_BORDER, bottom=THICK_BORDER)
    columnCursor = 2
    sheetStaff.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheetStaff, 1, columnCursor, "Старший", 14, "Times New Roman", True, False, "center", "center")
    sheetStaff.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=THICK_BORDER, bottom=THICK_BORDER)
    #--
    lengthOfDataBase = sheet.cell(row=1, column=1).value
    # Check the 3rd parameter
    if numOfSelectedSchedule == -1:
        selectedNumber = lengthOfDataBase
    else:
        selectedNumber = numOfSelectedSchedule
    if ((selectedNumber < 1) or (selectedNumber > lengthOfDataBase)):
        print(ERROR_STR_HEAD + "! (update_schedule_data_base_staff)\n\t\tOut of range!!!")  # it gives the last schedule in DB
        return None
    #--
    tableWidth = 1 + 7 + SPACE_BETWEEN_TABLES  # +1 - begin from 1; 7 - week length; +1 - space between tables
    startingPointColumn = 1 + ((selectedNumber-1) * tableWidth)
    startingPointRow = 4
    staffPointColumn = 1
    staffPointRow = 2 #Because we already have HEAD
    i = 0
    statusChar = "-"
    while sheet.cell(row=startingPointRow + i, column=startingPointColumn).value != None:
        formatting_cell(sheetStaff, staffPointRow + i, staffPointColumn, sheet.cell(row=startingPointRow+i, column=startingPointColumn).value,
                        14, "Times New Roman", sheet.cell(row=startingPointRow + i, column=startingPointColumn).font.b, True, "right", "center")
        if sheet.cell(row=startingPointRow + i, column=startingPointColumn).font.b:
            statusChar = "+"
        else:
            statusChar = "-"
        formatting_cell(sheetStaff, staffPointRow + i, staffPointColumn+1, statusChar,
                        14, "Times New Roman", False, False, "center", "center")
        sheetStaff.cell(row=staffPointRow + i, column=staffPointColumn).border = openpyxl.styles.Border(right=THIN_BORDER, bottom=THIN_BORDER)
        sheetStaff.cell(row=staffPointRow + i, column=staffPointColumn+1).border = openpyxl.styles.Border(right=THIN_BORDER, bottom=THIN_BORDER)
        i += 1
    wbSceduleDataBase.save(FILENAME_SCHEDULE_DATA_BASE)
    print(f"\nStaff was updated on Schedule Data Base '{filenameSceduleDataBase}'")

def update_schedule_data_base_track(filenameSceduleDataBase, poolOfStaffAndTrack=None,):
    # Check the 2nd parameter
    if poolOfStaffAndTrack != None:
        print(f"{ERROR_STR_HEAD}! (update_schedule_data_base_track)\n\t\tType of data wasn't selected")  # Todo Add update from array of data
        return None
    # --
    wbSceduleDataBase = openpyxl.load_workbook(filename=filenameSceduleDataBase)
    sheetStaff = wbSceduleDataBase.worksheets[1]
    sheetTrack = wbSceduleDataBase.worksheets[2]
    # Clear sheet of staff
    sheetTrack.delete_cols(1)
    sheetTrack.delete_cols(1)
    columnCursor = 1
    sheetTrack.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheetTrack, 1, columnCursor, "Имя", 14, "Times New Roman", True, False, "center", "center")
    sheetTrack.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=MEDIUM_BORDER, bottom=THICK_BORDER)
    columnCursor = 2
    sheetTrack.column_dimensions[openpyxl.utils.get_column_letter(columnCursor)].width = TABLE_1_COLUMN_WIDTH
    formatting_cell(sheetTrack, 1, columnCursor, "Кол-во Т", 14, "Times New Roman", True, False, "center", "center")
    sheetTrack.cell(row=1, column=columnCursor).border = openpyxl.styles.Border(right=THICK_BORDER, bottom=THICK_BORDER)
    #--
    staffPointRow = 2
    staffPointColumn = 1
    i = 0
    while sheetStaff.cell(row=staffPointRow + i, column=staffPointColumn).value != None:
        formatting_cell(sheetTrack, staffPointRow + i, staffPointColumn, sheetStaff.cell(row=staffPointRow + i, column=staffPointColumn).value,
                        14, "Times New Roman", sheetStaff.cell(row=staffPointRow + i, column=staffPointColumn).font.b, True, "right", "center")
        formatting_cell(sheetTrack, staffPointRow + i, staffPointColumn+1, 0,
                        14, "Times New Roman", False, False, "center", "center")
        sheetTrack.cell(row=staffPointRow + i, column=staffPointColumn).border = openpyxl.styles.Border(right=THIN_BORDER, bottom=THIN_BORDER)
        sheetTrack.cell(row=staffPointRow + i, column=staffPointColumn+1).border = openpyxl.styles.Border(right=THIN_BORDER, bottom=THIN_BORDER)
        i += 1
    if i == 0:
        print(ERROR_STR_HEAD + "! (update_schedule_data_base_track)\n\t\tList of staff is empty!")  # it gives the last schedule in DB
        return None
    wbSceduleDataBase.save(FILENAME_SCHEDULE_DATA_BASE)
    print(f"\nTrack was updated on Schedule Data Base '{filenameSceduleDataBase}'")

def get_schedule_data_base(filenameSceduleDataBase, numOfSelectedSchedule=-1):#TODO IsElder. How get status from table? Take it from Personal's sheet
    wbSceduleDataBase = openpyxl.load_workbook(filename=filenameSceduleDataBase)
    sheet = wbSceduleDataBase.worksheets[0]
    lengthOfDataBase = sheet.cell(row=1, column=1).value
    #Check parameters
    if numOfSelectedSchedule == -1:
        selectedNumber = lengthOfDataBase
    else:
        selectedNumber = numOfSelectedSchedule
    if ((selectedNumber < 1) or (selectedNumber > lengthOfDataBase)):
        print(ERROR_STR_HEAD + "! (get_schedule_data_base)\n\t\tOut of range!!!") #it gives the last schedule in DB
        return None
    print(f"\n\tTable's selected number: {selectedNumber}")
    #--
    #Init tables parameters
    tableWidth = 1 + 7 + SPACE_BETWEEN_TABLES  # +1 - begin from 1; 7 - week length; +1 - space between tables
    startingPointColumn = 1 + ((selectedNumber-1) * tableWidth)
    staffLength = 0
    r = 4
    while sheet.cell(row=r, column=startingPointColumn).value != None:
        staffLength += 1
        r += 1
    tableHeight = 1 + staffLength + SPACE_BETWEEN_TABLES
    startingPointRow = 3
    #--
    #get selected schedule from data base
    outputSchedule = WeekScheduleExcelType()
    # print("STARTING POINTS ", startingPointRow, startingPointColumn)
    for i in range(1, tableHeight-SPACE_BETWEEN_TABLES):
        # print(sheet.cell(row=startingPointRow+i, column=startingPointColumn).value, end="\t")
        bufferName = sheet.cell(row=startingPointRow+i, column=startingPointColumn).value
        bufferIsElder = sheet.cell(row=startingPointRow+i, column=startingPointColumn).font.b
        bufferShifts = list[ShiftType]()
        for j in range(1, tableWidth-SPACE_BETWEEN_TABLES):
            # print(sheet.cell(row=startingPointRow+i, column=startingPointColumn+j).value, end="\t")
            if sheet.cell(row=startingPointRow+i, column=startingPointColumn+j).value != CHAR_CROSS:
                if sheet.cell(row=startingPointRow+i, column=startingPointColumn+j).value == CHAR_TRACK:
                    bufferShifts.append((j, 1.0, "Truck"))
                elif sheet.cell(row=startingPointRow+i, column=startingPointColumn+j).value == CHAR_HALL:
                    bufferShifts.append((j, 1.0, "Hall"))
                else:
                    bufferShifts.append((j, 0.5, "Hall"))
        outputSchedule.append(EmployeeCard(name=bufferName, isElder=bufferIsElder, shifts=bufferShifts))
        # print()
    #--
    return outputSchedule

def get_schedule_data_base_staff(filenameSceduleDataBase):
    print("BEGIN STAFF")

def get_schedule_data_base_track(filenameSceduleDataBase):
    print("BEGIN TRACK")

def check_output_and_update_schedule(searchMode="fast"):#TEST FUNCTION!!!
    lengthOfPool = output_pool_of_schedule_to_excel(FILENAME_POOL_TIMETABLE, searchMode)  # "fast", "part", "full"
    while True:
        numChoosingTimetable = input("\nEnter number of choosing schedule: ")
        try:
            numChoosingTimetable = int(numChoosingTimetable)
        except:
            print(ERROR_STR_HEAD + " TYPE! (check_output_and_update_schedule)\n\t\tPlease, enter integer value.")
        else:
            if ((numChoosingTimetable < 1) or (numChoosingTimetable > lengthOfPool)):
                print(f"{ERROR_STR_HEAD} VALUE! (check_output_and_update_schedule)\n\t\tPlease, enter value from 1 to {lengthOfPool}.")
            else:
                break
    update_schedule_data_base(FILENAME_SCHEDULE_DATA_BASE, FILENAME_POOL_TIMETABLE, numChoosingTimetable)

def check_get_schdedule(numOfSelectedSchedule=-1):
    schedule = get_schedule_data_base(FILENAME_SCHEDULE_DATA_BASE, numOfSelectedSchedule)
    if schedule != None:
        for i in range(len(schedule)):
            print(schedule[i].Name, schedule[i].IsElder, schedule[i].Shifts)
    else:
        print(ERROR_STR_HEAD + "! (check_get_schdedule)\n\t\tEmpty data base!")

def check_full(searchMode="fast"):
    init_schedule_data_base(FILENAME_SCHEDULE_DATA_BASE)
    check_output_and_update_schedule(searchMode) # "fast", "part", "full"
    update_schedule_data_base_staff(FILENAME_SCHEDULE_DATA_BASE)
    update_schedule_data_base_track(FILENAME_SCHEDULE_DATA_BASE)
    check_get_schdedule()

if __name__ == "__main__":
    check_full("fast")
