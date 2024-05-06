from excel.ExcelCore import (
    output_pool_of_schedule_to_excel, update_schedule_data_base,
    get_schedule_list_data_base, get_schedule_list_staff, get_schedule_list_coefficients
)
from utility.DocumentControl import openExcelDocumentProcess, closeExcelDocumentProcess

ERROR_STR_HEAD = "\n\t**ERROR in excel.ExcelChecks"#! (numOfFunction)\n\t\t

def check_output_and_update_schedule(filenameScheduleDataBase, filenamePoolTimetable, searchMode="fast"):#TEST FUNCTION!!!
    closeExcelDocumentProcess(filenamePoolTimetable)
    lengthOfPool = output_pool_of_schedule_to_excel(filenameScheduleDataBase, filenamePoolTimetable, searchMode)  # "fast", "part", "full"
    openExcelDocumentProcess(filenamePoolTimetable)
    lengthOfPool = 12
    while True:
        numChoosingTimetable = input("\nEnter number of choosing schedule: ")
        try:
            numChoosingTimetable = int(numChoosingTimetable)
        except:
            print(ERROR_STR_HEAD + " TYPE! (check_output_and_update_schedule)\n\t\tPlease, enter integer value.")
        else:
            if (numChoosingTimetable < 1) or (numChoosingTimetable > lengthOfPool):
                print(f"{ERROR_STR_HEAD} VALUE! (check_output_and_update_schedule)\n\t\tPlease, enter value from 1 to {lengthOfPool}.")
            else:
                break

    closeExcelDocumentProcess(filenameScheduleDataBase)
    update_schedule_data_base(filenameScheduleDataBase, filenamePoolTimetable, numChoosingTimetable)
    openExcelDocumentProcess(filenameScheduleDataBase)

#BLOCK CHECK_GET
def check_get_list_DB(filenameScheduleDataBase, numOfSelectedSchedule=-1):
    schedule = get_schedule_list_data_base(filenameScheduleDataBase, numOfSelectedSchedule)
    if schedule != None:
        print("\n\t\t\tGet list DB:")
        print("\tThe last schedule in data base:")
        for i in range(len(schedule.EmployeeCards)):
            print(schedule.EmployeeCards[i].Name, schedule.EmployeeCards[i].IsElder, schedule.EmployeeCards[i].Shifts)
    else:
        print(ERROR_STR_HEAD + "! (check_get_list_DB)\n\t\tEmpty data base!!!")
        # return

def check_get_list_staff(filenameScheduleDataBase):
    #1 - NAME/IS_ELDER; 2 - NAME/NUM_OF_TRUCK; 3 - NAME/NUM_OF_HALL; 4 - NAME/PREFER_NUM_OF_HALL; 5 - NAME/UNDESIRABLE_DAYS.
    listOfLabels = [
        "1 - NAME/IS_ELDER",
        "2 - NAME/NUM_OF_TRUCK",
        "3 - NAME/NUM_OF_HALL",
        "4 - NAME/PREFER_NUM_OF_HALL",
        "5 - NAME/UNDESIRABLE_DAYS"
    ]
    maxIndex = 5
    print("\n\t\t\tGet list STAFF:")
    for i in range(1, maxIndex+1):
        print(f"{listOfLabels[i-1]}:")
        data = get_schedule_list_staff(filenameScheduleDataBase, i)
        print(f"\t{data}")

def check_get_list_coefficients(filenameScheduleDataBase):
    print("\n\t\t\tGet list COEFFICIENTS:")
    data = get_schedule_list_coefficients(filenameScheduleDataBase)
    print("Coefficients: ", data)

def check_get_block(filenameScheduleDataBase):
    check_get_list_DB(filenameScheduleDataBase)
    check_get_list_staff(filenameScheduleDataBase)
    check_get_list_coefficients(filenameScheduleDataBase)

#FULL CHECKs
def check_full(filenameScheduleDataBase, filenamePoolTimetable, searchMode="fast"):

    #MAIN CHECKS
    check_output_and_update_schedule(filenameScheduleDataBase, filenamePoolTimetable, searchMode) # "fast", "part", "full"

    #---GET_BLOCK---
    check_get_block(filenameScheduleDataBase)
