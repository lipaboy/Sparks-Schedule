import excel.ExcelChecks as eCheck
from excel.ExcelCore import FILENAME_SCHEDULE_DATA_BASE, FILENAME_POOL_TIMETABLE
from utility.DocumentControl import openExcelDocumentProcess, closeExcelDocumentProcess

if __name__ == "__main__":
    localFilenameScheduleDataBase = "../" + FILENAME_SCHEDULE_DATA_BASE
    localFilenamePoolTimetable = "../" + FILENAME_POOL_TIMETABLE

    closeExcelDocumentProcess(localFilenameScheduleDataBase)

    eCheck.check_full(localFilenameScheduleDataBase, localFilenamePoolTimetable, "fast")

    # init_schedule_data_base(localFilenameScheduleDataBase)
    openExcelDocumentProcess(localFilenameScheduleDataBase)
