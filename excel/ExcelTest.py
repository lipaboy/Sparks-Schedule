import excel.ExcelChecks as eCheck
import excel.ExcelCore as eCore
from utility.DocumentControl import openExcelDocumentProcess, closeExcelDocumentProcess

if __name__ == "__main__":
    localFilenameScheduleDataBase = "../" + eCore.FILENAME_SCHEDULE_DATA_BASE
    localFilenamePoolTimetable = "../" + eCore.FILENAME_POOL_TIMETABLE

    eCheck.check_full(localFilenameScheduleDataBase, localFilenamePoolTimetable, "fast")
    # print(eCore.get_truck_distribution_data_base(localFilenameScheduleDataBase))
