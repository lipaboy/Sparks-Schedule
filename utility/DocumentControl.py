import os
import psutil


def closeExcelDocumentProcess(excelFileName: str):
    processList = psutil.process_iter()
    for proc in processList:
        try:
            if 'excel' in proc.name().lower():
                # print(proc.name(), proc.pid, proc.cmdline(), proc.open_files())
                if any(excelFileName in path.path for path in proc.open_files()):
                    # print("close " + proc.name())
                    proc.kill()
        except psutil.AccessDenied:
            continue

def openExcelDocumentProcess(excelFileName: str):
    os.system('start EXCEL.exe ' + excelFileName)
