import datetime
import sys
import tkinter as tk
import os
import traceback
from multiprocessing import Process
import threading
from tkinter.constants import HORIZONTAL
from tkinter.ttk import Progressbar
import logging

import psutil

import excel.ExcelCore as ExcelCore
from tkcalendar import Calendar, DateEntry
from tkinter import ttk

import search.SparksScheduleSearch
from utility.DocumentControl import *


class MainWindow:
    def __init__(self, isDebug=False):
        self.window = tk.Tk()
        self.isDebug = isDebug

        # Set Height and Width of Window
        self.window.geometry("350x400")

        # Set the Title according to desire
        self.window.title("Расписание Спаркса")

        self.makeScheduleButton = tk.Button(
            text="Сформировать расписание",
            command=lambda: threading.Thread(target=self.makeScheduleRequest).start()
        )
        self.progressBar = Progressbar(self.window, orient=HORIZONTAL, length=100, mode='indeterminate')

        self.chooseScheduleButton = tk.Button(text="Выбрать", command=self.chooseScheduleRequest)
        # self.makeScheduleButton.grid(padx=20, pady=20)
        self.speedLabel = tk.Label(text='')
        self.speedSlider = ttk.Scale(
            self.window,
            from_=1,
            to=3,
            orient='horizontal',
            command=self.__showSlider
        )
        self.speedSlider.set(1)

        self.helloLabel = tk.Label(text="Здравствуй, Спаркс!")
        self.statusLabel = tk.Label(text='')

        self.chooseIdLabel = tk.Label(text="Выберете расписание (номер):")
        self.inputField = tk.Entry(fg="yellow", bg="purple", width=15)

        ttk.Label(text='Выберете неделю').pack(padx=10, pady=10)
        nextWeek = datetime.date.today() + datetime.timedelta(days=7)
        self.calendar = DateEntry(width=12,
                                  background='darkblue',
                                  foreground='white',
                                  borderwidth=2,
                                  year=nextWeek.year,
                                  month=nextWeek.month,
                                  day=nextWeek.day
                                  )
        # self.calendar.config(date)
        self.calendar.pack(padx=10, pady=10)
        print(self.calendar.get_date())

        self.helloLabel.pack(expand=True)
        self.speedLabel.pack()
        self.speedSlider.pack()
        self.makeScheduleButton.pack(expand=True)
        # self.inputField.pack(expand=True)
        # self.inputField.pack_forget()

    def __getMode(self):
        return search.SparksScheduleSearch.MODE_LIST[int(self.speedSlider.get()) - 1]

    def __showSlider(self, event):
        value = int(self.speedSlider.get())
        self.speedLabel.config(
            text=search.SparksScheduleSearch.MODE_LIST[value - 1]
        )

    def chooseScheduleRequest(self):
        scheduleNum = self.inputField.get()
        if scheduleNum.isdigit():
            if self.isDebug:
                closeExcelDocumentProcess(ExcelCore.FILENAME_SCHEDULE_DATA_BASE)
            try:
                ExcelCore.update_schedule_data_base(ExcelCore.FILENAME_SCHEDULE_DATA_BASE,
                                                    ExcelCore.FILENAME_POOL_TIMETABLE,
                                                    int(scheduleNum))
            except Exception:
                self.statusLabel.config(text='Произошли ошибки', fg='#f00')
                mainLogger.error(traceback.format_exc())
            else:
                # if self.isDebug:
                openExcelDocumentProcess(ExcelCore.FILENAME_SCHEDULE_DATA_BASE)

                self.chooseIdLabel.pack_forget()
                self.inputField.pack_forget()
                self.chooseScheduleButton.pack_forget()
                self.helloLabel.config(text='Данные успешно сохранены!')
                self.helloLabel.pack(expand=True)
                self.statusLabel.config(text='', fg='#000')
                self.statusLabel.pack_forget()
        else:
            self.statusLabel.config(text='Неверный номер расписания', fg='#f00')

    def makeScheduleRequest(self):
        """"""
        self.makeScheduleButton.config(state='disabled')
        self.progressBar.pack()

        closeExcelDocumentProcess(ExcelCore.FILENAME_POOL_TIMETABLE)

        if not self.isDebug:
            generateSchedules = Process(
                target=ExcelCore.output_pool_of_schedule_to_excel,
                args=(ExcelCore.FILENAME_SCHEDULE_DATA_BASE,
                      ExcelCore.FILENAME_POOL_TIMETABLE,
                      self.__getMode(),
                      self.calendar.get_date()),
                daemon=True
            )
            generateSchedules.start()
            while True:
                generateSchedules.join(0.7 if self.__getMode() == 'full' else 0.1)
                if generateSchedules.exitcode is not None:
                    break
                self.progressBar['value'] += 30
        else:
            ExcelCore.output_pool_of_schedule_to_excel(
                ExcelCore.FILENAME_SCHEDULE_DATA_BASE,
                ExcelCore.FILENAME_POOL_TIMETABLE,
                searchMode=self.__getMode(),
                currentDay=self.calendar.get_date())

        openExcelDocumentProcess(ExcelCore.FILENAME_POOL_TIMETABLE)

        if len(self.statusLabel.cget('text')) <= 0:
            self.makeScheduleButton.config(text='Сформировать заново')
            self.statusLabel.config(text='Варианты расписаний сформированы.')
            self.statusLabel.pack(pady=5)
            self.helloLabel.pack_forget()
            # self.makeScheduleButton.pack_forget()
            self.chooseIdLabel.pack(pady=10)
            self.inputField.pack(pady=5)
            self.chooseScheduleButton.pack(pady=5)

        self.makeScheduleButton.config(state='normal')
        self.progressBar.pack_forget()

    def mainloop(self):
        self.window.mainloop()


if __name__ == "__main__":
    fileLogHandler = logging.FileHandler(filename='sparks.log', encoding="utf-8", mode="a")
    fileLogHandler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    mainLogger = logging.getLogger(__name__)
    mainLogger.setLevel(logging.INFO)
    mainLogger.addHandler(logging.StreamHandler(sys.stdout))
    mainLogger.addHandler(fileLogHandler)

    window = MainWindow(True if len(sys.argv) > 1 and sys.argv[1] == 'debug' else False)
    window.mainloop()
