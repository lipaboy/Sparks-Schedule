import datetime
import tkinter as tk
import os
import excel.mainExcel as ExcelCore
from tkcalendar import Calendar, DateEntry
from tkinter import ttk


class MainWindow:
    def __init__(self):
        self.window = tk.Tk()

        # Set Height and Width of Window
        self.window.geometry("350x150")

        # Set the Title according to desire
        self.window.title("Расписание Спаркса")

        self.makeScheduleButton = tk.Button(text="Сформировать расписание", command=self.makeScheduleRequest)
        self.chooseScheduleButton = tk.Button(text="Выбрать", command=self.chooseScheduleRequest)
        # self.makeScheduleButton.grid(padx=20, pady=20)

        self.statusLabel = tk.Label(text="Здравствуй, Спаркс!")

        self.chooseIdLabel = tk.Label(text="Выберете расписание (номер):")
        self.inputField = tk.Entry(fg="yellow", bg="purple", width=15)

        ttk.Label(text='Choose date').pack(padx=10, pady=10)
        self.cal = DateEntry(width=12, background='darkblue',
                        foreground='white', borderwidth=2)
        self.cal.pack(padx=10, pady=10)
        print(self.cal.get_date())

        self.statusLabel.pack(expand=True)
        self.makeScheduleButton.pack(expand=True)
        # self.inputField.pack(expand=True)
        # self.inputField.pack_forget()

    def chooseScheduleRequest(self):

        ExcelCore.update_schedule_data_base(ExcelCore.FILENAME_SCHEDULE_DATA_BASE,
                                            ExcelCore.FILENAME_POOL_TIMETABLE,
                                            int(self.inputField.get()))
        self.chooseIdLabel.pack_forget()
        self.inputField.pack_forget()
        self.chooseScheduleButton.pack_forget()
        self.statusLabel.config(text='Заебись!')

    def makeScheduleRequest(self):
        """"""
        """ TODO: Вызвать функцию поиска расписания для вывода его в эксель"""
        ExcelCore.output_pool_of_schedule_to_excel(
            ExcelCore.FILENAME_POOL_TIMETABLE,
            'fast',
            # datetime.datetime.strptime(self.cal.get_date(), "%m-%d-%Y").date())
            self.cal.get_date())
        os.system('start EXCEL.exe ' + ExcelCore.FILENAME_POOL_TIMETABLE)

        self.statusLabel.config(text='Варианты расписаний сформированы.')
        self.makeScheduleButton.pack_forget()
        self.chooseIdLabel.pack(pady=10)
        self.inputField.pack(pady=5)
        self.chooseScheduleButton.pack(pady=5)
        pass

    def mainloop(self):
        self.window.mainloop()


if __name__ == "__main__":
    import os
    cwd = os.getcwd()
    print(cwd)
    window = MainWindow()
    window.mainloop()
