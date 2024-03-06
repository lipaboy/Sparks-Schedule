import tkinter as tk


class MainWindow:
    def __init__(self):
        self.window = tk.Tk()

        # Set Height and Width of Window
        self.window.geometry("350x150")

        # Set the Title according to desire
        self.window.title("Расписание Спаркса")

        self.makeScheduleButton = tk.Button(text="Сформировать расписание", command=self.makeScheduleRequest)
        # self.makeScheduleButton.grid(padx=20, pady=20)

        self.statusLabel = tk.Label(text="Здравствуй, Спаркс!")

        self.chooseIdLabel = tk.Label(text="Выберете расписание (номер):")
        self.inputField = tk.Entry(fg="yellow", bg="purple", width=15)

        self.statusLabel.pack(expand=True)
        self.makeScheduleButton.pack(expand=True)
        # self.inputField.pack(expand=True)
        # self.inputField.pack_forget()

    def makeScheduleRequest(self):
        self.statusLabel.config(text='Варианты расписаний сформированы.')
        self.makeScheduleButton.pack_forget()
        self.chooseIdLabel.pack(pady=10)
        self.inputField.pack(pady=5)
        pass

    def mainloop(self):
        self.window.mainloop()


if __name__ == "__main__":
    window = MainWindow()
    window.mainloop()
