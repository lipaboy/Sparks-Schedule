import sys
import time
import datetime
import excel.ExcelCore as eCore
import search.SparksScheduleSearch as SearchCore
from utility.DocumentControl import openExcelDocumentProcess, closeExcelDocumentProcess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow,
    QGridLayout, QVBoxLayout, QWidget,
    QLabel, QPushButton, QLineEdit, QSlider,
    QCalendarWidget, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, QMutex, pyqtSignal

WINDOW_SIZE_WIDTH = 500
WINDOW_SIZE_HEIGHT = 600
TEXT_SIZE_H1 = 30
TEXT_SIZE_H2 = 24
TEXT_SIZE_MAIN = 18
PROGRESS_BAR_MAX = 100
SPEED_DELAY = 0.01
MIN_DELAY = 0.1
MAX_DELAY = 1
COLOR_PALETTE = {
    #Yellow BLOCK
    "BACKGROUND": "#F4D03F",
    "TEXT": "#000000",
    "HEAD": "#DC7633",
    "UNIT": "#F5B041"
    #Blue, green, pink BLOCK
    # "BACKGROUND": "#514ED9",
    # "TEXT": "#008209",
    # "HEAD": "#CE0071",
    # "UNIT": "#E73A98"
}

mutex = QMutex()

def excepthook(cls, exception, traceback):
    print('calling excepthook...')
    # logger.error("{}".format(exception))

class TaskThread(QThread):

    started = pyqtSignal(str)
    progress_value = pyqtSignal(int)
    progress_text = pyqtSignal(str)
    finished_value = pyqtSignal(int)
    finished_text = pyqtSignal(str)

    def __init__(self, task_id = 1, mode = "fast", day = datetime.date.today()):
        super().__init__()
        self.localFilenameScheduleDataBase = "../" + eCore.FILENAME_SCHEDULE_DATA_BASE
        self.localFilenamePoolTimetable = "../" + eCore.FILENAME_POOL_TIMETABLE
        self.task_id = task_id
        self.mode = mode
        self.day = day

        modeSpeedDict = {
            SearchCore.MODE_LIST[0]: 0.3,
            SearchCore.MODE_LIST[1]: 0.5,
            SearchCore.MODE_LIST[2]: 2
        }
        self.modeSpeed = modeSpeedDict[self.mode]

    def run(self):
        if self.task_id == 1:
            self.started.emit(f"Task {self.task_id} *** START!")
            for i in range(1, PROGRESS_BAR_MAX):
                time.sleep(MAX_DELAY*self.modeSpeed)
                if mutex.tryLock():
                    self.progress_value.emit(i)
                    self.progress_text.emit(f"Task {self.task_id} >>> {i}")
                    mutex.unlock()
                else:
                    break
            self.finished_text.emit(f"Task {self.task_id} *** END!")
        elif self.task_id == 2:
            self.started.emit(f"Task {self.task_id} *** START!")
            lengthOfPool = eCore.output_pool_of_schedule_to_excel(self.localFilenameScheduleDataBase, self.localFilenamePoolTimetable, self.mode, self.day)
            mutex.lock()
            #TODO Race condition REWORK
            time.sleep(MAX_DELAY * self.modeSpeed)
            self.finished_value.emit(lengthOfPool)
            self.finished_text.emit(f"Task {self.task_id} *** END!")

class MainWindow(QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()

        self.lengthOfPool = -1

        self.setWindowTitle("Генератор расписаний")
        self.setMinimumSize(WINDOW_SIZE_WIDTH, WINDOW_SIZE_HEIGHT)
        self.setStyleSheet(
            "QMainWindow{background-color: " + COLOR_PALETTE["BACKGROUND"] + "};"
            "QPushButton{padding: 10px;}"
            "QLineEdit{padding: 10px;}"
        )

        self.threads = {}
        self.localFilenameScheduleDataBase = "../" + eCore.FILENAME_SCHEDULE_DATA_BASE
        self.localFilenamePoolTimetable = "../" + eCore.FILENAME_POOL_TIMETABLE
        self.modeList = list(map(lambda x: x.upper(), SearchCore.MODE_LIST))
        self.labelImformList = [
            ["Идёт генерация расписаний...", "Генерация завершена!"],
            ["Ожидайте...", "Варианты расписаний созданы!", "Введите номер расписания!", "Введите число от 1 до"],
            ["Расписание сохранено!", "Ошибка!\nРасписание НЕ добавлено!!!"]
        ]

        self.create_head()
        self.create_Top_widget()
        self.create_Middle_First_widget()
        self.create_Middle_Second_widget()
        self.create_Bottom_widget()

        self.widgetHead = QWidget()
        self.widgetHead.setStyleSheet(f'''
            background-color: {COLOR_PALETTE["HEAD"]};
            color: {COLOR_PALETTE["TEXT"]};
            border-radius: 10px;
        ''')
        self.widgetHead.setFixedHeight(int(self.window().height() / 5))
        self.widgetHead.setLayout(self.layoutHead)

        self.widgetTop = QWidget()
        self.widgetTop.setStyleSheet(f'''
            color: {COLOR_PALETTE["TEXT"]};
        ''')
        self.widgetTop.setLayout(self.layoutTop)

        self.widgetMiddleFirst = QWidget()
        self.widgetMiddleFirst.setStyleSheet(f'''
            color: {COLOR_PALETTE["TEXT"]};
        ''')
        self.widgetMiddleFirst.setFixedHeight(int(self.window().height() / 4))
        self.widgetMiddleFirst.setLayout(self.layoutMiddleFirst)
        self.widgetMiddleFirst.hide()

        self.widgetMiddleSecond = QWidget()
        self.widgetMiddleSecond.setStyleSheet(f'''
            color: {COLOR_PALETTE["TEXT"]};
        ''')
        self.widgetMiddleSecond.setLayout(self.layoutMiddleSecond)
        self.widgetMiddleSecond.hide()

        self.widgetBottom = QWidget()
        self.widgetBottom.setStyleSheet(f'''
            color: {COLOR_PALETTE["TEXT"]};
        ''')
        self.widgetBottom.setLayout(self.layoutBottom)
        self.widgetBottom.hide()

        mainLayout = QGridLayout()
        mainLayout.addWidget(self.widgetHead, 0, 0)
        mainLayout.addWidget(self.widgetTop, 1, 0)
        mainLayout.addWidget(self.widgetMiddleFirst, 1, 0)
        mainLayout.addWidget(self.widgetMiddleSecond, 3, 0)
        mainLayout.addWidget(self.widgetBottom, 1, 0)

        self.mainWidget = QWidget()
        self.mainWidget.setLayout(mainLayout)

        self.setCentralWidget(self.mainWidget)

    def create_head(self):
        self.layoutHead = QVBoxLayout()
        wLabelHead = QLabel("Генератор расписаний!")
        font = wLabelHead.font()
        font.setPointSize(TEXT_SIZE_H1)
        font.setBold(True)
        wLabelHead.setFont(font)
        wLabelHead.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.layoutHead.addWidget(wLabelHead)

    def create_Top_widget(self):
        self.layoutTop = QVBoxLayout()

        wLabelCalendar = QLabel("Выберете неделю:")
        font = wLabelCalendar.font()
        font.setPointSize(TEXT_SIZE_H2)
        wLabelCalendar.setFont(font)
        wLabelCalendar.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        self.wCalendar = QCalendarWidget()
        self.wCalendar.setStyleSheet(f'''
            background-color: {COLOR_PALETTE["UNIT"]};
        ''')
        font = self.wCalendar.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        self.wCalendar.setFont(font)

        self.wLabelSlider = QLabel()
        font = self.wLabelSlider.font()
        font.setPointSize(TEXT_SIZE_H2)
        font.setItalic(True)
        self.wLabelSlider.setFont(font)
        self.wLabelSlider.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.wSliderMode = QSlider()
        self.wSliderMode.setStyleSheet(f'''
            background-color: {COLOR_PALETTE["UNIT"]};
        ''')
        self.wSliderMode.setOrientation(Qt.Orientation.Horizontal)  # Горизонтальный, по стандарту вертикальный
        self.wSliderMode.setRange(0, 2)
        self.wSliderMode.setValue(0)
        self.wSliderMode.setSingleStep(1)
        self.wSliderMode.setPageStep(1)
        self.wSliderMode.setSizeIncrement(100, 100)
        self.wSliderMode.valueChanged.connect(self.slider_mode_value_changed)
        self.wLabelSlider.setText(f"Режим генератора: {self.modeList[self.wSliderMode.value()]}")

        self.wButtonCreate = QPushButton("Создать расписания!")
        font = self.wButtonCreate.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        self.wButtonCreate.setFont(font)
        self.wButtonCreate.clicked.connect(self.button_create_clicked)

        self.layoutTop.addWidget(wLabelCalendar)
        self.layoutTop.addWidget(self.wCalendar)
        self.layoutTop.addWidget(self.wLabelSlider)
        self.layoutTop.addWidget(self.wSliderMode)
        self.layoutTop.addWidget(self.wButtonCreate)

    def create_Middle_First_widget(self):
        self.layoutMiddleFirst = QVBoxLayout()


        self.wLabelInform1 = QLabel(self.labelImformList[0][0])
        font = self.wLabelInform1.font()
        font.setPointSize(TEXT_SIZE_H2)
        self.wLabelInform1.setFont(font)
        self.wLabelInform1.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        self.wProgressBar = QProgressBar()
        self.wProgressBar.setRange(0, PROGRESS_BAR_MAX)
        font = self.wProgressBar.font()
        font.setPointSize(TEXT_SIZE_H2)
        self.wProgressBar.setFont(font)

        self.layoutMiddleFirst.addWidget(self.wLabelInform1)
        self.layoutMiddleFirst.addWidget(self.wProgressBar)

    def create_Middle_Second_widget(self):
        self.layoutMiddleSecond = QVBoxLayout()

        self.wLabelInform2 = QLabel(self.labelImformList[1][0])
        font = self.wLabelInform2.font()
        font.setPointSize(TEXT_SIZE_H2)
        self.wLabelInform2.setFont(font)
        self.wLabelInform2.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        self.wLineEdit = QLineEdit()
        self.wLineEdit.setVisible(False)
        self.wLineEdit.setPlaceholderText("Введите номер расписания...")
        self.wLineEdit.setMaxLength(2)
        font = self.wLineEdit.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        self.wLineEdit.setFont(font)
        self.wLineEdit.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.wLineEdit.returnPressed.connect(self.button_select_clicked)

        self.wButtonSelect = QPushButton("Выбрать")
        self.wButtonSelect.setVisible(False)
        font = self.wButtonSelect.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        self.wButtonSelect.setFont(font)
        self.wButtonSelect.clicked.connect(self.button_select_clicked)

        self.wButtonBack = QPushButton("Назад")
        self.wButtonBack.setVisible(False)
        font = self.wButtonBack.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        self.wButtonBack.setFont(font)
        self.wButtonBack.clicked.connect(self.button_back_clicked)

        self.layoutMiddleSecond.addWidget(self.wLabelInform2)
        self.layoutMiddleSecond.addWidget(self.wLineEdit)
        self.layoutMiddleSecond.addWidget(self.wButtonSelect)
        self.layoutMiddleSecond.addWidget(self.wButtonBack)

    def create_Bottom_widget(self):
        self.layoutBottom = QVBoxLayout()

        self.wLabelInform3 = QLabel(self.labelImformList[2][0])
        font = self.wLabelInform3.font()
        font.setPointSize(TEXT_SIZE_H2)
        self.wLabelInform3.setFont(font)
        self.wLabelInform3.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        wButtonBackToMain = QPushButton("На главную")
        font = wButtonBackToMain.font()
        font.setPointSize(TEXT_SIZE_MAIN)
        wButtonBackToMain.setFont(font)
        wButtonBackToMain.clicked.connect(self.button_back_to_main_clicked)

        self.layoutBottom.addWidget(self.wLabelInform3)
        self.layoutBottom.addWidget(wButtonBackToMain)

    def slider_mode_value_changed(self, i):
        self.wLabelSlider.setText(f"Режим генератора: {self.modeList[i]}")

    def button_create_clicked(self):
        self.widgetTop.hide()
        self.widgetMiddleFirst.show()
        self.widgetMiddleSecond.show()
        self.wProgressBar.setValue(self.wProgressBar.minimum())

        selectedDay = datetime.date(self.wCalendar.selectedDate().year(), self.wCalendar.selectedDate().month(), self.wCalendar.selectedDate().day())
        thread1 = TaskThread(task_id=1, mode=self.modeList[self.wSliderMode.value()].lower())
        thread2 = TaskThread(task_id=2, mode=self.modeList[self.wSliderMode.value()].lower(), day=selectedDay)

        thread1.started.connect(self.thread_update_status)
        thread1.progress_value.connect(lambda value: self.thread_update_progress_bar(value, 1))
        thread1.progress_text.connect(self.thread_update_status)
        thread1.finished_value.connect(lambda value: self.thread_finish(value, 1))
        thread1.finished_text.connect(self.thread_update_status)

        thread2.started.connect(self.thread_update_status)
        thread2.finished_value.connect(lambda value: self.thread_finish(value, 2))
        thread2.finished_text.connect(self.thread_update_status)

        self.threads[1] = thread1
        self.threads[2] = thread2
        thread1.start()
        thread2.start()

    def button_select_clicked(self):
        if self.wLineEdit.text() == "":
            self.wLabelInform2.setText(self.labelImformList[1][2])
        else:
            numOfSelectSchedule = -1
            str = self.wLineEdit.text()
            try:
                numOfSelectSchedule = int(str)
            except:
                self.wLabelInform2.setText(f"{self.labelImformList[1][3]} {self.lengthOfPool}!")
            else:
                if ((numOfSelectSchedule < 1) or (numOfSelectSchedule > self.lengthOfPool)):
                    self.wLabelInform2.setText(f"{self.labelImformList[1][3]} {self.lengthOfPool}!")
                else:
                    self.widgetMiddleFirst.hide()
                    self.widgetMiddleSecond.hide()
                    self.widgetBottom.show()
                    closeExcelDocumentProcess(eCore.FILENAME_POOL_TIMETABLE)
                    closeExcelDocumentProcess(eCore.FILENAME_SCHEDULE_DATA_BASE)
                    errorFlag = eCore.update_schedule_data_base(self.localFilenameScheduleDataBase,
                                                                    self.localFilenamePoolTimetable,
                                                                    numOfSelectSchedule)
                    if errorFlag:
                        self.wLabelInform3.setText(self.labelImformList[2][0])
                        openExcelDocumentProcess(self.localFilenameScheduleDataBase)
                    else:
                        self.wLabelInform3.setText(self.labelImformList[2][1])

    def button_back_clicked(self):
        self.lengthOfPool = -1
        self.widgetMiddleFirst.hide()
        self.widgetMiddleSecond.hide()
        self.widgetTop.show()
        self.wLabelInform1.setText(self.labelImformList[0][0])
        self.wLabelInform2.setText(self.labelImformList[1][0])
        self.wLineEdit.setVisible(False)
        self.wButtonSelect.setVisible(False)
        self.wButtonBack.setVisible(False)
        self.wLineEdit.setText("")

    def button_back_to_main_clicked(self):
        self.lengthOfPool = -1
        self.widgetBottom.hide()
        self.widgetTop.show()
        self.wLabelInform1.setText(self.labelImformList[0][0])
        self.wLabelInform2.setText(self.labelImformList[1][0])
        self.wLineEdit.setVisible(False)
        self.wButtonSelect.setVisible(False)
        self.wButtonBack.setVisible(False)
        self.wLineEdit.setText("")

    def thread_update_status(self, text):
        print(text)

    def thread_update_progress_bar(self, value, task_id):
        if task_id == 1 and self.lengthOfPool == -1:
            self.wProgressBar.setValue(value)

    def thread_finish(self, value, task_id):
        if task_id == 1:
            pass
        elif task_id == 2:
            self.lengthOfPool = value
            while self.wProgressBar.value() != self.wProgressBar.maximum():
                time.sleep(SPEED_DELAY)
                self.wProgressBar.setValue(self.wProgressBar.value()+1)
            self.wLabelInform1.setText(self.labelImformList[0][1])
            self.wLabelInform2.setText(self.labelImformList[1][1])
            self.wLineEdit.setVisible(True)
            self.wButtonSelect.setVisible(True)
            self.wButtonBack.setVisible(True)
            closeExcelDocumentProcess(eCore.FILENAME_POOL_TIMETABLE)
            openExcelDocumentProcess(self.localFilenamePoolTimetable)
            mutex.unlock()


if __name__ == '__main__':
    sys.excepthook = excepthook
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
