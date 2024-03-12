""" Номер дня (1 - пн, 2 - вт, .. 7 - вск) """
DayType = int
""" Длина рабочей смены (ЧТ, ПТ - 0.5 (excel - 'С2'))"""
ShiftLength = float
""" Место, где работает сотрудник (Hall - в зале (excel - 'С'), Truck - тачка (excel - 'Т')) """
PlaceToWork = str
""" Карточка сотрудника """
# TODO: сделать как нормально
class EmployeeCard:
    def __init__(self, Name: str, IsElder: bool):
        self.Name = Name
        """ Является сотрудник старшим (True - старший) """
        self.IsElder = IsElder
        # self.lstDays = list[tuple]

""" Формат выгрузки в эксель модуль. Один ExcelType - это одно расписание """
ScheduleExtractionExcelType = dict[EmployeeCard, list[tuple[DayType, ShiftLength, PlaceToWork]]]

# todo:
# ScheduleExtractionExcelType = list[EmployeeCard]


