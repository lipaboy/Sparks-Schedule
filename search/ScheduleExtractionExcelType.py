from typing import TypeAlias

""" Номер дня (1 - пн, 2 - вт, .. 7 - вск) """
DayType: TypeAlias = int
""" Длина рабочей смены """
ShiftLength: TypeAlias = float
""" Место, где работает сотрудник (Hall - в зале, Truck - тачка) """
PlaceToWork: TypeAlias = str
""" Карточка сотрудника """
class EmployeeCard:
    def __init__(self, Name: str, IsElder: bool):
        self.Name = Name
        """ Является сотрудник старшим (True - старший) """
        self.IsElder = IsElder
""" Формат выгрузки в эксель модуль. Один ExcelType - это одно расписание """
ScheduleExtractionExcelType: TypeAlias = dict[EmployeeCard, list[tuple[DayType, ShiftLength, PlaceToWork]]]