""" Номер дня (1 - пн, 2 - вт, .. 7 - вск) """
DayType = int
""" Длина рабочей смены (ЧТ, ПТ - 0.5 (excel - 'С2'))"""
ShiftLength = float
""" Место, где работает сотрудник (Hall - в зале (excel - 'С'), Truck - тачка (excel - 'Т')) """
PlaceToWork = str

ShiftType = tuple[DayType, ShiftLength, PlaceToWork]

""" Карточка сотрудника """
class EmployeeCard:
    def __init__(self, name: str, isElder: bool, shifts: list[ShiftType]):
        self.Name = name
        """ Является сотрудник старшим (True - старший) """
        self.IsElder = isElder
        """ Смены сотрудника """
        self.Shifts = shifts

""" Формат выгрузки в эксель модуль. Один ExcelType - это одно расписание """

# employee name -> truck count
TruckDistributionType = dict[str, int]

class WeekScheduleExcelType:
    def __init__(self,
                 employeeCards=None,
                 trucks=None):
        if trucks is None:
            trucks = []
        if employeeCards is None:
            employeeCards = []
        self.EmployeeCards = employeeCards
        # employee name -> truck count
        self.Trucks = trucks


