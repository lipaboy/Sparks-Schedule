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

    def truckCount(self) -> int:
        return len(list(filter(lambda x: x[2] == 'Truck', self.Shifts)))

    def shiftCount(self) -> float:
        return sum(x[1] for x in self.Shifts)

""" Формат выгрузки в эксель модуль. Один ExcelType - это одно расписание """

class TruckElem:
    def __init__(self, trucksCount: int = 0, shiftsCount: float = 0):
        self.TruckCount = trucksCount
        self.ShiftCount = shiftsCount

    "Значение загруженности"
    def loadValue(self) -> float:
        return float(self.TruckCount) / self.ShiftCount if self.ShiftCount > 0.0 else 0.0

    def incTruck(self):
        self.TruckCount += 1
        self.ShiftCount += 1

    def incShift(self):
        self.ShiftCount += 1

# employee name -> truck count
TruckDistributionType = dict[str, TruckElem]

class WeekScheduleExcelType:
    def __init__(self,
                 employeeCards: list[EmployeeCard] = None,
                 trucks: TruckDistributionType = None):
        if employeeCards is None:
            employeeCards = list[EmployeeCard]()
        if trucks is None:
            trucks = TruckDistributionType()
        self.EmployeeCards = employeeCards
        # employee name -> truck count
        self.Trucks = trucks


