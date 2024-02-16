from Schedule import Schedule
from typing import TypeAlias

""" Номер дня (1 - пн, 2 - вт, .. 7 - вск) """
DayType: TypeAlias = int
""" Длина рабочей смены """
TurnLength: TypeAlias = float
""" Место, где работает сотрудник (Hall, Truck) """
PlaceToWork: TypeAlias = str
""" Формат выгрузки в эксель модуль """
ExcelType: TypeAlias = dict[str, list[tuple[DayType, TurnLength, PlaceToWork]]]


class EmployeeFavor:
    def __init__(self):
        self.ghostNames = {
            1: 'Даша',
            2: 'Маша',
            3: 'Саша',
            4: 'Лада',
            5: 'Артур'
        }
        self.eldersNames = {
            1: 'Вован',
            2: 'Люба'
        }

        # дни, в которых есть нецелые смены
        self.partTimeDays = {
            4: 0.5,
            5: 0.5
        }
        self.ghostLimits = {
            1: 3.5,
            2: 3.5,
            3: 3,
            4: 3,
            5: 1
        }
        self.undesirableGhostDays = {
            1: [1, 2],
            2: [3, 5, 6],
            3: [],
            4: [1, 2, 3, 7],
            5: []
        }

        self.week = range(1, 7 + 1)
        self.oneTimeWeek = range(1, 3 + 1)
        self.pairWeek = range(3 + 1, 7 + 1)
        self.turnDayLen = [self.partTimeDays.get(i)
                           if i in self.partTimeDays else 1.0 for i in self.week]

    def toExcel(self, schedule: Schedule) -> ExcelType:
        excelDict = ExcelType()

        for ghostId, ghostName in self.ghostNames.items():
            turnList = []

            for i, day in zip(schedule.ghostOneTime, self.oneTimeWeek):
                if i == ghostId:
                    turnList.append((day, 1.0, 'Hall'))

            for p, day in zip(schedule.ghostPair, self.pairWeek):
                if ghostId == p[0] or ghostId == p[1]:
                    turnList.append((day, self.turnDayLen[day - 1] if ghostId == p[0] else 1.0, 'Hall'))

            excelDict[ghostName] = turnList

        return excelDict

    def print(self, schedule: Schedule):
        week = range(1, 7 + 1)
        nameMaxLen = 6
        print(''.rjust(nameMaxLen + 1), end='')
        for day in ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']:
            print(day.rjust(4), end='')
        print()

        # todo: don't forget it
        # print(self.eldersNames[1].rjust(nameMaxLen) + ':', end='')
        # for i in week:
        #     turnName = 'C'
        #     print((turnName if i in self.schedule.vovan else 'x').rjust(4), end='')
        # print()
        # print(self.eldersNames[2].rjust(nameMaxLen) + ':', end='')
        # for i in week:
        #     turnName = 'C'
        #     print((turnName if i not in self.schedule.vovan else 'x').rjust(4), end='')
        # print()
        #
        # print('---------------------'.rjust((nameMaxLen + 1 + 7 * 4) // 5 * 4))

        oneTimeWeek = [1, 2, 3]
        for [ghostId, name] in self.ghostNames.items():
            print(name.rjust(nameMaxLen) + ':', end='')
            for i in week:
                if i in oneTimeWeek:
                    turnName = 'C'
                    print((turnName if schedule.ghostOneTime[i - 1] == ghostId else 'x')
                          .rjust(4), end='')
                else:
                    turnName = 'C2' if (i in self.partTimeDays
                                        and ghostId == schedule.ghostPair[i - 4][0]) else 'С'
                    print((turnName if ghostId in schedule.ghostPair[i - 4] else 'x')
                          .rjust(4), end='')
            print()
