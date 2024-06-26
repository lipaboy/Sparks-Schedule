import copy

from search.Schedule import Schedule
from search.WeekScheduleExcelType import WeekScheduleExcelType, EmployeeCard, TruckDistributionType, TruckElem


class EmployeeFavor:
    def __init__(self):
        ""

        """ TODO: 
            Гениальная идея рефакторинга: создать список всех (пока не удалять раздельные бдшки сотрудников)
            Вована с Любой перевести на 6, 7 
        """

        self.namesDB = {
            'Даша': 1,
            'Маша': 2,
            'Саша': 3,
            'Лада': 4,
            'Артур': 5,
            'Вован': 6,
            'Люба': 7,
        }

        self.whoElders = {'Вован', 'Люба'}

        "База духов"
        self.ghostNames = dict[str, int]()

        "База старших"
        self.elderNames = dict[str, int]()

        for name, id in self.namesDB.items():
            if name not in self.whoElders:
                self.ghostNames[name] = id

        maxGhostId = len(self.ghostNames)
        for name, id in self.namesDB.items():
            if name in self.whoElders:
                self.elderNames[name] = id - maxGhostId

        self._ghostNamesById = {v: k for k, v in self.ghostNames.items()}
        self._elderNamesById = {v: k for k, v in self.elderNames.items()}

        " Распределение тачек "
        self.truckDistribution = {
            name: TruckElem(0, 0) for name in (list(self.elderNames.keys()) + list(self.ghostNames.keys()))
        }

        # дни, в которых есть нецелые смены
        self.partTimeDays = {
            4: 0.5,
            5: 0.5
        }

        " Сколько каждый дух может работать в неделю дней "
        self.ghostLimits = {
            1: 3.5,
            2: 3.5,
            3: 3.0,
            4: 3.0,
            5: 1.0
        }
        " Нежелательные дни духов "
        self.undesirableGhostDays = {
            1: [1, 2],
            2: [3, 5, 6],
            3: [],
            4: [1, 2, 3, 7],
            5: []
        }

        self.elderLimits = {
            1: 4.0,
            2: 3.0,
        }
        " Нежелательные дни старших "
        self.undesirableElderDays = {
            1: [1, 4],
            2: [7],
        }

        self._week = range(1, 7 + 1)
        self._oneTimeWeek = range(1, 3 + 1)
        self._pairWeek = range(3 + 1, 7 + 1)
        self._shiftDayLen = [self.partTimeDays.get(i)
                           if i in self.partTimeDays else 1.0 for i in self._week]

        self._nameStringMaxLen = 6

    def loadEldermen(self, eldermen: list[str]):
        self.elderNames.clear()
        self._elderNamesById.clear()
        for name, id in zip(eldermen, range(1, len(eldermen) + 1)):
            # name = name.strip()
            if len(name) <= 0:
                continue
            self.elderNames[name] = id
            self._elderNamesById[id] = name

    def loadGhostmen(self, ghostmen: list[str]):
        self.ghostNames.clear()
        self._ghostNamesById.clear()
        for name, id in zip(ghostmen, range(1, len(ghostmen) + 1)):
            # name = name.strip()
            if len(name) <= 0:
                continue
            self.ghostNames[name] = id
            self._ghostNamesById[id] = name

    def loadShiftCountsPrefs(self, preferences: dict[str, float]):
        for name, shiftAverage in preferences.items():
            if name in self.elderNames.keys():
                self.elderLimits[ self.elderNames[name] ] = shiftAverage
            else:
                self.ghostLimits[ self.ghostNames[name] ] = shiftAverage

    def __getHalfShiftEmployee(self, schedule: Schedule, day: int) -> str:
        if day in self.partTimeDays.keys():
            return self._ghostNamesById[ schedule.getWorkersAtDay(day)[1][0] ]
        else:
            return ''

    def __getWorkerNamesAtDay(self,
                             schedule: Schedule,
                             day: int):

        elderId, ghostIdList = schedule.getWorkersAtDay(day)
        elder = self._elderNamesById[elderId]
        return [elder] + [self._ghostNamesById[w] for w in ghostIdList]

    def __parseWhoUnderTruck(self,
                             schedule: Schedule,
                             trucks: TruckDistributionType,
                             day: int) -> str:
        """ Правило: ставим того за тачку, кто меньше всего за ней стоял
            Нельзя ставить за тачку того, кто находится на полусмене (С2)
        """
        pretenderNames = self.__getWorkerNamesAtDay(schedule, day)
        halfEmployee = self.__getHalfShiftEmployee(schedule, day)
        if halfEmployee != '':
            pretenderNames.remove(halfEmployee)
        worker = min([(name, trucks[name]) for name in pretenderNames], key=lambda x: x[1].loadValue())
        return worker[0]

    def __getNameForTruck(self, name, weekTrucks, day):
        return 'Truck' if name == weekTrucks[day - 1] else 'Hall'

    def toExcel(self, schedule: Schedule) -> WeekScheduleExcelType:
        employeeCards = list[EmployeeCard]()
        trucks = copy.deepcopy(self.truckDistribution)

        weekTrucks = []
        for day in self._week:
            name = self.__parseWhoUnderTruck(schedule, trucks, day)
            weekTrucks.append(name)
            " Увеличиваем количество тачек для сотрудника на 1 с каждым днём "
            trucks[name].incTruck()
            workerNames = self.__getWorkerNamesAtDay(schedule, day)
            for w in workerNames:
                if name != w:
                    trucks[w].incShift()

        elderCards = {
            id: EmployeeCard(name, True, []) for name, id in self.elderNames.items()
        }
        for day in self._week:
            for id, elderSch in schedule.getElders().items():
                if day in elderSch:
                    elderCards[id].Shifts.append(
                        (day, 1.0, self.__getNameForTruck(self._elderNamesById[id], weekTrucks, day))
                    )
                    break

        for _, card in elderCards.items():
            employeeCards.append(card)

        for ghostName, ghostId in self.ghostNames.items():
            employeeCard = EmployeeCard(ghostName, False, [])

            for i, day in zip(schedule.ghostOneTime, self._oneTimeWeek):
                if i == ghostId:
                    employeeCard.Shifts.append(
                        (day, 1.0, self.__getNameForTruck(ghostName, weekTrucks, day))
                    )

            for p, day in zip(schedule.ghostPair, self._pairWeek):
                if ghostId == p[0] or ghostId == p[1]:
                    employeeCard.Shifts.append(
                        (day,
                         self._shiftDayLen[day - 1] if ghostId == p[0] else 1.0,
                         self.__getNameForTruck(ghostName, weekTrucks, day))
                    )

            employeeCards.append(employeeCard)

        return WeekScheduleExcelType(employeeCards=employeeCards, trucks=trucks)

    def pairDayStart(self):
        return self._pairWeek[0]

    def fromExcel(self, excelSchedule: WeekScheduleExcelType) -> Schedule:
        schedule = Schedule(self.pairDayStart())

        # todo: добавить старших

        for card in excelSchedule.EmployeeCards:
            # skip eldermen
            if card.IsElder:
                if card.Name not in self.elderNames.keys():
                    continue
                elderCard = card
                elderId = self.elderNames[elderCard.Name]
                for day, shiftLen, _ in elderCard.Shifts:
                        schedule.getElders()[elderId].append(day)
            else:
                if card.Name not in self.ghostNames.keys():
                    continue
                ghostCard = card
                ghostId = self.ghostNames[ghostCard.Name]
                for day, shiftLen, _ in ghostCard.Shifts:
                    if day < self.pairDayStart():
                        schedule.ghostOneTime[day - 1] = ghostId
                    elif day <= 7:
                        pair = schedule.ghostPair[day - self.pairDayStart()]
                        if shiftLen < 1.0 - 1e5 or pair[0] == 0:
                            pair = (ghostId, pair[1])
                        elif pair[1] == 0:
                            pair = (pair[0], ghostId)
                        else:
                            print("Error: schedule is not parsed correctly. "
                                  "Repeated schedule in one day is found.")

                        schedule.ghostPair[day - self.pairDayStart()] = pair

        if not schedule.isValid():
            print("Error: schedule is not parsed correctly.")

        return schedule

    def loadUndesirables(self, undesirables: dict[str, list[int]]):
        self.undesirableElderDays = {
            id: [] for _, id in self.elderNames.items()
        }
        self.undesirableGhostDays = {
            id: [] for _, id in self.ghostNames.items()
        }
        for name, days in undesirables.items():
            if name in self.elderNames.keys():
                self.undesirableElderDays[self.elderNames[name]] = days
            else:
                self.undesirableGhostDays[self.ghostNames[name]] = days

    def __printWeek(self):
        print(''.rjust(self._nameStringMaxLen + 1), end='')
        for day in ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']:
            print(day.rjust(4), end='')
        print()

    def printElder(self, schedule: Schedule):
        self.__printWeek()
        elders = schedule.getElders()

        for name, elderId in self.elderNames.items():
            print(name.rjust(self._nameStringMaxLen) + ':', end='')
            for day in self._week:
                turnName = 'C'
                if day in self.undesirableElderDays[elderId]:
                    turnName = 'S'
                print((turnName if day in elders[elderId] else 'x').rjust(4), end='')
            print()

    def print(self, schedule: Schedule):
        week = range(1, 7 + 1)

        self.printElder(schedule)

        print('---------------------'.rjust((self._nameStringMaxLen + 1 + 7 * 4) // 5 * 4))

        oneTimeWeek = [i for i in range(1, self.pairDayStart())]
        for [name, ghostId] in self.ghostNames.items():
            print(name.rjust(self._nameStringMaxLen) + ':', end='')
            for day in week:
                if day in oneTimeWeek:
                    turnName = 'C'
                    if day in self.undesirableGhostDays[ghostId]:
                        turnName = 'S'
                    print((turnName if schedule.ghostOneTime[day - 1] == ghostId else 'x')
                          .rjust(4), end='')
                else:
                    turnName = 'C2' if (day in self.partTimeDays
                                        and ghostId == schedule.ghostPair[day - 4][0]) else 'C'
                    if day in self.undesirableGhostDays[ghostId]:
                        turnName = turnName.replace('C', 'S')
                    print((turnName if ghostId in schedule.ghostPair[day - 4] else 'x')
                          .rjust(4), end='')
            print()
