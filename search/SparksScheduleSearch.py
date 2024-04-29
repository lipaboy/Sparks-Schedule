import copy
import itertools

from search.Schedule import Schedule
from search.EmployeeFavor import EmployeeFavor, WeekScheduleExcelType, TruckDistributionType


MODE_LIST = ['fast', 'part', 'full']
DEBATOV_LIMIT = float(1e7)


class SparksScheduleSearch:
    def calcNewTrucks(self, schedule: WeekScheduleExcelType) -> TruckDistributionType:
        newTrucks = schedule.Trucks
        for card in schedule.EmployeeCards:
            newTrucks[card.Name] += card.truckCount()
        return newTrucks

    def search(self,
               eldermen: list[str] = None, # по идее всегда не None
               ghostmen: list[str] = None,
               undesirableDays: dict[str, list[int]] = None,
               shiftCountPreferences: dict[str, float] = None,
               prevSchedule: WeekScheduleExcelType = None,
               coeficiencies: dict[str, float] = None,
               mode=MODE_LIST[1]) -> list[WeekScheduleExcelType]:

        if eldermen is not None and len(eldermen) > 0:
            self._favor.loadEldermen(eldermen)
        if ghostmen is not None and len(ghostmen) > 0:
            self._favor.loadGhostmen(ghostmen)
        if undesirableDays is not None:
            self._favor.loadUndesirables(undesirableDays)
        if shiftCountPreferences is not None:
            self._favor.loadShiftCountsPrefs(shiftCountPreferences)
        if prevSchedule is not None:
            self.__loadPreviousWeekSchedule(prevSchedule)
            if prevSchedule.Trucks is not None and len(prevSchedule.Trucks) > 0:
                self._favor.truckDistribution = prevSchedule.Trucks
        else:
            self.clear()
        if coeficiencies is not None:
            self.shiftRepeatCoef = coeficiencies['shiftRepeatCoef']
            self.differInShiftsCoef = coeficiencies['differInShiftsCoef']
            self.undesirableDayCoef = coeficiencies['undesirableDayCoef']

        minDebatov = DEBATOV_LIMIT
        ghostBestCount = 6
        ghostSchedulesBest = {minDebatov + i * 1.0: Schedule(self._favor.pairDayStart())
                              for i in range(ghostBestCount)}
        iterationId = 0
        self._schedule.setMode(mode)
        for _ in self._schedule.ghostTraversalGen():
            iterationId += 1

            """ прогресс бар """
            if iterationId % int(1e5) == 0:
                print('.', end='')

            currDebatov = self.calcGhostDebatov(self._schedule)

            if currDebatov < minDebatov:
                ghostSchedulesBest.pop(max(ghostSchedulesBest.keys()))
                while currDebatov in ghostSchedulesBest:
                    currDebatov += 0.0001
                ghostSchedulesBest[currDebatov] = copy.deepcopy(self._schedule)
                minDebatov = float(max(ghostSchedulesBest.keys()))

        print()

        """ Ищем расписание старших """

        minDebatov = DEBATOV_LIMIT
        elderBestCount = 2
        elderSchedulesBest = {minDebatov + i * 1.0: Schedule(self._favor.pairDayStart())
                              for i in range(elderBestCount)}
        for _ in self._schedule.elderTraversalGen():
            currDebatov = self.calcElderDebatov(self._schedule)

            if currDebatov < minDebatov:
                elderSchedulesBest.pop(max(elderSchedulesBest.keys()))
                while currDebatov in elderSchedulesBest:
                    currDebatov += 0.0001
                elderSchedulesBest[currDebatov] = copy.deepcopy(self._schedule)
                minDebatov = float(max(elderSchedulesBest.keys()))
            pass

        ghostSchedulesBest = dict(sorted(ghostSchedulesBest.items()))
        elderSchedulesBest = dict(sorted(elderSchedulesBest.items()))

        commonSchedules = []
        for g in ghostSchedulesBest.values():
            for e in elderSchedulesBest.values():
                s = copy.deepcopy(g)
                s.vovan = e.vovan
                commonSchedules.append(s)

        if self.debug:
            schedulesForPrint = reversed(commonSchedules)
            print(f"Количество итераций: {iterationId}")
            print()
            i = len(commonSchedules) - 1
            for s in schedulesForPrint:
                print(f"id: {i}")
                i -= 1
                self._favor.print(s)
                print()

        return [self._favor.toExcel(s) for s in commonSchedules]
    """"""

    """ Главный метод для поиска оптимальных расписаний """

    """ Посчитать количество дебатов для конкретного расписания старших """
    def calcElderDebatov(self,
                         schedule: Schedule) -> float:
        currDebatov = 0.0

        """ The Longest Turn Repeats """
        # вещественный, потому что длина смены может быть не целой величиной
        maxShiftRepeats = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        elders = schedule.getElders()

        for elderId, shifts in elders.items():
            prevDay = 0  # предыдущий день как будто есть, чтобы учитывать предыдущее расписание
            shiftRepeats = self._prevWeekElderShiftRepeat[1] \
                if elderId == self._prevWeekElderShiftRepeat[0] else 0.0

            for day in shifts:
                """ The Longest Shift Repeats """
                # если дни в расписании идут по порядку, значит есть непрерывная череда смен
                if day == prevDay + 1:
                    shiftRepeats += 1
                    maxShiftRepeats = max(shiftRepeats, maxShiftRepeats)
                else:
                    shiftRepeats = 1.0
                prevDay = day

                """ Undesirable Days """
                if day in self._favor.undesirableElderDays[elderId]:
                    undesirableDaysCount += 1

        """ The Longest Turn Repeats """
        if maxShiftRepeats >= 3.0:
            currDebatov += self.shiftRepeatCoef + (maxShiftRepeats - 3) * self.shiftRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    """ Посчитать количество дебатов для конкретного расписания духов """
    def calcGhostDebatov(self,
                         schedule: Schedule) -> float:
        currDebatov = 0.0

        """ Turn Count By Ghost """
        shiftCountDict = [0.0 for _ in self._favor.ghostNames]

        """ The Longest Turn Repeats """
        # вещественный, потому что длина смены может быть не целой величиной
        shiftRepeats = self._prevWeekGhostShiftRepeat.setdefault(schedule.ghostOneTime[0], 0.0)
        ghostIdRepeat = schedule.ghostOneTime[0] if shiftRepeats > 0.0 + 1e-5 else -1
        maxShiftRepeat = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        # пн, вт, ср
        for day, ghostId in zip(self._oneTimeWeek, schedule.ghostOneTime):
            """ Turn Count By Ghost """
            shiftCountDict[ghostId - 1] += 1

            """ The Longest Turn Repeats """
            if ghostIdRepeat == ghostId:
                shiftRepeats = shiftRepeats + 1
            else:
                maxShiftRepeat = max(maxShiftRepeat, shiftRepeats)
                shiftRepeats = 1
            ghostIdRepeat = ghostId

            """ Undesirable Days """
            if day in self._favor.undesirableGhostDays[ghostId]:
                undesirableDaysCount += 1

        """ The Longest Turn Repeats """
        secondShiftRepeats = 0.0
        secondGhostIdRepeat = -1

        # чт*, пт*, сб, вс
        for day, p in zip(self._pairWeek, schedule.ghostPair):
            # currDayLen = self.shiftLenByDay[day - 1]
            """ Turn Count By Ghost """
            shiftCountDict[p[0] - 1] += self.getShiftLenBy(day, True)
            shiftCountDict[p[1] - 1] += self.getShiftLenBy(day, False)

            """ The Longest Turn Repeats """
            # noinspection DuplicatedCode
            if secondGhostIdRepeat == p[0] or secondGhostIdRepeat == p[1]:
                secondShiftRepeats += self.getShiftLenBy(day, secondGhostIdRepeat == p[0])
            else:
                maxShiftRepeat = max(maxShiftRepeat, secondShiftRepeats)
                secondGhostIdRepeat = p[0] if ghostIdRepeat != p[0] else p[1]
                secondShiftRepeats = self.getShiftLenBy(day, ghostIdRepeat != p[0])

            # noinspection DuplicatedCode
            if ghostIdRepeat == p[0] or ghostIdRepeat == p[1]:
                shiftRepeats += self.getShiftLenBy(day, ghostIdRepeat == p[0])
            else:
                maxShiftRepeat = max(maxShiftRepeat, shiftRepeats)
                ghostIdRepeat = p[0] if secondGhostIdRepeat != p[0] else p[1]
                shiftRepeats = self.getShiftLenBy(day, secondGhostIdRepeat != p[0])

            """ Undesirable Days """
            if day in self._favor.undesirableGhostDays[p[0]]:
                undesirableDaysCount += self.getShiftLenBy(day, True)
            if day in self._favor.undesirableGhostDays[p[1]]:
                undesirableDaysCount += self.getShiftLenBy(day, False)

        """ Turn Count By Ghost """
        # смотрим разницу между желаемым количеством смен для каждого духа
        # и текущим рассматриваемым расписанием
        for ghostId, limit in self._favor.ghostLimits.items():
            currDebatov += self.differInShiftsCoef * abs(shiftCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxShiftRepeat = max(maxShiftRepeat, shiftRepeats, secondShiftRepeats)
        if maxShiftRepeat >= 3.0:
            currDebatov += self.shiftRepeatCoef + (maxShiftRepeat - 3) * self.shiftRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def __loadPreviousWeekSchedule(self,
                                   excelSchedule: WeekScheduleExcelType):
        prevSchedule = self._favor.fromExcel(excelSchedule)
        if self.debug:
            self._favor.print(prevSchedule)

        """ Старшие """
        elders = prevSchedule.getElders()
        elderShiftDayRepeats = 0
        elderId = 0
        for id, shifts in elders.items():
            if self._week[-1] in shifts:
                elderShiftDayRepeats = 1
                elderId = id
                break
        for day in list(reversed(self._week))[1:]:
            if day in elders[elderId]:
                elderShiftDayRepeats += 1
            else:
                break

        self._prevWeekElderShiftRepeat = (elderId, elderShiftDayRepeats)

        """ The Longest Turn Repeats """
        shiftRepeats = self._shiftLenByDay[-1]
        ghostIdRepeat = prevSchedule.ghostPair[-1][0]
        secondShiftRepeats = 1.0
        secondGhostIdRepeat = prevSchedule.ghostPair[-1][1]
        readyFirst = False
        readySecond = False

        for day, p in zip(itertools.islice(reversed(self._pairWeek), 1, None),
                          itertools.islice(reversed(prevSchedule.ghostPair), 1, None)):
            """ The Longest Turn Repeats """
            if (not readyFirst and
                    (ghostIdRepeat == p[0] or ghostIdRepeat == p[1])):
                shiftRepeats += self.getShiftLenBy(day, ghostIdRepeat == p[0])
            else:
                readyFirst = True

            if (not readySecond and
                    (secondGhostIdRepeat == p[0] or secondGhostIdRepeat == p[1])):
                secondShiftRepeats += self.getShiftLenBy(day, secondGhostIdRepeat == p[0])
            else:
                readySecond = True

            if readySecond and readyFirst:
                break

        self._prevWeekGhostShiftRepeat[ghostIdRepeat] = shiftRepeats
        self._prevWeekGhostShiftRepeat[secondGhostIdRepeat] = secondShiftRepeats

    def clear(self):
        self._prevWeekGhostShiftRepeat = dict[int, float]()

    def __init__(self):
        self._favor = EmployeeFavor()
        self._schedule = Schedule(self._favor.pairDayStart())

        """ Коэффициент дебатов для непрерывный череды смен"""
        self.shiftRepeatCoef = 10
        """ Коэффициент дебатов для разницы между реальным количеством смен и желаемым для сотрудника"""
        self.differInShiftsCoef = 2.1
        """ Коэффициент дебатов, когда сотрудник выходит на смену в не желаемый день"""
        self.undesirableDayCoef = 4

        self._prevWeekGhostShiftRepeat = dict[int, float]()
        self._prevWeekElderShiftRepeat = (-1, 0)

        self._week = range(1, 7 + 1)
        self._oneTimeWeek = range(1, len(self._schedule.ghostOneTime) + 1)
        self._pairWeek = range(len(self._schedule.ghostOneTime) + 1, 7 + 1)
        self._shiftLenByDay = [self._favor.partTimeDays.get(i)
                               if i in self._favor.partTimeDays else 1.0 for i in self._week]

        self.debug = False

    def getShiftLenBy(self,
                      day: int,
                      isFirstInShift: bool) -> float:
        return self._shiftLenByDay[day - 1] if isFirstInShift else 1.0
