import copy
import itertools

from search.Schedule import Schedule
from search.EmployeeFavor import EmployeeFavor, ScheduleExtractionExcelType


class SparksScheduleSearch:
    """"""
    """ Главный метод для поиска оптимальных расписаний """
    def search(self, mode='part') -> list[ScheduleExtractionExcelType]:
        minDebatov = float(1e5)
        bestCount = 12
        schedulesBest = {minDebatov + i * 1.0: Schedule(self.favor.pairDayStart())
                         for i in range(bestCount)}
        iterationId = 0
        self._schedule.setMode(mode)
        for _ in self._schedule.ghostTraversalGen():
            iterationId += 1

            """ прогресс бар """
            if iterationId % int(1e5) == 0:
                print('.', end='')

            currDebatov = self.calcDebatov(self._schedule)

            if currDebatov < minDebatov:
                schedulesBest.pop(max(schedulesBest.keys()))
                while currDebatov in schedulesBest:
                    currDebatov += 0.0001
                schedulesBest[currDebatov] = copy.deepcopy(self._schedule)
                minDebatov = float(max(schedulesBest.keys()))

        print()

        schedulesBest = dict(sorted(schedulesBest.items()))
        if self.debug:
            schedulesForPrint = dict(sorted(schedulesBest.items(), reverse=True))
            print(f"Количество итераций: {iterationId}")
            print()
            i = len(schedulesForPrint) - 1
            for debatov, s in schedulesForPrint.items():
                print(f"id: {i}")
                i -= 1
                print(f"Минимум дебатов: {debatov}")
                self.favor.print(s)
                print()

        return [self.favor.toExcel(s) for _, s in schedulesBest.items()]

    """ Посчитать количество дебатов для конкретного расписания """
    def calcDebatov(self,
                    schedule: Schedule) -> float:
        currDebatov = 0.0

        """ Turn Count By Ghost """
        turnCountDict = [0.0 for _ in self.favor.ghostNames]

        """ The Longest Turn Repeats """
        # вещественный, потому что длина смены может быть не целой величиной
        shiftRepeats = self._prevWeekShiftRepeat.setdefault(schedule.ghostOneTime[0], 0.0)
        ghostIdRepeat = schedule.ghostOneTime[0] if shiftRepeats > 0.0 + 1e-5 else -1
        maxShiftRepeat = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        # пн, вт, ср
        for day, ghostId in zip(self._oneTimeWeek, schedule.ghostOneTime):
            """ Turn Count By Ghost """
            turnCountDict[ghostId - 1] += 1

            """ The Longest Turn Repeats """
            if ghostIdRepeat == ghostId:
                shiftRepeats = shiftRepeats + 1
            else:
                maxShiftRepeat = max(maxShiftRepeat, shiftRepeats)
                shiftRepeats = 1
            ghostIdRepeat = ghostId

            """ Undesirable Days """
            if day in self.favor.undesirableGhostDays[ghostId]:
                undesirableDaysCount += 1

        """ The Longest Turn Repeats """
        secondShiftRepeats = 0.0
        secondGhostIdRepeat = -1

        # чт*, пт*, сб, вс
        for day, p in zip(self._pairWeek, schedule.ghostPair):
            # currDayLen = self.shiftLenByDay[day - 1]
            """ Turn Count By Ghost """
            turnCountDict[p[0] - 1] += self.getShiftLenBy(day, True)
            turnCountDict[p[1] - 1] += self.getShiftLenBy(day, False)

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
            if day in self.favor.undesirableGhostDays[p[0]]:
                undesirableDaysCount += self.getShiftLenBy(day, True)
            if day in self.favor.undesirableGhostDays[p[1]]:
                undesirableDaysCount += self.getShiftLenBy(day, False)

        """ Turn Count By Ghost """
        # смотрим разницу между желаемым количеством смен для каждого духа
        # и текущим рассматриваемым расписанием
        for ghostId, limit in self.favor.ghostLimits.items():
            currDebatov += self.differInShiftsCoef * abs(turnCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxShiftRepeat = max(maxShiftRepeat, shiftRepeats, secondShiftRepeats)
        if maxShiftRepeat >= 3.0:
            currDebatov += self.shiftRepeatCoef + (maxShiftRepeat - 3) * self.shiftRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def loadPreviousWeekSchedule(self,
                                 excelSchedule: ScheduleExtractionExcelType):
        prevSchedule = self.favor.fromExcel(excelSchedule)
        if self.debug:
            self.favor.print(prevSchedule)

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

        self._prevWeekShiftRepeat[ghostIdRepeat] = shiftRepeats
        self._prevWeekShiftRepeat[secondGhostIdRepeat] = secondShiftRepeats

    def clear(self):
        self._prevWeekShiftRepeat = dict[int, float]()

    def __init__(self):
        self.favor = EmployeeFavor()
        self._schedule = Schedule(self.favor.pairDayStart())

        """ Коэффициент дебатов для непрерывный череды смен"""
        self.shiftRepeatCoef = 10
        """ Коэффициент дебатов для разницы между реальным количеством смен и желаемым для сотрудника"""
        self.differInShiftsCoef = 2.1
        """ Коэффициент дебатов, когда сотрудник выходит на смену в не желаемый день"""
        self.undesirableDayCoef = 4

        self._prevWeekShiftRepeat = dict[int, float]()

        self._week = range(1, 7 + 1)
        self._oneTimeWeek = range(1, len(self._schedule.ghostOneTime) + 1)
        self._pairWeek = range(len(self._schedule.ghostOneTime) + 1, 7 + 1)
        self._shiftLenByDay = [self.favor.partTimeDays.get(i)
                               if i in self.favor.partTimeDays else 1.0 for i in self._week]

        self.debug = False

    def getShiftLenBy(self,
                      day: int,
                      isFirstInShift: bool) -> float:
        return self._shiftLenByDay[day - 1] if isFirstInShift else 1.0
