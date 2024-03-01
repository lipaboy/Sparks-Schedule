import copy
import itertools

from Schedule import Schedule
from EmployeeFavor import EmployeeFavor, ScheduleExtractionExcelType


class SparksScheduleSearch:
    """"""
    """ Главный метод для поиска оптимальных расписаний """
    def search(self) -> list[ScheduleExtractionExcelType]:
        minDebatov = 1e5
        bestCount = 12
        schedulesBest = {minDebatov + i * 1.0: Schedule(self._favor.pairDayStart())
                         for i in range(bestCount)}
        iterationId = 0
        for _ in self._schedule.ghostTraversalGen():
            iterationId += 1

            """ прогресс бар """
            if iterationId % int(1e5) == 0:
                print('.', end='')

            currDebatov = self.calcDebatov()

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
                self._favor.print(s)
                print()

        return [self._favor.toExcel(s) for _, s in schedulesBest.items()]

    def calcDebatov(self):
        currDebatov = 0.0

        """ Turn Count By Ghost """
        turnCountDict = [0.0 for _ in self._favor.ghostNames]

        """ The Longest Turn Repeats """
        shiftRepeats = 0.0       # вещественный, потому что длина смены может быть не целой величиной
        ghostIdRepeat = -1
        maxShiftRepeat = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        # пн, вт, ср
        for day, ghostId in zip(self._oneTimeWeek, self._schedule.ghostOneTime):
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
            if day in self._favor.undesirableGhostDays[ghostId]:
                undesirableDaysCount += 1

        """ The Longest Turn Repeats """
        secondShiftRepeats = 0.0
        secondGhostIdRepeat = -1

        # чт*, пт*, сб, вс
        for day, p in zip(self._pairWeek, self._schedule.ghostPair):
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
            if day in self._favor.undesirableGhostDays[p[0]]:
                undesirableDaysCount += self.getShiftLenBy(day, True)
            if day in self._favor.undesirableGhostDays[p[1]]:
                undesirableDaysCount += self.getShiftLenBy(day, False)

        """ Turn Count By Ghost """
        # смотрим разницу между желаемым количеством смен для каждого духа
        # и текущим рассматриваемым расписанием
        for ghostId, limit in self._favor.ghostLimits.items():
            currDebatov += self.differInTurnsCoef * abs(turnCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxShiftRepeat = max(maxShiftRepeat, shiftRepeats, secondShiftRepeats)
        if maxShiftRepeat >= 3.0:
            currDebatov += self.turnRepeatCoef + (maxShiftRepeat - 3) * self.turnRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def __init__(self):
        self._favor = EmployeeFavor()
        self._schedule = Schedule(self._favor.pairDayStart())

        self.turnRepeatCoef = 10
        self.differInTurnsCoef = 2.1
        self.undesirableDayCoef = 4

        self._week = range(1, 7 + 1)
        self._oneTimeWeek = range(1, len(self._schedule.ghostOneTime) + 1)
        self._pairWeek = range(len(self._schedule.ghostOneTime) + 1, 7 + 1)
        self._shiftLenByDay = [self._favor.partTimeDays.get(i)
                               if i in self._favor.partTimeDays else 1.0 for i in self._week]

        self.debug = False

    def loadPreviousWeekSchedule(self,
                                 excelSchedule: ScheduleExtractionExcelType):
        prevSchedule = self._favor.fromExcel(excelSchedule)
        if self.debug:
            self._favor.print(prevSchedule)

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

        d = dict()
        d[ghostIdRepeat] = shiftRepeats
        d[secondGhostIdRepeat] = secondShiftRepeats
        print(d)

    def getShiftLenBy(self,
                      day: int,
                      isFirstInShift: bool) -> float:
        return self._shiftLenByDay[day - 1] if isFirstInShift else 1.0
