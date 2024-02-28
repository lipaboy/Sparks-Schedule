import random
import copy
import itertools

from Schedule import Schedule
from EmployeeFavor import EmployeeFavor, ScheduleExtractionExcelType


class SparksScheduleSearch:
    def __init__(self):
        self.favor = EmployeeFavor()
        self.schedule = Schedule(self.favor.pairDayStart())
        self.__setBase()

        self.turnRepeatCoef = 10
        self.differInTurnsCoef = 2.1
        self.undesirableDayCoef = 4

        self.week = range(1, 7+1)
        self.oneTimeWeek = range(1, len(self.schedule.ghostOneTime) + 1)
        self.pairWeek = range(len(self.schedule.ghostOneTime) + 1, 7+1)
        self.shiftLenByDay = [self.favor.partTimeDays.get(i)
                           if i in self.favor.partTimeDays else 1.0 for i in self.week]

        self.debug = False

    def getShiftLenBy(self,
                      day: int,
                      isFirstInShift: bool) -> float:
        return self.shiftLenByDay[day - 1] if isFirstInShift else 1.0

    def loadPreviousWeekSchedule(self,
                                 excelSchedule: ScheduleExtractionExcelType):
        prevSchedule = self.favor.fromExcel(excelSchedule)
        if self.debug:
            self.favor.print(prevSchedule)

        """ The Longest Turn Repeats """
        shiftRepeats = self.shiftLenByDay[-1]
        ghostIdRepeat = prevSchedule.ghostPair[-1][0]
        secondShiftRepeats = 1.0
        secondGhostIdRepeat = prevSchedule.ghostPair[-1][1]
        readyFirst = False
        readySecond = False
        for day, p in zip(itertools.islice(reversed(self.pairWeek), 1, None),
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

    def calcDebatov(self):
        currDebatov = 0.0

        """ Turn Count By Ghost """
        turnCountDict = [0.0 for _ in self.favor.ghostNames]

        """ The Longest Turn Repeats """
        shiftRepeats = 0.0       # вещественный, потому что длина смены может быть не целой величиной
        ghostIdRepeat = -1
        maxShiftRepeat = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        # пн, вт, ср
        for day, ghostId in zip(self.oneTimeWeek, self.schedule.ghostOneTime):
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
        for day, p in zip(self.pairWeek, self.schedule.ghostPair):
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
            currDebatov += self.differInTurnsCoef * abs(turnCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxShiftRepeat = max(maxShiftRepeat, shiftRepeats, secondShiftRepeats)
        if maxShiftRepeat >= 3.0:
            currDebatov += self.turnRepeatCoef + (maxShiftRepeat - 3) * self.turnRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def findFirstGhost(self) -> list[ScheduleExtractionExcelType]:
        minDebatov = 1e5
        bestCount = 12
        schedulesBest = {minDebatov + i * 1.0: Schedule(self.favor.pairDayStart())
                         for i in range(bestCount)}
        iterationId = 0
        for _ in self.__ghostTraversalGen():
            iterationId += 1

            # прогресс бар
            if iterationId % int(1e5) == 0:
                print('.', end='')

            currDebatov = self.calcDebatov()

            if currDebatov < minDebatov:
                schedulesBest.pop(max(schedulesBest.keys()))
                while currDebatov in schedulesBest:
                    currDebatov += 0.0001
                schedulesBest[currDebatov] = copy.deepcopy(self.schedule)
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

    def __ghostTraversalGen(self):
        self.__setBase()
        while True:
            self.schedule.ghostPair = [(1, 2), (1, 2), (1, 2), (1, 2)]
            while True:
                yield
                if not self.schedule.nextGhostPair():
                    break
            if not self.schedule.nextGhostOneTime():
                break

    def calcTraverseLen(self):
        self.__setBase()
        traversalLen = 0
        while self.__ghostTraversalGen():
            traversalLen += 1
        return traversalLen

    def shuffle(self):
        self.__setBase()

        n = random.randint(1, 35)
        for i in range(1, n):
            # self.schedule.nextScheduleOfElderman()
            pass

        n = random.randint(1, 125)
        for i in range(1, n):
            self.schedule.nextGhostOneTime()

        # n = random.randint(1, int(1e4))
        # for i in range(1, n):
            # self.schedule.nextGhostPair()
            # self.schedule.nextPairSchedule_v2()

    def __setBase(self):
        # 35 вариантов
        self.schedule.vovan = [1, 2, 3, 4]
        # lubaSchedule === [5, 6, 7] соответственно

        # пн, вт, ср (125 variants)
        # хранит список духов по порядку дней когда у кого смена (Даша пн, Даша вт, Даша ср)
        self.schedule.ghostOneTime = [1, 1, 1]

        # чт, пт, сб, вс (10 000 variants - без С2, 20 000 - 1хС2, 40 000 - 2хС2 и т.д.)
        # хранит список пар духов по порядку дней у кого смена (Даша, Маша чт), (Даша, Маша пт)...
        # (1, 2) === (2, 1) - верно, когда нет дней с C2 или С18
        self.schedule.ghostPair = [(1, 2), (1, 2), (1, 2), (1, 2)]
