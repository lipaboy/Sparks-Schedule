import random
import copy

from Schedule import Schedule
from EmployeeFavor import EmployeeFavor, ScheduleExtractionExcelType


class SparksScheduleSearch:
    def __init__(self):
        self.schedule = Schedule()
        self.favor = EmployeeFavor()
        self.__setBase()

        self.turnRepeatCoef = 10
        self.differInTurnsCoef = 2.1
        self.undesirableDayCoef = 4

        self.week = range(1, 7+1)
        self.oneTimeWeek = range(1, len(self.schedule.ghostOneTime) + 1)
        self.pairWeek = range(len(self.schedule.ghostOneTime) + 1, 7+1)
        self.turnDayLen = [self.favor.partTimeDays.get(i)
                           if i in self.favor.partTimeDays else 1.0 for i in self.week]

        self.debug = False

    def calcDebatov(self):
        currDebatov = 0.0

        """ Turn Count By Ghost """
        turnCountDict = [0.0 for _ in self.favor.ghostNames]

        """ The Longest Turn Repeats """
        turnRepeats = 0.0
        ghostIdRepeat = -1
        maxTurnRepeat = 0.0

        """ Undesirable Days """
        undesirableDaysCount = 0

        # пн, вт, ср
        for day, ghostId in zip(self.oneTimeWeek, self.schedule.ghostOneTime):
            """ Turn Count By Ghost """
            turnCountDict[ghostId - 1] += 1

            """ The Longest Turn Repeats """
            if ghostIdRepeat == ghostId:
                turnRepeats = turnRepeats + 1
            else:
                maxTurnRepeat = max(maxTurnRepeat, turnRepeats)
                turnRepeats = 1
            ghostIdRepeat = ghostId

            """ Undesirable Days """
            if day in self.favor.undesirableGhostDays[ghostId]:
                undesirableDaysCount += 1

        """ The Longest Turn Repeats """
        secondTurnRepeats = 0.0
        secondGhostIdRepeat = -1

        # чт*, пт*, сб, вс
        for day, p in zip(self.pairWeek, self.schedule.ghostPair):
            currDayLen = self.turnDayLen[day - 1]
            """ Turn Count By Ghost """
            turnCountDict[p[0] - 1] += currDayLen
            turnCountDict[p[1] - 1] += 1

            """ The Longest Turn Repeats """
            # noinspection DuplicatedCode
            if secondGhostIdRepeat == p[0] or secondGhostIdRepeat == p[1]:
                secondTurnRepeats += currDayLen if secondGhostIdRepeat == p[0] else 1
            else:
                maxTurnRepeat = max(maxTurnRepeat, secondTurnRepeats)
                secondGhostIdRepeat, secondTurnRepeats = \
                    (p[0], currDayLen) if ghostIdRepeat != p[0] else (p[1], 1)

            # noinspection DuplicatedCode
            if ghostIdRepeat == p[0] or ghostIdRepeat == p[1]:
                turnRepeats += currDayLen if ghostIdRepeat == p[0] else 1
            else:
                maxTurnRepeat = max(maxTurnRepeat, turnRepeats)
                ghostIdRepeat, turnRepeats = \
                    (p[0], currDayLen) if secondGhostIdRepeat != p[0] else (p[1], 1)

            """ Undesirable Days """
            if day in self.favor.undesirableGhostDays[p[0]]:
                undesirableDaysCount += currDayLen
            if day in self.favor.undesirableGhostDays[p[1]]:
                undesirableDaysCount += 1

        """ Turn Count By Ghost """
        # смотрим разницу между желаемым количеством смен для каждого духа
        # и текущим рассматриваемым расписанием
        for ghostId, limit in self.favor.ghostLimits.items():
            currDebatov += self.differInTurnsCoef * abs(turnCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxTurnRepeat = max(maxTurnRepeat, turnRepeats, secondTurnRepeats)
        if maxTurnRepeat >= 3.0:
            currDebatov += self.turnRepeatCoef + (maxTurnRepeat - 3) * self.turnRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def findFirstGhost(self) -> list[ScheduleExtractionExcelType]:
        minDebatov = 1e5
        bestCount = 12
        schedulesBest = {minDebatov + i * 1.0: Schedule() for i in range(bestCount)}
        iterationId = 0
        for _ in self.__ghostTraversalGen():
            iterationId += 1

            currDebatov = self.calcDebatov()

            if currDebatov < minDebatov:
                schedulesBest.pop(max(schedulesBest.keys()))
                while currDebatov in schedulesBest:
                    currDebatov += 0.0001
                schedulesBest[currDebatov] = copy.deepcopy(self.schedule)
                minDebatov = float(max(schedulesBest.keys()))

        schedulesBest = dict(sorted(schedulesBest.items(), reverse=True))
        if self.debug:
            print(f"Количество итераций: {iterationId}")
            print()
            for debatov, s in schedulesBest.items():
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

        n = random.randint(1, int(1e4))
        for i in range(1, n):
            # self.schedule.nextGhostPair()
            self.schedule.nextPairSchedule_v2()

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
