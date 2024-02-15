import random
import copy

from Schedule import Schedule
from NextSchedule import *


class SparksScheduleSearch:
    def __init__(self):
        self.schedule = Schedule()
        self.__setBase()
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

        self.turnRepeatCoef = 10
        self.differInTurnsCoef = 2.1
        self.undesirableDayCoef = 4
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
        self.partTimeDays = {
            4: 0.5,
            5: 0.5
        }
        self.week = range(1, 7+1)
        self.oneTimeWeek = range(1, len(self.schedule.ghostOneTime) + 1)
        self.pairWeek = range(len(self.schedule.ghostOneTime) + 1, 7+1)
        self.turnDayLen = [self.partTimeDays.get(i) if i in self.partTimeDays else 1 for i in self.week]

        self.debug = False

    def calcDebatov(self):
        currDebatov = 0.0

        """ Turn Count By Ghost """
        turnCountDict = [0.0 for _ in self.ghostNames]

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
            if day in self.undesirableGhostDays[ghostId]:
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
            if day in self.undesirableGhostDays[p[0]]:
                undesirableDaysCount += currDayLen
            if day in self.undesirableGhostDays[p[1]]:
                undesirableDaysCount += 1

        """ Turn Count By Ghost """
        # смотрим разницу между желаемым количеством смен для каждого духа
        # и текущим рассматриваемым расписанием
        for ghostId, limit in self.ghostLimits.items():
            currDebatov += self.differInTurnsCoef * abs(turnCountDict[ghostId - 1] - limit)

        """ The Longest Turn Repeats """
        maxTurnRepeat = max(maxTurnRepeat, turnRepeats, secondTurnRepeats)
        if maxTurnRepeat >= 3.0:
            currDebatov += self.turnRepeatCoef + (maxTurnRepeat - 3) * self.turnRepeatCoef

        """ Undesirable Days """
        currDebatov += undesirableDaysCount * self.undesirableDayCoef

        return currDebatov

    def findFirstGhost(self):
        minDebatov = 1e5
        bestCount = 12
        scheduleBest = {minDebatov + i * 1.0: SparksScheduleSearch() for i in range(bestCount)}
        iterationId = 0
        for _ in self.__ghostTraversalGen():
            iterationId += 1

            currDebatov = self.calcDebatov()

            if currDebatov < minDebatov:
                scheduleBest.pop(max(scheduleBest.keys()))
                while currDebatov in scheduleBest:
                    currDebatov += 0.0001
                scheduleBest[currDebatov] = copy.deepcopy(self)
                minDebatov = float(max(scheduleBest.keys()))

        print(f"Количество итераций: {iterationId}")
        print()
        scheduleBest = dict(sorted(scheduleBest.items(), reverse=True))
        for debatov, s in scheduleBest.items():
            print(f"Минимум дебатов: {debatov}")
            s.print()
            print()

    def __ghostTraversalGen(self):
        self.__setBase()
        while True:
            self.schedule.ghostPair = [(1, 2), (1, 2), (1, 2), (1, 2)]
            while True:
                yield
                if not self.nextGhostPair():
                    break
            if not self.nextGhostOneTime():
                break

    def calcTraverseLen(self):
        self.__setBase()
        traversalLen = 0
        while self.__ghostTraversalGen():
            traversalLen += 1
        return traversalLen

    def nextGhostPair(self):
        return self.nextPairSchedule_v2(self.schedule.ghostPair, 5, 4)
        # return nextPairSchedule(self.schedule.ghostPair, 5, 4)

    def nextGhostOneTime(self):
        return nextOneTimeScheduleOfGhostman(self.schedule.ghostOneTime, 5, 3)

    def nextPairSchedule_v2(self,
                            ghostPairSchedule: list,
                            ghostCount: int,
                            daysCount: int):
        j = daysCount - 1
        dayLimit = max(self.partTimeDays.keys()) - 4
        while j >= 0 and ((j <= dayLimit and ghostPairSchedule[j] == (ghostCount, ghostCount - 1))
                          or (j > dayLimit and ghostPairSchedule[j] == (ghostCount - 1, ghostCount))):
            j -= 1

        if j < 0:
            return False

        pair = ghostPairSchedule[j]
        # проверка на четверг
        if j <= dayLimit:
            ghostPairSchedule[j] = (pair[0] + 1, 1) \
                if ghostPairSchedule[j][1] == ghostCount \
                else (pair[0], pair[1] + 1 if pair[0] != pair[1] + 1 else pair[0] + 1)
        else:
            ghostPairSchedule[j] = (pair[0] + 1, pair[0] + 2) \
                if ghostPairSchedule[j][1] == ghostCount else (pair[0], pair[1] + 1)

        for i in range(j + 1, daysCount):
            ghostPairSchedule[i] = (1, 2)

        return True

    def shuffle(self):
        self.__setBase()

        n = random.randint(1, 35)
        for i in range(1, n):
            nextScheduleOfElderman(self.schedule.vovan, 4, 7)

        n = random.randint(1, 125)
        for i in range(1, n):
            nextOneTimeScheduleOfGhostman(self.schedule.ghostOneTime, 5, 3)

        n = random.randint(1, int(1e4))
        for i in range(1, n):
            nextPairSchedule(self.schedule.ghostPair, 5, 4)

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

    def print(self):
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
                    print((turnName if self.schedule.ghostOneTime[i - 1] == ghostId else 'x')
                          .rjust(4), end='')
                else:
                    turnName = 'C2' if (i in self.partTimeDays
                                        and ghostId == self.schedule.ghostPair[i - 4][0]) else 'С'
                    print((turnName if ghostId in self.schedule.ghostPair[i - 4] else 'x')
                          .rjust(4), end='')
            print()
