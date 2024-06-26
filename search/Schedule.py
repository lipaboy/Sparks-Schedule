from search.NextSchedule import *


class Schedule:
    def __init__(self, pairDayStart: int):
        self.vovan = []
        self.ghostOneTime = [0 for _ in range(1, pairDayStart)]
        self.ghostPair = [(0, 0) for _ in range(pairDayStart, 7 + 1)]
        self.ghostPairAlgo = self.__nextGhostPair_vPart
        self.pairDayStart = pairDayStart
        self.ghostCount = 5

    " Сначала старшие, потом духи "
    def getWorkersAtDay(self, day: int) -> tuple[int, list[int]]:
        return (1 if day in self.vovan else 2,
                [self.ghostOneTime[day - 1]]
                if day < self.pairDayStart
                else list(self.ghostPair[day - self.pairDayStart])
                )

    def getElders(self) -> dict[int, list[int]]:
        return {
            1: self.vovan,
            2: self.calcLubaSchedule()
        }

    def setMode(self, mode: str):
        if mode == 'full':
            self.ghostPairAlgo = self.__nextGhostPair_vFull
        elif mode == 'part':
            self.ghostPairAlgo = self.__nextGhostPair_vPart
        elif mode == 'fast':
            self.ghostPairAlgo = self.__nextGhostPair_vFast

    def isValid(self):
        return (
            # todo: раскомментить когда придёт время
            # 0 not in self.vovan
            # and
                0 not in self.ghostOneTime
                and
                not any(0 in p for p in self.ghostPair))

    def ghostTraversalGen(self):
        self.__setBase()
        while True:
            self.ghostPair = [(1, 2), (1, 2), (1, 2), (1, 2)]
            while True:
                yield
                if not self.nextGhostPair():
                    break
            if not self.nextGhostOneTime():
                break

    def elderTraversalGen(self):
        self.__setBase()
        while True:
            yield
            if not self.nextElderman():
                break

    def calcLubaSchedule(self):
        luba = []
        for day in range(1, 8):
            if day not in self.vovan:
                luba.append(day)
        return luba

    def nextElderman(self):
        return nextScheduleOfElderman(self.vovan, 4, 7)

    def nextGhostPair(self):
        return self.ghostPairAlgo()

    def nextGhostOneTime(self):
        return nextOneTimeScheduleOfGhostman(self.ghostOneTime, self.ghostCount, 3)

    def __nextGhostPair_vFast(self):
        return nextPairSchedule_vFast(self.ghostPair, self.ghostCount, 4)

    def __nextGhostPair_vPart(self):
        return nextPairSchedule(self.ghostPair, self.ghostCount, 4)

    def __nextGhostPair_vFull(self):
        return self.__nextPairSchedule_v2(self.ghostCount, 4, 5)

    def __nextPairSchedule_v2(self,
                              ghostCount: int,
                              daysCount: int,
                              # Разграничение для дней с полусменами. Обычно они выпадают на пт, реже чт
                              dayLimit: int):
        """ Индекс дней здесь начинается с вск (daysCount - 1) и пока не дойдёт до 0"""
        j = daysCount - 1
        dayLimit -= 4
        while j >= 0 and ((j <= dayLimit and self.ghostPair[j] == (ghostCount, ghostCount - 1))
                          or (j > dayLimit and self.ghostPair[j] == (ghostCount - 1, ghostCount))):
            j -= 1

        if j < 0:
            return False

        pair = self.ghostPair[j]
        # проверка на четверг
        if j <= dayLimit:
            self.ghostPair[j] = (pair[0] + 1, 1) \
                if self.ghostPair[j][1] == ghostCount \
                else (pair[0], pair[1] + 1 if pair[0] != pair[1] + 1 else pair[0] + 1)
        else:
            self.ghostPair[j] = (pair[0] + 1, pair[0] + 2) \
                if self.ghostPair[j][1] == ghostCount else (pair[0], pair[1] + 1)

        for i in range(j + 1, daysCount):
            self.ghostPair[i] = (1, 2)

        return True

    def calcGhostTraverseLen(self):
        self.__setBase()
        traversalLen = 0
        for _ in self.ghostTraversalGen():
            traversalLen += 1
        return traversalLen

    def calcElderTraverseLen(self):
        self.__setBase()
        traversalLen = 0
        for _ in self.elderTraversalGen():
            traversalLen += 1
        return traversalLen

    def __setBase(self):
        # 35 вариантов
        self.vovan = [1, 2, 3, 4]
        # lubaSchedule === [5, 6, 7] соответственно

        # пн, вт, ср (125 variants)
        # хранит список духов по порядку дней когда у кого смена (Даша пн, Даша вт, Даша ср)
        self.ghostOneTime = [1, 1, 1]

        # чт, пт, сб, вс (10 000 variants - без С2, 20 000 - 1хС2, 40 000 - 2хС2 и т.д.)
        # хранит список пар духов по порядку дней у кого смена (Даша, Маша чт), (Даша, Маша пт)...
        # (1, 2) === (2, 1) - верно, когда нет дней с C2 или С18
        self.ghostPair = [(1, 2), (1, 2), (1, 2), (1, 2)]
