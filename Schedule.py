from NextSchedule import *


class Schedule:
    def __init__(self):
        self.vovan = []
        self.ghostOneTime = []
        self.ghostPair = []

    def nextPairSchedule_v2(self,
                            ghostCount: int,
                            daysCount: int,
                            dayLimit: int):
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

    def nextGhostPair(self):
        # return self.nextPairSchedule_v2(5, 4, max(self.favor.partTimeDays.keys()))
        return nextPairSchedule(self.ghostPair, 5, 4)

    def nextGhostOneTime(self):
        return nextOneTimeScheduleOfGhostman(self.ghostOneTime, 5, 3)
