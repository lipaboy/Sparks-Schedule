def nextScheduleOfElderman(schedule: list,
                           turnCount: int,  # количество смен у старшего по смене
                           daysCount: int  # количество дней, доступное для составления расписания
                           ) -> bool:
    k = turnCount
    for i in range(k - 1, -1, -1):
        if schedule[i] < daysCount - k + i + 1:
            schedule[i] += 1
            for j in range(i + 1, k):
                schedule[j] = schedule[j - 1] + 1
            return True
    return False


def nextOneTimeScheduleOfGhostman(schedule: list,
                                  ghostCount: int,  # количество духов
                                  daysCount: int  # количество дней для составления односменного расписания
                                  ) -> bool:
    j = daysCount - 1
    while j >= 0 and schedule[j] == ghostCount:
        j -= 1
    if j < 0:
        return False
    if schedule[j] >= ghostCount:
        j -= 1
    schedule[j] += 1
    if j == daysCount - 1:
        return True
    for k in range(j + 1, daysCount):
        schedule[k] = 1
    return True


def nextPairSchedule(ghostPairSchedule: list,
                     ghostCount: int,
                     daysCount: int):
    j = daysCount - 1
    while j >= 0 and ghostPairSchedule[j] == (ghostCount - 1, ghostCount):
        j -= 1

    if j < 0:
        return False

    pair = ghostPairSchedule[j]
    ghostPairSchedule[j] = (pair[0] + 1, pair[0] + 2) \
        if ghostPairSchedule[j][1] == ghostCount else (pair[0], pair[1] + 1)

    for i in range(j + 1, daysCount):
        ghostPairSchedule[i] = (1, 2)

    return True


