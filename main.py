from SparksSchedule import SparksSchedule
import time

if __name__ == "__main__":
    sparks = SparksSchedule()
    # sparks.shuffle()
    # sparks.print()
    # print(sparks.calcTraverseLen())

    find = True
    # find = False

    _startMeasure = time.time()
    if find:
        sparks.findFirstGhost()
    else:
        count = 1
        while sparks.nextGhostPair():
            print(sparks.ghostPair)
            count += 1
        print(count)

    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
