from SparksSchedule import SparksSchedule
import time

if __name__ == "__main__":
    sparks = SparksSchedule()
    # sparks.shuffle()
    # sparks.print()

    _startMeasure = time.time()
    # print(sparks.calcTraverseLen())
    sparks.findFirstGhost()
    # count = 1
    # while sparks.nextGhostPair():
    #     print(sparks.ghostPair)
    #     count += 1
    # print(count)
    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
