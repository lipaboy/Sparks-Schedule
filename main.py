from NextSchedule import *
from SparksSchedule import SparksSchedule
import time

if __name__ == "__main__":
    sparks = SparksSchedule()
    # sparks.shuffle()
    # sparks.print()

    _startMeasure = time.time()
    # print(sparks.calcTraverseLen())
    sparks.findFirstGhost()
    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
