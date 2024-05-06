from SparksScheduleSearch import SparksScheduleSearch, MODE_LIST
from Schedule import Schedule
import time

if __name__ == "__main__":
    sparks = SparksScheduleSearch()

    _startMeasure = time.time()

    # sparks.debug = True
    results = sparks.search(mode='fast')
    # for i in results:
    #     for j in i:
    #         print(j.Name, j.Shifts)
    #     print()

    sparks._favor.print(sparks._favor.fromExcel(results[10]))

    sparks.debug = True
    results = sparks.search(prevSchedule=results[10], mode='fast')
    # sparks.favor.print(sparks.favor.fromExcel(results[0]))

    print(f"Время работы: {time.time() - _startMeasure:.2f}s")

    s = Schedule(4)
    s.setMode(MODE_LIST[2])
    s.ghostCount = 6
    _startMeasure = time.time()
    kek = s.calcElderTraverseLen()
    print(kek)
    print(f"Время работы: {time.time() - _startMeasure:.2f}s")

    _startMeasure = time.time()
    print(s.calcGhostTraverseLen())
    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
