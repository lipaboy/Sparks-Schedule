from SparksScheduleSearch import SparksScheduleSearch
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
