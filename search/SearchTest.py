from SparksScheduleSearch import SparksScheduleSearch
import time

if __name__ == "__main__":
    sparks = SparksScheduleSearch()

    _startMeasure = time.time()

    sparks.debug = True
    results = sparks.search(mode='part', undesirableDays={'Маша': [1, 2, 5, 6, 7]})
    # for i in results:
    #     for j in i:
    #         print(j.Name, j.Shifts)
    #     print()

    sparks._favor.print(sparks._favor.fromExcel(results[11]))

    # sparks.debug = True
    # results = sparks.search(prevSchedule=results[11])
    # sparks.favor.print(sparks.favor.fromExcel(results[0]))

    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
