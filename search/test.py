from SparksScheduleSearch import SparksScheduleSearch
import time

if __name__ == "__main__":
    sparks = SparksScheduleSearch()

    _startMeasure = time.time()

    sparks.debug = True
    results = sparks.search()

    # sparks.favor.print(sparks.favor.fromExcel(results[0]))
    # sparks.loadPreviousWeekSchedule(results[0])

    # sparks.debug = True
    # results = sparks.search()
    # sparks.favor.print(sparks.favor.fromExcel(results[0]))

    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
