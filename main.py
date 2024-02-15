from SparksScheduleSearch import SparksScheduleSearch
import time

if __name__ == "__main__":
    sparks = SparksScheduleSearch()
    # sparks.shuffle()
    # sparks.print()
    # print(sparks.calcTraverseLen())

    find = True
    # find = False

    _startMeasure = time.time()
    if find:
        sparks.findFirstGhost()

        # sparks.debug = True
        # print()
        # print("То к чему стремимся (наверное):")
        # sparks.ghostOneTime = [3, 2, 1]
        # sparks.ghostPair = [(4, 2), (4, 1), (4, 1), (2, 3)]
        # sparks.print()
        # print(f"debatov = {sparks.calcDebatov()}")
    else:
        count = 1
        while sparks.nextGhostPair():
            print(sparks.ghostPair)
            count += 1
        print(count)

    print(f"Время работы: {time.time() - _startMeasure:.2f}s")
