from collections import OrderedDict


class RecentlyUpdatedCache(OrderedDict):
    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        # last=False will move the item to the start of the order
        self.move_to_end(key, last=False)


cache = RecentlyUpdatedCache()
cache[1.9] = 5
cache[2.4] = 6
cache[3.6] = 7

cache.popitem(min(cache.keys()))
print(min(cache.keys()))


# print the cache items in the order they were most recently updated
for item in cache.items():
    print(item)

# ('key2', 'value4'), ('key3', 'value3'), ('key1', 'value1')