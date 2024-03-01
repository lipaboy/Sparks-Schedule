# 7*53?3*1

def cmp(pair):
    return pair[0]

pair = (2, 'kek')
lst = []
lst.append((3, 'lol'))
lst.append(pair)

print(lst)

lst.sort(key=cmp)
# (3, 'lol') <= (2, 'kek')
# key((3, 'lol')) <= key((2, 'kek'))
# lst = sorted(lst, key=cmp)

# print(lst)