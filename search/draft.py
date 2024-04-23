with open('search/ege') as f:
    floor1, floor2 = [], []
    for j in range(12000):
        s = f.readline()
        a = s.split(' ')
        if int(a[0]) == 1:
            floor1.append((int(a[1]), int(a[2])))
        else:
            floor2.append((int(a[1]),int(a[2])))

floor1 = set(floor1)
floor1 = list(floor1)
floor2 = set(floor2)
floor2 = list(floor2)
floor1 = sorted(floor1)
floor2 = sorted(floor2)
place,needful = [floor1[0][0]],[]

for i in range(len(floor1) - 1):
    if floor1[i][0] != floor1[i+1][0]:
        if len(place) == 2 and (1 in place) and (6 in place):
            needful.append(floor1[i][0])
        place = []
    place.append(floor1[i+1][1])
if len(place) == 2 and (1 in place) and (6 in place):
    needful.append(floor1[-1][0])
place = [floor2[0][0]]
for i in range(len(floor2) - 1):
    if floor2[i][0] != floor2[i+1][0]:
        if len(place) == 2 and (1 in place) and (6 in place):
            needful.append(floor2[i][0])
        place = []
    place.append(floor2[i+1][1])
if len(place) == 2 and (1 in place) and (6 in place):
    needful.append(floor2[-1][0])
print(max(needful),len(needful))


