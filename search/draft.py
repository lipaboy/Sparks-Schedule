import itertools
alphabet = "РЕКА"
# ar = itertools.product(alphabet, repeat=4) #Размещение с повторением
ar = itertools.permutations(alphabet, 4)
l = sorted(list(ar))
for i in range(len(l)):
    print(i + 1, l[i])