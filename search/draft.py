with open('27.txt') as f:
    k = 166251
    n = 997506

    a = [int(x) for x in f]
    b = []
    for i in range(len(a)):
        b.append((a[i], i))

    b.sort(key=lambda x: x[0], reverse=True)
    for i in range(1000):
        print(b[i])