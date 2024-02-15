def gen(n):
    for i in range(n):
        yield i*i


if __name__ == '__main__':
    g = gen(5)
    # for i in g:
    #     print(i)

    # lst = list(g)
    # print(lst)

    print(g.__next__())
    print(g.__next__())
    print(g.__next__())
    print(g.__next__())