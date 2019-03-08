
def Get_V(elem):
    return elem[1]


list = [[1,2], [1, 1]]

list.sort(key = Get_V)

print(list)

