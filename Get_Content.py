import xlrd
import xlwt
import os

file = 'F:\Programming\Excel_operation\Input'

def Read_excel(number, people_name_row, people_name_col):
    # 挨个读入excel文件
    name_value = {}  # 名字对应内容，内容以列表形式存储
    nameList = []  # 读取对应的全部名字
    print('读取文件开始。。。。。。')
    os.chdir(file)
    for i in range(number):
        name = str(i + 1)
        wordbook = xlrd.open_workbook(name + '.xlsx')
        wordsheet = wordbook.sheet_by_index(0)
        for j in range(wordsheet.nrows - people_name_row):
            # 从开始行一直读取到最后一行
            list = []
            try:
                people_name = wordsheet.cell(people_name_row + j, 0).value  # 得到名字
                if not people_name in name_value:
                    for l in range(people_name_col + 1, wordsheet.ncols):
                        list.append(wordsheet.cell(people_name_row + j, l).value)
                        name_value[people_name] = list
                        nameList.append(people_name)
                else:
                    print('Key:', people_name, ' 已经被读取 取消读取')
            except:
                print(i)
    print('读取文件结束。')
    return name_value, nameList
