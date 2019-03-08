import xlrd
import xlwt
import os

file = 'F:\Programming\Excel_operation\Input'

def Read_excel(number, people_name_row, people_name_col):
    # 挨个读入excel文件
    name_value = {}  # 名字对应内容，内容以列表形式存储
    nameList = []  # 读取对应的全部名字
    os.chdir(file)
    for i in range(number):

        name = str(i + 1)
        wordbook = xlrd.open_workbook(name + '.xlsx')
        wordsheet = wordbook.sheet_by_index(0)

        list = []
        people_name = wordsheet.cell(people_name_row, people_name_col).value  # 得到名字
        nameList.append(people_name)
        for l in range(people_name_col + 1, wordsheet.ncols):
            list.append(wordsheet.cell(people_name_row, l).value)
        name_value[people_name] = list

    return name_value, nameList
