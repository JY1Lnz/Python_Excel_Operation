import xlrd
import xlwt
import os
from xlutils.copy import copy

file = 'F:\Programming\Excel_operation\Output'

def Write_Content(name_value, nameList, Row, Col):
    os.chdir(file)

    wordbook = xlrd.open_workbook('res.xlsx', formatting_info=True)  # 打开原本文件
    wordbook = copy(wordbook)  # 重新复制文件
    wordsheet = wordbook.get_sheet(0)  # 得到工作表

    for name in nameList:
        wordsheet.write(Row, Col, name)
        for i in range(len(name_value[name])):
            wordsheet.write(Row, Col + 1 + i, name_value[name][i])
        Row = Row + 1

    wordbook.save('res.xlsx')



