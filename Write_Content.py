import xlrd
import xlwt
import os
from xlutils.copy import copy

file = 'F:\Programming\Excel_operation\Output'


def Write_Content_Key(name_value, SRow, ERow, NCol):
    os.chdir(file)
    print('写入文件开始。。。。。。')
    wordbook = xlrd.open_workbook('2019.3.22.xlsx', formatting_info=True)  # 打开原本文件
    ws = wordbook.sheet_by_index(0)
    wordbook = copy(wordbook)  # 重新复制文件
    wordsheet = wordbook.get_sheet(0)  # 得到工作表

    for i in range(SRow, ERow):
        if ws.cell(i, 0).value in name_value:
            for j in range(NCol):
                wordsheet.write(i, j + 1, name_value[ws.cell(i, 0).value][j])
        else:
            print('Key: ',ws.cell(i, 0).value, '字典中不存在 无法写入')
    print('写入文件结束。')

    wordbook.save('one.xlsx')


def Write_Content_Name(name_value, NCol):
    """新建excel存储"""
    os.chdir(file)
    print('写入文件开始。。。。。。')
    wordbook = xlwt.Workbook(encoding= 'ascii')
    wordsheet = wordbook.add_sheet('Sheet1')

    line = 0
    for key in name_value:
        wordsheet.write(line, 0, key)
        print(line)
        for i in range(NCol):
            wordsheet.write(line, i + 1, name_value[key][i])
        line += 1
    print('写入文件结束。')

    wordbook.save('2019.3.22.xlsx')

