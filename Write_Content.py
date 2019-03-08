import xlrd
import xlwt
import os
from xlutils.copy import copy

file = 'F:\Programming\Excel_operation\Output'

def Write_Content(name_value, nameList, SRow, ERow, NCol):
    os.chdir(file)
    print('写入文件开始。。。。。。')
    wordbook = xlrd.open_workbook('one.xlsx', formatting_info=True)  # 打开原本文件
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



