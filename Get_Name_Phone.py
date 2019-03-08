import xlrd
import xlwt
import os

"""
python读取excel中单元格的内容返回的有5种类型，即上面例子中的ctype：
ctype : 0 empty，1 string，2 number， 3 date，4 boolean，5 error
"""

file = 'F:\Programming\Excel_operation\Input'

name_phone = {}



def Switch_Number(number):
    '''将得到的文本手机号转换'''
    number = str(number)
    number = number.split('.')
    number = number[0]
    return number


def Read_excel(number):
# 挨个读入excel文件
    os.chdir(file)
    for i in range(number):

        name = str(i + 1)
        wordbook = xlrd.open_workbook(name + '.xlsx')
        wordsheet = wordbook.sheet_by_index(0)

        value = str(wordsheet.cell(2, 4).value)
        phone = value.split('.')
        name_phone[wordsheet.cell(2, 1).value] = phone[0]
        '''
        for line in range(wordsheet.nrows):
            value = str(wordsheet.cell(line, 3).value)
            phone = value.split('.')
            name_phone[wordsheet.cell(line, 0).value] = phone[0]
        '''
    return name_phone

need_file = 'F:\Programming\Excel_operation\Output'

def Write_excel(name_phone):
    wordbook = xlwt.Workbook()
    wordsheet = wordbook.add_sheet('Sheet1')

    os.chdir(need_file)

    need_book = xlrd.open_workbook('need.xlsx')
    need_sheet = need_book.sheet_by_index(0)

    for i in range(need_sheet.nrows):
        if (i > 1):
            wordsheet.write(i, 1, need_sheet.cell(i, 1).value)
            if (need_sheet.cell(i, 1).value in name_phone):
                wordsheet.write(i, 4, name_phone[need_sheet.cell(i, 1).value])

    wordbook.save('res.xlsx')
