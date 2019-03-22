import Init
import Get_Content  # 全部信息读取
import Write_Content  # 全部信息写入
import Get_Name_Phone  # 电话号码读取录入

# 名字对应的输入输出文件位置
people_name_row_input = 11  # 文件中第一个名字对应的行
people_name_col_input = 1  # 文件中第一个名字对应的列

people_name_row_output = 7
people_name_col_output = 0

SRow = 11  # 写入范围的开始行
ERow = 40  # 写入范围的结束行
NCol = 5  # 写入范围的列长度

file_number = Init.Init()  # 文件名初始化

name_value = {}  # 名字对应内容
nameList = []  # 名字集合

name_value, nameList = Get_Content.Read_excel(file_number, people_name_row_input - 1, people_name_col_input - 1)
# Write_Content.Write_Content_Key(name_value, SRow - 1, ERow - 1, NCol)
Write_Content.Write_Content_Name(name_value, NCol)

for key in name_value:
    print(key)

