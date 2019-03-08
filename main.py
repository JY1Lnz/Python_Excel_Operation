import Init
import Get_Content  # 全部信息读取
import Write_Content  # 全部信息写入
import Get_Name_Phone  # 电话号码读取录入

# 名字对应的输入输出文件位置
people_name_row_input = 8  # 文件中第一个名字对应的行
people_name_col_input = 1  # 文件中第一个名字对应的列

people_name_row_output = 7
people_name_col_output = 0

SRow = 8  # 写入范围的开始行
ERow = 40  # 写入范围的结束行
NCol = 7  # 写入范围的列长度

file_number = Init.Init()#文件名初始化

#name_phone = {}
#name_phone = Get_Name_Phone.Read_excel(file_number)
#Get_Name_Phone.Write_excel(name_phone)

name_value = {}  # 名字对应内容
nameList = []  # 名字集合

name_value, nameList = Get_Content.Read_excel(file_number, people_name_row_input - 1, people_name_col_input - 1)

Write_Content.Write_Content(name_value, nameList, SRow - 1, ERow - 1, NCol)



