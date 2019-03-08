import Init
import Get_Content  # 全部信息读取
import Write_Content  # 全部信息写入
import Get_Name_Phone  # 电话号码读取录入

# 名字对应的输入输出文件位置
people_name_row_input = 8
people_name_col_input = 0

people_name_row_output = 8
people_name_col_output = 0

file_number = Init.Init()#文件名初始化

#name_phone = {}

#name_phone = Get_Name_Phone.Read_excel(file_number)

#Get_Name_Phone.Write_excel(name_phone)

name_value = {}  # 名字对应内容
nameList = []  # 名字集合

name_value, nameList = Get_Content.Read_excel(file_number, people_name_row_input, people_name_col_input)

Write_Content.Write_Content(name_value, nameList, people_name_row_output, people_name_col_output)

for name in nameList:
    print(name)
    print(name_value[name])

print('读取成功')
