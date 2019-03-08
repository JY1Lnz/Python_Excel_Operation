import os

operator_path = 'F:\Programming\Excel_operation\Input'

os.chdir(operator_path)

dir = os.listdir(operator_path) # 获取目录中全部文件名


def Init():

    # 将目录中的文件重命名1-n

    cnt = 1

    for i in dir:
        name = str(cnt)
        os.rename(i, name + '.xlsx')
        cnt = cnt + 1
    return cnt - 1

Init()