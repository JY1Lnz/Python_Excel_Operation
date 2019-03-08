import os

operator_path = 'F:\Programming\Excel_operation\Input'

os.chdir(operator_path)

dir = os.listdir(operator_path) # 获取目录中全部文件名

def get(elem):
    a = elem.split('.')
    return int(a[0])

dir.sort(key=get)

number = len(dir)

def Init():

    # 将目录中的文件重命名1-n
    print('文件名初始化。。。。。。')
    cnt = 1

    for i in dir:
        name = str(cnt)
        os.rename(i, name + '.xlsx')
        cnt = cnt + 1
    print('文件名初始化完成')
    print('共', cnt - 1, '个文件')
    return cnt - 1
