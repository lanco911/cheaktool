import os

input_text = "123"
path = r'/Users/huangwenyong/PycharmProjects/cheaktool/cheaktool/test'


def open_dir(input_text, path):
    '''打开指定路径'path'下的指定文件夹input_text'''
    path_list = os.listdir(path)
    #大到小排序path_list
    path_list.sort(reverse=True)
    dir_path = []  # 存放送审文件夹路径
    for i in path_list:
        if input_text in i and os.path.isdir(path + '/' + i):  # 判断input_text是否存在及是否为文件夹
            dir_path = path + '/' + i
            break
    #判断dir_path是否为空
    if dir_path:
        os.system('open ' + dir_path)
        print('打开成功')
        return dir_path
    else:
        print('该流水号不存在')
        return dir_path

def delete_file(file_name, dir_path):
    '''删除指定文件夹dir_path下的指定文件file_name'''
    dir_path_list = os.listdir(dir_path)
    for i in file_name:
        if i in dir_path_list:
            os.remove(dir_path + '/' + i)
            print('删除成功')
    for j in dir_path_list:
        #判断j是否为文件夹
        if os.path.isdir(dir_path + '/' + j):
            delete_file(file_name, dir_path + '/' + j)







