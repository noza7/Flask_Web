import glob
import os
import pandas as pd

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


# 创建目录
def creat_folder(folder_path):
    if not os.path.exists(folder_path):
        os.mkdir(folder_path)
        os.chmod(folder_path, os.O_RDWR)


# 删除文件
def delfile(path):
    #   read all the files under the folder
    fileNames = glob.glob(path + r'\*')
    for fileName in fileNames:
        try:
            #           delete file
            os.remove(fileName)
        except:
            try:
                #               delete empty folders
                os.rmdir(fileName)
            except:
                #               Not empty, delete files under folders
                delfile(fileName)
                #               now, folders are empty, delete it
                os.rmdir(fileName)


def multi_tables(path):
    '''
    先获取文件夹下文件名列表;
    用pandas拼接所有表格;
    '''
    # todo 获取表名
    path_ls = ['df1.xlsx', 'df2.xlsx', 'df3.xlsx']
    # todo 多表拼接
    # 表名列表
    df_ls = []
    for i in range(len(path_ls)):
        df_ls.append('df{}'.format(i))
    # 读取所有表
    for i in range(len(path_ls)):
        df_ls[i] = pd.read_excel(io=path + '/' + path_ls[i], sheet_name='Sheet1')
    # 表格拼接
    df = pd.concat(df_ls, ignore_index=True)
    df = df.astype(str)
    # todo ID补零
    df['ID'] = df['ID'].str.center(5, fillchar='0')
    # print(df['ID'])
    # 创建新表的路径文件名
    path_result = path + '/' + 'result.xlsx'
    # 写入数据
    writer = pd.ExcelWriter(path_result)
    df.to_excel(writer, '汇总', index=False)
    writer.save()
    print('程序顺利执行完成！')
