import glob
import os
import zipfile
from shutil import copy, rmtree
import pandas as pd
from flask_sqlalchemy import SQLAlchemy
from openpyxl import load_workbook
from openpyxl.styles import Side, Border

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
            rmtree(fileName)
            # try:
            #     #               delete empty folders
            #     os.rmdir(fileName)
            # except:
            #     #               Not empty, delete files under folders
            #     delfile(fileName)
            #     #               now, folders are empty, delete it
            #     os.rmdir(fileName)


def chengzhao_output_scores_tables(path):
    '''
    成招生成成绩单
    :param path: 模板所在目录
    :return:
    '''
    # todo  文件路径
    scores_path = f'{path}/scores.xlsx'
    templte_path = f'{path}/templte.xlsx'
    output_path = f'{path}/output/'
    creat_folder(output_path)

    # todo  提取数据
    df = pd.read_excel(io=scores_path, sheet_name=0, index_col='学号')
    # 数据结构 [['1601313030001', ['郭子铭', {'工程造价实训': '50', '数学': '30'}]], ['1601313030002', ...]]
    # 学号列表
    student_nums_ls = df.index.tolist()
    # 科目列表
    courses_ls = df.columns.tolist()[1:]
    # todo 获取学生数据
    data = []
    for i in student_nums_ls:
        ls = []
        student_name_ls = df.loc[i, ['姓名']].tolist()
        student_scores_ls = df.loc[i, '姓名':].tolist()[1:]
        scores_dict = dict(zip(courses_ls, student_scores_ls))
        student_name_ls.append(scores_dict)
        ls.append(str(i))
        ls.append(student_name_ls)
        data.append(ls)
    for i in range(len(data)):
        student_i = data[i]
        student_num = student_i[0]
        student_name = student_i[1][0]
        student_scores = student_i[1][1]
        path_output_ = '{}{}{}.xlsx'.format(output_path, student_num, student_name)
        copy(templte_path, path_output_)
        wb = load_workbook(path_output_)
        cols = wb['Sheet1'].max_column
        rows = wb['Sheet1'].max_row
        # 写入表格
        wb['Sheet1']['B6'] = student_num
        wb['Sheet1']['C6'] = '                      姓名：{}                        性别：男    '.format(student_name)
        # wb['Sheet1']['C{}'.format(rows)] = '                                            2019 年 3 月 1 日'
        for row in range(9, rows):
            key = wb['Sheet1'].cell(row=row, column=2).value
            if key in student_scores.keys():
                for col in range(5, cols + 1):
                    value = wb['Sheet1'].cell(row=row, column=col).value
                    if value == '*':
                        wb['Sheet1'].cell(row=row, column=col).value = student_scores[key]
                        # wb['Sheet1'].cell(row=row, column=col).border = border
        # 设置单元格样式
        thin = Side(border_style="thin", color='FF000000')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in range(7, rows + 1):
            for col in range(1, cols + 1):
                wb['Sheet1'].cell(row=row, column=col).border = border
        clean = Side(border_style=None)
        border_l = Border(left=clean, right=thin)
        border_r = Border(left=thin, right=clean)
        wb['Sheet1']['F{}'.format(rows - 1)].border = border_l
        wb['Sheet1']['J{}'.format(rows - 1)].border = border_l
        wb['Sheet1']['E{}'.format(rows - 1)].border = border_r
        wb['Sheet1']['I{}'.format(rows - 1)].border = border_r
        wb.save(path_output_)
        wb.close()


# 压缩文件
def to_zip(startdir):
    '''startdir:要压缩文件夹的绝对路径 '''
    file_news = startdir + '.zip'  # 压缩后文件夹的名字
    z = zipfile.ZipFile(file_news, 'w', zipfile.ZIP_DEFLATED)  # 参数一：文件夹名
    for dirpath, dirnames, filenames in os.walk(startdir):
        fpath = dirpath.replace(startdir, '')  # 这一句很重要，不replace的话，就从根目录开始复制
        fpath = fpath and fpath + os.sep or ''  # 这句话理解我也点郁闷，实现当前文件夹以及包含的所有文件的压缩
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), fpath + filename)
    print('压缩成功')
    z.close()
    return file_news


def part_print_do(path='考场情况表.xls', my_file='data.txt'):
    '''
    打印门贴签到表
    :param path:考场情况表路径
    :param my_file:生成文件路径
    :return:
    '''
    df = pd.read_excel(io=path, sheet_name='排考', dtype=str)
    exam_room_nums = df['考场号'].values.tolist()

    ls = []
    for exam_room_num in exam_room_nums:
        # first_num = str(exam_room_num)[0]
        print_num = int(str(exam_room_num)[1:])
        data = 'Set sh = Worksheets("sheet1"): sh.PrintOut {}, {}, , False'.format(print_num, print_num)
        ls.append(data)

    # todo 写入文件
    # 删除已有文件
    if os.path.exists(my_file):  # 如果文件存在
        # 删除文件，可使用以下两种方法。
        os.remove(my_file)  # 则删除
        # os.unlink(my_file)

    # 写入路径并输出
    def output_data(ls):
        with open(my_file, 'a+', encoding="utf-8") as f:
            # f.write('# -*- coding: utf-8 -*-\n')
            # f.write('# coding=gbk\n\n')
            f.write('Sub print1() \'定义打印过程' + '\n')
            f.write('Dim sh As Worksheet \'声明打印变量' + '\n')
            for i in ls:
                f.write(i + '\n')
            f.write('End Sub')

    output_data(ls)
