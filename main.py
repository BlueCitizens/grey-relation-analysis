# -*- coding: utf-8 -*-
# @Time    : 2021/12/5 12:12
# @Author  : Xiangyu Dai
# @Email   : bluecitizens@163.com
# @File    : main.py
# @Software: PyCharm

import xlrd
import xlwt
import numpy as np
import tkinter as tk
from tkinter import filedialog
import pprint as pp
import prettytable as pt


def read_excel(path):
    wb = xlrd.open_workbook(path)
    sheet_name = wb.sheet_names()[0]
    print("Current sheet:" + sheet_name)
    sheet1 = wb.sheet_by_index(0)  # 默认用第一张表
    row_num = sheet1.nrows
    col_num = sheet1.ncols
    # s = sheet1.cell(row_num - 1, col_num - 1).value
    # cols2 = sheet1.col_values(2)
    data = []
    for i in range(1, col_num):
        # print(i)
        data.append(sheet1.col_values(i)[1:row_num])
    table = pt.PrettyTable()
    table.add_rows(data)
    print("------原始数据集 Original Dataset------")
    print(table)
    wb.release_resources()
    del wb
    return data


def normalize_initial(x):
    x = np.array(x)  # to Array
    # print(1/(np.tile(np.array(x[:, 0]), (x.shape[1], 1))).T)
    return np.multiply(x, (1 / (np.tile(np.array(x[:, 0]), (x.shape[1], 1)))).T)  # 归一化 normalized by initial value


def normalize_average(x):
    x = np.array(x)  # to Array
    # print(np.sum(x[:], axis=1) / x.shape[1])
    # print(np.tile(np.sum(x[:], axis=1) / x.shape[1], (x.shape[1], 1)).T)
    return np.multiply(x, (
            1 / (np.tile(np.sum(x[:], axis=1) / x.shape[1], (x.shape[1], 1)))).T)  # 归一化 normalized by initial value


def gra_all(x_p):
    # ck = x_p[0, :]  # 参考序列X_0 reference sequence
    # cp = x_p[1:, :]  # 比较序列X_1,2,3 comparative sequences
    # print("------参考序列X_0 reference sequence------")
    # table_ck = pt.PrettyTable()
    # table_ck.add_row(ck)
    # print(table_ck)
    # print("------比较序列X_1,2,3 comparative sequences------")
    # table_cp = pt.PrettyTable()
    # table_cp.add_rows(cp)
    # print(table_cp)
    count = x_p.shape[0]
    res = []
    for i in range(0, count):
        temp = x_p
        ck = temp[i, :]  # 参考序列X_0 reference sequence
        print("------参考序列X_0 reference sequence------")
        table_ck = pt.PrettyTable()
        table_ck.add_row(ck)
        print(table_ck)
        cp = np.delete(temp, i, 0)  # 比较序列X_1,2,3 comparative sequences
        print("------比较序列X_1,2,3 comparative sequences------")
        table_cp = pt.PrettyTable()
        table_cp.add_rows(cp)
        print(table_cp)
        res.append(gra(ck, cp))
    table_res = pt.PrettyTable()
    table_res.add_rows(res)
    print("------模型结果------")
    print(table_res)


def gra(ck, cp):
    t = abs(cp - np.tile(ck, (cp.shape[0], 1)))  # 差序列Delta_i difference sequences
    table_t = pt.PrettyTable()
    table_t.add_rows(t)
    print("------差序列Delta_i difference sequences------")
    print(table_t)
    max_diff = t.max().max()  # 两级最大差 maximum difference
    min_diff = t.min().min()  # 两级最小差 minimum difference
    print("------两极最大差 最小差------")
    print(max_diff, min_diff)
    gcc = ((min_diff + 0.5 * max_diff) / (t + 0.5 * max_diff))  # 关联系数 grey correlation coefficient
    table_gcc = pt.PrettyTable()
    table_gcc.add_rows(gcc)
    print("------关联系数 correlation coefficient------")
    print(table_gcc)
    grg = np.sum(gcc, axis=1) / gcc.shape[1]  # 相关度 grey relational grade
    table_grg = pt.PrettyTable()
    table_grg.add_row(grg)
    print("------关联度 relational grade------")
    print(table_grg)
    return grg


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # 打开选择文件夹对话框
    root = tk.Tk()
    root.withdraw()
    excel_path = filedialog.askopenfilename()  # 获得excel文件
    print("File chosen! " + excel_path)
    dataset = read_excel(excel_path)
    normalizing_mode = int(input('选择归一化方式 Choose normalizing mode\n1 - 初值化initial\n2 - 均值化average\nInput(1/2):'))
    x_t = []
    if normalizing_mode == 1:
        print("------initial------")
        x_t = (normalize_initial(dataset))
    elif normalizing_mode == 2:
        print("------average------")
        x_t = (normalize_average(dataset))
    else:
        print("Undefined input, using default by initial")
        x_t = (normalize_initial(dataset))
    table_x = pt.PrettyTable()
    table_x.add_rows(x_t)
    print(table_x)
    gra_all(x_t)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
