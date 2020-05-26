#! /usr/bin/python3
# -*- coding: utf-8 -*-
import os


# 1. 从excel中读取数据，将身份证号码变成对应的考试号
def get_message(excel_file):
    d = {}
    with open(excel_file, 'rb') as ef:
        # 假设身份证号码与考生号分别在E、B列，从第二行开始
        for r in range(2, ef.max_row + 1):
            id_num = ef.cell(row=r, column=5).value
            candidate_num = ef.cell(row=r, column=2).value
            d[id_num] = candidate_num
        return d


def modify_file_names(d):
    
        for filename in os.listdir():
            if filename.endswith('.jpg') or filename.endswith('jpeg') or filename.endswith('png'):
                new_name = d[filename.rsplit('.')[-2]] + '.' + filename.rsplit('.')[-1]
                os.rename(filename, new_name)


if __name__ == "__main__":
    excel_file = input("Enter the excel file name:")
    d = get_message(excel_file)
    modify_file_names(d)