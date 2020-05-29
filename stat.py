#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@author mybsdc <mybsdc@gmail.com>
@date 2020/5/18
@time 12:17
"""

import sys
import time
import re
import string
from openpyxl import load_workbook
import xlwt
import xlrd
from xlutils.copy import copy as xlutils_copy


def letter2num(letter: str) -> int:
    """
    字母 a-z 映射到数字，从 0 开始
    """
    letter_lower = letter.lower()

    return string.ascii_lowercase.index(letter_lower)


origin_ws = load_workbook(
    r'C:\Users\luolongf\Desktop\学生名册修复版（所有文字可复制）.xlsx')['Sheet1']
all_origin_cells = origin_ws['A3':'K44']

target_file_name = r'C:\Users\luolongf\Desktop\一年级招生名册.xls'
workbook = xlrd.open_workbook(target_file_name, formatting_info=True)
target_wb = xlutils_copy(workbook)
target_ws = target_wb.get_sheet('中心校')
# tmp_ws = target_wb.get_sheet('Sheet1')

print('已读取所需表格信息，开始处理...')

birthday_regex = re.compile(
    r'^20[012][0-9](?:0[1-9]|1[0-2])(?:0[1-9]|[12][0-9]|3[0-2])$')
row_num = 3
error_ID = {}
ok_num = 0
for row in all_origin_cells:
    name = row[3].value.strip()
    sex = row[4].value.strip()
    ID = row[5].value

    birthday = ID[6:14]
    if birthday_regex.search(birthday) is None:
        print('{} 同学的生日异常 提取为：{} 身份证号：{} 请检查身份证号是否有误！'.format(name, birthday, ID))
        error_ID[ID] = name
        continue

    y = birthday[:4]
    m = birthday[4:6]
    d = birthday[-2:]

    # 2014-08-31 前出生的可上小学一年级
    if int(birthday) > 20140831:
        print('{} 同学太小了，还不能上一年级 （{}年{}月{}日出生）'.format(name, y, m, d))
        continue

    parent_name = row[6].value.strip()
    parent_ID = row[8].value

    father_name = ''
    father_ID = ''
    mother_name = ''
    mother_ID = ''

    relation = row[7].value.strip()
    if '父' in relation:
        father_name = parent_name
        father_ID = parent_ID
    elif '母' in relation:
        mother_name = parent_name
        mother_ID = parent_ID
    else:
        print('{} 同学与监护人关系为 {} 无法得知父母信息'.format(name, relation))
        pass

    addr = row[9].value.strip()
    tel = str(row[10].value)

    try:
        print('正在写入 {} 的信息到一年级报名册'.format(name))

        ok_num += 1

        target_ws.write(row_num, letter2num('A'), ok_num)
        target_ws.write(row_num, letter2num('B'), name)
        target_ws.write(row_num, letter2num('C'), sex)
        target_ws.write(row_num, letter2num('D'), birthday)
        target_ws.write(row_num, letter2num('E'), ID)
        target_ws.write(row_num, letter2num('F'), father_name)
        target_ws.write(row_num, letter2num('H'), father_ID)
        target_ws.write(row_num, letter2num('I'), mother_name)
        target_ws.write(row_num, letter2num('K'), mother_ID)
        target_ws.write(row_num, letter2num('L'), addr)
        target_ws.write(row_num, letter2num('M'), tel)

        row_num += 1
        time.sleep(0.2)
    except Exception as e:
        print(f'{name} 写入数据出错：{str(e)}')

target_wb.save(target_file_name)

print(f'大班共 {len(all_origin_cells)} 人，可就读一年级的共 {ok_num} 人，另外有 {len(error_ID)} 位同学身份证号有误，无法判断是否满足升学条件，他们分别是：')
for ID, name in error_ID.items():
    print(f'{name}：{ID}')

print('恭喜，所有处理已完成')
