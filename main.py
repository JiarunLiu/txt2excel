#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import xlwt
import argparse
import numpy as np

def txt2list(t):
    # list_ = t.split('\n')
    list_ = t

    lists = []
    keys = ['id', 'name', 'pos', 'stat', 'temp']
    for info in list_:
        info = info.split('. ')[1]
        infos = info.split('+')
        # try:
        info_dict = {'id': infos[0].strip(),
                     'name': infos[1].strip(),
                     'pos': infos[2].strip(),
                     'stat': infos[3].strip(),
                     'temp': infos[4].strip()}
        if "度" not in info_dict['temp']:
            info_dict['temp'] = (info_dict['temp'] + '度').strip()
        lists.append(info_dict)
        # except:
        #     print("Unsplitable text: {}\nPlease check your text format.".format(info))

    return lists

def sort(list):
    indexs = []
    for info in list:
        indexs.append(info['id'])
    sort_index = np.argsort(indexs)
    return np.array(list)[sort_index]

def make_excel(data):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)  # 创建一个Workbook对象，这就相当于创建了一个Excel文件
    sheet = book.add_sheet('test',
                           cell_overwrite_ok=True)  # # 其中的test是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False

    # 数据写入excel
    for row, val in enumerate(data):
        # print(val)
        sheet.write(row, 0, val['id'])  # 第二行开始
        sheet.write(row , 1, val['name'])  # 第二行开始
        sheet.write(row , 2, val['pos'])  # 第二行开始
        sheet.write(row, 3, val['stat'])  # 第二行开始
        sheet.write(row, 4, val['temp'])  # 第二行开始
        row = row + 1

    return book

def txt2excel(txt, exceldir):
    data = txt2list(txt)

    book = make_excel(sort(data))

    # 最后，将以上操作保存到指定的Excel文件中
    book.save(r'{}'.format(exceldir))  # 在字符串前加r，声明为raw字符串，这样就不会处理其中的转义了。否则，可能会报错


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='txt2excel')
    parser.add_argument('-i', '--input', dest='input', default='./src', type=str,
                        help='输入文本文件路径, 默认 ./text.txt')
    parser.add_argument('-o', '--output', dest='output', default='./result.xls', type=str,
                        help='输出文本文件路径, 默认 ./result.xls')
    args = parser.parse_args()

    assert os.path.isfile(args.input)
    with open(args.input) as f:
        txt = f.readlines()

    txt2excel(txt, args.output)
