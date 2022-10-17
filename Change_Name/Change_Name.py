#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
统计由Inventor导出的二维图纸量
"""

# =================
# imports
# =================
import os
import time

import win32com.client

import AutoPath_TypeConvert as Tc


# =========================================================
def getfilefame(path):
    dwg_list = []
    for root, dirs, files in os.walk(path):
        for i in files:
            if os.path.splitext(i)[1] == '.dwg':
                j = os.path.join(root, i)
                dwg_list.append(j)
    return dwg_list


def change_name(f_list, wincad):
    list_amount = 0
    drawing_amount = 0
    drawing_amount0_list = []
    for i in f_list:
        list_amount += 1
        print("第%d张，共%d张" % (list_amount, len(f_list)))
        name, number = None, None
        time.sleep(1)
        wincad.Documents.Open(i)
        doc = wincad.ActiveDocument
        time.sleep(0.5)
        # try:
        #     doc.SelectionSets.Item("SS1").Delete()
        # except:
        #     pass
        # slt = doc.SelectionSets.Add("SS1")
        # slt.Select(5)  # 全选
        # obj = slt[0]
        # obj.move(Tc.vtpnt(0, 0), Tc.vtpnt(0, 0))
        # slt.Delete()

        try:
            doc.SelectionSets.Item("SS2").Delete()
        except:
            pass
        slt = doc.SelectionSets.Add("SS2")
        filterType = [0]  # 定义过滤类型
        filterData = ["INSERT"]  # 设置过滤参数
        filterType = Tc.vtint(filterType)
        filterData = Tc.vtvariant(filterData)
        slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
        time.sleep(0.5)
        for entity in slt:
            if entity.Name == '标题栏 标题栏':
                attributes = entity.GetAttributes()
                for attri in attributes:
                    if attri.TagString == '零件代号':
                        if attri.TextString == '':
                            break
                        elif attri.TextString[0] == '\\':
                            name = attri.TextString.split(';', 1)[1]
                        else:
                            name = attri.TextString
                    if attri.TagString == '库存编号':
                        if attri.TextString == '':
                            break
                        elif attri.TextString[0] == '\\':
                            number = attri.TextString.split(';', 1)[1]
                        else:
                            number = attri.TextString
                break
        slt.Delete()
        doc.Close()
        if name and number:
            file_name_old = i
            file_name_new = os.path.split(i)[0] + '\\' + number + ' ' + name + '.dwg'
            os.rename(file_name_old, file_name_new)
        else:
            drawing_amount0_list.append(i)
    print("\n以下%d张图纸未重命名，请手动进行：" % len(drawing_amount0_list))
    for each in drawing_amount0_list:
        print(each, end=',\n')


if __name__ == '__main__':
    path = input("请输入要重命名的文件夹路径（默认包含所有子文件夹）：")
    # path = 'r"' + path + '"'
    f_list = getfilefame(path)
    print("共找到%d个dwg文件" % len(f_list))

    wincad = win32com.client.Dispatch("AutoCAD.Application")
    change_name(f_list, wincad)
    input()
