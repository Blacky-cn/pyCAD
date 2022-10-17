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


def cal_drawing(f_list, wincad):
    list_amount = 0
    drawing_amount = 0
    drawing_amount0_list = []
    for i in f_list:
        list_amount += 1
        print("第%d张，共%d张" % (list_amount, len(f_list)))
        time.sleep(1)
        wincad.Documents.Open(i)
        doc = wincad.ActiveDocument
        time.sleep(0.5)
        try:
            doc.SelectionSets.Item("SS1").Delete()
        except:
            pass
        slt = doc.SelectionSets.Add("SS1")
        filterType = [0]  # 定义过滤类型
        filterData = ["INSERT"]  # 设置过滤参数
        filterType = Tc.vtint(filterType)
        filterData = Tc.vtvariant(filterData)
        slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
        frame_exist, title_exist, framesize = 0, 0, 0
        time.sleep(0.5)
        for entity in slt:
            if entity.Name == 'sbtq' or entity.Name == 'tq5':
                point1, point2 = entity.GetBoundingBox()
                title_length = point2[0] - point1[0]
                title_height = point2[1] - point1[1]
                title_exist = 1
            elif entity.Name == '图框 A3至A4 图框':
                point3, point4 = entity.GetBoundingBox()
                frame_length = point3[0] - point4[0]
                frame_height = point3[1] - point4[1]
                frame_exist = 1
            elif entity.Name == '图框 A0至A2 图框':
                point3, point4 = entity.GetBoundingBox()
                frame_length = point3[0] - point4[0]
                frame_height = point3[1] - point4[1]
                frame_exist = 2
                framesize = 1
        slt.Delete()
        doc.Close()
        file_name = os.path.split(i)
        if frame_exist != 0 and title_exist != 0:
            if title_exist == 1:
                if round(frame_length / title_length, 2) == 1.17:  # A4
                    framesize = 0.125 * round((frame_height / frame_length) / (297 / 210), 1)
                else:  # A3
                    framesize = 0.25 * round((frame_length / frame_height) / (420 / 297), 1)
            else:
                if round((frame_height / title_height), 0) == 10:  # A2
                    framesize = 0.5 * round((frame_length / frame_height) / (594 / 420), 1)
                else:  # A1
                    framesize = 1 * round((frame_length / frame_height) / (841 / 594), 1)
            drawing_amount += framesize
        else:
            drawing_amount0_list.append([file_name[1], framesize])
        print("%s：%.3fA1" % (file_name[1], framesize))
    print("总图纸量为：%.3fA1" % drawing_amount)
    print("\n以下%d张图纸未统计，请手动计算" % len(drawing_amount0_list))
    for each in drawing_amount0_list:
        print(each, end=',\n')


if __name__ == '__main__':
    path = input("请输入要统计的文件夹路径（默认包含所有子文件夹）：")
    # path = 'r"' + path + '"'
    f_list = getfilefame(path)
    print("共找到%d个dwg文件" % len(f_list))

    wincad = win32com.client.Dispatch("AutoCAD.Application")
    cal_drawing(f_list, wincad)
    input()
