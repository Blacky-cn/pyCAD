#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: LZH
GUI Button调用函数
"""

# imports==========
import os
import sys
import time
from tkinter import ttk, messagebox as msg

import win32com.client

import AutoPath_TypeConvert as Tc


# =========================================================
class CALLBACKS:
    def __init__(self, oop):
        self.oop = oop

        wincad = win32com.client.Dispatch("AutoCAD.Application")
        self.doc = wincad.ActiveDocument
        time.sleep(0.1)
        self.doc.Utility.Prompt("Hello! AutoPath from LZH.\n")
        time.sleep(0.1)
        self.msp = self.doc.ModelSpace

        style = ttk.Style()
        style.configure('R.TButton', foreground='red')
        style.configure('G.TButton', foreground='green')

    # Radiobutton Callback
    def previousbutton(self):
        """返回初始GUI"""
        self.oop.main_frame.destroy()
        self.oop.page_menu()

    def nextbutton(self):
        """根据选择的工件类型、轨迹类型，生成对应初始化GUI"""
        if self.oop.cartype_value.get() == 1:  # 摆杆
            self.oop.main_frame.destroy()
            self.oop.page_pendulum()
        elif self.oop.cartype_value.get() == 2:  # 翻转机
            self.oop.main_frame.destroy()
            self.oop.page_shuttle()
        else:  # 台车
            if self.oop.pathtype_value.get() <= 2:  # 2台车
                self.oop.main_frame.destroy()
                self.oop.page_trolley2()
            else:  # 4台车
                self.oop.main_frame.destroy()
                self.oop.page_trolley4()

    def quit(self):
        """Exit GUI cleanly"""
        self.oop.win.quit()
        self.oop.win.destroy()
        exit()

    @staticmethod
    def readme():
        """readme Menu"""
        if hasattr(sys, '_MEIPASS'):
            path = os.path.join(sys._MEIPASS, 'docs/README.html')
        else:
            path = os.path.join(os.path.abspath('.'), 'docs/README.html')
        os.startfile(path)

    @staticmethod
    def changelog():
        """changlog Menu"""
        if hasattr(sys, '_MEIPASS'):
            path = os.path.join(sys._MEIPASS, 'docs/CHANGELOG.html')
        else:
            path = os.path.join(os.path.abspath('.'), 'docs/CHANGELOG.html')
        os.startfile(path)

    @staticmethod
    def aboutmsg():
        """About Menu"""
        msg.showinfo('关于 AE-Painting AutoPath',
                     'AutoPath 2023.1\n\n'
                     '\u00a9 Copyright by LZH. All Rights Reserved.\n\n'
                     'Supporting Autodesk CAD: v2012-v2024.')

    def click_block(self, block_name):
        """选择图块"""
        block = self.doc.Utility.GetEntity()  # 在cad中选取块
        time.sleep(0.1)
        while 'block' not in block[0].ObjectName.lower():  # 判断选取的是否为块
            msg.showerror('错误', '您选择的不是块，\n请重新选择！')
            block = self.doc.Utility.GetEntity()
            time.sleep(0.1)
        if block_name == 'car':
            self.car_name = block[0].Name
            self.oop.b_choose_car.configure(style='G.TButton', text='已选择')

    def click_pline(self, pline_name):
        """选择轨道线"""
        pline = self.doc.Utility.GetEntity()  # 在cad中选择多段线
        # time.sleep(0.1)
        while 'polyline' not in pline[0].ObjectName.lower():  # 判断选取的是否为多段线
            msg.showerror('错误', '您选择的不是多段线，\n请重新选择！')
            # time.sleep(0.1)
            pline = self.doc.Utility.GetEntity()
        if pline_name == 'chainpath' or pline_name == 'hinglepath':
            pline_pnt = list(pline[0].Coordinates[:])  # 获取多段线各顶点坐标
            pline_vertex = len(pline_pnt) // 2
            pllayer, pltype, plcw = pline[0].Layer, pline[0].Linetype, pline[0].ConstantWidth
            # 指定并调整轨迹线方向
            self.doc.Utility.Prompt("请指定多段线的起点，在相应端点上单击")
            # time.sleep(0.1)
            start_pnt = list(self.doc.Utility.GetPoint(Prompt=''))
            while start_pnt[:2] != pline_pnt[:2] and start_pnt[:2] != pline_pnt[-2:]:
                msg.showerror('错误', '未选择多段线的端点，请重新选择！')
                self.doc.Utility.Prompt("请指定多段线的起点，在相应端点上单击")
                start_pnt = list(self.doc.Utility.GetPoint(Prompt=''))
            pline1_pnt = []
            pline_bulge = [1] * pline_vertex
            pline1_bulge = [1] * pline_vertex
            if pline_pnt[0] > pline_pnt[-2] and pline_pnt[1] < pline_pnt[-1]:
                rev1 = -1
            else:
                rev1 = 2
            if start_pnt[:2] == pline_pnt[-2:]:
                rev2 = -1  # 指定起点相反
                for i in range(len(pline_pnt)):
                    pline1_pnt.append(pline_pnt[(pline_vertex - i // 2 - 1) * 2 + i % 2])
                for i in range(pline_vertex):
                    pline_bulge[i] = -pline[0].GetBulge(i)
                for i in range(pline_vertex):
                    if i == pline_vertex - 1:
                        pline1_bulge[i] = pline_bulge[i]
                    else:
                        pline1_bulge[i] = pline_bulge[pline_vertex - i - 2]
                pline1 = self.msp.AddLightWeightPolyline(Tc.vtfloat(pline1_pnt))
                for i in range(pline_vertex):
                    pline1.SetBulge(i, pline1_bulge[i])
                pline1.Layer, pline1.Linetype, pline1.ConstantWidth = pllayer, pltype, plcw
                pline1.Update()
                pline[0].Delete()
            else:
                rev2 = 1  # 指定起点相同
                pline1 = pline[0]
            rev = rev1 * rev2
            if pline_name == 'chainpath':
                self.chainpath, self.chainpath_vertex, self.chainpath_rev = pline1, pline_vertex, rev
                self.oop.b_choose_chainpath.configure(style='G.TButton', text='已选择')
            elif pline_name == 'hinglepath':
                self.hinglepath, self.hinglepath_vertex, self.hinglepath_rev = pline1, pline_vertex, rev
                self.oop.b_choose_hinglepath.configure(style='G.TButton', text='已选择')
        elif pline_name == 'rollerpath':
            self.rollerpath = pline[0]  # 获取滚轮所在轨迹线
            self.oop.b_choose_rollerpath.configure(style='G.TButton', text='已选择')

    def click_point(self, point_name):
        """选择包络点/全浸点"""
        if point_name == 'uenvelope':
            self.uenvelope = list(self.doc.Utility.GetPoint())
            self.oop.b_choose_uenvelope.configure(style='G.TButton', text='已选择')
        elif point_name == 'lenvelope':
            self.lenvelope = list(self.doc.Utility.GetPoint())
            self.oop.b_choose_lenvelope.configure(style='G.TButton', text='已选择')
        elif point_name == 'immersion':
            self.doc.SendCommand('_bedit\n')
            immersionpnt = list(self.doc.Utility.GetPoint())
            self.msp.AddLine(Tc.vtpnt(immersionpnt[0], immersionpnt[1]),
                             Tc.vtpnt(immersionpnt[0], immersionpnt[1] + 50))
            immersionpnt = Tc.vtpnt(immersionpnt[0], immersionpnt[1])
            self.msp.AddAttribute(30, 4, 'beginspoint', immersionpnt, 'beginspoint', 'beginspoint')
            leachingpnt = list(self.doc.Utility.GetPoint())
            self.msp.AddLine(Tc.vtpnt(leachingpnt[0], leachingpnt[1]),
                             Tc.vtpnt(leachingpnt[0], leachingpnt[1] - 50))
            leachingpnt = Tc.vtpnt(leachingpnt[0], leachingpnt[1])
            self.msp.AddAttribute(30, 4, 'endspoint', leachingpnt, 'endspoint', 'endspoint')
            self.doc.SendCommand('_bclose\n')
            self.doc.SendCommand('attsync n ' + 'Skid&Body' + '\n')
            self.oop.b_choose_immersion.configure(style='G.TButton', text='已选择')
        elif point_name == 'diptank_midpnt':
            self.diptank_midpnt = list(self.doc.Utility.GetPoint())
            self.oop.b_choose_diptank_midpnt.configure(style='G.TButton', text='已选择')

    def envolope(self, value):
        """包络线选项状态控制"""
        if value == 0:
            self.oop.b_choose_uenvelope['state'] = 'disabled'
            self.oop.b_choose_lenvelope['state'] = 'disabled'
        else:
            self.oop.b_choose_uenvelope['state'] = 'normal'
            self.oop.b_choose_lenvelope['state'] = 'normal'

    def delete_block_attributes(self):
        """删除选定块及其内部嵌套的块参照的所有属性"""
        if msg.askokcancel('提示',
                           '将删除选定块的属性，之后对该块的引用将不包含属性。\n提示完成前请勿进行其他操作，要执行此操作吗?'):
            block = self.doc.Utility.GetEntity()  # 在cad中选取块
            time.sleep(0.1)
            while 'block' not in block[0].ObjectName.lower():  # 判断选取的是否为块
                msg.showerror('错误', '您选择的不是块，\n请重新选择！')
                block = self.doc.Utility.GetEntity()
                time.sleep(0.1)
            block_definition = self.doc.Blocks(block[0].Name)
            for entity in block_definition:
                if entity.EntityName:
                    if entity.EntityName == 'AcDbAttributeDefinition':
                        entity.Delete()
                    elif entity.EntityName == 'AcDbBlockReference':
                        if entity.HasAttributes:
                            for attribute in entity.GetAttributes():
                                attribute.Erase()
            msg.showinfo('提示', '已删除选定块的属性')

    def delete_blockreference_attributes(self):
        """删除文件中所有块参照的属性"""
        if msg.askokcancel('提示',
                           '将删除文件中所有块参照的属性，但对块定义没有影响。\n提示完成前请勿进行其他操作，要执行此操作吗？'):
            for entity in self.msp:
                if entity.EntityName:
                    if entity.EntityName == 'AcDbBlockReference':
                        if entity.HasAttributes:
                            for attribute in entity.GetAttributes():
                                attribute.Erase()
                            entity.Update()
            msg.showinfo('提示', '已删除文件中所有块参照的属性')
