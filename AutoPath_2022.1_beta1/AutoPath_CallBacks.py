#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
GUI Button调用函数
"""

# imports==========
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
    def aboutmsg():
        """About Menu"""
        msg.showinfo('关于 AutoPath',
                     'AutoPath 2021.1.1 Beta\n\n'
                     'Copyright \u00a9 2021 LZH. All Rights Reserved.\n\n'
                     'Supporting Autodesk CAD: v2010-v2020. Other versions are not tested.')

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
        # elif block_name == 'chainplate':
        #     self.chainplate = block[0]
        #     self.b_choose_chainplate.configure(style='G.TButton', text='已选择')
        # elif block_name == 'fswing':
        #     self.fswing = block[0]
        #     self.b_choose_fswing.configure(style='G.TButton', text='已选择')
        # elif block_name == 'bswing':
        #     self.bswing = block[0]
        #     self.b_choose_bswing.configure(style='G.TButton', text='已选择')

    def click_pline(self, pline_name):
        """选择轨道线"""
        pline = self.doc.Utility.GetEntity()  # 在cad中选择多段线
        # time.sleep(0.1)
        while 'polyline' not in pline[0].ObjectName.lower():  # 判断选取的是否为多段线
            msg.showerror('错误', '您选择的不是多段线，\n请重新选择！')
            # time.sleep(0.1)
            pline = self.doc.Utility.GetEntity()
        if pline_name == 'chainpath':
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
            pline1_Pnt = []
            pline_bulge = [1] * pline_vertex
            pline1_bulge = [1] * pline_vertex
            if pline_pnt[0] > pline_pnt[-2] and pline_pnt[1] < pline_pnt[-1]:
                rev1 = -1
            else:
                rev1 = 2
            if start_pnt[:2] == pline_pnt[-2:]:
                rev2 = -1  # 指定起点相反
                for i in range(len(pline_pnt)):
                    pline1_Pnt.append(pline_pnt[(pline_vertex - i // 2 - 1) * 2 + i % 2])
                for i in range(pline_vertex):
                    pline_bulge[i] = -pline[0].GetBulge(i)
                for i in range(pline_vertex):
                    if i == pline_vertex - 1:
                        pline1_bulge[i] = pline_bulge[i]
                    else:
                        pline1_bulge[i] = pline_bulge[pline_vertex - i - 2]
                pline1 = self.msp.AddLightWeightPolyline(Tc.vtfloat(pline1_Pnt))
                for i in range(pline_vertex):
                    pline1.SetBulge(i, pline1_bulge[i])
                pline1.Layer, pline1.Linetype, pline1.ConstantWidth = pllayer, pltype, plcw
                pline1.Update()
                pline[0].Delete()
            else:
                rev2 = 1  # 指定起点相同
                pline1 = pline[0]
            rev = rev1 * rev2
            self.chainpath, self.chainpath_vertex, self.chainpath_rev = pline1, pline_vertex, rev
            self.oop.b_choose_chainpath.configure(style='G.TButton', text='已选择')
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
            self.immersion = list(self.doc.Utility.GetPoint())
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
