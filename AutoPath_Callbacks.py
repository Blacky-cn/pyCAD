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

from AutoPath_PathWidgets import PathWidgets


# =========================================================
class CallBacks:
    def __init__(self, oop):
        wincad = win32com.client.Dispatch("AutoCAD.Application.22")
        self.doc = wincad.ActiveDocument
        time.sleep(0.1)
        self.doc.Utility.Prompt("Hello! AutoPath from LZH.\n")
        time.sleep(0.1)
        self.msp = self.doc.ModelSpace
        self.oop = oop
        self.tab2 = ttk.Frame(self.oop.tabcontrol)  # Create Tab2
        self.oop.tabcontrol.add(self.tab2, text='输入参数', state='disabled')  # Add Tab2
        self.oop.tabcontrol.pack(expand=1, fill='both')  # Pack to make visible
        self.pathwidgets = PathWidgets(self, self.tab2, self.doc, self.msp)

    # Radiobutton Callback
    def nexten(self):
        """当工件类型、轨迹类型都选择后，'下一步'按钮变为可用"""
        if (self.oop.cartype_value.get() == 0) or (self.oop.pathtype_value.get() == 0):
            self.oop.b_next.configure(state='disabled')
        else:
            self.oop.b_next.configure(state='normal')

    # Button Click Function
    def donext(self):
        """根据选择的工件类型、轨迹类型，生成对应初始化GUI"""
        self.oop.tabcontrol.tab(1, state='normal')
        self.oop.tabcontrol.select(self.tab2)
        for child in self.tab2.winfo_children():
            child.destroy()
        self.doselect = self.oop.cartype_value.get() * 10 + self.oop.pathtype_value.get()
        if self.doselect // 10 == 1:  # 摆杆
            self.pathwidgets.donext_pendulum(self.doselect)
        elif self.doselect // 10 == 2:  # 翻转机
            pass
        else:  # 台车
            if self.doselect % 10 <= 2:  # 2台车
                self.pathwidgets.donext_2trolley(self.doselect)
            else:  # 4台车
                self.pathwidgets.donext_4trolley(self.doselect)

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
