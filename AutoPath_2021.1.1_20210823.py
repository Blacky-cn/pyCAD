#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
"""

# imports==========
import tkinter as tk
from tkinter import ttk, Menu

from AutoPath_CallBacks import CallBacks


# =========================================================
class OOP(object):
    def __init__(self):
        # Create Instance
        self.win = tk.Tk()
        self.win.title("AutoPath")
        self.win.geometry('580x410')
        self.win.resizable(False, False)
        self.create_widgets()

    def create_widgets(self):
        """选择想要实现的功能，并准备进入下一步，生成对应初始化GUI"""
        # Create Tab Control
        self.tabcontrol = ttk.Notebook(self.win)
        # Create Tab1
        self.tab1 = ttk.Frame(self.tabcontrol)
        # Add Tab1_功能选择
        self.tabcontrol.add(self.tab1, text='功能选择')
        self.callbacks = CallBacks(self)

        # Add Frame1_工件类型=======================================
        self.tab1_frame1 = ttk.LabelFrame(self.tab1, text='1.请选择工件类型:')
        self.tab1_frame1.grid(column=0, row=0, sticky='WE', padx=8, pady=4)

        # Add Radiobutton_工件类型
        self.cartype_value = tk.IntVar()
        cartype = {'摆杆': 1, '翻转机': 2, '台车': 3}
        for key, values in cartype:
            self.a_cartype = ttk.Radiobutton(self.tab1_frame1, text=key, value=values, variable=self.cartype_value,
                                             command=self.choose_pathtype)
            self.a_cartype.grid(column=values, row=0, sticky=tk.W)
        self.cartype_value.set(0)

        for child in self.tab1_frame1.winfo_children():
            child.grid_configure(padx=8, pady=4)

        # Add Frame2_轨迹类型=======================================
        self.tab1_frame2 = ttk.LabelFrame(self.tab1, text='2.请选择轨迹类型:')
        self.tab1_frame2.grid(column=0, row=1, sticky='WE', padx=8, pady=4)

        # Add Frame3_确定与退出======================================
        tab1_frame3 = ttk.LabelFrame(self.tab1, text='')
        tab1_frame3.grid(column=0, row=2, padx=8, pady=4)
        self.b_next = ttk.Button(tab1_frame3, text='下一步', state='disabled', command=self.callbacks.donext)
        self.b_next.grid(column=0, row=0, padx=8, pady=8)
        ttk.Button(tab1_frame3, text='退出', command=self.callbacks.quit).grid(column=1, row=0, padx=8, pady=8)

        # Add Menu==============================================
        menu_bar = Menu(self.win)
        self.win.config(menu=menu_bar)
        # Add File Menu
        file_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='文件', menu=file_menu)
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self.callbacks.quit)
        # Add Help Menu
        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='帮助', menu=help_menu)
        help_menu.add_command(label='入门指南')
        help_menu.add_separator()
        help_menu.add_command(label='关于', command=self.callbacks.aboutmsg)

    def choose_pathtype(self):
        """根据工件类型，选择轨迹类型"""
        for child in self.tab1_frame2.winfo_children():
            child.destroy()
        # Add Radiobutton_轨迹类型
        self.pathtype_value = tk.IntVar()
        if self.cartype_value.get() == 1:
            pathtype = {
                '工艺段-绘制轨迹': 1,
                '工艺段-仿真动画': 2,
                '返回段-绘制轨迹': 3,
                '返回段-仿真动画': 4,
                '浸入即出槽分析': 5,
                '工艺时间计算': 6
            }
        elif self.cartype_value.get() == 2:
            pathtype = {
                '工艺段-绘制轨迹': 1,
                '工艺段-仿真动画': 2,
                '返回段-绘制轨迹': 3,
                '返回段-仿真动画': 4,
                '浸入即出槽分析': 5,
                '工艺时间计算': 6
            }
        else:
            pathtype = {
                '2台车-绘制轨迹': 1,
                '2台车-动画仿真': 2,
                '4台车-绘制轨迹': 3,
                '4台车-动画仿真': 4
            }
        for key, values in pathtype.items():
            self.a_pathtype = ttk.Radiobutton(self.tab1_frame2, text=key, value=values, variable=self.pathtype_value,
                                              command=self.callbacks.nexten)
            self.a_pathtype.grid(column=(values - 1) % 4, row=(values - 1) // 4, sticky=tk.W)
        self.pathtype_value.set(0)

        for child in self.tab1_frame2.winfo_children():
            child.grid_configure(padx=8, pady=4)


# =========================================================
if __name__ == '__main__':
    oop = OOP()
    # oop.win.iconbitmap('car.ico')
    oop.win.mainloop()
