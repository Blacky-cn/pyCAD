#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
"""

import tkinter as tk
from tkinter import ttk, Menu
from AutoPath_Callbacks import Callbacks


class OOP(object):
    def __init__(self):
        # Create Instance
        self.win = tk.Tk()
        self.win.title("AutoPath")
        self.win.geometry('660x380')
        self.create_widgets()

    ##########################################################################
    def create_widgets(self):
        # Create Tab Control
        self.tabControl = ttk.Notebook(self.win)
        tab1 = ttk.Frame(self.tabControl)  # Create Tab1
        self.tabControl.add(tab1, text='功能选择')  # Add Tab1
        self.callBacks = Callbacks(self)

        # Adding a Label
        ttk.Label(tab1, text='请选择工件类型: ').grid(column=0, row=0, sticky=tk.W)
        # Adding Radiobutton
        self.cartype_value = tk.IntVar()
        cartype = [('摆杆', 1), ('翻转机', 2), ('台车', 3)]
        for i, j in cartype:
            self.a1 = ttk.Radiobutton(tab1, text=i, value=j, variable=self.cartype_value, command=self.callBacks.nexten)
            self.a1.grid(column=j, row=0, sticky=tk.W)
        # Adding a Label
        ttk.Label(tab1, text='请选择轨迹类型: ').grid(column=0, row=1, sticky=tk.W)
        # Adding Radiobutton
        self.pathtype_value = tk.IntVar()
        pathtype = [('绘制轨迹', 1), ('仿真动画', 2)]
        for i, j in pathtype:
            self.a2 = ttk.Radiobutton(tab1, text=i, value=j, variable=self.pathtype_value,
                                      command=self.callBacks.nexten)
            self.a2.grid(column=j, row=1, sticky=tk.W)

        # LabelFrame using Tab1 as the parent
        tab1_Button = ttk.LabelFrame(tab1, text='')
        tab1_Button.grid(column=0, row=2, columnspan=4)
        self.b11 = ttk.Button(tab1_Button, text='下一步', state='disabled', command=self.callBacks.donext)
        self.b11.grid(column=0, row=0, padx=8, pady=8)
        self.b12 = ttk.Button(tab1_Button, text='退出', command=self.callBacks.quit)
        self.b12.grid(column=1, row=0, padx=8, pady=8)

        for child in tab1.winfo_children():
            child.grid_configure(padx=8, pady=4)

        # Adding Menu
        menu_bar = Menu(self.win)
        self.win.config(menu=menu_bar)
        # Adding File Menu
        file_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='文件', menu=file_menu)
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self.callBacks.quit)
        # Adding Help Menu
        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='帮助', menu=help_menu)
        help_menu.add_command(label='入门指南')
        help_menu.add_separator()
        help_menu.add_command(label='关于', command=self.callBacks.aboutmsg)


# ============================================================
if __name__ == '__main__':
    oop = OOP()
    # oop.win.iconbitmap('car.ico')
    oop.win.mainloop()
