#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
GUI Button调用函数
"""

import win32com.client
import time
from tkinter import ttk, messagebox as msg
from AutoPath_Dopath import Dopath


class Callbacks(object):
    def __init__(self, oop):
        wincad = win32com.client.Dispatch("AutoCAD.Application")
        self.doc = wincad.ActiveDocument
        time.sleep(0.1)
        self.doc.Utility.Prompt("Hello! Autocad from pywin32com.\n")
        time.sleep(0.1)
        self.msp = self.doc.ModelSpace
        self.oop = oop
        self.tab2 = ttk.Frame(self.oop.tabControl)  # Create Tab2
        self.oop.tabControl.add(self.tab2, text='输入参数', state='disabled')  # Add Tab2
        self.oop.tabControl.pack(expand=1, fill='both')  # Pack to make visible
        self.doPath = Dopath(self, self.tab2, self.doc, self.msp)

    # Button Click Function
    def donext(self):
        self.oop.tabControl.tab(1, state='normal')
        self.oop.tabControl.select(self.tab2)
        for child in self.tab2.winfo_children():
            child.destroy()
        self.doselect = self.oop.cartype_value.get() * 10 + self.oop.pathtype_value.get()
        if self.doselect // 10 == 1:
            self.doPath.donext1(self.doselect)
        elif self.doselect // 10 == 3:
            self.doPath.donext3(self.doselect)

    # Radiobutton Callback
    def nexten(self):
        if (self.oop.cartype_value.get() == 0) or (self.oop.pathtype_value.get() == 0):
            self.oop.b11.configure(state='disabled')
        else:
            self.oop.b11.configure(state='normal')

    # Exit GUI cleanly
    def quit(self):
        self.oop.win.quit()
        self.oop.win.destroy()
        exit()

    # About Menu
    def aboutmsg(self):
        msg.showinfo('关于 AutoPath',
                     'AutoPath 2021.1.1 Beta\n\n'
                     'Copyright \u00a9 2021 LZH. All Rights Reserved.\n\n'
                     'Supporting Autodesk CAD: v2010-v2020. Other versions are not tested.')
