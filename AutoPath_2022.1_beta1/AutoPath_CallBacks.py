#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
GUI Button调用函数
"""

# imports==========
from tkinter import messagebox as msg


# =========================================================
class CallBacks:
    def __init__(self, oop):
        self.oop = oop

    # Radiobutton Callback
    def nexten(self):
        """当工件类型、轨迹类型都选择后，'下一步'按钮变为可用"""
        if (self.oop.cartype_value.get() == 0) or (self.oop.pathtype_value.get() == 0):
            self.oop.b_next.configure(state='disabled')
        else:
            self.oop.b_next.configure(state='normal')

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
