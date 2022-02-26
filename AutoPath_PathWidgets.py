#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
根据所选择的功能，生成对应初始化GUI
"""

# =================
# imports
# =================
import time
import tkinter as tk
from tkinter import ttk, messagebox as msg

import win32com.client

import AutoPath_ToopTip as tt
import AutoPath_TypeConvert as tc
from AutoPath_CallBacks import CallBacks
from AutoPath_DoPath import DoPath


# =========================================================
class PathWidgets:
    def __init__(self, oop):
        wincad = win32com.client.Dispatch("AutoCAD.Application")
        self.doc = wincad.ActiveDocument
        time.sleep(0.1)
        self.doc.Utility.Prompt("Hello! AutoPath from LZH.\n")
        time.sleep(0.1)
        self.msp = self.doc.ModelSpace

        self.oop = oop
        self.tab2 = ttk.Frame(self.oop.tabcontrol)  # Create Tab2
        self.oop.tabcontrol.add(self.tab2, text='输入参数', state='disabled')  # Add Tab2
        self.oop.tabcontrol.pack(expand=1, fill='both')  # Pack to make visible
        self.callbacks = CallBacks(self)
        self.doPath = DoPath(self.oop, self.doc, self.msp)

        style = ttk.Style()
        style.configure('R.TButton', foreground='red')
        style.configure('G.TButton', foreground='green')

    def donext(self):
        """根据选择的工件类型、轨迹类型，生成对应初始化GUI"""
        self.oop.tabcontrol.tab(1, state='normal')
        self.oop.tabcontrol.select(self.tab2)
        for child in self.tab2.winfo_children():
            child.destroy()
        self.doselect = self.oop.cartype_value.get() * 10 + self.oop.pathtype_value.get()
        if self.doselect // 10 == 1:  # 摆杆
            self.donext_pendulum(self.doselect)
        elif self.doselect // 10 == 2:  # 翻转机
            pass
        else:  # 台车
            if self.doselect % 10 <= 2:  # 2台车
                self.donext_2trolley(self.doselect)
            else:  # 4台车
                self.donext_4trolley(self.doselect)

    def donext_pendulum(self, select):
        """选择摆杆程序，生成对应初始化GUI"""
        # Add Frame1_图块选择==================================
        tab2_frame1 = ttk.LabelFrame(self.tab2, text='图块选择')
        tab2_frame1.grid(column=0, row=1, sticky='WNS', padx=8, pady=4)

        # Add a Label_选择轨迹线
        ttk.Label(tab2_frame1, text='轨迹线及起点: ').grid(column=0, row=0)
        # Add a Button_选择轨迹线
        self.b_choose_chainpath = ttk.Button(tab2_frame1, text='单击选择', style='R.TButton',
                                             command=lambda: self.click_pline('chainpath'))
        self.b_choose_chainpath.grid(column=1, row=0)

        if select % 10 == 3 or select % 10 == 4:
            # Add a Label_选择轨迹线
            ttk.Label(tab2_frame1, text='摆杆滚轮所在轨迹线: ').grid(column=0, row=1)
            # Add a Button_选择轨迹线
            self.b_choose_rollerpath = ttk.Button(tab2_frame1, text='单击选择', style='R.TButton',
                                                  command=lambda: self.click_pline('rollerpath'))
            self.b_choose_rollerpath.grid(column=1, row=1)

        if select % 10 == 5:
            # Add a Label_选择轨迹中心点
            ttk.Label(tab2_frame1, text='轨迹中心点: ').grid(column=0, row=1)
            # Add a Button_选择轨迹中心点
            self.b_choose_diptank_midpnt = ttk.Button(tab2_frame1, text='单击选择', style='R.TButton',
                                                      command=lambda: self.click_point('diptank_midpnt'))
            self.b_choose_diptank_midpnt.grid(column=1, row=1)

        # Add a Label_选择工件方向
        l_dir = ttk.Label(tab2_frame1, text='工件/摆杆块方向: ')
        l_dir.grid(column=0, row=2)
        # Add a Tooltip_提示框
        tt.create_tooltip(l_dir, '工件块、链板块与摆杆块方向需一致')
        # Add Radiobutton_选择工件方向
        self.dirvalue = tk.IntVar()
        self.dirvalue.set(1)
        cardirs = {'右': 1, '左': 2}
        for cardir, cardir_num in cardirs.items():
            rb_dir = ttk.Radiobutton(tab2_frame1, text=cardir, value=cardir_num, variable=self.dirvalue)
            rb_dir.grid(column=cardir_num, row=2)

        if select % 10 == 1 or select % 10 == 2:
            # Add a Label_选择摆杆状态
            ttk.Label(tab2_frame1, text='摆杆状态: ').grid(column=0, row=3)
            # Add Radiobutton_选择摆杆状态
            self.swingstate_value = tk.IntVar()
            self.swingstate_value.set(1)
            swingstates = {'前摆杆竖直': 1, '后摆杆竖直': 2}
            for swingstate, swingstate_num in swingstates.items():
                rb_swingstate = ttk.Radiobutton(tab2_frame1, text=swingstate, value=swingstate_num,
                                                variable=self.swingstate_value)
                rb_swingstate.grid(column=swingstate_num, row=3)

            # Add a Label_是否绘制包络线
            c_envelope_value = tk.IntVar()
            c_envelope = tk.Checkbutton(tab2_frame1, text='提取包络线: ', variable=c_envelope_value,
                                        command=lambda: self.envolope(c_envelope_value.get()))
            c_envelope.grid(column=0, row=4)
            c_envelope.deselect()
            # Add a Button_选择上包络线特征点
            self.b_choose_uenvelope = ttk.Button(tab2_frame1, text='选择上特征点', state='disabled', style='R.TButton',
                                                 command=lambda: self.click_point('uenvelope'))
            self.b_choose_uenvelope.grid(column=1, row=4)
            # Add a Button_选择下包络线特征点
            self.b_choose_lenvelope = ttk.Button(tab2_frame1, text='选择下特征点', state='disabled', style='R.TButton',
                                                 command=lambda: self.click_point('lenvelope'))
            self.b_choose_lenvelope.grid(column=2, row=4)
            # 若为仿真动画，禁用包络线选项
            if select % 10 == 2:
                c_envelope.configure(state='disabled')
                c_envelope.deselect()

        for child in tab2_frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2_基本参数==================================
        tab2_frame2 = ttk.LabelFrame(self.tab2, text='基本参数')
        tab2_frame2.grid(column=1, row=1, sticky='WNS', padx=8, pady=4)

        # Add a Label_链板节距
        ttk.Label(tab2_frame2, text='链板节距(mm): ').grid(column=0, row=0)
        # Add an Entry_链板节距
        self.chainbracing = tk.DoubleVar()
        ttk.Entry(tab2_frame2, width=12, textvariable=self.chainbracing).grid(column=1, row=0)
        self.chainbracing.set(250)

        # Add a Label_摆杆间距
        ttk.Label(tab2_frame2, text='摆杆间距(mm): ').grid(column=0, row=1)
        self.bracing = tk.DoubleVar()
        # Add an Entry_摆杆间距
        ttk.Entry(tab2_frame2, width=12, textvariable=self.bracing).grid(column=1, row=1)
        self.bracing.set(3250)

        # Add a Label_摆杆长度
        l_swingleng = tk.Label(tab2_frame2, text='摆杆长度(mm): ')
        l_swingleng.grid(column=0, row=2)
        # Add a Tooltip_提示框
        tt.create_tooltip(l_swingleng, '套筒中心至摆杆底部圆管中心距离')
        # Add an Entry_摆杆长度
        self.swingleng = tk.DoubleVar()
        ttk.Entry(tab2_frame2, width=12, textvariable=self.swingleng).grid(column=1, row=2)
        self.swingleng.set(2950)

        # Add a Label_轨迹步长
        ttk.Label(tab2_frame2, text='轨迹步长(mm): ').grid(column=0, row=4)
        # Add an Entry_轨迹步长
        self.step = tk.IntVar()
        e_step = ttk.Entry(tab2_frame2, width=12, textvariable=self.step)
        e_step.grid(column=1, row=4)
        self.step.set(500)

        # Add a Label_工件数量
        l_num = ttk.Label(tab2_frame2, text='工件数量: ')
        l_num.grid(column=0, row=5)
        # Add an Entry_工件数量
        self.carnum = tk.IntVar()
        e_num = ttk.Entry(tab2_frame2, width=12, textvariable=self.carnum)
        e_num.grid(column=1, row=5)

        # Add a Label_工件节距
        l_pitch = ttk.Label(tab2_frame2, text='工件节距(mm): ')
        l_pitch.grid(column=0, row=6)
        # Add an Entry_工件节距
        self.pitch = tk.DoubleVar()
        e_pitch = ttk.Entry(tab2_frame2, width=12, textvariable=self.pitch)
        e_pitch.grid(column=1, row=6)
        self.pitch.set(6750)

        # 若为浸入即出槽分析，禁用工件数量、步长选项
        if select % 10 == 5:
            e_step.configure(state='disabled')
            e_num.configure(state='disabled')
        # 若为绘制轨迹，禁用工件数量、节距选项
        elif select % 10 != 2 and select % 10 != 4:
            e_num.configure(state='disabled')
            self.carnum.set(1)
            e_pitch.configure(state='disabled')
            self.pitch.set(0)

        for child in tab2_frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # 其他功能：浸入即出，工艺时间
        if select % 10 == 5:
            # Add Frame3_浸入即出==============================
            tab2_frame3 = ttk.LabelFrame(self.tab2, text='浸入即出槽状态分析')
            tab2_frame3.grid(column=0, row=3, sticky='WE', columnspan=2, padx=8, pady=4)

            ttk.Label(tab2_frame3, text='选择分析模式: ').grid(column=0, row=0)
            self.swingmode_value = tk.IntVar()
            self.swingmode_value.set(1)
            swingmodes = {'内侧摆杆竖直': 1, '外侧摆杆竖直': 2}
            for swingmode, swingmode_num in swingmodes.items():
                rb_swingmode = ttk.Radiobutton(tab2_frame3, text=swingmode, value=swingmode_num,
                                               variable=self.swingmode_value)
                rb_swingmode.grid(column=swingmode_num, row=0)

            for child in tab2_frame3.winfo_children():
                child.grid_configure(sticky=tk.W, padx=8, pady=4)
        elif select % 10 == 6:
            # Add Frame4_工艺时间==============================
            tab2_frame4 = ttk.LabelFrame(self.tab2, text='工艺时间计算')
            tab2_frame4.grid(column=0, row=4, sticky='WE', columnspan=2, padx=8, pady=4)

            # Add a Label_工件高度
            ttk.Label(tab2_frame4, text='工件高度(mm): ').grid(column=0, row=0)
            # Add an Entry_工件高度
            self.carheight = tk.StringVar()
            ttk.Entry(tab2_frame4, width=12, textvariable=self.carheight).grid(column=1, row=0)

            # Add a Label_液面高度
            ttk.Label(tab2_frame4, text='液面距轨道中心(mm): ').grid(column=2, row=0)
            # Add an Entry_液面高度
            self.liquidheight = tk.StringVar()
            ttk.Entry(tab2_frame4, width=12, textvariable=self.liquidheight).grid(column=3, row=0)

            # Add a Label_链速
            ttk.Label(tab2_frame4, text='链速(m/min): ').grid(column=0, row=1)
            # Add an Entry_链速
            self.chainspeed = tk.StringVar()
            ttk.Entry(tab2_frame4, width=12, textvariable=self.chainspeed).grid(column=1, row=1)

            # Add a Label_选择全浸点
            ttk.Label(tab2_frame4, text='工件全浸点: ').grid(column=2, row=1)
            # Add a Button_选择全浸点
            self.b_choose_immersion = ttk.Button(tab2_frame4, text='单击选择', style='R.TButton',
                                                 command=lambda: self.click_point('immersion'))
            self.b_choose_immersion.grid(column=3, row=1)

            for child in tab2_frame4.winfo_children():
                child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame5_退出====================================
        tab2_frame5 = ttk.LabelFrame(self.tab2, text='')
        tab2_frame5.grid(column=0, row=5, columnspan=2)
        if select % 10 <= 4:
            self.b_entry = ttk.Button(tab2_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum(select, self.dirvalue.get(),
                                                                              self.swingstate_value.get(),
                                                                              self.chainpath,
                                                                              self.step.get(),
                                                                              self.chainbracing.get(),
                                                                              self.bracing.get(),
                                                                              self.carnum.get(),
                                                                              self.pitch.get(),
                                                                              self.swingleng.get()) - 252.75)
        elif select % 10 == 5:
            self.b_entry = ttk.Button(tab2_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_diptank(self.swingmode_value.get(),
                                                                                      self.chainpath,
                                                                                      self.diptank_midpnt,
                                                                                      self.chainbracing.get(),
                                                                                      self.bracing.get(),
                                                                                      self.pitch.get(),
                                                                                      self.swingleng.get()) - 252.75)
        self.b_entry.grid(column=0, row=0, padx=20, pady=8)
        self.b_quit = ttk.Button(tab2_frame5, text='退出', command=self.callbacks.quit)
        self.b_quit.grid(column=1, row=0, padx=20, pady=8)

    def donext_2trolley(self, select):
        """选择2台车程序，生成对应初始化GUI"""
        # Add Frame1_图块选择==================================
        tab2_frame1 = ttk.LabelFrame(self.tab2, text='图块选择')
        tab2_frame1.grid(column=0, row=0, sticky='WNS', padx=8, pady=4)

        # Add a Label_选择工件
        ttk.Label(tab2_frame1, text='工件: ').grid(column=0, row=0)
        # Add a Button_选择工件
        self.b_choose_car = ttk.Button(tab2_frame1, text='单击选择', style='R.TButton',
                                       command=lambda: self.click_block('car'))
        self.b_choose_car.grid(column=1, row=0)

        # Add a Label_选择轨迹线
        ttk.Label(tab2_frame1, text='轨迹线及起点: ').grid(column=0, row=1)
        # Add a Button_选择轨迹线
        self.b_choose_chainpath = ttk.Button(tab2_frame1, text='单击选择', style='R.TButton',
                                             command=lambda: self.click_pline('chainpath'))
        self.b_choose_chainpath.grid(column=1, row=1)

        # Add a Label_选择工件方向
        ttk.Label(tab2_frame1, text='工件块方向: ').grid(column=0, row=2)
        # Add a Radiobutton_选择工件方向
        self.dirvalue = tk.IntVar()
        self.dirvalue.set(1)
        dirs = {'右': 1, '左': 2}
        for cardir, cardir_num in dirs.items():
            a31 = ttk.Radiobutton(tab2_frame1, text=cardir, value=cardir_num, variable=self.dirvalue)
            a31.grid(column=cardir_num, row=2)

        for child in tab2_frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2_基本参数==================================
        tab2_frame2 = ttk.LabelFrame(self.tab2, text='基本参数')
        tab2_frame2.grid(column=1, row=0, sticky='WNS', padx=8, pady=4)

        # Add a Label_工件前后支撑距离
        ttk.Label(tab2_frame2, text='工件前后支撑距离(mm): ').grid(column=0, row=0)
        # Add a Entry_工件前后支撑距离
        self.bracing = tk.DoubleVar()
        ttk.Entry(tab2_frame2, width=12, textvariable=self.bracing).grid(column=1, row=0)

        # Add a Label_轨迹步长
        ttk.Label(tab2_frame2, text='轨迹步长(mm): ').grid(column=0, row=1)
        # Add an Entry_轨迹步长
        self.step = tk.DoubleVar()
        ttk.Entry(tab2_frame2, width=12, textvariable=self.step).grid(column=1, row=1)

        # Add a Label_工件数量
        l_carnum = ttk.Label(tab2_frame2, text='工件数量: ')
        l_carnum.grid(column=0, row=2)
        # Add an Entry_工件数量
        self.carnum = tk.IntVar()
        e_carnum = ttk.Entry(tab2_frame2, width=12, textvariable=self.carnum)
        e_carnum.grid(column=1, row=2)
        # Add a Label_工件节距
        l_pitch = ttk.Label(tab2_frame2, text='工件节距(mm): ')
        l_pitch.grid(column=0, row=3)
        # Add an Entry_工件节距
        self.pitch = tk.DoubleVar()
        e_pitch = ttk.Entry(tab2_frame2, width=12, textvariable=self.pitch)
        e_pitch.grid(column=1, row=3)
        # 若为绘制轨迹，则禁用工件数量、节距选项
        if select % 10 == 1:
            l_carnum.configure(state='disabled')
            e_carnum.configure(state='disabled')
            self.carnum.set(1)
            l_pitch.configure(state='disabled')
            e_pitch.configure(state='disabled')
            self.pitch.set(0)

        for child in tab2_frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame3_开始与退出========================================
        tab2_frame3 = ttk.LabelFrame(self.tab2, text='')
        tab2_frame3.grid(column=0, row=1, columnspan=2)
        self.b_entry = ttk.Button(tab2_frame3, text='确定',
                                  command=lambda: self.doPath.do_2trolley(select, self.dirvalue.get(), self.car_name,
                                                                          self.chainpath, self.step.get(),
                                                                          self.bracing.get(), self.carnum.get(),
                                                                          self.pitch.get()))
        self.b_entry.grid(column=0, row=0, padx=20, pady=8)
        self.b_quit = ttk.Button(tab2_frame3, text='退出', command=self.oop.quit)
        self.b_quit.grid(column=1, row=0, padx=20, pady=8)

    def donext_4trolley(self, select):
        """选择4台车程序，生成对应初始化GUI"""
        pass

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
            self.b_choose_car.configure(style='G.TButton', text='已选择')
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
                pline1 = self.msp.AddLightWeightPolyline(tc.vtfloat(pline1_Pnt))
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
            self.b_choose_chainpath.configure(style='G.TButton', text='已选择')
        elif pline_name == 'rollerpath':
            self.rollerpath = pline[0]  # 获取滚轮所在轨迹线
            self.b_choose_rollerpath.configure(style='G.TButton', text='已选择')

    def click_point(self, point_name):
        """选择包络点/全浸点"""
        if point_name == 'uenvelope':
            self.uenvelope = list(self.doc.Utility.GetPoint())
            self.b_choose_uenvelope.configure(style='G.TButton', text='已选择')
        elif point_name == 'lenvelope':
            self.lenvelope = list(self.doc.Utility.GetPoint())
            self.b_choose_lenvelope.configure(style='G.TButton', text='已选择')
        elif point_name == 'immersion':
            self.immersion = list(self.doc.Utility.GetPoint())
            self.b_choose_immersion.configure(style='G.TButton', text='已选择')
        elif point_name == 'diptank_midpnt':
            self.diptank_midpnt = list(self.doc.Utility.GetPoint())
            self.b_choose_diptank_midpnt.configure(style='G.TButton', text='已选择')

    def envolope(self, value):
        """包络线选项状态控制"""
        if value == 0:
            self.b_choose_uenvelope['state'] = 'disabled'
            self.b_choose_lenvelope['state'] = 'disabled'
        else:
            self.b_choose_uenvelope['state'] = 'normal'
            self.b_choose_lenvelope['state'] = 'normal'
