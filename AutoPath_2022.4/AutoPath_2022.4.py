#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: LZH
"""
import time
# imports==========
import tkinter as tk
from tkinter import ttk, Menu

import AutoPath_ToopTip as Tt
from AutoPath_CallBacks import CALLBACKS
from AutoPath_DoPath import DOPATH


# =========================================================
class OOP(object):
    def __init__(self):
        # Create Instance
        self.win = tk.Tk()
        self.win.title("AutoPath")
        self.win.geometry('580x450')
        self.win.resizable(False, False)
        self.callbacks = CALLBACKS(self)
        self.doPath = DOPATH(self, self.callbacks.doc, self.callbacks.msp)
        self.page_menu()
        style = ttk.Style()
        style.configure('R.TButton', foreground='red')
        style.configure('G.TButton', foreground='green')

    def page_menu(self):
        """选择想要实现的轨迹类型，生成对应初始化GUI"""
        self.main_frame = ttk.LabelFrame(self.win, text='1.请选择轨迹类型:')
        self.main_frame.grid(column=0, row=0)
        # Add Frame1_工件类型=======================================
        menu_frame1 = ttk.LabelFrame(self.main_frame, text='请选择工件类型')
        menu_frame1.grid(column=0, row=0, sticky='WE', padx=8, pady=4)

        # Add Radiobutton_工件类型
        self.cartype_value = tk.IntVar()
        cartypes = {'摆杆': 1, '翻转机': 2, '台车': 3}
        for cartype, cartype_num in cartypes.items():
            self.a_cartype = ttk.Radiobutton(menu_frame1, text=cartype, value=cartype_num,
                                             variable=self.cartype_value, command=self._choose_pathtype)
            self.a_cartype.grid(column=cartype_num, row=0, sticky=tk.W)
        self.cartype_value.set(0)

        for child in menu_frame1.winfo_children():
            child.grid_configure(padx=8, pady=4)

        # Add Frame2_轨迹类型=======================================
        self.menu_frame2 = ttk.LabelFrame(self.main_frame, text='请选择轨迹类型')
        self.menu_frame2.grid(column=0, row=1, sticky='WE', padx=8, pady=4)

        # Add Frame3_确定与退出======================================
        menu_frame3 = ttk.LabelFrame(self.main_frame, text='')
        menu_frame3.grid(column=0, row=2, padx=8, pady=4)
        self.b_next = ttk.Button(menu_frame3, text='下一步', state='disabled', command=self.callbacks.nextbutton)
        self.b_next.grid(column=0, row=0, padx=8, pady=8)
        ttk.Button(menu_frame3, text='退出', command=self.callbacks.quit).grid(column=1, row=0, padx=8, pady=8)

        # Add Menu==============================================
        self._menubar()

    def _choose_pathtype(self):
        """根据工件类型，选择轨迹类型"""
        for child in self.menu_frame2.winfo_children():
            child.destroy()
        # Add Radiobutton_轨迹类型
        self.pathtype_value = tk.IntVar()
        if self.cartype_value.get() == 1:
            pathtypes = {
                '工艺段-绘制轨迹': 1,
                '工艺段-仿真动画': 2,
                '返回段-绘制轨迹': 3,
                '返回段-仿真动画': 4,
                '浸入即出槽分析': 5,
                '工艺时间计算': 6
            }
        elif self.cartype_value.get() == 2:
            pathtypes = {
                '工艺段-绘制轨迹': 1,
                '工艺段-仿真动画': 2,
                '返回段-绘制轨迹': 3,
                '返回段-仿真动画': 4,
                '浸入即出槽分析': 5,
                '工艺时间计算': 6
            }
        else:
            pathtypes = {
                '2台车-绘制轨迹': 1,
                '2台车-动画仿真': 2,
                '4台车-绘制轨迹': 3,
                '4台车-动画仿真': 4
            }
        for pathtype, pathtype_num in pathtypes.items():
            self.a_pathtype = ttk.Radiobutton(self.menu_frame2, text=pathtype, value=pathtype_num,
                                              variable=self.pathtype_value, command=self._next_enabled)
            self.a_pathtype.grid(column=(pathtype_num - 1) % 4, row=(pathtype_num - 1) // 4, sticky=tk.W)
        self.pathtype_value.set(0)

        for child in self.menu_frame2.winfo_children():
            child.grid_configure(padx=8, pady=4)

    def page_pendulum(self):
        """选择摆杆程序，生成对应初始化GUI"""
        self.main_frame = ttk.LabelFrame(self.win, text='2.请输入参数:')
        self.main_frame.grid(column=0, row=0)
        # Add Frame1_图块选择==================================
        pendulum_frame1 = ttk.LabelFrame(self.main_frame, text='图块选择')
        pendulum_frame1.grid(column=0, row=1, sticky='WNS', padx=8, pady=4)

        # Add a Label_选择轨迹线
        ttk.Label(pendulum_frame1, text='轨迹线及起点: ').grid(column=0, row=0)
        # Add a Button_选择轨迹线
        self.b_choose_chainpath = ttk.Button(pendulum_frame1, text='单击选择', style='R.TButton',
                                             command=lambda: self.callbacks.click_pline('chainpath'))
        self.b_choose_chainpath.grid(column=1, row=0)

        if self.pathtype_value.get() in [3, 4]:
            # Add a Label_选择轨迹线
            ttk.Label(pendulum_frame1, text='摆杆滚轮所在轨迹线: ').grid(column=0, row=1)
            # Add a Button_选择轨迹线
            self.b_choose_rollerpath = ttk.Button(pendulum_frame1, text='单击选择', style='R.TButton',
                                                  command=lambda: self.callbacks.click_pline('rollerpath'))
            self.b_choose_rollerpath.grid(column=1, row=1)

        if self.pathtype_value.get() == 5:
            # Add a Label_选择轨迹中心点
            ttk.Label(pendulum_frame1, text='轨迹中心点: ').grid(column=0, row=1)
            # Add a Button_选择轨迹中心点
            self.b_choose_diptank_midpnt = ttk.Button(pendulum_frame1, text='单击选择', style='R.TButton',
                                                      command=lambda: self.callbacks.click_point('diptank_midpnt'))
            self.b_choose_diptank_midpnt.grid(column=1, row=1)

        # Add a Label_选择工件方向
        l_dir = ttk.Label(pendulum_frame1, text='工件/摆杆块方向: ')
        l_dir.grid(column=0, row=2)
        # Add a Tooltip_提示框
        Tt.create_tooltip(l_dir, '工件块、链板块与摆杆块方向需一致')
        # Add Radiobutton_选择工件方向
        self.dirvalue = tk.IntVar()
        self.dirvalue.set(2)
        cardirs = {'右': 1, '左': 2}
        for cardir, cardir_num in cardirs.items():
            rb_dir = ttk.Radiobutton(pendulum_frame1, text=cardir, value=cardir_num, variable=self.dirvalue)
            rb_dir.grid(column=cardir_num, row=2)

        if self.pathtype_value.get() in [1, 2, 6]:
            # Add a Label_选择摆杆状态
            ttk.Label(pendulum_frame1, text='摆杆状态: ').grid(column=0, row=3)
            # Add Radiobutton_选择摆杆状态
            self.swingstate_value = tk.IntVar()
            self.swingstate_value.set(2)
            swingstates = {'前摆杆竖直': 1, '后摆杆竖直': 2}
            for swingstate, swingstate_num in swingstates.items():
                rb_swingstate = ttk.Radiobutton(pendulum_frame1, text=swingstate, value=swingstate_num,
                                                variable=self.swingstate_value)
                rb_swingstate.grid(column=swingstate_num, row=3)

        if self.pathtype_value.get() == 1:
            # Add a Label_是否绘制包络线
            c_envelope_value = tk.IntVar()
            c_envelope = tk.Checkbutton(pendulum_frame1, text='提取包络线: ', variable=c_envelope_value,
                                        command=lambda: self.callbacks.envolope(c_envelope_value.get()))
            c_envelope.grid(column=0, row=4)
            c_envelope.deselect()
            # Add a Button_选择上包络线特征点
            self.b_choose_uenvelope = ttk.Button(pendulum_frame1, text='选择上特征点', state='disabled',
                                                 style='R.TButton',
                                                 command=lambda: self.callbacks.click_point('uenvelope'))
            self.b_choose_uenvelope.grid(column=1, row=4)
            # Add a Button_选择下包络线特征点
            self.b_choose_lenvelope = ttk.Button(pendulum_frame1, text='选择下特征点', state='disabled',
                                                 style='R.TButton',
                                                 command=lambda: self.callbacks.click_point('lenvelope'))
            self.b_choose_lenvelope.grid(column=2, row=4)
            c_envelope.configure(state='disabled')
            c_envelope.deselect()

        if self.pathtype_value.get() in [3, 4]:
            # Add a Label_选择摆杆塔
            ttk.Label(pendulum_frame1, text='摆杆塔: ').grid(column=0, row=5)
            # Add Radiobutton_选择摆杆塔
            self.tower_value = tk.IntVar()
            self.tower_value.set(1)
            towers = {'入口塔': 1, '出口塔': 2}
            for tower, tower_num in towers.items():
                rb_tower = ttk.Radiobutton(pendulum_frame1, text=tower, value=tower_num, variable=self.tower_value)
                rb_tower.grid(column=tower_num, row=5)

        for child in pendulum_frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2_基本参数==================================
        pendulum_frame2 = ttk.LabelFrame(self.main_frame, text='基本参数')
        pendulum_frame2.grid(column=1, row=1, sticky='WNS', padx=8, pady=4)

        # Add a Label_链板节距
        ttk.Label(pendulum_frame2, text='链板节距(mm): ').grid(column=0, row=0)
        # Add an Entry_链板节距
        self.chainbracing = tk.DoubleVar()
        ttk.Entry(pendulum_frame2, width=12, textvariable=self.chainbracing).grid(column=1, row=0)
        self.chainbracing.set(250)

        # Add a Label_摆杆间距
        l_bracing = ttk.Label(pendulum_frame2, text='摆杆间距(mm): ')
        l_bracing.grid(column=0, row=1)
        self.bracing = tk.DoubleVar()
        # Add an Entry_摆杆间距
        e_bracing = ttk.Entry(pendulum_frame2, width=12, textvariable=self.bracing)
        e_bracing.grid(column=1, row=1)
        self.bracing.set(3500)

        # Add a Label_摆杆长度
        l_swingleng = tk.Label(pendulum_frame2, text='摆杆长度(mm): ')
        l_swingleng.grid(column=0, row=2)
        # Add a Tooltip_提示框
        Tt.create_tooltip(l_swingleng, '套筒中心至摆杆底部圆管中心距离')
        # Add an Entry_摆杆长度
        self.swingleng = tk.DoubleVar()
        ttk.Entry(pendulum_frame2, width=12, textvariable=self.swingleng).grid(column=1, row=2)
        self.swingleng.set(3000)

        # Add a Label_轨迹步长
        l_step = ttk.Label(pendulum_frame2, text='轨迹步长(mm): ')
        l_step.grid(column=0, row=4)
        # Add an Entry_轨迹步长
        self.step = tk.IntVar()
        e_step = ttk.Entry(pendulum_frame2, width=12, textvariable=self.step)
        e_step.grid(column=1, row=4)
        self.step.set(500)

        # Add a Label_工件数量
        l_carnum = ttk.Label(pendulum_frame2, text='工件数量: ')
        l_carnum.grid(column=0, row=5)
        # Add an Entry_工件数量
        self.carnum = tk.IntVar()
        self.carnum.set(1)
        e_carnum = ttk.Entry(pendulum_frame2, width=12, textvariable=self.carnum)
        e_carnum.grid(column=1, row=5)

        # Add a Label_工件节距
        l_pitch = ttk.Label(pendulum_frame2, text='工件节距(mm): ')
        l_pitch.grid(column=0, row=6)
        # Add an Entry_工件节距
        self.pitch = tk.DoubleVar()
        e_pitch = ttk.Entry(pendulum_frame2, width=12, textvariable=self.pitch)
        e_pitch.grid(column=1, row=6)
        self.pitch.set(7000)

        if self.pathtype_value.get() in [1, 3]:  # 若为轨迹，禁用工件数量、工件节距
            e_carnum.configure(state='disabled')
            e_pitch.configure(state='disabled')
            if self.pathtype_value.get() == 3:  # 若为返回段轨迹，设定摆杆间距并禁用
                e_bracing.configure(state='disabled')
                self.bracing.set(0)
        elif self.pathtype_value.get() in [2, 4]:  # 若为动画，工件数量默认为2
            self.carnum.set(2)
        elif self.pathtype_value.get() in [5, 6]:  # 若为浸入即出槽、工艺时间分析，禁用步长、工件数量
            e_step.configure(state='disabled')
            e_carnum.configure(state='disabled')
            if self.pathtype_value.get() == 5:  # 若为浸入即出槽，设定工件数量
                self.carnum.set(2)
            else:  # 若为工艺时间，禁用工件节距
                e_pitch.configure(state='disabled')

        for child in pendulum_frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # 其他功能：浸入即出，工艺时间
        if self.pathtype_value.get() == 5:
            # Add Frame3_浸入即出==============================
            pendulum_frame3 = ttk.LabelFrame(self.main_frame, text='浸入即出槽状态分析')
            pendulum_frame3.grid(column=0, row=3, sticky='WE', columnspan=2, padx=8, pady=4)

            ttk.Label(pendulum_frame3, text='选择分析模式: ').grid(column=0, row=0)
            self.swingmode_value = tk.IntVar()
            self.swingmode_value.set(1)
            swingmodes = {'内侧摆杆竖直': 1, '外侧摆杆竖直': 2}
            for swingmode, swingmode_num in swingmodes.items():
                rb_swingmode = ttk.Radiobutton(pendulum_frame3, text=swingmode, value=swingmode_num,
                                               variable=self.swingmode_value)
                rb_swingmode.grid(column=swingmode_num, row=0)

            for child in pendulum_frame3.winfo_children():
                child.grid_configure(sticky=tk.W, padx=8, pady=4)
        elif self.pathtype_value.get() == 6:
            self.swingmode_value = 1
            # Add Frame4_工艺时间==============================
            pendulum_frame4 = ttk.LabelFrame(self.main_frame, text='工艺时间计算')
            pendulum_frame4.grid(column=0, row=4, sticky='WE', columnspan=2, padx=8, pady=4)

            # Add a Label_液面高度
            ttk.Label(pendulum_frame4, text='液面距轨道中心(mm): ').grid(column=0, row=0)
            # Add an Entry_液面高度
            self.liquidheight = tk.DoubleVar()
            ttk.Entry(pendulum_frame4, width=12, textvariable=self.liquidheight).grid(column=1, row=0)
            self.liquidheight.set(550)

            # Add a Label_链速
            ttk.Label(pendulum_frame4, text='链速(m/min): ').grid(column=2, row=0)
            # Add an Entry_链速
            self.chainspeed = tk.DoubleVar()
            ttk.Entry(pendulum_frame4, width=12, textvariable=self.chainspeed).grid(column=3, row=0)
            self.chainspeed.set(8)

            # Add a Label_选择全浸点
            ttk.Label(pendulum_frame4, text='工件全浸点: ').grid(column=0, row=1)
            # Add a Button_选择全浸点
            self.b_choose_immersion = ttk.Button(pendulum_frame4, text='单击选择', style='R.TButton',
                                                 command=lambda: self.callbacks.click_point('immersion'))
            self.b_choose_immersion.grid(column=1, row=1)

            # Add a Button_删除块属性
            self.b_del_block_attributes = ttk.Button(pendulum_frame4, text='删除块属性',
                                                     command=lambda: self.callbacks.delete_block_attributes())
            self.b_del_block_attributes.grid(column=2, row=1)
            # Add a Button_删除块参照属性
            self.b_del_blockreference_attributes = ttk.Button(pendulum_frame4, text='删除块参照属性',
                                                              command=lambda: self.callbacks.delete_blockreference_attributes())
            self.b_del_blockreference_attributes.grid(column=3, row=1)

            for child in pendulum_frame4.winfo_children():
                child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame5_确认与退出====================================
        pendulum_frame5 = ttk.LabelFrame(self.main_frame, text='')
        pendulum_frame5.grid(column=0, row=5, columnspan=2)
        self._progressbar(pendulum_frame5)
        self.b_previous = ttk.Button(pendulum_frame5, text='上一步', command=self.callbacks.previousbutton)
        self.b_previous.grid(column=0, row=1, padx=20, pady=8)
        if self.pathtype_value.get() in [1, 2]:  # 工艺段
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_process(self.pathtype_value.get(),
                                                                                      self.dirvalue.get(),
                                                                                      self.swingstate_value.get(),
                                                                                      self.callbacks.chainpath,
                                                                                      self.step.get(),
                                                                                      self.chainbracing.get(),
                                                                                      self.bracing.get(),
                                                                                      self.carnum.get(),
                                                                                      self.pitch.get(),
                                                                                      self.swingleng.get() - 252.75))
        elif self.pathtype_value.get() in [3, 4]:  # 返回段
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_return(self.pathtype_value.get(),
                                                                                     self.dirvalue.get(),
                                                                                     self.tower_value.get(),
                                                                                     self.callbacks.chainpath,
                                                                                     self.callbacks.rollerpath,
                                                                                     self.step.get(),
                                                                                     self.chainbracing.get(),
                                                                                     self.bracing.get(),
                                                                                     self.carnum.get(),
                                                                                     self.pitch.get(),
                                                                                     self.swingleng.get()))
        elif self.pathtype_value.get() == 5:  # 浸入即出槽
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_diptank(self.dirvalue.get(),
                                                                                      self.swingmode_value.get(),
                                                                                      self.callbacks.chainpath,
                                                                                      self.callbacks.diptank_midpnt,
                                                                                      self.chainbracing.get(),
                                                                                      self.bracing.get(),
                                                                                      self.pitch.get(),
                                                                                      self.swingleng.get() - 252.75))
        elif self.pathtype_value.get() == 6:  # 工艺时间
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_immersiontime(self.dirvalue.get(),
                                                                                            self.swingstate_value.get(),
                                                                                            self.callbacks.chainpath,
                                                                                            self.chainbracing.get(),
                                                                                            self.bracing.get(),
                                                                                            self.swingleng.get() - 252.75,
                                                                                            self.liquidheight.get(),
                                                                                            self.chainspeed.get() * 1000 / 60))
        self.b_entry.grid(column=1, row=1, padx=20, pady=8)
        self.b_quit = ttk.Button(pendulum_frame5, text='退出', command=self.callbacks.quit)
        self.b_quit.grid(column=2, row=1, padx=20, pady=8)

        # Add Menu==============================================
        self._menubar()

    def page_shuttle(self):
        """选择翻转机程序，生成对应初始化GUI"""
        self.main_frame = ttk.LabelFrame(self.win, text='2.请输入参数:')
        self.main_frame.grid(column=0, row=0)
        # Add Frame5_退出====================================
        pendulum_frame5 = ttk.LabelFrame(self.main_frame, text='')
        pendulum_frame5.grid(column=0, row=5, columnspan=2)
        self._progressbar(pendulum_frame5)
        self.b_previous = ttk.Button(pendulum_frame5, text='上一步', command=self.callbacks.previousbutton)
        self.b_previous.grid(column=0, row=1, padx=20, pady=8)
        if self.pathtype_value.get() <= 4:
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_process(self.pathtype_value.get(),
                                                                                      self.dirvalue.get(),
                                                                                      self.swingstate_value.get(),
                                                                                      self.callbacks.chainpath,
                                                                                      self.step.get(),
                                                                                      self.chainbracing.get(),
                                                                                      self.bracing.get(),
                                                                                      self.carnum.get(),
                                                                                      self.pitch.get(),
                                                                                      self.swingleng.get()) - 252.75)
        elif self.pathtype_value.get() == 5:
            self.b_entry = ttk.Button(pendulum_frame5, text='确定',
                                      command=lambda: self.doPath.do_pendulum_diptank(self.dirvalue.get(),
                                                                                      self.swingmode_value.get(),
                                                                                      self.callbacks.chainpath,
                                                                                      self.callbacks.diptank_midpnt,
                                                                                      self.chainbracing.get(),
                                                                                      self.bracing.get(),
                                                                                      self.pitch.get(),
                                                                                      self.swingleng.get()) - 252.75)
        self.b_entry.grid(column=1, row=1, padx=20, pady=8)
        self.b_quit = ttk.Button(pendulum_frame5, text='退出', command=self.callbacks.quit)
        self.b_quit.grid(column=2, row=1, padx=20, pady=8)

        # Add Menu==============================================
        self._menubar()

    def page_trolley2(self):
        """选择2台车程序，生成对应初始化GUI"""
        self.main_frame = ttk.LabelFrame(self.win, text='2.请输入参数:')
        self.main_frame.grid(column=0, row=0)
        # Add Frame1_图块选择==================================
        trolley2_frame1 = ttk.LabelFrame(self.main_frame, text='图块选择')
        trolley2_frame1.grid(column=0, row=0, sticky='WNS', padx=8, pady=4)

        # Add a Label_选择工件
        ttk.Label(trolley2_frame1, text='工件: ').grid(column=0, row=0)
        # Add a Button_选择工件
        self.b_choose_car = ttk.Button(trolley2_frame1, text='单击选择', style='R.TButton',
                                       command=lambda: self.callbacks.click_block('car'))
        self.b_choose_car.grid(column=1, row=0)

        # Add a Label_选择链条中心线
        ttk.Label(trolley2_frame1, text='链条中心线及起点: ').grid(column=0, row=1)
        # Add a Button_选择链条中心线
        self.b_choose_chainpath = ttk.Button(trolley2_frame1, text='单击选择', style='R.TButton',
                                             command=lambda: self.callbacks.click_pline('chainpath'))
        self.b_choose_chainpath.grid(column=1, row=1)

        # Add a Label_选择铰接点中心线
        ttk.Label(trolley2_frame1, text='铰接点中心线及起点: ').grid(column=0, row=2)
        # Add a Button_选择铰接点中心线
        self.b_choose_hinglepath = ttk.Button(trolley2_frame1, text='单击选择', style='R.TButton',
                                              command=lambda: self.callbacks.click_pline('hinglepath'))
        self.b_choose_hinglepath.grid(column=1, row=2)
        # Add a Tooltip_提示框
        Tt.create_tooltip(self.b_choose_hinglepath, '链条与铰接点中心线起点需为同一方向')

        # Add a Label_选择工件方向
        ttk.Label(trolley2_frame1, text='工件块方向: ').grid(column=0, row=3)
        # Add a Radiobutton_选择工件方向
        self.dirvalue = tk.IntVar()
        self.dirvalue.set(1)
        dirs = {'右': 1, '左': 2}
        for cardir, cardir_num in dirs.items():
            a31 = ttk.Radiobutton(trolley2_frame1, text=cardir, value=cardir_num, variable=self.dirvalue)
            a31.grid(column=cardir_num, row=3)

        for child in trolley2_frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2_基本参数==================================
        trolley2_frame2 = ttk.LabelFrame(self.main_frame, text='基本参数')
        trolley2_frame2.grid(column=1, row=0, sticky='WNS', padx=8, pady=4)

        # Add a Label_工件前后支撑距离
        ttk.Label(trolley2_frame2, text='工件前后支撑距离(mm): ').grid(column=0, row=0)
        # Add a Entry_工件前后支撑距离
        self.bracing = tk.DoubleVar()
        ttk.Entry(trolley2_frame2, width=12, textvariable=self.bracing).grid(column=1, row=0)
        self.bracing.set(2150)

        # Add a Label_轨迹步长
        ttk.Label(trolley2_frame2, text='轨迹步长(mm): ').grid(column=0, row=1)
        # Add an Entry_轨迹步长
        self.step = tk.DoubleVar()
        ttk.Entry(trolley2_frame2, width=12, textvariable=self.step).grid(column=1, row=1)
        self.step.set(200)

        # Add a Label_工件数量
        l_carnum = ttk.Label(trolley2_frame2, text='工件数量: ')
        l_carnum.grid(column=0, row=2)
        # Add an Entry_工件数量
        self.carnum = tk.IntVar()
        e_carnum = ttk.Entry(trolley2_frame2, width=12, textvariable=self.carnum)
        e_carnum.grid(column=1, row=2)
        # Add a Label_工件节距
        l_pitch = ttk.Label(trolley2_frame2, text='工件节距(mm): ')
        l_pitch.grid(column=0, row=3)
        # Add an Entry_工件节距
        self.pitch = tk.DoubleVar()
        e_pitch = ttk.Entry(trolley2_frame2, width=12, textvariable=self.pitch)
        e_pitch.grid(column=1, row=3)
        # 若为绘制轨迹，则禁用工件数量、节距选项
        if self.pathtype_value.get() == 1:
            l_carnum.configure(state='disabled')
            e_carnum.configure(state='disabled')
            self.carnum.set(1)
            l_pitch.configure(state='disabled')
            e_pitch.configure(state='disabled')
            self.pitch.set(0)

        for child in trolley2_frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame3_开始与退出========================================
        trolley2_frame3 = ttk.LabelFrame(self.main_frame, text='')
        trolley2_frame3.grid(column=0, row=1, columnspan=2)
        self._progressbar(trolley2_frame3)
        self.b_previous = ttk.Button(trolley2_frame3, text='上一步', command=self.callbacks.previousbutton)
        self.b_previous.grid(column=0, row=1, padx=20, pady=8)
        self.b_entry = ttk.Button(trolley2_frame3, text='确定',
                                  command=lambda: self.doPath.do_2trolley(self.pathtype_value.get(),
                                                                          self.dirvalue.get(), self.callbacks.car_name,
                                                                          self.callbacks.chainpath,
                                                                          self.callbacks.hinglepath, self.step.get(),
                                                                          self.bracing.get(), self.carnum.get(),
                                                                          self.pitch.get()))
        self.b_entry.grid(column=1, row=1, padx=20, pady=8)
        self.b_quit = ttk.Button(trolley2_frame3, text='退出', command=self.callbacks.quit)
        self.b_quit.grid(column=2, row=1, padx=20, pady=8)

        # Add Menu==============================================
        self._menubar()

    def page_trolley4(self):
        """选择4台车程序，生成对应初始化GUI"""
        self.main_frame = ttk.LabelFrame(self.win, text='2.请输入参数:')
        self.main_frame.grid(column=0, row=0)
        # Add Frame3_开始与退出========================================
        trolley2_frame3 = ttk.LabelFrame(self.main_frame, text='')
        trolley2_frame3.grid(column=0, row=1, columnspan=2)
        self._progressbar(trolley2_frame3)
        self.b_previous = ttk.Button(trolley2_frame3, text='上一步', command=self.callbacks.previousbutton)
        self.b_previous.grid(column=0, row=1, padx=20, pady=8)
        self.b_entry = ttk.Button(trolley2_frame3, text='确定')
        self.b_entry.grid(column=1, row=1, padx=20, pady=8)
        self.b_quit = ttk.Button(trolley2_frame3, text='退出', command=self.callbacks.quit)
        self.b_quit.grid(column=2, row=1, padx=20, pady=8)

        # Add Menu==============================================
        self._menubar()

    def _menubar(self):
        """添加菜单栏"""
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

    def _next_enabled(self):
        """当工件类型、轨迹类型都选择后，'下一步'按钮变为可用"""
        if (self.cartype_value.get() == 0) or (self.pathtype_value.get() == 0):
            self.b_next.configure(state='disabled')
        else:
            self.b_next.configure(state='normal')

    def _progressbar(self, frame):
        """初始化进度条"""
        self.pb = ttk.Progressbar(frame, length=210, mode='determinate', value=0)
        self.pb.grid(column=0, row=0, columnspan=2, padx=20, pady=8)
        self.pb_value = tk.StringVar()
        self.pb_value.set('0%')
        self.l_pb = ttk.Label(frame, textvariable=self.pb_value)
        self.l_pb.grid(column=2, row=0, padx=20, pady=8)

    def update_progressbar(self, process_value, finish=0):
        """更新进度条"""
        self.pb['value'] = process_value
        if finish:
            self.pb_value.set('完成')
        else:
            self.pb_value.set(f"{process_value}%")
        self.win.update()
        time.sleep(0.1)


# =========================================================
if __name__ == '__main__':
    oop = OOP()
    # oop.win.iconbitmap('car.ico')
    oop.win.mainloop()
