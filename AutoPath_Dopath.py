#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
cadpath_20210817重构使用class，合并轨迹与仿真选项卡
"""

import math
import tkinter as tk
from tkinter import ttk, messagebox as msg
import time
from AutoPath_Typeconvert import Typeconvert


class Dopath(object):
    def __init__(self, oop, tab2, doc, msp):
        self.oop = oop
        self.tab2 = tab2
        self.doc = doc
        self.msp = msp
        self.typeConvert = Typeconvert()

    # 选择工件
    def click_block(self):
        self.carbody = self.get_block()

    def get_block(self):
        block = self.doc.Utility.GetEntity()  # 在cad中选取块
        time.sleep(0.1)
        while 'Block' not in block[0].ObjectName:  # 判断选取的是否为块
            msg.showerror('错误', '您选择的不是块，\n请重新选择！')
            block = self.doc.Utility.GetEntity()
            time.sleep(0.1)
        self.carselect.configure(text='已选择工件', foreground='green')
        return block[0]

    # 选择轨迹线
    def click_pline(self):
        self.track, self.track_vertex, self.track_rev = self.get_pline()  # 获取轨道线、顶点数

    def get_pline(self):
        pline = self.doc.Utility.GetEntity()  # 在cad中选择多段线
        while 'Polyline' not in pline[0].ObjectName:  # 判断选取的是否为多段线
            msg.showerror('错误', '您选择的不是多段线，\n请重新选择！')
            pline = self.doc.Utility.GetEntity()
        pline_pnt = list(pline[0].Coordinates[:])  # 获取多段线各顶点坐标
        pline_vertex = len(pline_pnt) // 2
        pllayer, pltype, plcw = pline[0].Layer, pline[0].Linetype, pline[0].ConstantWidth
        # 指定并调整轨迹线方向
        self.doc.Utility.Prompt("请指定多段线的起点，在相应端点上单击")
        time.sleep(0.1)
        start_pnt = list(self.doc.Utility.GetPoint())
        while start_pnt[:2] != pline_pnt[:2] and start_pnt[:2] != pline_pnt[-2:]:
            msg.showerror('错误', '未选择多段线的端点，请重新选择！')
            start_pnt = list(self.doc.Utility.GetPoint())
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
            pline1 = self.msp.AddLightWeightPolyline(self.typeConvert.vtfloat(pline1_Pnt))
            for i in range(pline_vertex):
                pline1.SetBulge(i, pline1_bulge[i])
            pline1.Layer, pline1.Linetype, pline1.ConstantWidth = pllayer, pltype, plcw
            pline1.Update()
            pline[0].Delete()
        else:
            rev2 = 1  # 指定起点相同
            pline1 = pline[0]
        rev = rev1 * rev2
        self.pathselect.configure(text='已选择轨迹线', foreground='green')
        return pline1, pline_vertex, rev

    # 轨迹线各段长度
    def pline_length(self, pline1):
        leng = []
        pline2 = pline1.Explode()
        for i in pline2:
            if 'Arc' in i.ObjectName:
                leng.append(i.ArcLength)
            else:
                leng.append(i.Length)
            i.Delete()
        return leng

    # 求插入点
    def pathpnt(self, pline, step):
        try:
            self.doc.SelectionSets.Item("SS1").Delete()
        except:
            print("Delete selection failed")
        slt = self.doc.SelectionSets.Add("SS1")
        LayerObj = self.doc.Layers.Add("Path_pnt")
        self.doc.ActiveLayer = LayerObj
        LayerObj.LayerOn = False
        handle_str = '_measure'
        handle_str += ' (handent "' + pline.Handle + '")'
        handle_str += '\n'
        step1 = str(step.get())
        handle_str += step1
        handle_str += '\n'
        self.doc.SendCommand(handle_str)
        filterType = [8]  # 定义过滤类型
        filterData = ["Path_pnt"]  # 设置过滤参数
        filterType = self.typeConvert.vtint(filterType)
        filterData = self.typeConvert.vtvariant(filterData)
        slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
        steppnt = [1] * len(slt)
        for i in range(len(slt)):
            steppnt[i] = slt[len(slt) - i - 1].Coordinates[:2]
        # print(steppnt)
        slt.Erase()
        slt.Delete()
        self.doc.ActiveLayer = self.doc.Layers("0")
        LayerObj.Delete()
        return steppnt

    # 根据距离求点
    def getdistpt(self, pline, leng, nowdist, frtpnt, dist):
        sndpnt = [0, 0]
        for i in leng:
            if i >= (nowdist - dist):
                vertex = leng.index(i)
                bulge = pline.GetBulge(vertex)
                if vertex == 0:
                    preleng = 0
                else:
                    preleng = leng[vertex - 1]
                if bulge == 0:
                    if i >= nowdist:
                        sndpnt[0] = frtpnt[0] - dist * (frtpnt[0] - pline.Coordinates[vertex * 2]) / (
                                nowdist - preleng)
                        sndpnt[1] = frtpnt[1] - dist * (frtpnt[1] - pline.Coordinates[vertex * 2 + 1]) / (
                                nowdist - preleng)
                    else:
                        sndpnt[0] = pline.Coordinates[(vertex + 1) * 2] - (dist - nowdist + i) * (
                                pline.Coordinates[(vertex + 1) * 2] - pline.Coordinates[vertex * 2]) / (i - preleng)
                        sndpnt[1] = pline.Coordinates[(vertex + 1) * 2 + 1] - (dist - nowdist + i) * (
                                pline.Coordinates[(vertex + 1) * 2 + 1] - pline.Coordinates[vertex * 2 + 1]) / (
                                            i - preleng)
                else:
                    pline1 = pline.Explode()
                    centpnt = pline1[vertex].Center  # 圆弧圆心坐标
                    centpnt = self.typeConvert.vtpnt(centpnt[0], centpnt[1])
                    startpnt = self.typeConvert.vtpnt(pline.Coordinates[vertex * 2],
                                                      pline.Coordinates[vertex * 2 + 1])
                    ray = self.msp.AddRay(centpnt, startpnt)
                    arcang = pline1[vertex].TotalAngle  # 圆弧弧度
                    roang = (nowdist - preleng - dist) * arcang / (i - preleng)
                    if pline1[vertex].StartAngle < pline1[vertex].EndAngle:  # 逆时针
                        if (pline1[vertex].StartPoint[0] != pline.Coordinates[vertex * 2]) and (
                                pline1[vertex].StartPoint[1] != pline.Coordinates[vertex * 2 + 1]):
                            roang = -roang
                    ray.Rotate(centpnt, roang)
                    sndpnt = pline1[vertex].IntersectWith(ray, 0)
                    for j in pline1:
                        j.Delete()
                    ray.Delete()
                break
        return sndpnt

    ##############################
    # Do '摆杆'
    ##############################
    def donext1(self, select):
        # Add Frame1_图块选择================================================================
        self.tab2_Frame1 = ttk.LabelFrame(self.tab2, text='图块选择')
        self.tab2_Frame1.grid(column=0, row=0, sticky='WN', padx=8, pady=4)

        # Add a Label
        ttk.Label(self.tab2_Frame1, text='选择前摆杆: ').grid(column=0, row=0)
        # Add a Label
        self.carselect = ttk.Label(self.tab2_Frame1, text='未选择前摆杆', width=12)
        self.carselect.grid(column=1, row=0)
        self.carselect.configure(foreground='red')
        # Add a Button
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_block).grid(column=2, row=0)

        # Add a Label
        ttk.Label(self.tab2_Frame1, text='选择后摆杆: ').grid(column=0, row=1)
        # Add a Label
        self.carselect = ttk.Label(self.tab2_Frame1, text='未选择后摆杆', width=12)
        self.carselect.grid(column=1, row=1)
        self.carselect.configure(foreground='red')
        # Add a Button
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_block).grid(column=2, row=1)

        # Add a Label
        ttk.Label(self.tab2_Frame1, text='选择工件: ').grid(column=0, row=2)
        # Add a Label
        self.carselect = ttk.Label(self.tab2_Frame1, text='未选择工件', width=12)
        self.carselect.grid(column=1, row=2)
        self.carselect.configure(foreground='red')
        # Add a Button
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_block).grid(column=2, row=2)

        ttk.Label(self.tab2_Frame1, text='选择轨迹线并指定起点: ').grid(column=0, row=3)
        self.pathselect = ttk.Label(self.tab2_Frame1, text='未选择轨迹线', width=12)
        self.pathselect.grid(column=1, row=3)
        self.pathselect.configure(foreground='red')
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_pline).grid(column=2, row=3)

        ttk.Label(self.tab2_Frame1, text='选择工件及摆杆块方向: ').grid(column=0, row=4)
        self.cardir_value = tk.IntVar()
        self.cardir_value.set(1)
        cardir = [('右', 1), ('左', 2)]
        for i, j in cardir:
            a31 = ttk.Radiobutton(self.tab2_Frame1, text=i, value=j, variable=self.cardir_value)
            a31.grid(column=j, row=4)

        for child in self.tab2_Frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2_基本参数================================================================
        self.tab2_Frame2 = ttk.LabelFrame(self.tab2, text='基本参数')
        self.tab2_Frame2.grid(column=1, row=0, sticky='WN', padx=8, pady=4)

        ttk.Label(self.tab2_Frame2, text='链板节距(mm): ').grid(column=0, row=0)
        self.chainbracing = tk.StringVar()
        ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.chainbracing).grid(column=1, row=0)

        ttk.Label(self.tab2_Frame2, text='前后摆杆间距(mm): ').grid(column=0, row=1)
        self.bracing = tk.StringVar()
        ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.bracing).grid(column=1, row=1)

        ttk.Label(self.tab2_Frame2, text='轨迹步长(mm): ').grid(column=0, row=2)
        self.step = tk.StringVar()
        ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.step).grid(column=1, row=2)

        # 若为仿真动画，则启用工件数量、节距选项
        l31 = ttk.Label(self.tab2_Frame2, text='工件数量: ')
        l31.grid(column=0, row=3)
        self.carnum = tk.StringVar()
        e31 = ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.carnum)
        e31.grid(column=1, row=3)
        l32 = ttk.Label(self.tab2_Frame2, text='工件节距(mm): ')
        l32.grid(column=0, row=4)
        self.pitch = tk.StringVar()
        e32 = ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.pitch)
        e32.grid(column=1, row=4)
        if select % 10 == 1:
            l31.configure(state='disabled')
            e31.configure(state='disabled')
            l32.configure(state='disabled')
            e32.configure(state='disabled')

        for child in self.tab2_Frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame3_浸入即出================================================================
        self.tab2_Frame3 = ttk.LabelFrame(self.tab2, text='浸入即出状态分析')
        self.tab2_Frame3.grid(column=0, row=1, sticky='WN', columnspan=2, padx=8, pady=4)

        ttk.Label(self.tab2_Frame3, text='选择摆杆状态: ').grid(column=0, row=0)
        self.swingstate_value = tk.IntVar()
        self.swingstate_value.set(1)
        swingstate = [('前摆杆竖直', 1), ('后摆杆竖直', 2)]
        for i, j in swingstate:
            a32 = ttk.Radiobutton(self.tab2_Frame3, text=i, value=j, variable=self.swingstate_value)
            a32.grid(column=j, row=0)

        ttk.Button(self.tab2_Frame3, text='分析', command=lambda: self.do_diptank(self.swingstate_value.get())).grid(
            column=3, row=0)

        for child in self.tab2_Frame3.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame4=====================================================================
        tab2_Button = ttk.LabelFrame(self.tab2, text='')
        tab2_Button.grid(column=0, row=2, columnspan=2)
        self.b21 = ttk.Button(tab2_Button, text='确定', command=lambda: self.do_path3(select))
        self.b21.grid(column=0, row=0, padx=20, pady=8)
        self.b22 = ttk.Button(tab2_Button, text='退出', command=self.oop.quit)
        self.b22.grid(column=1, row=0, padx=20, pady=8)

    # 浸入即出槽分析
    def do_diptank(self, swing):
        pass

    ##############################
    # Do '台车'
    ##############################
    def donext3(self, select):
        # Add Frame1=====================================================================
        self.tab2_Frame1 = ttk.LabelFrame(self.tab2, text='图块选择')
        self.tab2_Frame1.grid(column=0, row=0, sticky='WN', padx=8, pady=4)

        # Add a Label
        ttk.Label(self.tab2_Frame1, text='选择工件: ').grid(column=0, row=0)
        # Add a Label
        self.carselect = ttk.Label(self.tab2_Frame1, text='未选择工件', width=12)
        self.carselect.grid(column=1, row=0)
        self.carselect.configure(foreground='red')
        # Add a Button
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_block).grid(column=2, row=0)

        ttk.Label(self.tab2_Frame1, text='选择轨迹线并指定起点: ').grid(column=0, row=1)
        self.pathselect = ttk.Label(self.tab2_Frame1, text='未选择轨迹线', width=12)
        self.pathselect.grid(column=1, row=1)
        self.pathselect.configure(foreground='red')
        ttk.Button(self.tab2_Frame1, text='选择', command=self.click_pline).grid(column=2, row=1)

        ttk.Label(self.tab2_Frame1, text='选择工件块方向: ').grid(column=0, row=2)
        self.cardir_value = tk.IntVar()
        self.cardir_value.set(1)
        cardir = [('右', 1), ('左', 2)]
        for i, j in cardir:
            a31 = ttk.Radiobutton(self.tab2_Frame1, text=i, value=j, variable=self.cardir_value)
            a31.grid(column=j, row=2)

        for child in self.tab2_Frame1.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame2=====================================================================
        self.tab2_Frame2 = ttk.LabelFrame(self.tab2, text='基本参数')
        self.tab2_Frame2.grid(column=1, row=0, sticky='WN', padx=8, pady=4)

        ttk.Label(self.tab2_Frame2, text='工件前后支撑距离(mm): ').grid(column=0, row=0)
        self.bracing = tk.StringVar()
        ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.bracing).grid(column=1, row=0)

        ttk.Label(self.tab2_Frame2, text='轨迹步长(mm): ').grid(column=0, row=1)
        self.step = tk.StringVar()
        ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.step).grid(column=1, row=1)

        # 若为仿真动画，则启用工件数量、节距选项
        l31 = ttk.Label(self.tab2_Frame2, text='工件数量: ')
        l31.grid(column=0, row=2)
        self.carnum = tk.StringVar()
        e31 = ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.carnum)
        e31.grid(column=1, row=2)
        l32 = ttk.Label(self.tab2_Frame2, text='工件节距(mm): ')
        l32.grid(column=0, row=3)
        self.pitch = tk.StringVar()
        e32 = ttk.Entry(self.tab2_Frame2, width=12, textvariable=self.pitch)
        e32.grid(column=1, row=3)
        if select % 10 == 1:
            l31.configure(state='disabled')
            e31.configure(state='disabled')
            l32.configure(state='disabled')
            e32.configure(state='disabled')

        for child in self.tab2_Frame2.winfo_children():
            child.grid_configure(sticky=tk.W, padx=8, pady=4)

        # Add Frame3=====================================================================
        tab2_Button = ttk.LabelFrame(self.tab2, text='')
        tab2_Button.grid(column=0, row=1, columnspan=2)
        self.b21 = ttk.Button(tab2_Button, text='确定', command=lambda: self.do_path3(select))
        self.b21.grid(column=0, row=0, padx=20, pady=8)
        self.b22 = ttk.Button(tab2_Button, text='退出', command=self.oop.quit)
        self.b22.grid(column=1, row=0, padx=20, pady=8)

    # '台车' - 轨迹/动画
    def do_path3(self, select):
        self.steppnt = self.pathpnt(self.track, self.step)
        if select % 10 == 1:
            self.pathcarbody31(self.carbody, self.track, self.steppnt, self.step, self.bracing)
        else:
            self.pathcarbody32(self.carbody, self.track, self.steppnt, self.step, self.bracing, self.carnum,
                               self.pitch)

    # '台车' - '绘制轨迹'
    def pathcarbody31(self, block, pline, steppnt, step, bracing):
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [self.doc.Layers.Add(i) for i in layernames]  # 批量创建图层
        for lay in layerobjs:
            lay.LayerOn = False  # 关闭图层
        for i in range(len(layerobjs)):
            layerobjs[i].color = clrnums[i]
        leng = self.pline_length(pline)
        for i in range(1, len(leng)):
            leng[i] += leng[i - 1]
        j = 1
        jj = 0  # 判断是否为第一个插入点
        step0 = ((pline.Coordinates[0] - steppnt[0][0]) ** 2 + (pline.Coordinates[1] - steppnt[0][1]) ** 2) ** 0.5
        step1 = ((pline.Coordinates[-2] - steppnt[-1][0]) ** 2 + (
                pline.Coordinates[-1] - steppnt[-1][1]) ** 2) ** 0.5
        if step0 == 0:
            frtstep = 0
        elif step0 > step1:
            frtstep = float(step.get())
        else:
            frtstep = leng[-1] - len(steppnt) * float(step.get())
        # 插入文字说明
        inserttext = "轨迹线长度：" + str(leng[-1] // 1)[:-2] + "mm\n" + "轨迹步长：" + step.get() + "mm"
        textpnt = self.typeConvert.vtpnt((pline.Coordinates[0] + pline.Coordinates[-2]) / 2,
                                         max(pline.Coordinates[0], pline.Coordinates[1]) + 3000)
        mt = self.msp.AddMText(textpnt, 3000, inserttext)
        mt.height = 200
        for i in steppnt:
            if ((j - 1) * float(step.get()) + frtstep) > float(bracing.get()):
                # print((j - 1) * float(step.get()) + frtstep)
                layernum = (j - 1) % len(clrnums)
                self.doc.ActiveLayer = self.doc.Layers(layernames[layernum])
                inspnt = self.typeConvert.vtpnt(i[0], i[1])
                circlebr = self.msp.AddCircle(inspnt, float(bracing.get()))
                intpnts = pline.IntersectWith(circlebr, 0)
                angbr = []
                for k in range(len(intpnts) // 3):
                    brpnt = self.typeConvert.vtpnt(intpnts[k * 3], intpnts[k * 3 + 1])
                    angbr.append(self.doc.Utility.AngleFromXAxis(brpnt, inspnt))
                for k in range(len(leng)):
                    if ((j - 1) * float(step.get()) + frtstep) < leng[k]:
                        jj += 1
                        # print(leng[k])
                        plpnt = pline.Coordinates[(k * 2):(k * 2 + 2)]
                        plpnt = self.typeConvert.vtpnt(plpnt[0], plpnt[1])
                        angpl = self.doc.Utility.AngleFromXAxis(plpnt, inspnt)
                        if jj == 1:  # 判断是否为第一个插入点
                            if i[0] < pline.Coordinates[k * 2]:  # 起始向左
                                mirmark = 1
                            else:
                                mirmark = 0  # 起始向右
                        break
                angdif = []
                for k in range(len(angbr)):
                    angabs = abs(angbr[k] - angpl)
                    if angabs > math.pi:
                        angabs = math.pi * 2 - angabs
                    angdif.append(angabs)
                angmin = min(angdif)
                index1 = angdif.index(angmin)
                if self.cardir_value.get() == 1:  # 工件向右
                    a = angbr[index1]
                    car1 = self.msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                    if mirmark == 1:  # 起始向左
                        fstpnt = self.typeConvert.vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                        car1.Mirror(fstpnt, inspnt)
                        car1.Delete()
                else:  # 工件向左
                    if angbr[index1] < math.pi:
                        a = angbr[index1] + math.pi
                    else:
                        a = angbr[index1] - math.pi
                    car1 = self.msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                    if mirmark == 0:  # 起始向右
                        fstpnt = self.typeConvert.vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                        car1.Mirror(fstpnt, inspnt)
                        car1.Delete()
                circlebr.Delete()
            j += 1
        for lay in layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

    # '台车' - '仿真动画'
    def pathcarbody32(self, block, pline, steppnt, step, bracing, carnum, pitch):
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [self.doc.Layers.Add(i) for i in layernames]  # 批量创建图层
        for lay in layerobjs:
            lay.LayerOn = False  # 关闭图层
        for i in range(len(layerobjs)):
            layerobjs[i].color = clrnums[i]
        leng = self.pline_length(pline)
        for i in range(1, len(leng)):
            leng[i] += leng[i - 1]
        j = 1
        jj = 0  # 判断是否为第一个插入点
        step0 = ((pline.Coordinates[0] - steppnt[0][0]) ** 2 + (pline.Coordinates[1] - steppnt[0][1]) ** 2) ** 0.5
        step1 = ((pline.Coordinates[-2] - steppnt[-1][0]) ** 2 + (
                pline.Coordinates[-1] - steppnt[-1][1]) ** 2) ** 0.5
        if step0 == 0:
            frtstep = 0
        elif step0 > step1:
            frtstep = float(step.get())
        else:
            frtstep = leng[-1] - len(steppnt) * float(step.get())
        # 插入文字说明
        inserttext = "轨迹线长度：" + str(leng[-1] // 1)[:-2] + "mm\n" + "工件节距：" + pitch.get() + "mm"
        textpnt = self.typeConvert.vtpnt((pline.Coordinates[0] + pline.Coordinates[-2]) / 2,
                                         max(pline.Coordinates[0], pline.Coordinates[1]) + 3000)
        mt = self.msp.AddMText(textpnt, 3000, inserttext)
        mt.height = 200
        for i in steppnt:
            car = [1] * int(carnum.get())
            if ((j - 1) * float(step.get()) + frtstep) > (
                    (int(carnum.get()) - 1) * float(pitch.get()) + float(bracing.get())):
                # print((j - 1) * float(step.get()) + frtstep)
                for num in range(int(carnum.get())):
                    layernum = num % len(clrnums)
                    self.doc.ActiveLayer = self.doc.Layers(layernames[layernum])
                    if num == 0:
                        inspnt0 = i
                    else:
                        nowdist = (j - 1) * float(step.get()) + frtstep - (num - 1) * float(pitch.get())
                        inspnt0 = self.getdistpt(pline, leng, nowdist, inspnt0, float(pitch.get()))
                    inspnt = self.typeConvert.vtpnt(inspnt0[0], inspnt0[1])
                    circlebr = self.msp.AddCircle(inspnt, float(bracing.get()))
                    intpnts = pline.IntersectWith(circlebr, 0)
                    angbr = []
                    for k in range(len(intpnts) // 3):
                        brpnt = self.typeConvert.vtpnt(intpnts[k * 3], intpnts[k * 3 + 1])
                        angbr.append(self.doc.Utility.AngleFromXAxis(brpnt, inspnt))
                    for k in range(len(leng)):
                        if ((j - 1) * float(step.get()) + frtstep - num * float(pitch.get())) < leng[k]:
                            jj += 1
                            # print(leng[k])
                            plpnt = pline.Coordinates[(k * 2):(k * 2 + 2)]
                            plpnt = self.typeConvert.vtpnt(plpnt[0], plpnt[1])
                            angpl = self.doc.Utility.AngleFromXAxis(plpnt, inspnt)
                            if jj == 1:  # 判断是否为第一个插入点
                                if i[0] < pline.Coordinates[k * 2]:  # 起始向左
                                    mirmark = 1
                                else:
                                    mirmark = 0  # 起始向右
                            break
                    angdif = []
                    for k in range(len(angbr)):
                        angabs = abs(angbr[k] - angpl)
                        if angabs > math.pi:
                            angabs = math.pi * 2 - angabs
                        angdif.append(angabs)
                    angmin = min(angdif)
                    index1 = angdif.index(angmin)
                    if self.cardir_value.get() == 1:  # 工件向右
                        a = angbr[index1]
                        car[num] = self.msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                        if mirmark == 1:  # 起始向左
                            fstpnt = self.typeConvert.vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                            car1 = car[num].Mirror(fstpnt, inspnt)
                            car[num].Delete()
                            car[num] = car1
                    else:  # 工件向左
                        if angbr[index1] < math.pi:
                            a = angbr[index1] + math.pi
                        else:
                            a = angbr[index1] - math.pi
                        car[num] = self.msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                        if mirmark == 0:  # 起始向右
                            fstpnt = self.typeConvert.vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                            car1 = car[num].Mirror(fstpnt, inspnt)
                            car[num].Delete()
                            car[num] = car1
                    circlebr.Delete()
                for lay in layerobjs:
                    lay.LayerOn = True  # 打开图层
                # 选择是否保留当前结果
                self.doc.Utility.InitializeUserInput(0, "C, c")
                doconti = 'a'
                while doconti != 'C' and doconti != 'c' and doconti != '':
                    doconti = self.doc.Utility.GetKeyword("继续仿真下一工件位置或 [保留当前工件位置(C)]: ")
                time.sleep(0.3)
                if doconti == '':
                    for eachcar in car:
                        eachcar.Delete()
                for lay in layerobjs:
                    lay.LayerOn = False  # 关闭图层
            j += 1
        self.doc.ActiveLayer = self.doc.Layers("0")
