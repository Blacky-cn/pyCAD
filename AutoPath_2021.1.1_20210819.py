#!/usr/bin/env python
# -*- coding: utf-8 -*-
# by: LZH
# cadpath_20210817重构使用class，合并轨迹与仿真选项卡

import math
import tkinter as tk
from tkinter import ttk
from tkinter import Menu
from tkinter import messagebox as msg
import pythoncom
import win32com.client
import time


def vtpnt(x, y, z=0):
    """坐标点转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):
    """转化为对象数组"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtfloat(list1):
    """列表转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list1)


def vtint(list1):
    """列表转化为整数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list1)


def vtvariant(list1):
    """列表转化为变体"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list1)


class Dotab2(object):
    def __init__(self, widget, select):
        self.widget = widget
        self.select = select
        for child in self.widget.winfo_children():
            child.destroy()
        if self.select // 10 == 3:
            self.donext3(self.select)

    # Do '台车'
    def donext3(self, select):
        # Add a Label
        ttk.Label(self.widget, text='选择工件: ').grid(column=0, row=0)
        # Add a Label
        self.carselect = ttk.Label(self.widget, text='未选择工件', width=12)
        self.carselect.grid(column=1, row=0)
        self.carselect.configure(foreground='red')
        # Add a Button
        ttk.Button(self.widget, text='选择', command=self.click_block).grid(column=2, row=0)

        ttk.Label(self.widget, text='选择工件块方向: ').grid(column=0, row=1)
        self.cardir_value = tk.IntVar()
        self.cardir_value.set(1)
        cardir = [('右', 1), ('左', 2)]
        for i, j in cardir:
            a31 = ttk.Radiobutton(self.widget, text=i, value=j, variable=self.cardir_value)
            a31.grid(column=j, row=1)

        ttk.Label(self.widget, text='工件前后支撑距离(mm): ').grid(column=0, row=2)
        self.bracing = tk.StringVar()
        ttk.Entry(self.widget, width=12, textvariable=self.bracing).grid(column=1, row=2)

        ttk.Label(self.widget, text='选择轨迹线并指定起点: ').grid(column=0, row=3)
        self.pathselect = ttk.Label(self.widget, text='未选择轨迹线', width=12)
        self.pathselect.grid(column=1, row=3)
        self.pathselect.configure(foreground='red')
        ttk.Button(self.widget, text='选择', command=self.click_pline).grid(column=2, row=3)

        ttk.Label(self.widget, text='轨迹步长(mm): ').grid(column=0, row=4)
        self.step = tk.StringVar()
        ttk.Entry(self.widget, width=12, textvariable=self.step).grid(column=1, row=4)

        # 若为仿真动画，则启用工件数量、节距选项
        l31 = ttk.Label(self.widget, text='工件数量: ')
        l31.grid(column=0, row=5)
        self.carnum = tk.StringVar()
        e31 = ttk.Entry(self.widget, width=12, textvariable=self.carnum)
        e31.grid(column=1, row=5)
        l32 = ttk.Label(self.widget, text='工件节距(mm): ')
        l32.grid(column=0, row=6)
        self.pitch = tk.StringVar()
        e32 = ttk.Entry(self.widget, width=12, textvariable=self.pitch)
        e32.grid(column=1, row=6)
        if self.select % 10 == 1:
            l31.configure(state='disabled')
            e31.configure(state='disabled')
            l32.configure(state='disabled')
            e32.configure(state='disabled')

        tab2_Button = ttk.LabelFrame(self.widget, text='')
        tab2_Button.grid(column=0, row=7, columnspan=3)
        self.b21 = ttk.Button(tab2_Button, text='确定', command=self.do_path3)
        self.b21.grid(column=0, row=0, padx=8, pady=8)
        self.b22 = ttk.Button(tab2_Button, text='退出', command=oop.quit)
        self.b22.grid(column=1, row=0, padx=8, pady=8)

        for child in self.widget.winfo_children():
            if child != tab2_Button:
                child.grid_configure(sticky=tk.W, padx=8, pady=4)

    # 选择工件
    def click_block(self):
        self.carbody = self.get_block()

    def get_block(self):
        block = doc.Utility.GetEntity()  # 在cad中选取块
        time.sleep(0.1)
        while 'Block' not in block[0].ObjectName:  # 判断选取的是否为块
            msg.showerror('错误', '您选择的不是块，\n请重新选择！')
            block = doc.Utility.GetEntity()
            time.sleep(0.1)
        self.carselect.configure(text='已选择工件', foreground='green')
        return block[0]

    # 选择轨迹线
    def click_pline(self):
        self.track, self.track_vertex, self.track_rev = self.get_pline()  # 获取轨道线、顶点数

    def get_pline(self):
        pline = doc.Utility.GetEntity()  # 在cad中选择多段线
        while 'Polyline' not in pline[0].ObjectName:  # 判断选取的是否为多段线
            msg.showerror('错误', '您选择的不是多段线，\n请重新选择！')
            pline = doc.Utility.GetEntity()
        pline_pnt = list(pline[0].Coordinates[:])  # 获取多段线各顶点坐标
        pline_vertex = len(pline_pnt) // 2
        pllayer, pltype, plcw = pline[0].Layer, pline[0].Linetype, pline[0].ConstantWidth
        # 指定并调整轨迹线方向
        doc.Utility.Prompt("请指定多段线的起点，在相应端点上单击")
        time.sleep(0.1)
        start_pnt = list(doc.Utility.GetPoint())
        while start_pnt[:2] != pline_pnt[:2] and start_pnt[:2] != pline_pnt[-2:]:
            msg.showerror('错误', '未选择多段线的端点，请重新选择！')
            start_pnt = list(doc.Utility.GetPoint())
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
            pline1 = msp.AddLightWeightPolyline(vtfloat(pline1_Pnt))
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

    # 绘制轨迹
    def do_path3(self):
        self.steppnt = self.pathpnt(self.track, self.step)
        if self.select % 10 == 1:
            self.pathcarbody31(self.carbody, self.track, self.steppnt, self.step, self.bracing)
        else:
            self.pathcarbody32(self.carbody, self.track, self.steppnt, self.step,
                               self.bracing, self.carnum, self.pitch)

    # 求插入点
    def pathpnt(self, pline, step):
        try:
            doc.SelectionSets.Item("SS1").Delete()
        except:
            print("Delete selection failed")
        slt = doc.SelectionSets.Add("SS1")
        LayerObj = doc.Layers.Add("Path_pnt")
        doc.ActiveLayer = LayerObj
        LayerObj.LayerOn = False
        handle_str = '_measure'
        handle_str += ' (handent "' + pline.Handle + '")'
        handle_str += '\n'
        step1 = str(step.get())
        handle_str += step1
        handle_str += '\n'
        doc.SendCommand(handle_str)
        filterType = [8]  # 定义过滤类型
        filterData = ["Path_pnt"]  # 设置过滤参数
        filterType = vtint(filterType)
        filterData = vtvariant(filterData)
        slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
        steppnt = [1] * len(slt)
        for i in range(len(slt)):
            steppnt[i] = slt[len(slt) - i - 1].Coordinates[:2]
        print(steppnt)
        slt.Erase()
        slt.Delete()
        doc.ActiveLayer = doc.Layers("0")
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
                        sndpnt[0] = frtpnt[0] - dist * (frtpnt[0] -
                                                        pline.Coordinates[vertex * 2]) / (nowdist - preleng)
                        sndpnt[1] = frtpnt[1] - dist * (frtpnt[1] -
                                                        pline.Coordinates[vertex * 2 + 1]) / (nowdist - preleng)
                    else:
                        sndpnt[0] = pline.Coordinates[(vertex + 1) * 2] - \
                                    (dist - nowdist + i) * (pline.Coordinates[(vertex + 1) * 2] -
                                                            pline.Coordinates[vertex * 2]) / (i - preleng)
                        sndpnt[1] = pline.Coordinates[(vertex + 1) * 2 + 1] - \
                                    (dist - nowdist + i) * (pline.Coordinates[(vertex + 1) * 2 + 1] -
                                                            pline.Coordinates[vertex * 2 + 1]) / (i - preleng)
                else:
                    pline1 = pline.Explode()
                    centpnt = pline1[vertex].Center  # 圆弧圆心坐标
                    centpnt = vtpnt(centpnt[0], centpnt[1])
                    startpnt = vtpnt(pline.Coordinates[vertex * 2], pline.Coordinates[vertex * 2 + 1])
                    ray = msp.AddRay(centpnt, startpnt)
                    arcang = pline1[vertex].TotalAngle  # 圆弧弧度
                    roang = (nowdist - preleng - dist) * arcang / (i - preleng)
                    if pline1[vertex].StartAngle < pline1[vertex].EndAngle:  # 逆时针
                        if (pline1[vertex].StartPoint[0] != pline.Coordinates[vertex * 2]) and \
                                (pline1[vertex].StartPoint[1] != pline.Coordinates[vertex * 2 + 1]):
                            roang = -roang
                    ray.Rotate(centpnt, roang)
                    sndpnt = pline1[vertex].IntersectWith(ray, 0)
                    for j in pline1:
                        j.Delete()
                    ray.Delete()
                break
        return sndpnt

    # '台车' - '绘制轨迹'
    def pathcarbody31(self, block, pline, steppnt, step, bracing):
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [doc.Layers.Add(i) for i in layernames]  # 批量创建图层
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
        step1 = ((pline.Coordinates[-2] - steppnt[-1][0]) ** 2 + (pline.Coordinates[-1] - steppnt[-1][1]) ** 2) ** 0.5
        if step0 == 0:
            frtstep = 0
        elif step0 > step1:
            frtstep = float(step.get())
        else:
            frtstep = leng[-1] - len(steppnt) * float(step.get())
        # 插入文字说明
        inserttext = "轨迹线长度：" + str(leng[-1] // 1)[:-2] + "mm\n" + "轨迹步长：" + step.get() + "mm"
        textpnt = vtpnt((pline.Coordinates[0] + pline.Coordinates[-2]) / 2,
                        max(pline.Coordinates[0], pline.Coordinates[1]) + 3000)
        mt = msp.AddMText(textpnt, 3000, inserttext)
        mt.height = 200
        for i in steppnt:
            if ((j - 1) * float(step.get()) + frtstep) > float(bracing.get()):
                print((j - 1) * float(step.get()) + frtstep)
                layernum = (j - 1) % len(clrnums)
                doc.ActiveLayer = doc.Layers(layernames[layernum])
                inspnt = vtpnt(i[0], i[1])
                circlebr = msp.AddCircle(inspnt, float(bracing.get()))
                intpnts = pline.IntersectWith(circlebr, 0)
                angbr = []
                for k in range(len(intpnts) // 3):
                    brpnt = vtpnt(intpnts[k * 3], intpnts[k * 3 + 1])
                    angbr.append(doc.Utility.AngleFromXAxis(brpnt, inspnt))
                for k in range(len(leng)):
                    if ((j - 1) * float(step.get()) + frtstep) < leng[k]:
                        jj += 1
                        print(leng[k])
                        plpnt = pline.Coordinates[(k * 2):(k * 2 + 2)]
                        plpnt = vtpnt(plpnt[0], plpnt[1])
                        angpl = doc.Utility.AngleFromXAxis(plpnt, inspnt)
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
                    car1 = msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                    if mirmark == 1:  # 起始向左
                        fstpnt = vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                        car1.Mirror(fstpnt, inspnt)
                        car1.Delete()
                else:  # 工件向左
                    if angbr[index1] < math.pi:
                        a = angbr[index1] + math.pi
                    else:
                        a = angbr[index1] - math.pi
                    car1 = msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                    if mirmark == 0:  # 起始向右
                        fstpnt = vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                        car1.Mirror(fstpnt, inspnt)
                        car1.Delete()
                circlebr.Delete()
            j += 1
        for lay in layerobjs:
            lay.LayerOn = True  # 打开图层
        doc.ActiveLayer = doc.Layers("0")

    # '台车' - '仿真动画'
    def pathcarbody32(self, block, pline, steppnt, step, bracing, carnum, pitch):
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [doc.Layers.Add(i) for i in layernames]  # 批量创建图层
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
        step1 = ((pline.Coordinates[-2] - steppnt[-1][0]) ** 2 + (pline.Coordinates[-1] - steppnt[-1][1]) ** 2) ** 0.5
        if step0 == 0:
            frtstep = 0
        elif step0 > step1:
            frtstep = float(step.get())
        else:
            frtstep = leng[-1] - len(steppnt) * float(step.get())
        # 插入文字说明
        inserttext = "轨迹线长度：" + str(leng[-1] // 1)[:-2] + "mm\n" + "工件节距：" + pitch.get() + "mm"
        textpnt = vtpnt((pline.Coordinates[0] + pline.Coordinates[-2]) / 2,
                        max(pline.Coordinates[0], pline.Coordinates[1]) + 3000)
        mt = msp.AddMText(textpnt, 3000, inserttext)
        mt.height = 200
        for i in steppnt:
            car = [1] * int(carnum.get())
            if ((j - 1) * float(step.get()) + frtstep) > (
                    (int(carnum.get()) - 1) * float(pitch.get()) + float(bracing.get())):
                print((j - 1) * float(step.get()) + frtstep)
                for num in range(int(carnum.get())):
                    layernum = num % len(clrnums)
                    doc.ActiveLayer = doc.Layers(layernames[layernum])
                    if num == 0:
                        inspnt0 = i
                    else:
                        nowdist = (j - 1) * float(step.get()) + frtstep - (num - 1) * float(pitch.get())
                        inspnt0 = self.getdistpt(pline, leng, nowdist, inspnt0, float(pitch.get()))
                    inspnt = vtpnt(inspnt0[0], inspnt0[1])
                    circlebr = msp.AddCircle(inspnt, float(bracing.get()))
                    intpnts = pline.IntersectWith(circlebr, 0)
                    angbr = []
                    for k in range(len(intpnts) // 3):
                        brpnt = vtpnt(intpnts[k * 3], intpnts[k * 3 + 1])
                        angbr.append(doc.Utility.AngleFromXAxis(brpnt, inspnt))
                    for k in range(len(leng)):
                        if ((j - 1) * float(step.get()) + frtstep - num * float(pitch.get())) < leng[k]:
                            jj += 1
                            print(leng[k])
                            plpnt = pline.Coordinates[(k * 2):(k * 2 + 2)]
                            plpnt = vtpnt(plpnt[0], plpnt[1])
                            angpl = doc.Utility.AngleFromXAxis(plpnt, inspnt)
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
                        car[num] = msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                        if mirmark == 1:  # 起始向左
                            fstpnt = vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                            car1 = car[num].Mirror(fstpnt, inspnt)
                            car[num].Delete()
                            car[num] = car1
                    else:  # 工件向左
                        if angbr[index1] < math.pi:
                            a = angbr[index1] + math.pi
                        else:
                            a = angbr[index1] - math.pi
                        car[num] = msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                        if mirmark == 0:  # 起始向右
                            fstpnt = vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                            car1 = car[num].Mirror(fstpnt, inspnt)
                            car[num].Delete()
                            car[num] = car1
                    circlebr.Delete()
                for lay in layerobjs:
                    lay.LayerOn = True  # 打开图层
                # 选择是否保留当前结果
                doc.Utility.InitializeUserInput(0, "C, c")
                doconti = 'a'
                while doconti != 'C' and doconti != 'c' and doconti != '':
                    doconti = doc.Utility.GetKeyword("继续仿真下一工件位置或 [保留当前工件位置(C)]: ")
                time.sleep(0.3)
                if doconti == '':
                    for eachcar in car:
                        eachcar.Delete()
                for lay in layerobjs:
                    lay.LayerOn = False  # 关闭图层
            j += 1
        doc.ActiveLayer = doc.Layers("0")


class OOP:
    def __init__(self):
        # Create Instance
        self.win = tk.Tk()
        self.win.title("AutoPath")
        self.win.geometry('360x320')
        self.win.resizable(False, False)
        self.create_widgets()

    # Button Click Function
    def donext(self):
        self.tabControl.tab(1, state='normal')
        self.tabControl.select(self.tab2)
        doselect = self.cartype_value.get() * 10 + self.pathtype_value.get()
        Dotab2(self.tab2, doselect)

    # Radiobutton Callback
    def nexten(self):
        if (self.cartype_value.get() == 0) or (self.pathtype_value.get() == 0):
            self.b11.configure(state='disabled')
        else:
            self.b11.configure(state='normal')

    # Exit GUI cleanly
    def quit(self):
        self.win.quit()
        self.win.destroy()
        exit()

    # About Menu
    def aboutmsg(self):
        msg.showinfo('关于 AutoPath',
                     'AutoPath 2021.1.1 Beta\n\n'
                     'Copyright \u00a9 2021 LZH. All Rights Reserved.\n\n'
                     'Supporting Autodesk CAD: v2010-v2020. Other versions are not tested.')

    ##########################################################################
    def create_widgets(self):
        # Adding Menu
        menu_bar = Menu(self.win)
        self.win.config(menu=menu_bar)
        # Adding File Menu
        file_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='文件', menu=file_menu)
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self.quit)
        # Adding Help Menu
        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label='帮助', menu=help_menu)
        help_menu.add_command(label='入门指南')
        help_menu.add_separator()
        help_menu.add_command(label='关于', command=self.aboutmsg)

        # Create Tab Control
        self.tabControl = ttk.Notebook(self.win)
        tab1 = ttk.Frame(self.tabControl)  # Create Tab1
        self.tabControl.add(tab1, text='功能选择')  # Add Tab1
        self.tab2 = ttk.Frame(self.tabControl)  # Create Tab2
        self.tabControl.add(self.tab2, text='输入参数', state='disabled')  # Add Tab2
        self.tabControl.pack(expand=1, fill='both')  # Pack to make visible

        # Adding a Label
        ttk.Label(tab1, text='请选择工件类型: ').grid(column=0, row=0, sticky=tk.W)
        # Adding Radiobutton
        self.cartype_value = tk.IntVar()
        cartype = [('摆杆', 1), ('翻转机', 2), ('台车', 3)]
        for i, j in cartype:
            self.a1 = ttk.Radiobutton(tab1, text=i, value=j, variable=self.cartype_value, command=self.nexten)
            self.a1.grid(column=j, row=0, sticky=tk.W)
        # Adding a Label
        ttk.Label(tab1, text='请选择轨迹类型: ').grid(column=0, row=1, sticky=tk.W)
        # Adding Radiobutton
        self.pathtype_value = tk.IntVar()
        pathtype = [('绘制轨迹', 1), ('仿真动画', 2)]
        for i, j in pathtype:
            self.a2 = ttk.Radiobutton(tab1, text=i, value=j, variable=self.pathtype_value, command=self.nexten)
            self.a2.grid(column=j, row=1, sticky=tk.W)

        # LabelFrame using Tab1 as the parent
        tab1_Button = ttk.LabelFrame(tab1, text='')
        tab1_Button.grid(column=0, row=2, columnspan=4)
        self.b11 = ttk.Button(tab1_Button, text='下一步', state='disabled', command=self.donext)
        self.b11.grid(column=0, row=0, padx=8, pady=8)
        self.b12 = ttk.Button(tab1_Button, text='退出', command=self.quit)
        self.b12.grid(column=1, row=0, padx=8, pady=8)

        for child in tab1.winfo_children():
            child.grid_configure(padx=8, pady=4)


if __name__ == '__main__':
    wincad = win32com.client.Dispatch("AutoCAD.Application")
    doc = wincad.ActiveDocument
    time.sleep(0.1)
    doc.Utility.Prompt("Hello! Autocad from pywin32com.\n")
    time.sleep(0.1)
    msp = doc.ModelSpace
    oop = OOP()
    # oop.win.iconbitmap('car.ico')
    oop.win.mainloop()
