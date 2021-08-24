#!/usr/bin/env python
# -*- coding: utf-8 -*-
# by: LZH

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


def wintk():
    def get_block():
        global carselect
        block = doc.Utility.GetEntity()  # 在cad中选取块
        time.sleep(0.1)
        while 'Block' not in block[0].ObjectName:  # 判断选取的是否为块
            msg.showerror('错误', '您选择的不是块，\n请重新选择！')
            block = doc.Utility.GetEntity()
            time.sleep(0.1)
        carselect.configure(text='已选择工件')
        return block[0]

    def click_block():
        global carbody
        carbody = get_block()

    def pathpnt(pline, step):
        try:
            doc.SelectionSets.Item("SS1").Delete()
        except:
            print("Delete selection failed")
        slt = doc.SelectionSets.Add("SS1")
        LayerObj = doc.Layers.Add("Path_pnt")
        doc.ActiveLayer = LayerObj
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
        # print(steppnt)
        slt.Erase()
        slt.Delete()
        doc.ActiveLayer = doc.Layers("0")
        LayerObj.Delete()
        return steppnt

    def get_pline():
        global pathselect
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
            doc.Utility.Prompt("未选择多段线的端点，请重新选择")
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
        pathselect.configure(text='已选择轨迹线')
        return pline1, pline_vertex, rev

    def click_pline():
        global track, track_vertex, track_rev
        track, track_vertex, track_rev = get_pline()  # 获取轨道线、顶点数

    def pline_length(pline1):
        leng = []
        pline2 = pline1.Explode()
        for i in pline2:
            if 'Arc' in i.ObjectName:
                leng.append(i.ArcLength)
            else:
                leng.append(i.Length)
            i.Delete()
        return leng

    def pathcarbody(block, pline, steppnt, step, bracing):
        global cardir_value
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [doc.Layers.Add(i) for i in layernames]  # 批量创建图层
        for lay in layerobjs:
            lay.LayerOn = False  # 关闭图层
        for i in range(len(layerobjs)):
            layerobjs[i].color = clrnums[i]
        leng = pline_length(pline)
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
        for i in steppnt:
            if ((j - 1) * float(step.get()) + frtstep) > float(bracing.get()):
                print((j - 1) * float(step.get()) + frtstep)
                layernum = (j - 1) % len(clrnums)
                doc.ActiveLayer = doc.Layers(layernames[layernum])
                inspnt = vtpnt(i[0], i[1])
                center = vtpnt(i[0], i[1])
                circlebr = msp.AddCircle(center, float(bracing.get()))
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
                if cardir_value.get() == 1:  # 工件向右
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

    def do_path():
        global track, track_rev, carbody, step, bracing
        steppnt = pathpnt(track, step)
        pathcarbody(carbody, track, steppnt, step, bracing)

    def _quit():
        win.quit()
        win.destroy()
        exit()

    def _Aboutmsg():
        msg.showinfo('关于 AutoPath',
                     'AutoPath 2021.08.1\n\n'
                     'Rights to LZH\n\n'
                     '支持Autodesk CAD版本：CAD2010-2020，其他版本未经测试\n')

    def dotab31():  # 执行'台车'-‘绘制轨迹’
        global carselect, pathselect, cardir_value, step, bracing
        ttk.Label(tab2, text='选择工件: ').grid(column=0, row=0)
        carselect = ttk.Label(tab2, text='未选择工件', width=12)
        carselect.grid(column=1, row=0)
        carselect.configure(background='white')
        ttk.Button(tab2, text='选择', command=click_block).grid(column=2, row=0)

        ttk.Label(tab2, text='选择工件块方向: ').grid(column=0, row=1)
        cardir_value = tk.IntVar()
        cardir_value.set(1)
        cardir = [('右', 1), ('左', 2)]
        for i, j in cardir:
            a3 = ttk.Radiobutton(tab2, text=i, value=j, variable=cardir_value)
            a3.grid(column=j, row=1)

        ttk.Label(tab2, text='工件前后支撑距离: ').grid(column=0, row=2)
        bracing = tk.StringVar()
        ttk.Entry(tab2, width=12, textvariable=bracing).grid(column=1, row=2)

        ttk.Label(tab2, text='选择轨迹线并指定起点: ').grid(column=0, row=3)
        pathselect = ttk.Label(tab2, text='未选择轨迹线', width=12)
        pathselect.grid(column=1, row=3)
        pathselect.configure(background='white')
        ttk.Button(tab2, text='选择', command=click_pline).grid(column=2, row=3)

        ttk.Label(tab2, text='轨迹步长: ').grid(column=0, row=4)
        step = tk.StringVar()
        ttk.Entry(tab2, width=12, textvariable=step).grid(column=1, row=4)

        tab2_Button = ttk.LabelFrame(tab2, text='')
        tab2_Button.grid(column=0, row=5, columnspan=3)
        b1 = ttk.Button(tab2_Button, text='确定', command=do_path)
        b1.grid(column=0, row=0, padx=8, pady=8)
        b2 = ttk.Button(tab2_Button, text='退出', command=_quit)
        b2.grid(column=1, row=0, padx=8, pady=8)

        for child in tab2.winfo_children():
            if child != tab2_Button:
                child.grid_configure(sticky=tk.W, padx=8, pady=4)

    def nextenable():
        if (cartype_value.get() == 0) or (pathtype_value.get() == 0):
            b11.configure(state='disabled')
        elif (cartype_value.get() == 3) and (pathtype_value.get() == 1):
            b11.configure(state='normal')
            dotab31()

    def donext():
        tabControl.tab(1, state='normal')
        tabControl.select(tab2)

    win = tk.Tk()
    win.title('AutoPath')
    win.geometry('360x300')
    # win.resizable(False, False)
    tabControl = ttk.Notebook(win)
    tab1 = ttk.Frame(tabControl)
    tabControl.add(tab1, text='功能选择')
    tab2 = ttk.Frame(tabControl)
    tabControl.add(tab2, text='输入参数', state='disabled')
    tabControl.pack(expand=1, fill='both')

    ttk.Label(tab1, text='请选择工件类型: ').grid(column=0, row=0, sticky=tk.W)
    cartype_value = tk.IntVar()
    cartype = [('摆杆', 1), ('翻转机', 2), ('台车', 3)]
    for i, j in cartype:
        a1 = ttk.Radiobutton(tab1, text=i, value=j, variable=cartype_value, command=nextenable)
        a1.grid(column=j, row=0, sticky=tk.W)
    ttk.Label(tab1, text='请选择轨迹类型: ').grid(column=0, row=1, sticky=tk.W)
    pathtype_value = tk.IntVar()
    pathtype = [('绘制轨迹', 1), ('仿真动画', 2)]
    for i, j in pathtype:
        a2 = ttk.Radiobutton(tab1, text=i, value=j, variable=pathtype_value, command=nextenable)
        a2.grid(column=j, row=1, sticky=tk.W)

    tab1_Button = ttk.LabelFrame(tab1, text='')
    tab1_Button.grid(column=0, row=2, columnspan=4)
    b11 = ttk.Button(tab1_Button, text='下一步', state='disabled', command=donext)
    b11.grid(column=0, row=0, padx=8, pady=8)
    b12 = ttk.Button(tab1_Button, text='退出', command=_quit)
    b12.grid(column=1, row=0, padx=8, pady=8)

    for child in tab1.winfo_children():
        child.grid_configure(padx=8, pady=4)

    menu_bar = Menu(win)
    win.config(menu=menu_bar)
    file_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label='文件', menu=file_menu)
    file_menu.add_separator()
    file_menu.add_command(label='退出', command=_quit)

    help_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label='帮助', menu=help_menu)
    help_menu.add_command(label='入门指南')
    help_menu.add_separator()
    help_menu.add_command(label='关于', command=_Aboutmsg)

    win.mainloop()


if __name__ == '__main__':
    wincad = win32com.client.Dispatch("AutoCAD.Application")
    doc = wincad.ActiveDocument
    time.sleep(0.1)
    doc.Utility.Prompt("Hello! Autocad from pywin32com.\n")
    time.sleep(0.1)
    msp = doc.ModelSpace
    wintk()
