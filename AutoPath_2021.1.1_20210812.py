#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
import win32com.client
import win32api
import win32con
import math


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


def get_block():
    block = doc.Utility.GetEntity()  # 在cad中选取块
    while 'Block' not in block[0].ObjectName:  # 判断选取的是否为块
        win32api.MessageBox(0, "您选择的不是块，\n请重新选择！", "警告", win32con.MB_ICONWARNING)
        block = doc.Utility.GetEntity()
    return block[0]


def revdirection(pnt, rev):
    if rev == 4:
        revpnt = [1] * len(pnt)
        for i in range(len(pnt)):
            revpnt[i] = pnt[len(pnt) - i - 1]
    else:
        revpnt = pnt
    print(revpnt)
    return revpnt


def get_pline():
    pline = doc.Utility.GetEntity()  # 在cad中选择多段线
    while 'Polyline' not in pline[0].ObjectName:  # 判断选取的是否为多段线
        win32api.MessageBox(0, "您选择的不是多段线，\n请重新选择！", "警告", win32con.MB_ICONWARNING)
        pline = doc.Utility.GetEntity()
    pline_pnt = list(pline[0].Coordinates[:])  # 获取多段线各顶点坐标
    pline_vertex = len(pline_pnt) // 2
    pllayer, pltype, plcw = pline[0].Layer, pline[0].Linetype, pline[0].ConstantWidth
    # 指定并调整轨迹线方向
    print("请指定多段线的起点，在相应端点上单击")
    start_pnt = list(doc.Utility.GetPoint())
    while start_pnt[:2] != pline_pnt[:2] and start_pnt[:2] != pline_pnt[-2:]:
        print("未选择多段线的端点，请重新选择")
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
    return pline1, pline_vertex, rev


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
    step = str(step)
    handle_str += step
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


def pathcarbody(block, pline, steppnt, step, bracing, rev):
    clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
    layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
    layerobjs = [doc.Layers.Add(i) for i in layernames]  # 批量创建图层
    # for lay in layerobjs:
    #     lay.LayerOn = False  # 关闭图层
    for i in range(len(layerobjs)):
        layerobjs[i].color = clrnums[i]
    leng = pline_length(pline)
    for i in range(1, len(leng)):
        leng[i] += leng[i - 1]
    j = 1
    jj = 0  # 判断是否为第一个插入点
    for i in steppnt:
        if (j * float(step)) > bracing:
            print(j * float(step))
            layernum = (j - 1) % len(clrnums)
            doc.ActiveLayer = doc.Layers(layernames[layernum])
            inspnt = vtpnt(i[0], i[1])
            center = vtpnt(i[0], i[1])
            circlebr = msp.AddCircle(center, bracing)
            intpnts = pline.IntersectWith(circlebr, 0)
            angbr = []
            for k in range(len(intpnts) // 3):
                brpnt = vtpnt(intpnts[k * 3], intpnts[k * 3 + 1])
                angbr.append(doc.Utility.AngleFromXAxis(brpnt, inspnt))
            for k in range(len(leng)):
                if (j * float(step)) < leng[k]:
                    jj += 1
                    print(leng[k])
                    plpnt = pline.Coordinates[(k * 2):(k * 2 + 2)]
                    plpnt = vtpnt(plpnt[0], plpnt[1])
                    angpl = doc.Utility.AngleFromXAxis(plpnt, inspnt)
                    if jj == 1:  # 判断是否为第一个插入点
                        if i[0] < pline.Coordinates[k * 2]:  # 起始向左
                            mirmark = 1
                        else:  # 起始向右
                            mirmark = 0
                    break
            angdif = []
            for k in range(len(angbr)):
                angabs = abs(angbr[k] - angpl)
                if angabs > math.pi:
                    angabs = math.pi * 2 - angabs
                angdif.append(angabs)
            angmin, angmax = min(angdif), max(angdif)
            index1, index2 = angdif.index(angmin), angdif.index(angmax)
            if rev == 1:  # 工件向右
                a = angbr[index1]
                car1 = msp.InsertBlock(inspnt, block.Name, 1, 1, 1, a)
                if mirmark == 1:  # 起始向左
                    fstpnt = vtpnt(intpnts[index1 * 3], intpnts[index1 * 3 + 1])
                    car1.Mirror(fstpnt, inspnt)
                    car1.Delete()
            else:  # 工件向左
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


wincad = win32com.client.Dispatch("AutoCAD.Application")
doc = wincad.ActiveDocument
doc.Utility.Prompt("Hello! Autocad from pywin32com.\n")
msp = doc.ModelSpace
# doc.SetVariable("PDMODE", 0)
# print(doc.Name)

doc.Utility.InitializeUserInput(0, "L, l")
cardir = "a"
while cardir != "L" and cardir != "l" and cardir != "":
    cardir = doc.Utility.GetKeyword("工件方向向右或 [向左(L)]: ")
if cardir == "":
    car_rev = 1
else:
    car_rev = -1

carbody = get_block()  # 获取carbody块
# bracing = doc.Utility.GetReal("请输入前后支撑间距: ")
bracing = 20
print("请选择轨道线: ")
track, track_vertex, track_rev = get_pline()  # 获取轨道线、顶点数
# print("请选择链条线: ")
# chain, chain_vertex, chain_rev = get_pline()  # 获取链条线、顶点数

# step = doc.Utility.GetString(0, "请输入轨迹步长: ")
step = 50
steppnt = pathpnt(track, step)
steppnt = revdirection(steppnt, track_rev)
pathcarbody(carbody, track, steppnt, step, bracing, car_rev)
