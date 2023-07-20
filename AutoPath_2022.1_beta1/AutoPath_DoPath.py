#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
轨迹函数
"""

# imports==========
import math
import time
from tkinter import messagebox as msg

import AutoPath_TypeConvert as Tc


# =========================================================
class DOPATH:
    def __init__(self, oop, doc, msp):
        self.oop = oop
        self.doc = doc
        self.msp = msp
        self.layerobjs = self._create_layer()

    def do_pendulum(self, select, dirvalue, swingstate_value, chainpath, step, chainbracing, bracing, carnum, pitch,
                    swingleng):
        """摆杆 - 轨迹/动画"""
        self.dirvalue = dirvalue
        steppnt = self._pathpnt(chainpath, step)
        for lay in self.layerobjs:
            lay.LayerOn = False  # 关闭图层
        self.chainleng = self._pline_length(chainpath)
        self.frtstep = self._frtsteppnt(chainpath, steppnt, step, self.chainleng)
        if swingstate_value == 1:
            swingstate_value_text = '模式：前摆杆竖直'
        else:
            swingstate_value_text = '模式：后摆杆竖直'
        if select % 10 == 1:
            inserttext = swingstate_value_text + '\n轨道长度：' + str(
                round(self.chainleng[-1], 2)) + 'mm\n' + '轨迹步长：' + str(
                step) + 'mm\n' + '摆杆间距：' + str(bracing) + 'mm\n' + '摆杆长度：' + str(swingleng + 252.75) + 'mm'
        else:
            inserttext = swingstate_value_text + '\n轨道长度：' + str(
                round(self.chainleng[-1], 2)) + 'mm\n' + '工件节距：' + str(
                pitch) + 'mm\n' + '摆杆间距：' + str(bracing) + 'mm\n' + '摆杆长度：' + str(swingleng + 252.75) + 'mm'
        self._insert_mt(chainpath, inserttext)
        self.j = 1
        self.jj = 0  # 判断是否为第一次插入
        for i in steppnt:
            # 更新进度条==============
            progress_value = round(self.j / len(steppnt) * 100, 0)
            self.oop.update_progressbar(progress_value)

            chainplate = []
            fswing = []
            bswing = []
            car = []
            if ((self.j - 1) * float(step) + self.frtstep) > ((carnum - 1) * pitch + bracing + 2 * chainbracing):
                layernum = self.j % len(self.layerobjs)
                self.doc.ActiveLayer = self.doc.Layers(self.layerobjs[layernum].Name)
                for num in range(carnum):
                    nowdist = (self.j - 1) * float(step) + self.frtstep
                    if num == 0:
                        inspnt0 = i
                    else:
                        inspnt0 = self._find_distpnt(chainpath, self.chainleng, nowdist, i, num * pitch)
                    self.swingpnt = []
                    for swingnum in range(2):
                        if swingnum == 0:
                            inspnt1 = inspnt0
                        if swingnum == 1:
                            inspnt1 = self._find_distpnt(chainpath, self.chainleng, nowdist - num * pitch, inspnt0, bracing)
                        self._find_insertpnt(chainpath, chainbracing, nowdist - num * pitch - swingnum * bracing,
                                             inspnt1, i)
                        self._insert_block('ChainPlate', chainplate, swingnum)
                        self.circlebr.Delete()
                        self.swingpnt.append(self._find_isotrianglepnt(inspnt1))
                        if swingnum == 0 and swingstate_value == 1:  # 前摆杆竖直
                            self._insert_swing(fswing, self.swingpnt[0], swingleng, bracing, 'FSwing')
                        elif swingnum == 1 and swingstate_value == 2:  # 后摆杆竖直
                            self._insert_swing(bswing, self.swingpnt[1], swingleng, bracing, 'BSwing')
                        else:
                            swingpnt0 = Tc.vtpnt(self.swingpnt[-1][0], self.swingpnt[-1][1])
                            self.circleswing = self.msp.AddCircle(swingpnt0, swingleng)
                    if swingstate_value == 1:  # 前摆杆竖直，插入后摆杆和工件
                        self._insert_car(swingstate_value, car, bswing)
                    else:  # 后摆杆竖直，插入前摆杆和工件
                        self._insert_car(swingstate_value, car, fswing)
                    self.circlecar.Delete()
                    self.circleswing.Delete()
                if select % 10 == 2:
                    self._animation_next_or_copy(chainplate, fswing, bswing, car)
            self.j += 1
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

        # 更新进度条==============
        self.oop.update_progressbar(progress_value, finish=1)

    def do_pendulum_diptank(self, dirvalue, swingmode_value, chainpath, diptank_midpnt, chainbracing, bracing, pitch,
                            swingleng):
        """摆杆 - 浸入即出槽分析"""
        for lay in self.layerobjs:
            lay.LayerOn = False  # 关闭图层
        self.dirvalue = dirvalue
        self.chainleng = self._pline_length(chainpath)
        pline_segment = self._point_on_pline(chainpath, diptank_midpnt)
        dist_of_plinestart = self._find_dist(chainpath, diptank_midpnt[:2], pline_segment)
        if (dist_of_plinestart <= (bracing + (pitch - bracing) / 2 + chainbracing)) or (
                (self.chainleng[-1] - dist_of_plinestart) <= (bracing + (pitch - bracing) / 2 + chainbracing)):
            if msg.showerror('错误', '请加长轨迹线！'):
                self.oop.menubutton.quit()
        if swingmode_value == 1:
            swingmode_value_text = '模式：内侧摆杆竖直'
        else:
            swingmode_value_text = '模式：外侧摆杆竖直'
        inserttext = '浸入即出槽分析\n' + swingmode_value_text + '\n摆杆间距：' + str(
            bracing) + 'mm\n' + '摆杆长度：' + str(
            swingleng) + 'mm\n' + '工件节距：' + str(pitch) + 'mm'
        self._insert_mt(chainpath, inserttext)
        self.jj = 0  # 判断是否为第一次插入
        for num in range(2):
            chainplate = []
            fswing = []
            bswing = []
            car = []
            if num == 0:  # 后车
                nowdist = (pitch - bracing) / 2
            else:  # 前车
                nowdist = -(pitch + bracing) / 2
            layernum = (num + 1) % len(self.layerobjs)
            self.doc.ActiveLayer = self.doc.Layers(self.layerobjs[layernum].Name)
            inspnt0 = self._find_distpnt(chainpath, self.chainleng, dist_of_plinestart, diptank_midpnt[:2], nowdist)
            self.swingpnt = []
            for swingnum in range(2):
                # 更新进度条==============
                progress_value = round((num * 2 + swingnum + 1) / 4 * 100, 0)
                self.oop.update_progressbar(progress_value)

                if swingnum == 0:
                    inspnt1 = inspnt0  # 前链板插入点
                else:
                    inspnt1 = self._find_distpnt(chainpath, self.chainleng, dist_of_plinestart - nowdist, inspnt0,
                                                 bracing)  # 后链板插入点
                self._find_insertpnt(chainpath, chainbracing, dist_of_plinestart - nowdist - swingnum * bracing,
                                     inspnt1, inspnt0)
                self._insert_block('ChainPlate', chainplate, swingnum)
                self.circlebr.Delete()
                self.swingpnt.append(self._find_isotrianglepnt(inspnt1))
                if swingmode_value == 1:  # 内侧摆杆竖直，即前车后摆杆、后车前摆杆竖直
                    if num == 1 and swingnum == 1:  # 插入前车后摆杆
                        self._insert_swing(fswing, self.swingpnt[1], swingleng, bracing, 'FSwing')
                    elif num == 0 and swingnum == 0:  # 插入后车前摆杆
                        self._insert_swing(bswing, self.swingpnt[0], swingleng, bracing, 'BSwing')
                    else:
                        swingpnt0 = Tc.vtpnt(self.swingpnt[-1][0], self.swingpnt[-1][1])
                        self.circleswing = self.msp.AddCircle(swingpnt0, swingleng)
                else:  # 外侧摆杆竖直，即前车前摆杆、后车后摆杆竖直
                    if num == 1 and swingnum == 0:  # 插入前车前摆杆
                        self._insert_swing(fswing, self.swingpnt[1], swingleng, bracing, 'FSwing')
                    elif num == 0 and swingnum == 1:  # 插入后摆杆竖直
                        self._insert_swing(bswing, self.swingpnt[0], swingleng, bracing, 'BSwing')
                    else:
                        swingpnt0 = Tc.vtpnt(self.swingpnt[-1][0], self.swingpnt[-1][1])
                        self.circleswing = self.msp.AddCircle(swingpnt0, swingleng)
            if swingmode_value == 1:  # 内侧摆杆竖直，即前车后摆杆、后车前摆杆竖直
                if num == 1:  # 插入前车前摆杆和工件
                    swingstate_value = 2
                    self._insert_car(swingstate_value, car, fswing)
                else:  # 插入后车后摆杆和工件
                    swingstate_value = 1
                    self._insert_car(swingstate_value, car, bswing)
            else:  # 外侧摆杆竖直，即前车前摆杆、后车后摆杆竖直
                if num == 1:  # 插入前车后摆杆和工件
                    swingstate_value = 1
                    self._insert_car(swingstate_value, car, bswing)
                else:  # 插入后车前摆杆和工件
                    swingstate_value = 2
                    self._insert_car(swingstate_value, car, fswing)
            self.circlecar.Delete()
            self.circleswing.Delete()
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

        # 更新进度条==============
        self.oop.update_progressbar(progress_value, finish=1)

    def do_2trolley(self, select, dirvalue, car_name, chainpath, hinglepath, step, bracing, carnum, pitch):
        """台车 - 轨迹/动画"""
        self.dirvalue = dirvalue
        steppnt = self._pathpnt(chainpath, step)
        # for lay in self.layerobjs:
        #     lay.LayerOn = False  # 关闭图层
        self.chainleng = self._pline_length(chainpath)
        self.frtstep = self._frtsteppnt(chainpath, steppnt, step, self.chainleng)
        if select % 10 == 1:
            inserttext = '轨道长度：' + str(int(self.chainleng[-1])) + 'mm\n' + '轨迹步长：' + str(step) + 'mm'
        else:
            inserttext = '轨道长度：' + str(int(self.chainleng[-1])) + 'mm\n' + '工件节距：' + str(pitch) + 'mm'
        self._insert_mt(chainpath, inserttext)
        self.j = 1
        self.jj = 0  # 判断是否为第一次插入
        for i in steppnt:
            # 更新进度条==============
            progress_value = round(self.j / len(steppnt) * 100, 0)
            self.oop.update_progressbar(progress_value)

            car = []
            nowdist = (self.j - 1) * step + self.frtstep
            if nowdist > ((carnum - 1) * pitch + bracing):
                layernum = self.j % len(self.layerobjs)
                self.doc.ActiveLayer = self.doc.Layers(self.layerobjs[layernum].Name)
                for num in range(carnum):
                    if num == 0:
                        inspnt0 = i
                    else:
                        inspnt0 = self._find_distpnt(chainpath, self.chainleng, nowdist, i, num * pitch)
                    if chainpath.Handle == hinglepath.Handle:  # 链条线和轨道线重合
                        self._find_insertpnt(chainpath, bracing, nowdist - num * pitch, inspnt0, i)
                    else:  # 链条线和轨道线不重合
                        inspnt1 = self._find_line2pnt(chainpath, hinglepath, self.chainleng, inspnt0, nowdist)
                        self.rollerleng = self._pline_length(hinglepath)
                        self._find_insertpnt(hinglepath, bracing, nowdist - num * pitch, inspnt1, i)
                    self._insert_block(car_name, car, num)
                    self.circlebr.Delete()
                if select % 10 == 2:
                    self._animation_next_or_copy(car)
            self.j += 1
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

        # 更新进度条==============
        self.oop.update_progressbar(progress_value, finish=1)

    def _find_insertpnt(self, chainpath, bracing, nowdist, inspnt0, i):
        self.inspnt = Tc.vtpnt(inspnt0[0], inspnt0[1])
        self.circlebr = self.msp.AddCircle(self.inspnt, bracing)
        self.intpnts = chainpath.IntersectWith(self.circlebr, 0)
        self.brpnt = []
        self.angbr = []
        for k in range(len(self.intpnts) // 3):
            self.brpnt.append([self.intpnts[k * 3], self.intpnts[k * 3 + 1]])
            self.angbr.append(
                self.doc.Utility.AngleFromXAxis(Tc.vtpnt(self.brpnt[-1][0], self.brpnt[-1][1]), self.inspnt))
        for k in range(len(self.chainleng)):
            if nowdist < self.chainleng[k]:
                self.jj += 1
                # print(chainleng[k])
                plpnt = chainpath.Coordinates[(k * 2):(k * 2 + 2)]
                plpnt = Tc.vtpnt(plpnt[0], plpnt[1])
                angpl = self.doc.Utility.AngleFromXAxis(plpnt, self.inspnt)
                if self.jj == 1:  # 判断是否为第一次插入
                    if i[0] < chainpath.Coordinates[k * 2]:  # 起始向左
                        self.mirmark = 1
                    else:
                        self.mirmark = 0  # 起始向右
                break
        angdif = []
        for k in self.angbr:
            angabs = abs(k - angpl)
            if angabs > math.pi:
                angabs = math.pi * 2 - angabs
            angdif.append(angabs)
        self.index = angdif.index(min(angdif))

    def _insert_block(self, block_name, blocklist, num):
        """插入块并调整角度"""
        if self.dirvalue == 1:  # 工件向右
            a = self.angbr[self.index]
            blocklist.append(self.msp.InsertBlock(self.inspnt, block_name, 1, 1, 1, a))
            if self.mirmark == 1:  # 起始向左
                fstpnt = Tc.vtpnt(self.intpnts[self.index * 3], self.intpnts[self.index * 3 + 1])
                car1 = blocklist[num].Mirror(fstpnt, self.inspnt)
                blocklist[num].Delete()
                blocklist[num] = car1
        else:  # 工件向左
            if self.angbr[self.index] < math.pi:
                a = self.angbr[self.index] + math.pi
            else:
                a = self.angbr[self.index] - math.pi
            blocklist.append(self.msp.InsertBlock(self.inspnt, block_name, 1, 1, 1, a))
            if self.mirmark == 0:  # 起始向右
                fstpnt = Tc.vtpnt(self.intpnts[self.index * 3], self.intpnts[self.index * 3 + 1])
                car1 = blocklist[num].Mirror(fstpnt, self.inspnt)
                blocklist[num].Delete()
                blocklist[num] = car1

    def _insert_swing(self, swinglist, swingpnt, swingleng, bracing, swing_name):
        """插入摆杆"""
        swingpnt0 = Tc.vtpnt(swingpnt[0], swingpnt[1])
        swingpnt1 = Tc.vtpnt(swingpnt[0], swingpnt[1] + 1)
        swinglist.append(self.msp.InsertBlock(swingpnt0, swing_name, 1, 1, 1, 0))
        self._mirror_block(swingpnt0, swingpnt1, swinglist)
        car_fpnt0 = [swingpnt[0], swingpnt[1] - swingleng]
        car_fpnt = Tc.vtpnt(car_fpnt0[0], car_fpnt0[1])
        self.circlecar = self.msp.AddCircle(car_fpnt, bracing)

    def _insert_car(self, state_value, blocklist0, blocklist1, block_name='Skid&Body'):
        """插入工件及另一根摆杆"""
        car_pnt = self.circlecar.IntersectWith(self.circleswing, 0)
        car_pntdif = []
        for k in range(len(car_pnt) // 3):
            if state_value == 1:  # 前摆杆竖直，找后摆杆位置
                car_pntdif.append(abs(car_pnt[k * 3] - self.swingpnt[1][0]))
            else:  # 后摆杆竖直，找前摆杆位置
                car_pntdif.append(abs(car_pnt[k * 3] - self.swingpnt[0][0]))
        car_pnt_index = car_pntdif.index(min(car_pntdif))
        car_inspnt0 = Tc.vtpnt(self.circlecar.Center[0], self.circlecar.Center[1])
        car_inspnt1 = Tc.vtpnt(car_pnt[car_pnt_index * 3], car_pnt[car_pnt_index * 3 + 1])
        if state_value == 1:  # 前摆杆竖直，以circlecar圆心为插入点，第二点为circlecar与circleswing交点
            bswing_pnt = Tc.vtpnt(self.swingpnt[1][0], self.swingpnt[1][1])
            car_angle = self.doc.Utility.AngleFromXAxis(car_inspnt0, car_inspnt1)
            swing_angel = self.doc.Utility.AngleFromXAxis(bswing_pnt, car_inspnt1)
            blocklist0.append(self.msp.InsertBlock(car_inspnt0, block_name, 1, 1, 1, car_angle))
            self._mirror_block(car_inspnt0, car_inspnt1, blocklist0)
            blocklist1.append(self.msp.InsertBlock(bswing_pnt, 'BSwing', 1, 1, 1, swing_angel - math.pi * 3 / 2))
            self._mirror_block(bswing_pnt, car_inspnt1, blocklist1)
        else:  # 后摆杆竖直，以circlecar与circleswing交点为插入点，第二点为circlecar圆心
            fswing_pnt = Tc.vtpnt(self.swingpnt[0][0], self.swingpnt[0][1])
            car_angle = self.doc.Utility.AngleFromXAxis(car_inspnt1, car_inspnt0)
            swing_angel = self.doc.Utility.AngleFromXAxis(fswing_pnt, car_inspnt1)
            blocklist0.append(self.msp.InsertBlock(car_inspnt1, block_name, 1, 1, 1, car_angle))
            self._mirror_block(car_inspnt1, car_inspnt0, blocklist0)
            blocklist1.append(self.msp.InsertBlock(fswing_pnt, 'FSwing', 1, 1, 1, swing_angel - math.pi * 3 / 2))
            self._mirror_block(fswing_pnt, car_inspnt1, blocklist1)

    def _mirror_block(self, blockpnt0, blockpnt1, blocklist):
        """若工件方向与轨迹起始方向相反，则镜像摆杆、工件"""
        if (self.dirvalue == 1 and self.mirmark == 1) or (self.dirvalue == 2 and self.mirmark == 0):
            block_mirror = blocklist[-1].Mirror(blockpnt0, blockpnt1)
            blocklist[-1].Delete()
            blocklist[-1] = block_mirror

    def _pathpnt(self, pline, step):
        """沿轨道线按步长求插入点"""
        steppnt = []
        try:
            self.doc.SelectionSets.Item("SS1").Delete()
        except BaseException as e:
            if e.args[0] == -2147352567:
                print("Delete selection failed")
        slt = self.doc.SelectionSets.Add("SS1")
        LayerObj = self.doc.Layers.Add("Path_pnt")
        self.doc.ActiveLayer = LayerObj
        LayerObj.LayerOn = False
        handle_str = '_measure'
        handle_str += ' (handent "' + pline.Handle + '")'
        handle_str += '\n'
        step1 = str(step)
        handle_str += step1
        handle_str += '\n'
        self.doc.SendCommand(handle_str)
        filterType = [8]  # 定义过滤类型
        filterData = ["Path_pnt"]  # 设置过滤参数
        filterType = Tc.vtint(filterType)
        filterData = Tc.vtvariant(filterData)
        slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
        steppnt = [1] * len(slt)
        for i in range(len(slt)):
            steppnt[i] = slt[len(slt) - i - 1].Coordinates[:2]
        slt.Erase()
        slt.Delete()
        self.doc.ActiveLayer = self.doc.Layers("0")
        LayerObj.Delete()
        return steppnt

    def _find_distpnt(self, pline, leng, nowdist, frtpnt, dist):
        """根据两点沿轨道线的距离，求轨道线上点位置"""
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
                    if (i >= nowdist and dist >= 0) or (preleng <= nowdist and dist < 0):
                        sndpnt[0] = frtpnt[0] - dist * (frtpnt[0] - pline.Coordinates[vertex * 2]) / (nowdist - preleng)
                        sndpnt[1] = frtpnt[1] - dist * (frtpnt[1] - pline.Coordinates[vertex * 2 + 1]) / (
                                nowdist - preleng)
                    elif (i < nowdist and dist >= 0) or (preleng > nowdist and dist < 0):
                        sndpnt[0] = pline.Coordinates[(vertex + 1) * 2] - (i - nowdist + dist) * (
                                pline.Coordinates[(vertex + 1) * 2] - pline.Coordinates[vertex * 2]) / (i - preleng)
                        sndpnt[1] = pline.Coordinates[(vertex + 1) * 2 + 1] - (i - nowdist + dist) * (
                                pline.Coordinates[(vertex + 1) * 2 + 1] - pline.Coordinates[vertex * 2 + 1]) / (
                                            i - preleng)
                else:
                    eachpline = pline.Explode()
                    centpnt = eachpline[vertex].Center  # 圆弧圆心坐标
                    centpnt = Tc.vtpnt(centpnt[0], centpnt[1])
                    startpnt = Tc.vtpnt(pline.Coordinates[vertex * 2],
                                        pline.Coordinates[vertex * 2 + 1])
                    ray = self.msp.AddRay(centpnt, startpnt)
                    arcang = eachpline[vertex].TotalAngle  # 圆弧弧度
                    roang = (nowdist - preleng - dist) * arcang / (i - preleng)
                    if eachpline[vertex].StartAngle < eachpline[vertex].EndAngle:  # 逆时针
                        if (eachpline[vertex].StartPoint[0] != pline.Coordinates[vertex * 2]) and (
                                eachpline[vertex].StartPoint[1] != pline.Coordinates[vertex * 2 + 1]):
                            roang = -roang
                    ray.Rotate(centpnt, roang)
                    sndpnt = eachpline[vertex].IntersectWith(ray, 0)
                    for j in eachpline:
                        j.Delete()
                    ray.Delete()
                break
        return sndpnt

    def _find_isotrianglepnt(self, inspnt0, height=140, direction=0):
        """已知等腰三角形两底角坐标和高，求顶角坐标"""
        frtpnt = [inspnt0[0], inspnt0[1]]
        sndpnt = self.brpnt[self.index]
        if direction == 0:
            if frtpnt[0] == sndpnt[0]:
                trdpnt_x = frtpnt[0] - height
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2
            elif frtpnt[1] == sndpnt[1]:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2
                trdpnt_y = frtpnt[1] - height
            elif (frtpnt[0] < sndpnt[0] and frtpnt[1] < sndpnt[1]) or (frtpnt[0] > sndpnt[0] and frtpnt[1] > sndpnt[1]):
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 + height / (
                        ((sndpnt[0] - frtpnt[0]) / (sndpnt[1] - frtpnt[1])) ** 2 + 1) ** 0.5
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2 - height / (
                        ((sndpnt[1] - frtpnt[1]) / (sndpnt[0] - frtpnt[0])) ** 2 + 1) ** 0.5
            else:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 - height / (
                        ((sndpnt[0] - frtpnt[0]) / (sndpnt[1] - frtpnt[1])) ** 2 + 1) ** 0.5
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2 - height / (
                        ((sndpnt[1] - frtpnt[1]) / (sndpnt[0] - frtpnt[0])) ** 2 + 1) ** 0.5
        else:
            if frtpnt[0] == sndpnt[0]:
                trdpnt_x = frtpnt[0] + height
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2
            elif frtpnt[1] == sndpnt[1]:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2
                trdpnt_y = frtpnt[1] + height
            elif frtpnt[0] < sndpnt[0] and frtpnt[1] < sndpnt[1]:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 - height / (
                        ((sndpnt[0] - frtpnt[0]) / (sndpnt[1] - frtpnt[1])) ** 2 + 1) ** 0.5
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2 + height / (
                        ((sndpnt[1] - frtpnt[1]) / (sndpnt[0] - frtpnt[0])) ** 2 + 1) ** 0.5
            else:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 + height / (
                        ((sndpnt[0] - frtpnt[0]) / (sndpnt[1] - frtpnt[1])) ** 2 + 1) ** 0.5
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2 + height / (
                        ((sndpnt[1] - frtpnt[1]) / (sndpnt[0] - frtpnt[0])) ** 2 + 1) ** 0.5
        trdpnt = [trdpnt_x, trdpnt_y]
        return trdpnt

    def _point_on_pline(self, pline, point, precision=0.0001):
        """判断点是否在多段线上"""
        point1 = Tc.vtpnt(point[0], point[1])
        circle = self.msp.AddCircle(point1, precision)
        if not pline.IntersectWith(circle, 0):
            circle.Delete()
            if msg.showerror('错误', '选择的中点不在轨迹线上！'):
                self.oop.menubutton.quit()
        pline1 = pline.Explode()
        for i in range(len(pline1)):
            if pline1[i].IntersectWith(circle, 0):
                circle.Delete()
                del pline1
                return i

    def _find_dist(self, pline, point, vertex):
        """求多段线上一点到起点的距离"""
        bulge = pline.GetBulge(vertex)
        if vertex == 0:
            preleng = 0
        else:
            preleng = self.chainleng[vertex - 1]
        if bulge == 0:
            dist = ((point[0] - pline.Coordinates[vertex * 2]) ** 2 + (
                    point[1] - pline.Coordinates[vertex * 2 + 1]) ** 2) ** 0.5
        else:
            startpnt = (pline.Coordinates[vertex * 2], pline.Coordinates[vertex * 2 + 1])
            arc_radius = pline[vertex].Radius  # 圆弧半径
            dist = 2 * arc_radius * math.asin(
                ((startpnt[0] - point[0]) ** 2 + (startpnt[1] - point[1]) ** 2) ** 0.5 / 2 / arc_radius)
        return preleng + dist

    def _find_line2pnt(self, pline1, pline2, leng, point1, nowdist):
        """已知pline1上的点point1，从point1做pline1的垂线，求垂线与pline2的焦点"""
        for i in leng:
            if i >= nowdist:
                vertex = leng.index(i)
                bulge = pline1.GetBulge(vertex)
                eachpline = pline1.Explode()
                point1 = Tc.vtpnt(point1[0], point1[1])
                if bulge == 0:
                    endpnt = Tc.vtpnt(pline1.Coordinates[vertex * 2], pline1.Coordinates[vertex * 2 + 1])
                    xline = self.msp.Addxline(point1, endpnt)
                    xline.Rotate(point1, eachpline[vertex].Angle + math.pi / 2)
                    point2 = pline2.IntersectWith(xline, 0)
                else:
                    centpnt = eachpline[vertex].Center  # 圆弧圆心坐标
                    centpnt = Tc.vtpnt(centpnt[0], centpnt[1])
                    xline = self.msp.Addxline(point1, centpnt)
                    point2 = pline2.IntersectWith(xline, 0)
                for j in eachpline:
                    j.Delete()
                xline.Delete()
                break
        return point2

    def _create_layer(self):
        """创建绘图层"""
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [self.doc.Layers.Add(i) for i in layernames]  # 批量创建图层
        for i in range(len(layerobjs)):
            layerobjs[i].color = clrnums[i]
        return layerobjs

    def _insert_mt(self, chainpath, inserttext):
        """绘制轨迹/仿真时，在轨道线上方插入文字说明"""
        textpnt = Tc.vtpnt((min(chainpath.Coordinates[0:-2:2]) + max(chainpath.Coordinates[0:-2:2])) / 2 - 1500,
                           max(chainpath.Coordinates[1:-1:2]) + 3000)
        mt = self.msp.AddMText(textpnt, 3000, inserttext)
        mt.Height = 200
        mt.styleName = 'STANDARD'

    def _animation_next_or_copy(self, *blocks):
        """仿真生成一组工件后，选择继续仿真/保留当前结果"""
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        # 选择是否保留当前结果
        self.doc.Utility.InitializeUserInput(0, "C, c")
        docontinue = 'a'
        while docontinue.lower() != 'c' and docontinue != '':
            docontinue = self.doc.Utility.GetKeyword("继续仿真下一工件位置或 [保留当前工件位置(C)]: ")
        time.sleep(1)
        if docontinue == '':
            for block in blocks:
                for eachblock in block:
                    eachblock.Delete()
        for lay in self.layerobjs:
            lay.LayerOn = False  # 关闭图层

    @staticmethod
    def _pline_length(pline):
        """求轨道线各段长度"""
        leng = []
        pline1 = pline.Explode()
        for i in pline1:
            if 'Arc' in i.ObjectName:
                leng.append(i.ArcLength)
            else:
                leng.append(i.Length)
            i.Delete()
        for i in range(1, len(leng)):
            leng[i] += leng[i - 1]
        return leng

    @staticmethod
    def _frtsteppnt(chainpath, steppnt, step, leng):
        """求轨道线上第一个插入点位置"""
        step0 = ((chainpath.Coordinates[0] - steppnt[0][0]) ** 2 + (
                chainpath.Coordinates[1] - steppnt[0][1]) ** 2) ** 0.5
        step1 = ((chainpath.Coordinates[-2] - steppnt[-1][0]) ** 2 + (
                chainpath.Coordinates[-1] - steppnt[-1][1]) ** 2) ** 0.5
        if step0 == 0:
            frtstep = 0
        elif step0 > step1:
            frtstep = float(step)
        else:
            frtstep = leng[-1] - len(steppnt) * float(step)
        return frtstep
