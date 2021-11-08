#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
轨迹函数
"""

# imports==========
import math
import time

import AutoPath_TypeConvert as Tc


# =========================================================
class DoPath:
    def __init__(self, oop, doc, msp):
        self.oop = oop
        self.doc = doc
        self.msp = msp
        self.layerobjs = self.create_layer()

    def do_pendulum(self, select, dirvalue, swingstate_value, chainpath, step, chainbracing, bracing, carnum, pitch,
                    swingleng):
        """摆杆 - 轨迹/动画"""
        steppnt = self.pathpnt(chainpath, step)
        # for lay in self.layerobjs:
        #     lay.LayerOn = False  # 关闭图层
        self.leng = self.pline_length(chainpath)
        self.frtstep = self.frtsteppnt(chainpath, steppnt, step, self.leng)
        if select % 10 == 1:
            inserttext = '轨道长度：' + str(int(self.leng[-1])) + 'mm\n' + '轨迹步长：' + step + 'mm\n' + \
                         '摆杆间距：' + str(bracing) + 'mm\n' + '摆杆长度：' + str(swingleng) + 'mm'
        else:
            inserttext = '轨道长度：' + str(int(self.leng[-1])) + 'mm\n' + '工件节距：' + str(pitch) + 'mm\n' + \
                         '摆杆间距：' + str(bracing) + 'mm\n' + '摆杆长度：' + str(swingleng + 252.75) + 'mm'
        self.insert_mt(chainpath, inserttext)
        self.j = 1
        self.jj = 0  # 判断是否为第一次插入
        for i in steppnt:
            chainplate = []
            fswing = []
            bswing = []
            car = []
            if ((self.j - 1) * float(step) + self.frtstep) > (
                    (carnum - 1) * pitch + bracing + 2 * chainbracing):
                # print((j - 1) * float(step.get()) + self.frtstep)
                for num in range(carnum):
                    layernum = num % len(self.layerobjs)
                    self.doc.ActiveLayer = self.doc.Layers(self.layerobjs[layernum].Name)
                    nowdist = (self.j - 1) * float(step) + self.frtstep - (num - 1) * pitch
                    if num == 0:
                        inspnt0 = i
                    else:
                        inspnt0 = self.find_distpnt(chainpath, self.leng, nowdist, inspnt0, pitch)
                    self.swingpnt = []
                    for swingnum in range(2):
                        if swingnum == 1:
                            inspnt0 = self.find_distpnt(chainpath, self.leng, nowdist, inspnt0, bracing)
                        self.find_insertpnt(chainpath, step, chainbracing, bracing, inspnt0, swingnum, i)
                        self.insert_block('ChainPlate', chainplate, swingnum, dirvalue)
                        self.circlebr.Delete()
                        self.swingpnt.append(self.find_isotrianglepnt(inspnt0))
                        if swingnum == 0 and swingstate_value == 1:  # 前摆杆竖直
                            swingpnt0 = Tc.vtpnt(self.swingpnt[0][0], self.swingpnt[0][1])
                            fswing.append(self.msp.InsertBlock(swingpnt0, 'FSwing', 1, 1, 1, 0))
                            car_fpnt0 = [self.swingpnt[0][0], self.swingpnt[0][1] - swingleng]
                            car_fpnt = Tc.vtpnt(car_fpnt0[0], car_fpnt0[1])
                            self.circlecar = self.msp.AddCircle(car_fpnt, bracing)
                        elif swingnum == 1 and swingstate_value == 2:  # 后摆杆竖直
                            swingpnt0 = Tc.vtpnt(self.swingpnt[1][0], self.swingpnt[1][1])
                            bswing.append(self.msp.InsertBlock(swingpnt0, 'BSwing', 1, 1, 1, 0))
                            car_bpnt0 = [self.swingpnt[1][0], self.swingpnt[1][0] - swingleng]
                            car_bpnt = Tc.vtpnt(car_bpnt0[0], car_bpnt0[1])
                            self.circlecar = self.msp.AddCircle(car_bpnt, bracing)
                        else:
                            swingpnt0 = Tc.vtpnt(self.swingpnt[-1][0], self.swingpnt[-1][1])
                            self.circleswing = self.msp.AddCircle(swingpnt0, swingleng)
                    if swingstate_value == 1:  # 前摆杆竖直，插入后摆杆和工件
                        self.insert_car(swingstate_value, car, bswing)
                    else:  # 后摆杆竖直，插入前摆杆和工件
                        self.insert_car(swingstate_value, car, fswing)
                    self.circlecar.Delete()
                    self.circleswing.Delete()
                if select % 10 == 2:
                    self.animation_next_or_copy(chainplate, fswing, bswing, car)
            self.j += 1
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

    def do_pendulum_diptank(self, swing):
        """摆杆 - 浸入即出槽分析"""
        pass

    def do_2trolley(self, select, dirvalue, car_name, chainpath, step, bracing, carnum, pitch):
        """台车 - 轨迹/动画"""
        steppnt = self.pathpnt(chainpath, step)
        # for lay in self.layerobjs:
        #     lay.LayerOn = False  # 关闭图层
        self.leng = self.pline_length(chainpath)
        self.frtstep = self.frtsteppnt(chainpath, steppnt, step, self.leng)
        if select % 10 == 1:
            inserttext = '轨道长度：' + str(int(self.leng[-1])) + 'mm\n' + '轨迹步长：' + step + 'mm'
        else:
            inserttext = '轨道长度：' + str(int(self.leng[-1])) + 'mm\n' + '工件节距：' + str(pitch) + 'mm'
        self.insert_mt(chainpath, inserttext)
        self.j = 1
        self.jj = 0  # 判断是否为第一次插入
        for i in steppnt:
            car = []
            if ((self.j - 1) * float(step) + self.frtstep) > ((carnum - 1) * pitch + bracing):
                # print((j - 1) * float(step.get()) + self.frtstep)
                for num in range(carnum):
                    layernum = num % len(self.layerobjs)
                    self.doc.ActiveLayer = self.doc.Layers(self.layerobjs[layernum].Name)
                    nowdist = (self.j - 1) * float(step) + self.frtstep - (num - 1) * pitch
                    if num == 0:
                        inspnt0 = i
                    else:
                        inspnt0 = self.find_distpnt(chainpath, self.leng, nowdist, inspnt0, pitch)
                    self.find_insertpnt(chainpath, step, bracing, pitch, inspnt0, num, i)
                    self.insert_block(car_name, car, num, dirvalue)
                    self.circlebr.Delete()
                if select % 10 == 2:
                    self.animation_next_or_copy(car)
            self.j += 1
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        self.doc.ActiveLayer = self.doc.Layers("0")

    def find_insertpnt(self, chainpath, step, bracing, pitch, inspnt0, num, i):
        self.inspnt = Tc.vtpnt(inspnt0[0], inspnt0[1])
        self.circlebr = self.msp.AddCircle(self.inspnt, bracing)
        self.intpnts = chainpath.IntersectWith(self.circlebr, 0)
        self.brpnt = []
        self.angbr = []
        for k in range(len(self.intpnts) // 3):
            self.brpnt.append([self.intpnts[k * 3], self.intpnts[k * 3 + 1]])
            self.angbr.append(
                self.doc.Utility.AngleFromXAxis(Tc.vtpnt(self.brpnt[-1][0], self.brpnt[-1][1]), self.inspnt))
        for k in range(len(self.leng)):
            if ((self.j - 1) * float(step) + self.frtstep - num * pitch) < self.leng[k]:
                self.jj += 1
                # print(leng[k])
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

    def insert_block(self, block_name, blocklist, num, dirvalue):
        """插入块并调整角度"""
        if dirvalue == 1:  # 工件向右
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

    def insert_car(self, state_value, blocklist0, blocklist1, block_name='Skid&Body'):
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
            blocklist1.append(self.msp.InsertBlock(bswing_pnt, 'BSwing', 1, 1, 1, math.pi * 3 / 2 - swing_angel))
        else:  # 后摆杆竖直，以circlecar与circleswing交点为插入点，第二点为circlecar圆心
            fswing_pnt = Tc.vtpnt(self.swingpnt[0][0], self.swingpnt[0][1])
            car_angle = self.doc.Utility.AngleFromXAxis(car_inspnt1, car_inspnt0)
            swing_angel = self.doc.Utility.AngleFromXAxis(fswing_pnt, car_inspnt0)
            blocklist0.append(self.msp.InsertBlock(car_inspnt1, block_name, 1, 1, 1, car_angle))
            blocklist1.append(self.msp.InsertBlock(fswing_pnt, 'FSwing', 1, 1, 1, math.pi * 3 / 2 - swing_angel))

    def pathpnt(self, pline, step):
        """沿轨道线按步长求插入点"""
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
        # print(steppnt)
        slt.Erase()
        slt.Delete()
        self.doc.ActiveLayer = self.doc.Layers("0")
        LayerObj.Delete()
        return steppnt

    def find_distpnt(self, pline, leng, nowdist, frtpnt, dist):
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
                    centpnt = Tc.vtpnt(centpnt[0], centpnt[1])
                    startpnt = Tc.vtpnt(pline.Coordinates[vertex * 2],
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

    def find_isotrianglepnt(self, inspnt0, height=140, direction=0):
        frtpnt = [inspnt0[0], inspnt0[1]]
        sndpnt = self.brpnt[self.index]
        if direction == 0:
            if frtpnt[0] == sndpnt[0]:
                trdpnt_x = frtpnt[0] - height
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2
            elif frtpnt[1] == sndpnt[1]:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2
                trdpnt_y = frtpnt[1] - height
            else:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 + height / (
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
            else:
                trdpnt_x = (frtpnt[0] + sndpnt[0]) / 2 - height / (
                        ((sndpnt[0] - frtpnt[0]) / (sndpnt[1] - frtpnt[1])) ** 2 + 1) ** 0.5
                trdpnt_y = (frtpnt[1] + sndpnt[1]) / 2 + height / (
                        ((sndpnt[1] - frtpnt[1]) / (sndpnt[0] - frtpnt[0])) ** 2 + 1) ** 0.5
        trdpnt = [trdpnt_x, trdpnt_y]
        return trdpnt

    def create_layer(self):
        """创建绘图层"""
        clrnums = [1, 2, 3, 4, 6]  # 图层颜色列表
        layernames = ["轨迹层_1", "轨迹层_2", "轨迹层_3", "轨迹层_4", "轨迹层_5"]  # 图层名称列表
        layerobjs = [self.doc.Layers.Add(i) for i in layernames]  # 批量创建图层
        for i in range(len(layerobjs)):
            layerobjs[i].color = clrnums[i]
        return layerobjs

    def insert_mt(self, chainpath, inserttext):
        """绘制轨迹/仿真时，在轨道线上方插入文字说明"""
        textpnt = Tc.vtpnt((min(chainpath.Coordinates[0:-2:2]) + max(chainpath.Coordinates[0:-2:2])) / 2 - 1500,
                           max(chainpath.Coordinates[1:-1:2]) + 3000)
        mt = self.msp.AddMText(textpnt, 3000, inserttext)
        mt.Height = 200
        mt.styleName = 'STANDARD'

    def animation_next_or_copy(self, *blocks):
        """生成一组仿真后，选择继续仿真/保留当前结果"""
        for lay in self.layerobjs:
            lay.LayerOn = True  # 打开图层
        # 选择是否保留当前结果
        self.doc.Utility.InitializeUserInput(0, "C, c")
        doconti = 'a'
        while doconti.lower() != 'c' and doconti != '':
            doconti = self.doc.Utility.GetKeyword("继续仿真下一工件位置或 [保留当前工件位置(C)]: ")
        time.sleep(0.3)
        if doconti == '':
            for block in blocks:
                for eachblock in block:
                    eachblock.Delete()
        for lay in self.layerobjs:
            lay.LayerOn = False  # 关闭图层

    @staticmethod
    def pline_length(pline):
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
    def frtsteppnt(chainpath, steppnt, step, leng):
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
