#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
cadpath_20210817重构使用class，合并轨迹与仿真选项卡
"""

import pythoncom
import win32com.client


class Typeconvert(object):
    def vtpnt(self, x, y, z=0):
        """坐标点转化为浮点数"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def vtobj(self, obj):
        """转化为对象数组"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

    def vtfloat(self, list1):
        """列表转化为浮点数"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list1)

    def vtint(self, list1):
        """列表转化为整数"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list1)

    def vtvariant(self, list1):
        """列表转化为变体"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list1)
