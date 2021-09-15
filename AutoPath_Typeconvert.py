#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: LZH
win32com与cad中数据类型转换
"""

# imports==========
import pythoncom
import win32com.client


# =========================================================
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
