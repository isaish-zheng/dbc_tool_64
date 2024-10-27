#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : ZYD
# @Time    : 2024/7/12 上午8:34
# @version : V1.0.0
# @function:

##############################
# Module imports
##############################

from tkinter import BooleanVar, IntVar, StringVar
from typing import NamedTuple, List

from tkui.tktypes import TkFrame


##############################
# Constant definitions
##############################


##############################
# Model API function declarations
##############################

class FramesIdx(NamedTuple):
    """
    excel2dbc和dbc2excel界面的索引
    e2d: 0
    d2e: 1
    """
    E2D: int = 0
    D2E: int = 1


class UIModel(object):
    """
    界面的数据模型
    """

    def __init__(self):
        """构造函数"""

        # 常量数据
        self.FRAMES = FramesIdx()  # 界面切换按钮可选的索引

        # 持久数据
        self.opened_excel_filepath: str = '' # 打开的excel文件路径
        self.opened_excel_tpl_filepath: str = '' # 打开的excel模板文件路径
        self.opened_dbc_filepath: str = '' # 打开的dbc文件路径
        self.current_frame_idx: int = 0  # 当前显示的界面索引
        self.radiobutton_dest_frame_idx = IntVar(value=0)  # 界面切换单选按钮关联的状态变量
        self.radiobutton_sort_signals = StringVar(value='')  # 信号排序单选按钮关联的状态变量
        self.radiobutton_sort_messages = StringVar(value='')  # 消息排序单选按钮关联的状态变量
        self.checkbutton_excel_auto_fit = BooleanVar(value=True)  # Excel自适应复选按钮关联的状态变量

        # 临时数据
        self.frames: List[TkFrame] = []  # e2d和d2e界面的实例

    @property
    def is_excel_auto_fit(self) -> bool:
        """
        是否自动调整excel

        :return: 是否自动调整excel
        :rtype: bool
        """
        return self.checkbutton_excel_auto_fit.get()

    @is_excel_auto_fit.setter
    def is_excel_auto_fit(self, value: bool) -> None:
        """
        设置是否自动调整excel

        :param value: 是否自动调整excel
        :type value: bool
        """
        self.checkbutton_excel_auto_fit.set(value)

    @property
    def mode_sort_signals(self) -> str:
        """
        信号排序方式

        :return: 信号排序方式
        :rtype: str
        """
        return self.radiobutton_sort_signals.get()

    @mode_sort_signals.setter
    def mode_sort_signals(self, value: str) -> None:
        """
        设置信号排序方式

        :param value: 信号排序方式
        :type value: str
        """
        self.radiobutton_sort_signals.set(value)

    @property
    def mode_sort_messages(self) -> str:
        """
        消息排序方式

        :return: 消息排序方式
        :rtype: str
        """
        return self.radiobutton_sort_messages.get()

    @mode_sort_messages.setter
    def mode_sort_messages(self, value: str) -> None:
        """
        设置消息排序方式

        :param value: 消息排序方式
        :type value: str
        """
        self.radiobutton_sort_messages.set(value)
