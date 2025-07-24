#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : ZYD
# @Time    : 2024/7/12 上午8:35
# @version : V1.0.0
# @function:

##############################
# Module imports
##############################


import configparser # 用于处理配置文件
import copy
from math import pi, sin
import time
from tkinter import END, filedialog  # 类型提示
import traceback  # 用于获取异常详细信息
from typing import Callable, Optional, Union, Dict, List

import cantools
from cantools.database import utils as ct_utils, Database
from cantools.database.can.attribute import Attribute
from cantools.database.can.attribute_definition import AttributeDefinition
from cantools.database.can.formats.dbc_specifics import DbcSpecifics
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment  # 导入格式
from openpyxl.comments import Comment  # 导入批注
from openpyxl.worksheet.datavalidation import DataValidation  # 导入验证单元
import pandas as pd
import win32com.client as win32
from pandas import DataFrame
from pandas.core.dtypes.inference import is_float

from tkui.tktypes import *
from .view import UIModel, UIView


##############################
# Controller API function declarations
##############################

class GenError(Exception):
    """
    自定义Gen异常类
    
    :param msg: 要显示的异常消息
    :type msg: str
    """

    def __init__(self, msg: str):
        """
        构造函数
        
        """
        self.msg = msg

    def __str__(self):
        return f"{self.msg}"


class UICtrl(object):
    """
    视图的业务逻辑处理

    :param model: 视图的数据模型
    :type model: UIModel
    :param view: 视图
    :type view: UIView
    :param cfg_path: 配置文件路径
    :type cfg_path: str
    """

    def __init__(self,
                 model: UIModel,
                 view: UIView,
                 cfg_path: str) -> None:
        """
        构造函数
        
        """
        self.model = model
        self.view = view
        self.view.set_presenter(self)

        self.__cfg_path = cfg_path

        # 初始化配置
        self.ini_config()
        # 加载配置
        self.load_config()
        # 配置文件中的current_frame_idx为str，model中为int，所以此处要特别处理
        self.model.radiobutton_dest_frame_idx.set(model.current_frame_idx)  # 配置文件中加载的值类型为字符串
        self.model.current_frame_idx = int(model.current_frame_idx)  # 转换为整型

    def text_log(self, txt: str, *args, **kwargs) -> None:
        """
        向文本显示框写入信息

        :param txt: 待写入的信息
        :type txt: str
        :param args: 位置参数，第一个参数为文字颜色
                    None-灰色,'done'-绿色,'warning'-黄色,'error'-红色
        :param kwargs: 关键字参数（未使用）
        """

        def get_str_time() -> str:
            """
            获取当前时间，格式化字符串"%Y/%m/%d %H:%M:%S"

            :return: 当前时间
            :rtype: str
            """
            time_now = time.strftime("%Y/%m/%d %H:%M:%S", time.localtime())  # 时间戳
            return str(time_now)

        # color = COLOR_LABEL_FG
        color = '#787878'
        if args:
            if args[0] == 'done':
                color = 'green'
            elif args[0] == 'warning':
                color = '#cc5d20'
            elif args[0] == 'error':
                color = 'red'
        self.view.text_info.config(state='normal')
        self.view.text_info.insert(END, get_str_time() + ' ' + txt + '\n', color)
        self.view.text_info.config(state='disabled')
        self.view.text_info.see(END)
        self.view.text_info.tag_config(tagName=color, foreground=color)

    def ini_config(self) -> None:
        """
        初始化配置文件，若配置文件不存在，则新建配置

        """
        try:
            if not os.path.isfile(self.__cfg_path):
                conf = configparser.ConfigParser()
                section = 'user'
                conf.add_section(section)
                conf.set(section, 'opened_excel_filepath', '')
                conf.set(section, 'opened_excel_tpl_filepath', '')
                conf.set(section, 'opened_dbc_filepath', '')
                conf.set(section, 'current_frame_idx', '0')
                conf.set(section, 'mode_sort_signals', 'SOURCE_FILE')
                conf.set(section, 'mode_sort_messages', 'SOURCE_FILE')
                conf.set(section, 'is_excel_auto_fit', 'True')
                # 保存配置文件
                with open(self.__cfg_path, 'w', encoding='utf-8') as f:
                    conf.write(f)
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    def load_config(self) -> None:
        """
        加载配置文件中的配置保存到界面的model中

        """
        try:
            conf = configparser.ConfigParser()
            # 读取配置文件中的各项配置,通过ui显示
            conf.read(filenames=self.__cfg_path, encoding='utf-8')
            sections = conf.sections()
            for section in sections:
                for option in conf.options(section):
                    if hasattr(self.model, option):
                        setattr(self.model, option, conf.get(section, option))
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    def save_config(self) -> None:
        """
        保存需要的配置项到配置文件
        """
        try:
            conf = configparser.ConfigParser()
            conf.read(self.__cfg_path, encoding='utf-8')
            sections = conf.sections()
            for section in sections:
                for option in conf.options(section):
                    if hasattr(self.model, option):
                        conf.set(section, option, str(getattr(self.model, option)))
            with open(self.__cfg_path, 'w', encoding='utf-8') as f:
                conf.write(f)
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    def handler_on_closing(self) -> None:
        """
        执行指定任务后关闭界面窗口

        """
        try:
            self.save_config()
            self.view.quit()
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    def handler_change_frame(self) -> None:
        """
        切换frame界面

        """

        dest_idx = self.model.radiobutton_dest_frame_idx.get()
        if self.model.current_frame_idx == dest_idx:
            return
        else:
            def fluent_change() -> None:
                nonlocal i
                i = i + 1
                if dest_idx == self.model.FRAMES.E2D:
                    self.model.frames[dest_idx + 1].place(
                        x=self.view.get_dpi(-WIDTH_ROOT_WINDOW) * sin(i * pi / 180))
                    self.model.frames[dest_idx].place(
                        x=self.view.get_dpi(WIDTH_ROOT_WINDOW) - self.view.get_dpi(WIDTH_ROOT_WINDOW) * sin(
                            i * pi / 180))
                elif dest_idx == self.model.FRAMES.D2E:
                    self.model.frames[dest_idx - 1].place(
                        x=self.view.get_dpi(-WIDTH_ROOT_WINDOW) * sin(i * pi / 180))
                    self.model.frames[dest_idx].place(
                        x=self.view.get_dpi(WIDTH_ROOT_WINDOW) - self.view.get_dpi(WIDTH_ROOT_WINDOW) * sin(
                            i * pi / 180))
                if i < 90:
                    self.view.after(8, fluent_change, )  # 8ms移动一次

            self.model.current_frame_idx = dest_idx
            i = 0
            fluent_change()

    def handler_open_file(self, filetype: str, **kwargs) -> str:
        """
        弹出文件对话框，打开指定类型的文件

        :param filetype: 文件类型
        :type filetype: str
        :param kwargs: 关键字参数，lbl为显示已打开文件路径的标签控件，dir为要打开文件的初始路径
        :type kwargs: dict
        :return: 打开的文件路径
        :rtype: str
        """
        if filetype == 'DBC':
            file_format = '.dbc'
        elif filetype == 'Excel模版':
            file_format = '.xlsx'
        else:
            file_format = '.xlsx'
        lbl = kwargs.get('lbl')
        directory = str(kwargs.get('dir'))

        # 输出日志
        self.text_log(f'打开{filetype}文件中...')
        # 打开文件对话框
        file_path = filedialog.askopenfilename(
            # 默认扩展名，.号可带可不带
            defaultextension=file_format,
            # 文件类型选项
            filetypes=[(filetype + '文件', file_format)],
            # 初始目录，默认当前目录
            initialdir=os.path.dirname(directory) if directory else os.getcwd(),
            # 初始文件名，默认为空
            # initialfile='test',
            # 打开的位置，默认是根窗口
            parent=self.view,
            # 窗口标题
            title='打开' + filetype + '文件')
        # 将打开的文件路径写入标签
        if lbl:
            lbl.config(text=file_path)
        if filetype == 'DBC':
            self.model.opened_dbc_filepath = file_path
        elif filetype == 'Excel模版':
            self.model.opened_excel_tpl_filepath = file_path
        else:
            self.model.opened_excel_filepath = file_path
        if file_path:
            # 输出日志
            self.text_log('已打开' + filetype + '文件' + file_path, 'done')
        else:
            # 输出日志
            self.text_log('路径无效，未选择' + filetype + '文件' + file_path, 'warning')
        return file_path

    def handler_save_e2d(self) -> None:
        """
        生成DBC按钮的回调函数

        """
        try:
            self.__save_e2d()
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    def handler_save_d2e(self) -> None:
        """
        生成Excel按钮的回调函数

        """
        try:
            self.__save_d2e()
        except Exception as e:
            self.text_log(f'发生异常 {e}', 'error')
            self.text_log(f"{traceback.format_exc()}", 'error')

    @staticmethod
    def handler_open_folder(path: str, mode: str) -> None:
        """
        从资源管理器打开文件或文件夹

        :param path: 文件路径
        :type path: str
        :param mode: 打开模式，'file'为打开文件，'folder'为打开文件夹
        :type mode: str
        """

        if os.path.isfile(path) and mode == 'folder':
            path = os.path.dirname(path)
        os.startfile(path)

    def __get_sort_signals(self) -> Optional[Callable]:
        """
        获取保存文件的消息中的信号排序方式

        :return: 排序函数
        :rtype: Callable | None
        """
        mode_sort_signals = self.model.mode_sort_signals
        if mode_sort_signals == 'START_BIT_ASC':
            return ct_utils.sort_signals_by_start_bit
        elif mode_sort_signals == 'START_BIT_DESC':
            return ct_utils.sort_signals_by_start_bit_reversed
        else:
            return None

    def __sort_messages(self, msg_list: list):
        """
        对保存到文件的消息进行排序，返回排序后的消息列表

        :param msg_list: 消息列表
        :type msg_list: list
        :return: 排序后的消息列表
        :rtype: list
        """
        mode_sort_messages = self.model.mode_sort_messages
        if mode_sort_messages == 'NAME_ASC':
            msg_list.sort(key=lambda msg: (msg.name, msg.frame_id))
        elif mode_sort_messages == 'ID_ASC':
            msg_list.sort(key=lambda msg: (msg.frame_id, msg.name))
        return msg_list

    def __save_file(self, filetype: str, data: Union[Database, DataFrame], save_fun: Callable, **kwargs) -> str:
        """
        根据指定文件类型、路径、数据和保存方法弹出保存文件对话框，将结果赋值给标签显示

        :param filetype: 文件类型,'DBC'或'Excel'
        :type filetype: str
        :param data: 数据,Database或DataFrame
        :type data: Database | DataFrame
        :param save_fun: 保存方法
        :type save_fun: Callable
        :param kwargs: 关键字参数,lbl为标签，dir为路径
        :type kwargs: dict
        :return: 保存的文件路径
        :rtype: str
        """
        # 获取参数
        lbl = kwargs.get('lbl')
        directory = kwargs.get('dir')

        if filetype == 'DBC':
            fileformat = '.dbc'
            initialfile = directory[directory.rfind('/') + 1:directory.rfind('.xlsx')] + '_New' + filetype
        else:
            fileformat = '.xlsx'
            initialfile = directory[directory.rfind('/') + 1:directory.rfind('.dbc')] + '_New' + filetype

        # 输出日志
        self.text_log('保存' + filetype + '文件中...')
        # 保存文件对话框
        savepath = filedialog.asksaveasfilename(
            # 默认扩展名，.号可带可不带
            defaultextension=fileformat,
            # 文件类型选项
            filetypes=[(filetype + ' Files', fileformat)],
            # 初始目录，默认当前目录
            initialdir=dir if dir else os.getcwd(),
            # 初始文件名，默认为空
            initialfile=initialfile,
            # 打开的位置，默认是根窗口
            parent=self.view,
            # 窗口标题
            title='保存' + filetype)
        if savepath:
            # 保存文件
            isDone = save_fun(data, savepath)
            if isDone:
                # 将保存的文件路径写入标签
                if lbl:
                    lbl.config(text=savepath)
                # 输出日志
                self.text_log('已保存' + filetype + '文件' + savepath, 'done')
                # 询问是否从资源管理器打开文件夹
                isok = messagebox.askyesnocancel(title='是否打开' + filetype + '？',
                                                 message='是(Y)：打开文件\n否(N)：打开文件路径')
                if isok is True:
                    self.handler_open_folder(savepath, 'file')
                elif isok is False:
                    self.handler_open_folder(savepath[0: savepath.rfind('/')], 'folder')
        else:
            # 将保存的文件路径写入标签
            lbl.config(text=savepath)
            # 输出日志
            self.text_log('路径无效，未保存' + filetype + '文件' + savepath, 'warning')

        return savepath

    @staticmethod
    def __autofit_Excel(path: str, row_sum: int, column_sum: int) -> None:
        """
        自动调整Excel格式

        :param path: 文件路径
        :type path: str
        :param row_sum: 行数
        :type row_sum: int
        :param column_sum: 列数
        :type column_sum: int
        """

        # 设置VBA宏并调用，自动调整列宽和自动换行
        xl = win32.Dispatch('Excel.Application')
        # xl = win32.gencache.EnsureDispatch('Excel.Application')
        # xl = win32.EnsureDispatch('Excel.Application')
        xl.Visible = True
        xlbook = xl.Workbooks.Open(path)
        xlmodule = xlbook.VBProject.VBComponents.Add(1)
        xlmodule.Name = '我的模块'
        sel_zone = '\"' + 'A1:' + chr(ord('A') + column_sum - 1) + str(row_sum) + '\"'
        code = f"""Sub MyVBA()
                    Range({sel_zone}).Select
                    Selection.Columns.AutoFit
                    With Selection
                        .WrapText = True
                    End With
                    Selection.Columns.AutoFit
                    With ActiveWindow
                        .WindowState = xlMaximized
                        .SplitColumn = 0
                        .SplitRow = 1
                    End With
                    ActiveWindow.FreezePanes = True
                    ActiveWorkbook.Save
                    End Sub"""
        xlmodule.CodeModule.AddFromString(code)
        xlbook.Application.Run('我的模块.MyVBA')
        xl.Application.Quit()

    def __save_e2d(self) -> None:
        """
        从excel数据生成和保存DBC文件，将结果赋值给标签显示

        """

        def _load_excel(filepath: str) -> pd.DataFrame:
            """
            加载Excel,返回pandas.DataFrame,默认加载第一张Sheet

            :param filepath: excel文件路径
            :type filepath: str
            :return: pandas.DataFrame对象
            :rtype: pandas.DataFrame
            """
            # 输出日志
            self.text_log('加载Excel数据中...')

            df = pd.read_excel(filepath)  # 读取Excel文件，excel首行作为df的列索引，第二行即首个数据行的行索引0
            # print(df.to_string())  # 打印DataFrame，完整打印
            # print(df.shape)  # 打印DataFrame的行数和列数,1基
            # print(df.columns)  # 打印DataFrame的列索引，是一个含有列索引列表的Index对象
            # print(df.loc[0, 'Msg_ID'])  # 按照行和列进行数据筛选
            # print(df.loc[0])  # 打印DataFrame的行索引为0的数据
            # 输出日志
            self.text_log('加载Excel数据完成', 'done')
            return df

        def _generate_dbc(df: pd.DataFrame) -> cantools.db.Database:
            """
            处理加载excel后的DataFrame数据，生成Database

            :param df: pandas.DataFrame对象
            :type df: pandas.DataFrame
            :return: Database对象
            :rtype: cantools.db.Database
            """
            # 输出日志
            self.text_log('生成DBC数据库中...')

            # 自定义属性：消息发送方式，由此属性定义可生成关联到消息的属性
            msg_send_type_list = ['Cycle', 'Event']
            attr_def_dict = dict()
            # attr_def_dict['GenMsgCycleTime'] = AttributeDefinition('GenMsgCycleTime',
            #                                                        default_value=0,
            #                                                        kind='BO_',
            #                                                        type_name='INT',
            #                                                        minimum=0,
            #                                                        maximum=2 ** 16 - 1)
            attr_def_dict['GenMsgSendType'] = AttributeDefinition('GenMsgSendType',
                                                                  default_value='',
                                                                  kind='BO_',
                                                                  type_name='ENUM',
                                                                  choices=msg_send_type_list)
            # attr_dict['GenMsgCycleTime'] = Attribute(value=0,
            #                                          definition=attr_def_dict['GenMsgCycleTime'])

            msg_idx_in_df_list = []  # df表中消息的行索引列表
            sig_obj_of_msg_list = []  # 某个消息下所有的信号对象列表,cantools.database.can.signal.Signal
            msg_obj_list = []  # 数据库中所有的消息对象列表, cantools.database.can.message.Message

            # 查询所有消息在df表中的行索引，保存在列表中
            for row_idx in range(0, df.shape[0]):  # 遍历DataFrame的所有数据行(不含首行列索引)0基，shape为(rows,cols)1基
                row = df.loc[row_idx]  # 行数据
                if str(row['Msg_ID']).startswith('0x'):  # 行数据为消息
                    msg_idx_in_df_list.append(row_idx)

            # 根据df表内容生成消息对象列表
            for idx in range(0, len(msg_idx_in_df_list)):
                # 查询一个消息下的信号区域起始索引，范围为[start,stop-1)
                sig_row_start = 0
                sig_row_stop = 0
                sig_obj_of_msg_list.clear()
                if idx < len(msg_idx_in_df_list) - 1:
                    # 非最后一个消息
                    sig_row_start = msg_idx_in_df_list[idx] + 1  # start为当前消息行的下一行索引
                    sig_row_stop = msg_idx_in_df_list[idx + 1]  # stop下一消息行索引,即当前消息区域的最后一行索引+1
                else:
                    # 最后一个消息
                    sig_row_start = msg_idx_in_df_list[idx] + 1  # start为当前消息行的下一行索引
                    sig_row_stop = df.shape[0]  # stop为df表的末尾行索引+1,即当前消息区域的最后一行索引+1

                # 将当前消息范围[start,stop-1)里的信号添加到信号对象列表中
                for row_idx in range(sig_row_start, sig_row_stop):  # 遍历信号区域的每行
                    row = df.loc[row_idx]  # 行数据
                    # 判断消息格式'Intel'或‘Motorola’
                    if str(row['Byte_Order']) == 'Intel':
                        byte_order = 'little_endian'
                    else:
                        byte_order = 'big_endian'
                    # 判断数据类型
                    is_signed = False
                    is_float =False
                    if str(row['Data_Type']) == 'Signed':
                        is_signed = True
                    elif str(row['Data_Type']) == 'Unsigned':
                        is_signed = False
                    else:
                        is_signed = False
                        is_float = True
                    # 判断Unit是否为NaN
                    if str(row['Unit']) == 'nan':
                        # print(f"unit:{row['Unit']}")
                        unit = None
                    else:
                        unit = str(row['Unit'])
                    # 判断Sender_Receiver是否为NaN
                    if str(row['Sender_Receiver']) == 'nan':
                        receivers = None
                    else:
                        # 去除每个字符串两端的空格，对于信号为接收节点列表
                        receivers = list()
                        raw_string = str(row['Sender_Receiver']).strip()
                        s_list = raw_string.split(';')
                        for s in s_list:
                            s = s.strip()
                            if s:
                                receivers.append(s)
                    # 判断Value_table是否为NaN
                    if str(row['Value_table']) == 'nan':
                        choices = None
                    else:
                        # 去除每个字符串两端的空格，分别处理十进制和十六进制数
                        choices = dict()
                        raw_string = str(row['Value_table']).strip()
                        s_list = raw_string.split(';')
                        for s in s_list:
                            s = s.strip()
                            if s:
                                k = s.split('>')[0].strip()
                                v = s.split('>')[1].strip()
                                if k.lower().startswith('0x') or k.lower().startswith('-0x'):
                                    choices[int(k, 16)] = v
                                else:
                                    choices[int(k)] = v
                    # 判断Comment是否为NaN
                    if str(row['Comment']) == 'nan':
                        cmt = None
                    else:
                        cmt = str(row['Comment'])
                    # 创建信号对象
                    sig_obj = cantools.db.Signal(
                        name=str(row['Signal_Name']),
                        start=int(row['Start_Bit']),
                        length=int(row['Signal_Size']),
                        byte_order=byte_order,
                        is_signed=is_signed,
                        initial=int(row['Init_Value']), # 初始值，物理量
                        # invalid = None,
                        scale=float(row['Factor']),
                        offset=float(row['Offset']),
                        minimum=float(row['Minimum']),
                        maximum=float(row['Maximum']),
                        unit=unit,
                        choices=choices,
                        comment=cmt,
                        receivers=receivers,
                        # is_multiplexer = False,
                        # multiplexer_ids = None,
                        # multiplexer_signal = None,
                        is_float = is_float, # 是否为浮点数,根据length为32或64自动判断为IEEE Float或IEEE
                    )
                    sig_obj_of_msg_list.append(sig_obj)  # 添加信号对象到列表

                # 生成消息列表
                row = df.loc[msg_idx_in_df_list[idx]]  # 消息行内容
                # 消息关联自定义属性：发送方式设定
                attr_dict = dict()
                attr_dict['GenMsgSendType'] = Attribute(value=msg_send_type_list.index(str(row['Msg_Send_Type'])),
                                                        definition=attr_def_dict['GenMsgSendType'])
                msg_dbcspecifics = DbcSpecifics(attributes=attr_dict,
                                                attribute_definitions=None,
                                                environment_variables=None,
                                                value_tables=None,
                                                attributes_rel=None,
                                                attribute_definitions_rel=None)
                # 生成消息
                # 判断Comment是否为NaN
                if str(row['Comment']) == 'nan':
                    cmt = None
                else:
                    cmt = str(row['Comment'])
                # 判断Sender_Receiver是否为NaN
                if str(row['Sender_Receiver']) == 'nan':
                    senders = None
                else:
                    # 去除每个字符串两端的空格，对于消息为发送节点列表
                    senders = list()
                    raw_string = str(row['Sender_Receiver']).strip()
                    s_list = raw_string.split(';')
                    for s in s_list:
                        s = s.strip()
                        if s:
                            senders.append(s)
                msg_obj = cantools.db.Message(
                    frame_id=int(row['Msg_ID'], 16),
                    name=str(row['Msg_Name']),
                    length=int(row['Msg_Length']),
                    # header_byte_order = 'big_endian',
                    # unused_bit_pattern = 0x00,
                    comment=cmt,
                    senders=senders,
                    # send_type=str(row['Msg_Send_Type']),
                    cycle_time=int(row['Cycle_ms']),
                    dbc_specifics=copy.deepcopy(msg_dbcspecifics),
                    is_extended_frame=bool(row['Is_Extended']),
                    is_fd=bool(row['Is_FD']),
                    # bus_name='PFC',
                    sort_signals=self.__get_sort_signals(),
                    signals=copy.deepcopy(sig_obj_of_msg_list)
                )
                msg_obj_list.append(msg_obj)  # 添加消息对象到列表
            # 定义节点
            # node1 = cantools.db.Node(
            #     name='VCU'
            # )

            # 自定义属性：消息发送方式
            dbc_dbcspecifics = DbcSpecifics(attributes=None,
                                            attribute_definitions=attr_def_dict,
                                            environment_variables=None,
                                            value_tables=None,
                                            attributes_rel=None,
                                            attribute_definitions_rel=None)
            # 定义数据库
            database = cantools.db.Database(messages=self.__sort_messages(msg_obj_list),
                                            # nodes=[node1, ],
                                            # version='1.0'
                                            dbc_specifics=dbc_dbcspecifics,
                                            sort_signals=self.__get_sort_signals(),
                                            )
            # 输出日志
            self.text_log('生成DBC数据库完成', 'done')
            return database

        # 函数_将数据库保存为DBC文件
        def _save_dbc(database: cantools.db.Database, savepath: str) -> bool:
            cantools.db.dump_file(database=database,
                                  filename=savepath,
                                  database_format='dbc',
                                  # sort_signals=ct_utils.sort_signals_by_start_bit,
                                  sort_signals=self.__get_sort_signals(),
                                  encoding='ANSI')
            return True

        # 执行
        lbl = self.view.lbl_save_dbc
        openpath = self.model.opened_excel_filepath

        df = _load_excel(filepath=openpath)
        db = _generate_dbc(df=df)
        self.__save_file(filetype='DBC',
                         data=db,
                         save_fun=_save_dbc,
                         lbl=lbl,
                         dir=openpath)

    def __save_d2e(self) -> None:
        """
        从dbc数据生成和保存Excel文件，将结果赋值给标签显示

        """
        # 渲染Excel消息数据区域，首元素[rowsum,colsum]区域大小，1基
        # 除首个元素外其余每个元素为[rowstart,colsum,rowstop]，为消息内容的起始行、所占列数、结束行，1基
        render_area = []

        # 定义数据库
        # database = cantools.db.load_file(filename=obj_tkui.opened_dbc_filepath,
        #                                  database_format='dbc',
        #                                  encoding='ANSI',
        #                                  sort_signals=__get_sort_signals(obj_tkui))

        def _load_dbc(path: str) -> cantools.db.Database:
            """
            加载DBC数据，返回Database

            :param path: DBC文件路径
            :type path: str
            :return: cantools.db.Database对象
            :rtype: cantools.db.Database
            """
            # 输出日志
            self.text_log('加载DBC数据中...')
            db = cantools.db.load_file(filename=path,
                                       database_format='dbc',
                                       encoding='ANSI',
                                       sort_signals=self.__get_sort_signals())
            if db is not None:
                self.text_log('加载DBC数据完成', 'done')
                return db
            else:
                msg = f'加载DBC数据失败，请检查文件路径是否正确：{path}'
                raise GenError(msg)

        def _load_tpl(path: str) -> pd.DataFrame:
            """
            加载Excel模版数据，返回DataFrame，默认加载第一张Sheet

            :param path: Excel文件路径
            :type path: str
            :return: pd.DataFrame对象
            :rtype: pd.DataFrame
            """
            # 输出日志
            self.text_log('加载Excel模版数据中...')
            df = pd.read_excel(path)
            if df is not None:
                self.text_log('加载Excel模版数据完成', 'done')
                return df
            else:
                msg = f'加载Excel模版数据失败，请检查文件路径是否正确：{path}'
                raise GenError(msg)

        # 函数_生成Excel，
        def _generate_excel(self, db: cantools.db.Database, df: pd.DataFrame) -> pd.DataFrame:
            """
            生成Excel数据，返回Database

            :param db: 加载dbc生成的Database对象
            :type db: cantools.db.Database
            :param df: 加载excel模板生成的DataFrame对象
            :type df: pd.DataFrame
            :return: pd.DataFrame对象，由dbc数据和excel模板生成的excel数据对象
            :rtype: pd.DataFrame
            """
            nonlocal render_area
            render_area = []
            # 输出日志
            self.text_log('生成Excel数据中...')
            # print(df.to_string())
            # print(df.shape[1])
            df.drop(index=df.index, axis=0, inplace=True)  # 删除行(axis=0),删除index指定的行，inplace=True在源数据操作
            # new_row = pd.Series({'Msg_ID': 0x18F0000})
            # df.append(new_row, ignore_index=True)
            # 通过赋值为每列设置数据类型
            data = {'Msg_ID': 0x00,
                    'Msg_Name': '',
                    'Msg_Length': 0,
                    'Msg_Send_Type': '',
                    'Cycle_ms': 0,
                    'Is_Extended': 'False',
                    'Is_FD': 'False',
                    'Signal_Name': '',
                    'Start_Bit': 0,
                    'Signal_Size': 0,
                    'Byte_Order': '',
                    'Data_Type': '',
                    'Init_Value': 0,
                    'Factor': 0.0,
                    'Offset': 0.0,
                    'Minimum': 0,
                    'Maximum': 0,
                    'Unit': '',
                    'Sender_Receiver': '',
                    'Value_table': '',
                    'Comment': '',
                    }
            df = df.append(data, ignore_index=True)  # 将索引数据添加到对象，不使用索引标签
            row_idx = 0
            # 按照指定方式进行消息排序
            self.__sort_messages(db.messages)
            # 遍历消息和信号，生成数据表
            for msg in db.messages:
                df.at[row_idx, 'Msg_ID'] = hex(msg.frame_id).upper().replace('X', 'x', 1) if msg.frame_id else 0x00
                df.at[row_idx, 'Msg_Name'] = str(msg.name) if msg.name else ''
                df.at[row_idx, 'Msg_Length'] = int(msg.length) if msg.length else 0
                df.at[row_idx, 'Msg_Send_Type'] = str(msg.send_type) if msg.send_type else ''
                df.at[row_idx, 'Cycle_ms'] = int(msg.cycle_time) if msg.cycle_time else 0
                df.at[row_idx, 'Is_Extended'] = 'True' if msg.is_extended_frame else 'False'
                df.at[row_idx, 'Is_FD'] = 'True' if msg.is_fd else 'False'
                df.at[row_idx, 'Signal_Name'] = None
                df.at[row_idx, 'Start_Bit'] = None
                df.at[row_idx, 'Signal_Size'] = None
                df.at[row_idx, 'Byte_Order'] = None
                df.at[row_idx, 'Data_Type'] = None
                df.at[row_idx, 'Init_Value'] = None
                df.at[row_idx, 'Factor'] = None
                df.at[row_idx, 'Offset'] = None
                df.at[row_idx, 'Minimum'] = None
                df.at[row_idx, 'Maximum'] = None
                df.at[row_idx, 'Unit'] = None
                # 将节点列表格式化为字符串
                senders = list(msg.senders) if msg.senders else []
                s_list = []
                if senders:
                    for sender in senders:
                        s_list.append(str(sender) + ';')
                df.at[row_idx, 'Sender_Receiver'] = ''.join(s_list)[:-1]  # 不含末尾的';'
                df.at[row_idx, 'Value_table'] = None
                df.at[row_idx, 'Comment'] = str(msg.comment) if msg.comment else ''
                row_idx = row_idx + 1  # 0基，不含索引行
                render_area.append([row_idx + 1, df.shape[1]])  # 1基
                # print(msg.name)
                for sig in msg.signals:
                    df.at[row_idx, 'Signal_Name'] = str(sig.name) if sig.name else ''
                    df.at[row_idx, 'Start_Bit'] = int(sig.start) if sig.start else 0
                    df.at[row_idx, 'Signal_Size'] = int(sig.length) if sig.length else 0
                    df.at[row_idx, 'Byte_Order'] = 'Motorola' if sig.byte_order == 'big_endian' else 'Intel'
                    if sig.is_float and sig.length == 32:
                        data_type = 'IEEE Float'
                    elif sig.is_float and sig.length == 64:
                        data_type = 'IEEE Double'
                    elif sig.is_signed:
                        data_type = 'Signed'
                    else:
                        data_type = 'Unsigned'
                    # print(data_type)
                    df.at[row_idx, 'Data_Type'] = data_type
                    # print(f'sig.is_signed:{sig.is_signed}')
                    # print(f'sig.is_float:{sig.is_float}')
                    df.at[row_idx, 'Init_Value'] = float(sig.initial) if sig.initial else 0
                    df.at[row_idx, 'Factor'] = float(sig.scale) if sig.scale else 0
                    df.at[row_idx, 'Offset'] = float(sig.offset) if sig.offset else 0
                    df.at[row_idx, 'Minimum'] = float(sig.minimum) if sig.minimum else 0
                    df.at[row_idx, 'Maximum'] = float(sig.maximum) if sig.maximum else 0
                    df.at[row_idx, 'Unit'] = str(sig.unit) if sig.unit else ''
                    # 将节点列表格式化为字符串
                    receivers = list(sig.receivers) if msg.receivers else []
                    s_list = []
                    if receivers:
                        for receiver in receivers:
                            s_list.append(str(receiver) + ';')
                    df.at[row_idx, 'Sender_Receiver'] = ''.join(s_list)[:-1]  # 不含末尾的';'
                    # 将真值表格式化为字符串
                    choices = dict(sig.choices) if sig.choices else {}
                    s_list = []
                    if choices:
                        for k, v in choices.items():
                            k_hex = hex(k)
                            if k_hex.startswith('0x'):
                                s_list.append('0x' + k_hex[2:].upper() + '>' + str(v) + ';\n')
                            elif k_hex.startswith('-0x'):
                                s_list.append('-0x' + k_hex[3:].upper() + '>' + str(v) + ';\n')
                            else:
                                self.text_log(f"真值表'{k}':'{v}'无法处理", 'error')
                                raise ValueError
                    df.at[row_idx, 'Value_table'] = ''.join(s_list)[:-1]  # 不含末尾的'\n'
                    df.at[row_idx, 'Comment'] = str(sig.comment) if sig.comment else ''
                    row_idx = row_idx + 1
                render_area[-1].append(row_idx + 1)  # 添加消息中最后一个信号所在的行
            render_area.insert(0, [df.shape[0] + 1, df.shape[1]])  # 添加首元素(行索引,列索引),1基

            # 输出日志
            self.text_log('生成Excel数据完成', 'done')
            return df

        # 根据给定范围渲染表格
        # area[0]为表格范围
        # 其他为各消息行位置区域
        def _render_excel(self, path: str, area: list, validate: Dict[str, List[str]], comment: Dict[str, List[str]]):
            """
            根据给定范围渲染表格

            :param path: 文件路径
            :type path: str
            :param area: 表格范围
            :type area: list
            :param validate: 验证规则
            :type validate: Dict[str, List[str]]
            :param comment: 标题栏批注
            :type comment: Dict[str, List[str]]
            """
            # 输出日志
            self.text_log('渲染Excel表格中...')
            wk = openpyxl.load_workbook(path)  # 加载已经存在的excel
            sheet = wk.active  # 打开excel中激活的表格
            sheet_props = sheet.sheet_properties
            sheet_props.outlinePr.summaryBelow = False  # 分级显示标识禁止显示在下方

            # 设置单元格边框
            thin = Side(border_style='thin', color='000000')  # 边框
            border = Border(left=thin, right=thin, top=thin, bottom=thin)  # 设置边框
            # for row in sheet['A1:E12']:
            #     for cell in row: cell.border = border
            row_sum, column_sum = area[0]
            for row in range(1, row_sum + 1):
                for column in range(1, column_sum + 1):
                    sheet.cell(row=row, column=column).border = border  # 填充单元格背景色

            # 设置消息行单元格颜色
            filler = PatternFill('solid', fgColor='00b0f0')  # 设置填充颜色
            area.pop(0)  # 移除area首元素
            for i in area:
                row, last, row_stop = i
                sheet.row_dimensions.group(row + 1, row_stop, hidden=True)  # 分级显示消息
                for column in range(1, last + 1):
                    sheet.cell(row=row, column=column).fill = filler  # 填充单元格背景色

            # 设置标题行单元格颜色和文字
            filler = PatternFill('solid', fgColor='8db4e2')  # 设置必填填充颜色
            align = Alignment(horizontal='center', vertical='bottom',  # 创建一个对齐对象并设置属性
                              wrap_text=False, indent=0,
                              text_rotation=90)
            for column in range(1, column_sum + 1):
                sheet.cell(row=1, column=column).fill = filler  # 填充单元格背景色
                sheet.cell(row=1, column=column).alignment = align  # 将对齐对象应用到单元格
            filler = PatternFill('solid', fgColor='daeef3')  # 设置选填填充颜色
            col_sum_optional = 4  # 选填列总数
            for i in range(0, col_sum_optional):
                sheet.cell(row=1, column=column - i).fill = filler  # 填充单元格背景色
                # sheet.cell(row=1, column=column).fill = filler  # 填充单元格背景色

            # 数据验证
            # Dict[str, List[str]],列表末尾附加此列名再Excel表中的列索引，如'D'、'L'
            for col_idx_name in list(validate.keys()):
                col_num = len(validate[col_idx_name]) - 1
                # print(col_num)
                col_idx = validate[col_idx_name][col_num]
                # print(col_idx)
                cell = col_idx + str(2) + ':' + col_idx + str(row_sum)
                formula1 = ','.join(validate[col_idx_name][0:col_num])
                formula1 = "\"" + formula1 + "\""
                # print(formula1)
                dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
                dv.add(cell)
                sheet.add_data_validation(dv)

            # 添加标题批注
            # Dict[str, List[str]], 列表末尾附加要批注的单元格在Excel中的索引，如'D1'、'L1'
            for col_idx_name in list(comment.keys()):
                cell = comment[col_idx_name][1]
                text = Comment(text=comment[col_idx_name][0],
                               author='ZYD',
                               height=200,
                               width=160)
                sheet[cell].comment = text

            # 保存文件
            wk.save(path)

            # 自动调整列宽和换行
            if self.model.is_excel_auto_fit:
                self.__autofit_Excel(path, row_sum, column_sum)

            # 输出日志
            self.text_log('渲染Excel表格完成', 'done')

        # 函数_将Excel数据保存为Excel文件
        def _save_excel(dataframe, savepath: str) -> bool:
            """
            将Excel数据保存为Excel文件

            :param dataframe: 数据
            :param savepath: 保存路径
            :type savepath: str
            :return: 若保存完成返回True
            :rtype: bool
            """
            nonlocal render_area
            dataframe.to_excel(savepath,
                               index=False,
                               encoding='ANSI')
            # 数据验证
            # Dict[str, List[str]],列表末尾附加此列名再Excel表中的列索引，如'D'、'L'
            validate = {'Msg_Send_Type': ['Cycle', 'Event'],
                        'Is_Extended': ['True', 'False'],
                        'Is_FD': ['True', 'False'],
                        'Byte_Order': ['Intel', 'Motorola'],
                        'Data_Type': ['Signed', 'Unsigned', 'IEEE Float', 'IEEE Double']
                        }
            for i in range(0, len(list(validate.keys()))):
                col_idx_name = list(validate.keys())[i]
                if col_idx_name in list(df.columns):
                    validate[col_idx_name].append(chr(ord('A') + list(dataframe.columns).index(col_idx_name)))
                else:
                    pass
            # print(validate)
            # 标题批注
            #Dict[str, List[str]], 列表末尾附加要批注的单元格在Excel中的索引，如'D1'、'L1'
            comment = {'Msg_ID': ["1、必填\n2、须为以0x开头表示的十六进制"],
                       'Msg_Name': ["1、必填\n2、不允许出现空格\n3、须为以C语言规范字符开头（不允许为数字、标点符号开头）"],
                       'Msg_Length': ["1、必填\n2、长度单位为Byte"],
                       'Msg_Send_Type': ["1、必填\n2、支持周期型（Cycle）和事件型(Event)"],
                       'Cycle_ms': ["1、必填\n2、周期"],
                       'Is_Extended': ["1、必填\n2、是否为扩展帧，只允许填True或False"],
                       'Is_FD': ["1、必填\n2、是否为CANFD帧，只允许填True或False，暂时无效"],
                       'Signal_Name': [
                           "1、必填\n2、不允许出现空格\n3、须为以C语言规范字符开头（不允许为数字、标点符号开头）"],
                       'Start_Bit': [
                           "1、必填\n2、Intel格式为小端格式，数据LSB-MSB为低位地址-高位地址，start填LSB\n3、Motorola格式为大端格式，数据LSB-MSB为高位地址-低位地址，start填MSB"],
                       'Signal_Size': ["1、必填\n2、注意Intel和Motorola的区别，填写错误会造成重叠"],
                       'Byte_Order': ["1、必填\n2、Intel小端格式、Motorola大端格式"],
                       'Data_Type': ["1、必填\n2、数据类型，只允许填Signed、Unsigned、IEEE Float、IEEE Double"],
                       'Init_Value': ["1、必填\n2、注意区别物理初始值还是总线初始值，暂时无效"],
                       'Factor': ["1、必填\n2、系数"],
                       'Offset': ["1、必填\n2、偏移量"],
                       'Minimum': ["1、必填\n2、物理值"],
                       'Maximum': ["1、必填\n2、物理值"],
                       'Unit': ["1、选填\n2、单位"],
                       'Sender_Receiver': ["1、选填\n2、节点（对于消息为发送节点，对于信号为接收节点）"],
                       'Value_table': ["1、选填\n2、真值表"],
                       'Comment': ["1、选填\n2、注释"],
                       }
            for i in range(0, len(list(comment.keys()))):
                col_idx_name = list(comment.keys())[i]
                if col_idx_name in list(df.columns):
                    comment[col_idx_name].append(chr(ord('A') + list(dataframe.columns).index(col_idx_name)) + '1')
                else:
                    pass
            # print(comment)
            _render_excel(self, savepath, render_area, validate, comment)
            return True

        # 执行
        lbl = self.view.lbl_save_excel
        openpath = self.model.opened_dbc_filepath
        tpl_path = self.handler_open_file('Excel模版', dir=openpath)
        if not tpl_path:
            return
        df = _load_tpl(tpl_path)
        db = _load_dbc(openpath)
        df = _generate_excel(self, db, df)
        savepath = self.__save_file('Excel', df, _save_excel, lbl=lbl, dir=openpath)
