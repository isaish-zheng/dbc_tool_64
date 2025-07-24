#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : ZYD
# @Time    : 2024/7/11 下午5:31
# @version : V1.0.0
# @function:


##############################
# Module imports
##############################

import base64  # Base64编码解码模块
import subprocess
from typing import Any

from tkui import icon
from tkui.tktypes import *

from .model import UIModel


##############################
# View API function declarations
##############################

class UIView(tk.Tk, GetDpiMixIn):
    """
    根窗口
    """

    WIDTH_OPERATION_FRAME = 800
    HEIGHT_TEXT_INFO = 340
    WIDTH_LABEL = 690
    HEIGHT_LABEL = 40

    def __init__(self) -> None:
        """构造函数"""
        # 操作系统使用程序自身的dpi适配
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        super().__init__()
        # 显示根窗口
        self.set_root()

    def set_presenter(self, presenter: Any) -> None:
        """
        设置presenter，presenter中含一系列方法，用于处理界面事件

        :param presenter: presenter
        :type presenter: Any
        """
        self.presenter = presenter
        # 窗口点击关闭触发的功能
        self.protocol('WM_DELETE_WINDOW', lambda: self.presenter.handler_on_closing())

    @staticmethod
    def show_warning(msg: str) -> None:
        """
        显示警告弹窗

        :param msg: 警告信息
        :type msg: str
        """
        messagebox.showwarning(title='警告', message=msg)

    def set_root(self):
        """
        设置根窗口

        """
        # root = tk.Tk()
        self.title("DBC Tool")
        # 窗口尺寸和位置
        x_pos = int((self.winfo_screenwidth() -
                     super().get_dpi(WIDTH_ROOT_WINDOW)) / 2)
        y_pos = int((self.winfo_screenheight() -
                     super().get_dpi(HEIGHT_ROOT_WINDOW + HEIGHT_WINDOW_MENU_BAR)) / 2)
        self.geometry(
            f"{super().get_dpi(WIDTH_ROOT_WINDOW)}x"
            f"{super().get_dpi(HEIGHT_ROOT_WINDOW + HEIGHT_WINDOW_MENU_BAR)}"
            f"+{x_pos}+{y_pos}")
        self.resizable(width=False, height=False)
        # root.overrideredirect(True)  # 去除标题栏
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(base64.b64decode(icon.img))
        self.iconbitmap('tmp.ico')
        os.remove('tmp.ico')

    def set_root_menu(self, model: UIModel) -> None:
        """
        设置根窗口菜单
        """
        # 窗口绑定菜单栏
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        """菜单栏绑定文件菜单"""
        file_menu = tk.Menu(menu_bar, tearoff=False, font=FONT_BUTTON,
                            bg=COLOR_FRAME_BG, fg='black',
                            activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG)
        self.bind("<F>", lambda e: file_menu.post(self.winfo_x() + 10, self.winfo_y() + 75))
        menu_bar.add_cascade(label="文件(F)", menu=file_menu)

        # 文件菜单中的界面切换单选按钮
        file_menu.add_radiobutton(label="Excel2DBC",
                                  variable=model.radiobutton_dest_frame_idx,
                                  value=model.FRAMES.E2D,
                                  command=lambda: self.presenter.handler_change_frame(),
                                  )
        file_menu.add_radiobutton(label="DBC2Excel",
                                  variable=model.radiobutton_dest_frame_idx,
                                  value=model.FRAMES.D2E,
                                  command=lambda: self.presenter.handler_change_frame(),
                                  )

        # 添加分割线
        file_menu.add_separator()
        # 文件菜单绑定Excel自适应按钮
        file_menu.add_checkbutton(label="Excel AutoFit",
                                  variable=model.checkbutton_excel_auto_fit,
                                  command=lambda: None,
                                  )

        # 添加分割线
        file_menu.add_separator()
        # 文件菜单绑定设置
        file_settings_menu = tk.Menu(file_menu, tearoff=False, font=FONT_BUTTON,
                                     bg=COLOR_FRAME_BG, fg='black',
                                     activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG)
        file_menu.add_cascade(menu=file_settings_menu, label="设置")
        # 文件_设置绑定2Excel信号排序单选按钮
        # signals排序
        file_settings_menu.add_radiobutton(label="Sort Signals By Source File",
                                           variable=model.radiobutton_sort_signals,
                                           value='SOURCE_FILE',
                                           command=lambda: None)
        file_settings_menu.add_radiobutton(label="Sort Signals By Start Bit ASC",
                                           variable=model.radiobutton_sort_signals,
                                           value='START_BIT_ASC',
                                           command=lambda: None)
        file_settings_menu.add_radiobutton(label="Sort Signals By Start Bit DESC",
                                           variable=model.radiobutton_sort_signals,
                                           value='START_BIT_DESC',
                                           command=lambda: None)
        # 添加分割线
        file_settings_menu.add_separator()
        # messages排序
        file_settings_menu.add_radiobutton(label="Sort Messages By Source File",
                                           variable=model.radiobutton_sort_messages,
                                           value='SOURCE_FILE',
                                           command=lambda: None)
        file_settings_menu.add_radiobutton(label="Sort Messages By Name ASC",
                                           variable=model.radiobutton_sort_messages,
                                           value='NAME_ASC',
                                           command=lambda: None)
        file_settings_menu.add_radiobutton(label="Sort Messages By ID ASC",
                                           variable=model.radiobutton_sort_messages,
                                           value='ID_ASC',
                                           command=lambda: None)

        # 添加分割线
        file_menu.add_separator()
        # 文件菜单绑定退出按钮
        file_menu.add_command(label="退出", command=lambda: self.presenter.handler_on_closing())

        """菜单栏绑定帮助菜单"""
        help_menu = tk.Menu(menu_bar, tearoff=False, font=FONT_BUTTON,
                            bg=COLOR_FRAME_BG, fg='black',
                            activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG)
        self.bind("<H>", lambda e: help_menu.post(self.winfo_x() + 85, self.winfo_y() + 75))
        menu_bar.add_cascade(label="帮助(H)", menu=help_menu)
        help_menu.add_command(label="帮助", command=lambda: self.__show_help())
        help_menu.add_command(label="关于", command=lambda: self.__show_about())
        help_menu.add_command(label="Intel&Motorola", command=lambda: self.__show_exp_img())

    def set_operation_frame(self, model: UIModel):
        """
        设置操作frame界面
        """

        """
        窗口中的界面
        """
        # frame界面
        frame_e2d = TkFrame(master=self,
                            bg=COLOR_FRAME_BG, borderwidth=0,
                            x=(WIDTH_ROOT_WINDOW * (1 if model.current_frame_idx else 0)), y=0,
                            width=WIDTH_ROOT_WINDOW, height=HEIGHT_ROOT_WINDOW)
        frame_d2e = TkFrame(master=self,
                            bg=COLOR_FRAME_BG, borderwidth=0,
                            x=(WIDTH_ROOT_WINDOW * (0 if model.current_frame_idx else 1)), y=0,
                            width=WIDTH_ROOT_WINDOW, height=HEIGHT_ROOT_WINDOW)
        """
        添加控件到frame_e2d界面
        """
        # 标签_显示excel文件路径
        lbl_open_excel = TkLabel(master=frame_e2d,
                                 bg=COLOR_LABEL_BG, fg=COLOR_LABEL_FG, borderwidth=0,
                                 text="1、打开Excel文件", font=FONT_LABEL,
                                 relief="sunken", justify='left',
                                 anchor='w', wraplength=self.WIDTH_LABEL,
                                 x=10, y=10, width=self.WIDTH_LABEL, height=self.HEIGHT_LABEL)
        # 标签_显示已保存DBC文件路径
        self.lbl_save_dbc = TkLabel(master=frame_e2d,
                               bg=COLOR_LABEL_BG, fg=COLOR_LABEL_FG, borderwidth=0,
                               text="2、生成DBC文件", font=FONT_LABEL,
                               relief="sunken", justify='left',
                               anchor='w', wraplength=self.WIDTH_LABEL,
                               x=10, y=60, width=self.WIDTH_LABEL, height=self.HEIGHT_LABEL)
        # 绑定菜单
        lbl_save_dbc_menu = tk.Menu(self, tearoff=False, font=FONT_BUTTON,
                                    bg=COLOR_FRAME_BG, fg='black',
                                    activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG)
        lbl_save_dbc_menu.add_command(label="打开文件",
                                      command=lambda: self.presenter.handler_open_folder(self.lbl_save_dbc.cget('text'),
                                                                                         'file')
                                      )
        lbl_save_dbc_menu.add_command(label="打开文件夹",
                                      command=lambda: self.presenter.handler_open_folder(self.lbl_save_dbc.cget('text'),
                                                                                         'folder')
                                      )
        self.lbl_save_dbc.bind("<Button-3>", lambda e: lbl_save_dbc_menu.post(e.x_root, e.y_root))

        # 按钮_打开Excel
        btn_open_excel = TkButton(master=frame_e2d,
                                  bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                  activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG,
                                  borderwidth=0,
                                  text="打开Excel", font=FONT_BUTTON,
                                  command=lambda: self.presenter.handler_open_file(filetype='Excel',
                                                                                   lbl=lbl_open_excel,
                                                                                   dir=model.opened_excel_filepath),
                                  x=self.WIDTH_LABEL + 20, y=15, width=WIDTH_BUTTON, height=HEIGHT_BUTTON)
        # 按钮_生成DBC
        btn_generate_dbc = TkButton(master=frame_e2d,
                                    bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                    activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG,
                                    borderwidth=0,
                                    text="生成DBC", font=FONT_BUTTON,
                                    command=lambda: self.presenter.handler_save_e2d(),
                                    x=self.WIDTH_LABEL + 20, y=65, width=WIDTH_BUTTON, height=HEIGHT_BUTTON)

        """
        添加控件到frame_d2e界面
        """
        # 标签_显示dbc文件路径
        lbl_open_dbc = TkLabel(master=frame_d2e,
                               bg=COLOR_LABEL_BG, fg=COLOR_LABEL_FG, borderwidth=0,
                               text="1、打开DBC文件", font=FONT_LABEL,
                               relief="sunken", justify='left',
                               anchor='w', wraplength=self.WIDTH_LABEL,
                               x=10, y=10, width=self.WIDTH_LABEL, height=self.HEIGHT_LABEL)
        # 标签_显示已保存Excel文件路径
        self.lbl_save_excel = TkLabel(master=frame_d2e,
                                 bg=COLOR_LABEL_BG, fg=COLOR_LABEL_FG, borderwidth=0,
                                 text="2、生成Excel文件", font=FONT_LABEL,
                                 relief="sunken", justify='left',
                                 anchor='w', wraplength=self.WIDTH_LABEL,
                                 x=10, y=60, width=self.WIDTH_LABEL, height=self.HEIGHT_LABEL)

        # 绑定菜单
        lbl_save_excel_menu = tk.Menu(self, tearoff=False, font=FONT_BUTTON,
                                      bg=COLOR_FRAME_BG, fg='black',
                                      activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG)
        lbl_save_excel_menu.add_command(label="打开文件",
                                        command=lambda: self.presenter.handler_open_folder(self.lbl_save_excel.cget('text'),
                                                                                           'file')
                                        )
        lbl_save_excel_menu.add_command(label="打开文件夹",
                                        command=lambda: self.presenter.handler_open_folder(self.lbl_save_excel.cget('text'),
                                                                                           'folder')
                                        )
        self.lbl_save_excel.bind("<Button-3>", lambda e: lbl_save_excel_menu.post(e.x_root, e.y_root))

        # 按钮_打开DBC
        btn_open_dbc = TkButton(master=frame_d2e,
                                bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG,
                                borderwidth=0,
                                text="打开DBC", font=FONT_BUTTON,
                                command=lambda: self.presenter.handler_open_file(filetype='DBC',
                                                                                 lbl=lbl_open_dbc,
                                                                                 dir=model.opened_dbc_filepath),
                                x=self.WIDTH_LABEL + 20, y=15, width=WIDTH_BUTTON, height=HEIGHT_BUTTON)
        # 按钮_生成Excel
        btn_generate_excel = TkButton(master=frame_d2e,
                                      bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                      activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_ACTIVE_FG,
                                      borderwidth=0,
                                      text="生成Excel", font=FONT_BUTTON,
                                      command=lambda: self.presenter.handler_save_d2e(),
                                      x=self.WIDTH_LABEL + 20, y=65, width=WIDTH_BUTTON, height=HEIGHT_BUTTON)

        # 信息提示
        self.text_info = TkText(master=self,
                                bg=COLOR_LABEL_BG, fg=COLOR_LABEL_FG,
                                borderwidth=5,
                                font=FONT_LABEL,
                                relief='flat', wrap='none', state='disabled',
                                x=10, y=110, width=self.WIDTH_OPERATION_FRAME - 38, height=self.HEIGHT_TEXT_INFO)
        self.text_info.creat_scrollbar()
        # 保存e2d和d2e界面的实例到model
        model.frames = [frame_e2d, frame_d2e]

    @staticmethod
    def __show_about():
        """
        显示关于弹窗

        """
        msg = ('产品信息: \n'
               '    DBC Tool V2025.1.0\n'
               '    Author: ZYD\n')
        messagebox.showinfo(title='关于', message=msg)


    @staticmethod
    def __show_help():
        """
        显示帮助说明弹窗

        """
        msg = ('!!!须按照模板填写数据，否则无法生成DBC!!!\n'
               '!!!生成后打开.dbc文件再保存检查是否正常!!!\n'
               '1、打开文件\n'
               '2、点击生成按钮\n'
               '3、等待生成完成\n')
        messagebox.showinfo(title='帮助', message=msg)


    @staticmethod
    def __show_exp_img():
        """
        显示解释intel&motorola格式的图片

        """
        with open('exp_img.jpg', 'wb') as tmp:
            tmp.write(base64.b64decode(icon.exp_img))
        p = subprocess.Popen(["start", "exp_img.jpg"], shell=True)
