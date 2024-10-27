#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : ZYD
# @Time    : 2024/7/11 下午2:58
# @version : V1.0.0
# @function:

##############################
# Module imports
##############################

import traceback  # 用于获取异常详细信息

from app import UIModel, UIView, UICtrl


if __name__ == '__main__':
    # 配置文件路径
    CONFIG_PATH = 'cfg.ini'
    try:
        # 创建view
        ui_view = UIView()
        # 创建model
        ui_model = UIModel()
        # 创建controller
        ui_ctrl = UICtrl(model=ui_model,
                         view=ui_view,
                         cfg_path=CONFIG_PATH)

        # 显示窗口内容
        ui_view.set_root_menu(ui_model)
        ui_view.set_operation_frame(ui_model)

        # 建立窗口事件循环
        ui_view.mainloop()
    except Exception as e:
        ui_ctrl.text_log(f'发生异常 {e}', 'error')
        ui_ctrl.text_log(f"{traceback.format_exc()}", 'error')
        print(f'发生异常 {e}')
        print(f"{traceback.format_exc()}")
    finally:
        pass
