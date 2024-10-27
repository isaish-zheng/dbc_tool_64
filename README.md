> # 产品信息
> **DBC Tool**


> # 本产品包含以下内容
> * **Excel2DBC**
> * **DBC2Excel**


> # 功能说明
> ## Excel2DBC
> 1. 将Excel文件转换为DBC文件
> ## DBC2Excel
> 1. 将DBC文件转换为Excel文件


> # 使用说明
> ## 操作系统
> **Windows7及以上版本(x64)**
> ## Python版本(64位)
> * **Python3.7.9** [Python解释器下载](https://www.python.org/ftp/python/3.7.9/python-3.7.9-amd64.exe)
> ## 依赖模块
> 1. 参照requirements.txt安装依赖模块
> ## 打包方式
> **PyInstaller**
> * 安装pyinstaller模块，打开cmd窗口，运行以下命令
> 
>   `pip install pyinstaller`
> * 打包成exe，按住shift+鼠标右键弹出菜单，点击打开powershell，运行以下命令
> 
>   `pyinstaller main.py -F -w -i .\icons\Z.ico -n "DBC Tool"`
> 
> ## 运行exe
> * 以管理员身份运行DBC Tool.exe
> 
> ### 使用Excel2DBC
> 1. 菜单栏：文件->勾选"Excel2DBC"
> 2. 单击"打开Excel"按钮，选择.xlsx文件
> 3. 单击"生成DBC"按钮，保存.dbc文件
>
> ### 使用DBC2Excel
> 1. 菜单栏：文件->勾选"DBC2Excel"
> 2. 单击"打开DBC"按钮，选择.dbc文件
> 3. 单击"生成Excel"按钮，选择tpl.xlsx模板，然后按照模板保存.xlsx文件

> # 其它
> * 本产品仅用于学习交流，请勿用于商业用途
