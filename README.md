# 批量打印（BatchPrint）

把 **Excel / Word / PDF** 一次性拖进来、排好顺序，逐项确认「打不打、打几份、什么方向」，一键交给本机 WPS 与打印机，全程零弹窗。

> 面向 Windows + 金山 WPS 的桌面小工具，单文件 exe，免安装。界面是苹果磨砂毛玻璃 + 极光粉彩风格（PySide6 自绘窗口）。

## 功能

| 能力 | 说明 |
|---|---|
| Excel 按工作表 | 展开每个 `.xlsx/.xls/.xlsm/.xlsb/.et` 的所有工作表，逐表勾选；隐藏表默认不勾，打印时临时显示、打完自动恢复隐藏 |
| Word / PDF 整篇 | `.doc/.docx/.docm/.rtf/.wps/.wpt/.pdf` 各占一行，页数从文档真实页面读取 |
| 逐项份数 / 方向 | 每行独立设份数（玻璃 `− n +` 步进器）与方向：保持原样 / 横向 / 纵向 / 适应一页宽 |
| 拖拽 + 顺序 | 支持拖文件、拖文件夹（递归最多 3 层），左侧列表可拖动排序，Ctrl 多选 / Ctrl+A 全选 / Delete 删除 |
| 发送即完成 | 状态列：`待打印` → `正在发送` → 任务进打印队列立刻 `已发送`；不再逐个等纸出炉，任务之间零等待（实测 6 项 × 2 份 8 秒送完）；出错 `失败` |
| 全程静默 | 不弹确认框、不弹进度层；WPS 自带的「打印进程」小窗会被后台守护线程隐藏（只隐藏不关闭，不会变成取消打印） |
| 份数逐份提交 | WPS 的 `PrintOut(Copies=n)` 在表格/文字里都会被忽略，所以份数 n 就提交 n 次、每次 1 份 |

## 快速开始

1. **直接用**：双击 `dist\批量打印.exe`（Windows 10/11，无需安装 Python）。
2. **跑源码**：
   ```bash
   pip install PySide6 pywin32 pypdfium2 pillow
   python gui.py
   ```
   或双击 `启动.bat`。
3. **重新打包**：
   ```bash
   python 打包exe.py
   ```
   产物在 `dist\批量打印.exe`。

## 文件说明

| 文件 | 作用 |
|---|---|
| `gui.py` | 桌面主程序：自绘无边框窗口、毛玻璃/3D 控件、拖拽列表、任务表、实时状态 |
| `wps_print.py` | Excel 打印核心 + 批量调度、打印机枚举/端口/状态、打印面板隐藏、队列核验 |
| `doc_print.py` | Word（WPS 文字 COM）与 PDF（GDI 直接送打印机）打印核心 |
| `打包exe.py` | PyInstaller 打包脚本（单文件、裁掉多余 Qt 模块，61 MB 左右） |
| `app.py` / `templates/index.html` / `pick_files.py` | 更早的浏览器界面版本，备用，可不用 |
| `图标.ico` / `图标.png` | 程序与窗口图标 |
| `使用说明.md` | 详细使用说明与踩坑记录（WPS 版本、打印机名写法、份数等） |

## 环境与坑

- Windows + WPS Office（实测 12.1.x），Python 3.10+；打包用 PyInstaller。
- WPS 表格设置打印机只认 `打印机名 在 NeXX:`，WPS 文字反过来只认纯打印机名——代码里是「候选列表逐个试 + 回读校验」。
- 表格 `PrintOut` 不认 `Background` 关键字、不认 `Copies`；`PrintToFile` 在 12.1 不出文件。
- 后台线程必须先 `pythoncom.CoInitialize()`。

## 许可

自用/学习为主，代码可自由参考。
