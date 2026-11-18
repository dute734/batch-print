# -*- coding: utf-8 -*-
"""打包成单个免安装桌面程序：dist/批量打印.exe

轻量化要点：
- PDF 引擎用 pypdfium2（约 8 MB），不用 PyMuPDF（约 98 MB）
- 只带 QtCore / QtGui / QtWidgets，其余 Qt 模块全部排除
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))

QT_EXCLUDES = [
    "PySide6.QtNetwork", "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtXml",
    "PySide6.QtConcurrent", "PySide6.QtPrintSupport", "PySide6.QtDBus",
    "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets", "PySide6.QtSvg",
    "PySide6.QtSvgWidgets", "PySide6.QtUiTools", "PySide6.QtHelp",
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebSockets",
    "PySide6.QtQuick", "PySide6.QtQml", "PySide6.Qt3DCore", "PySide6.QtMultimedia",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtDesigner",
    "pymupdf", "fitz", "tkinter",
]


def main():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("未安装 PyInstaller，正在安装…")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile",
        "--windowed",
        "--name", "批量打印",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32timezone",
        "--hidden-import", "win32gui",
        "--hidden-import", "PIL.ImageWin",
        "--collect-submodules", "win32com",
        "--collect-all", "pypdfium2",
        "--hidden-import", "pypdfium2_raw",
        "--icon", os.path.join(BASE, "图标.ico"),
        "--add-data", os.path.join(BASE, "图标.png") + ";.",
    ]
    for mod in QT_EXCLUDES:
        args += ["--exclude-module", mod]
    args.append("gui.py")
    print(" ".join(args))
    subprocess.check_call(args, cwd=BASE)
    dist = os.path.join(BASE, "dist", "批量打印.exe")
    if os.path.exists(dist):
        size = round(os.path.getsize(dist) / 1024 / 1024, 1)
        print("\n打包完成：%s（%s MB）" % (dist, size))
        print("把这个 exe 单独拷走即可，双击就是一个桌面程序，不会开浏览器。")
    else:
        print("打包失败，请查看上面的报错。")


if __name__ == "__main__":
    main()
