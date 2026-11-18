# -*- coding: utf-8 -*-
"""批量打印 · 原生桌面窗口（PySide6 + 本机 WPS 表格 COM）

打印全程静默，不弹任何窗口，进度只在主界面显示。
"""
import os
import sys
import threading
import time

from PySide6.QtCore import Qt, QThread, Signal, QSize, QTimer, QRect, QPoint
from PySide6.QtGui import (QColor, QFont, QIcon, QPixmap, QPainter, QLinearGradient,
                           QBrush, QFontMetrics, QPainterPath, QRadialGradient, QPen)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QListWidget, QListWidgetItem, QAbstractItemView, QTextEdit,
    QFrame, QSizePolicy, QCheckBox, QSplitter, QStyledItemDelegate, QStyle,
    QGraphicsDropShadowEffect, QSizeGrip,
)

import doc_print
import wps_print


def _res_path(name):
    """打包成 exe 后资源被解压到临时目录，这里统一找"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


EXCEL_EXT = wps_print.EXCEL_EXT
WORD_EXT = wps_print.WORD_EXT
PDF_EXT = wps_print.PDF_EXT
SUPPORTED_EXT = EXCEL_EXT + WORD_EXT + PDF_EXT
KIND_LABEL = {"excel": "Excel", "word": "Word", "pdf": "PDF"}

QSS = """
QLabel { color: #2c2a3a; font-family: "Microsoft YaHei UI"; background: transparent; }
QLabel#title { font-size: 17px; font-weight: 700; letter-spacing: 1px; }
QLabel#sub { color: #9a95b4; font-size: 11px; }
QLabel#section { font-size: 13px; font-weight: 600; color: #3a3550; }
QLabel#muted { color: #9a95b4; font-size: 11px; }
QLabel#chip {
    background: rgba(255,255,255,0.72); border: 1px solid rgba(255,255,255,0.9);
    border-radius: 11px; padding: 3px 10px; color: #615d78; font-size: 11px;
}
QLabel#badge {
    background: rgba(255,255,255,0.65); border: 1px solid rgba(255,255,255,0.95);
    border-radius: 12px; padding: 5px 12px; color: #615d78; font-size: 11px;
}

/* 3D 凸起按钮：上浅下深的渐变 + 立体边 */
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255,255,255,0.98), stop:1 rgba(240,236,250,0.92));
    border: 1px solid rgba(255,255,255,0.95);
    border-bottom: 2px solid rgba(196,184,226,0.75);
    border-radius: 10px; padding: 6px 14px; color: #3a3550;
    font-family: "Microsoft YaHei UI"; font-size: 12px;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #ffffff, stop:1 #f4eeff);
    border-bottom: 2px solid rgba(183,157,240,0.9);
}
QPushButton:pressed {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #eee6ff, stop:1 #f7f2ff);
    border-top: 2px solid rgba(196,184,226,0.8);
    border-bottom: 1px solid rgba(255,255,255,0.9);
    padding-top: 8px; padding-bottom: 4px;
}
QPushButton:disabled { color: #b9b4cc; background: rgba(255,255,255,0.55);
    border-bottom: 2px solid rgba(226,220,240,0.7); }
QPushButton#ghost { padding: 4px 10px; font-size: 11px; }
QPushButton#danger { color: #b93a45; }
QPushButton#primary {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #ffb3d0, stop:0.5 #b9a0f2, stop:1 #86c9f3);
    color: white; font-weight: 700; border: 1px solid rgba(255,255,255,0.9);
    border-bottom: 3px solid rgba(120,100,180,0.55);
    padding: 10px 28px; border-radius: 14px; font-size: 14px; letter-spacing: 1px;
}
QPushButton#primary:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #ffc2da, stop:0.5 #c6b1f7, stop:1 #97d3f6);
}
QPushButton#primary:pressed {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #f0a4c4, stop:0.5 #a891e8, stop:1 #7bbde8);
    padding-top: 12px; padding-bottom: 8px;
}
QPushButton#winBtn {
    background: transparent; border: none; border-radius: 9px;
    color: #8b87a3; font-size: 13px; padding: 0px; min-width: 28px; max-width: 28px;
    min-height: 26px; max-height: 26px;
}
QPushButton#winBtn:hover { background: rgba(255,255,255,0.75); color: #3a3550; }
QPushButton#winClose:hover { background: rgba(255,120,140,0.85); color: white; }

QComboBox, QSpinBox {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #ffffff, stop:1 #f4f1fb);
    border: 1px solid rgba(255,255,255,0.95);
    border-bottom: 2px solid rgba(199,188,228,0.7);
    border-radius: 9px; padding: 3px 8px; font-family: "Microsoft YaHei UI";
    font-size: 12px; min-height: 20px; max-height: 26px;
}
QComboBox:hover, QSpinBox:hover { border-bottom: 2px solid #b79df0; }
QComboBox::drop-down { border: none; width: 16px; }
QComboBox QAbstractItemView {
    background: #ffffff; border: 1px solid rgba(200,190,225,0.6);
    selection-background-color: rgba(183,157,240,0.25); selection-color: #2c2a3a;
}
QListWidget#chips { background: transparent; border: none; outline: none; }

QSplitter#mainSplit::handle {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255,255,255,0.65), stop:1 rgba(200,190,235,0.45));
    height: 8px; border-radius: 4px; margin: 4px 0px;
}
QSplitter#mainSplit::handle:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #ffb3d0, stop:0.5 #b9a0f2, stop:1 #86c9f3);
}

QTableWidget {
    background: rgba(255,255,255,0.52); border: 1px solid rgba(255,255,255,0.85);
    border-radius: 12px; gridline-color: rgba(210,203,232,0.35);
    font-family: "Microsoft YaHei UI"; font-size: 12px; color: #2c2a3a;
    selection-background-color: rgba(183,157,240,0.2); outline: none;
}
QTableWidget::item { padding: 3px 4px; border: none; }
QHeaderView::section {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255,214,232,0.35), stop:1 rgba(214,206,246,0.35));
    color: #5f5a78; border: none; border-bottom: 1px solid rgba(255,255,255,0.9);
    padding: 7px 6px; font-family: "Microsoft YaHei UI"; font-size: 11px; font-weight: 600;
}
QTextEdit#log {
    background: rgba(255,255,255,0.55); border: 1px solid rgba(255,255,255,0.9);
    border-radius: 10px; color: #6a6580; font-family: Consolas, "Microsoft YaHei UI";
    font-size: 11px; padding: 4px 6px;
}
QScrollBar:vertical { background: transparent; width: 9px; margin: 2px; }
QScrollBar::handle:vertical {
    background: rgba(183,157,240,0.55); border-radius: 4px; min-height: 24px;
}
QScrollBar::handle:vertical:hover { background: rgba(150,120,225,0.8); }
QScrollBar::add-line, QScrollBar::sub-line { height: 0px; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
"""


# ------------------------------------------------------------------ 背景 / 玻璃壳
class AuroraShell(QFrame):
    """整窗底板：极光光斑 + 磨砂玻璃圆角"""

    def __init__(self):
        super().__init__()
        self.setObjectName("shell")
        self._radius = 22

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        rect = self.rect()
        path = QPainterPath()
        path.addRoundedRect(rect.adjusted(0.5, 0.5, -0.5, -0.5), self._radius, self._radius)

        # 底色
        base = QLinearGradient(rect.topLeft(), rect.bottomRight())
        base.setColorAt(0.0, QColor("#fdf7ff"))
        base.setColorAt(0.35, QColor("#f6f2fd"))
        base.setColorAt(0.7, QColor("#eef4fd"))
        base.setColorAt(1.0, QColor("#fdf5f8"))
        p.setClipPath(path)
        p.fillPath(path, QBrush(base))

        # 极光光斑
        w, h = rect.width(), rect.height()
        blobs = [
            (0.14 * w, 0.06 * h, 0.55 * w, QColor(255, 176, 208, 120)),
            (0.78 * w, 0.02 * h, 0.60 * w, QColor(178, 158, 246, 140)),
            (0.55 * w, 0.92 * h, 0.65 * w, QColor(150, 205, 246, 130)),
            (0.03 * w, 0.55 * h, 0.45 * w, QColor(168, 240, 226, 90)),
        ]
        for cx, cy, r, col in blobs:
            g = QRadialGradient(cx, cy, r)
            g.setColorAt(0.0, col)
            mid = QColor(col)
            mid.setAlpha(0)
            g.setColorAt(1.0, mid)
            p.fillPath(path, QBrush(g))

        # 磨砂玻璃罩（提高通透感的乳白层）
        p.fillPath(path, QColor(255, 255, 255, 62))

        # 外描边 + 顶部高光（3D 凸起感）
        p.setPen(QPen(QColor(255, 255, 255, 210), 1))
        p.drawPath(path)
        hi = QLinearGradient(rect.topLeft(), rect.bottomLeft())
        hi.setColorAt(0.0, QColor(255, 255, 255, 120))
        hi.setColorAt(0.25, QColor(255, 255, 255, 0))
        p.fillPath(path, QBrush(hi))


class GlassCard(QFrame):
    """半透明磨砂玻璃卡 + 立体投影"""

    def __init__(self, deep=False):
        super().__init__()
        self.setObjectName("cardDeep" if deep else "card")
        self.setAttribute(Qt.WA_StyledBackground, True)
        eff = QGraphicsDropShadowEffect(self)
        eff.setBlurRadius(28)
        eff.setOffset(0, 6)
        eff.setColor(QColor(150, 132, 190, 70))
        self.setGraphicsEffect(eff)
        self.setStyleSheet(
            "QFrame#card, QFrame#cardDeep {"
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "  stop:0 rgba(255,255,255,0.80), stop:0.55 rgba(255,255,255,0.66),"
            "  stop:1 rgba(243,239,252,0.72));"
            "border: 1px solid rgba(255,255,255,0.98);"
            "border-bottom: 2px solid rgba(206,196,232,0.55);"
            "border-radius: 16px; }")


class WinButton(QPushButton):
    def __init__(self, text, close=False):
        super().__init__(text)
        self.setObjectName("winClose" if close else "winBtn")
        self.setCursor(Qt.PointingHandCursor)


class GlassStepper(QFrame):
    """玻璃风份数调节器：− 数字 + ，替代带黑三角的 QSpinBox"""

    valueChanged = Signal(int)

    def __init__(self, value=1, minimum=1, maximum=99):
        super().__init__()
        self._min = minimum
        self._max = maximum
        self._value = max(minimum, min(maximum, int(value)))
        self.setObjectName("stepper")
        self.setFixedHeight(22)
        self.setMinimumWidth(72)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(
            "QFrame#stepper { background: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "  stop:0 #ffffff, stop:1 #f4f1fb);"
            "  border: 1px solid rgba(255,255,255,0.95);"
            "  border-bottom: 2px solid rgba(199,188,228,0.65);"
            "  border-radius: 10px; }")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(2, 0, 2, 0)
        lay.setSpacing(0)
        self.minus = self._btn("−", lambda: self._step(-1))
        lay.addWidget(self.minus)
        self.label = QLabel(str(self._value))
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color:#3a3550; font-size:12px; font-weight:600;")
        lay.addWidget(self.label, 1)
        self.plus = self._btn("+", lambda: self._step(1))
        lay.addWidget(self.plus)

    def _btn(self, text, slot):
        b = QPushButton(text)
        b.setCursor(Qt.PointingHandCursor)
        b.setFixedSize(20, 18)
        b.setStyleSheet(
            "QPushButton { border: none; background: transparent; color: #8b87a3;"
            "  font-size: 13px; font-weight: 600; padding: 0px; border-radius: 6px; }"
            "QPushButton:hover { background: rgba(183,157,240,0.35); color: #4a4370; }"
            "QPushButton:pressed { background: rgba(183,157,240,0.55); }")
        b.clicked.connect(slot)
        return b

    def _step(self, delta):
        new = max(self._min, min(self._max, self._value + delta))
        if new != self._value:
            self.setValue(new)

    def value(self):
        return self._value

    def setValue(self, v):
        v = max(self._min, min(self._max, int(v)))
        if v == self._value:
            return
        self._value = v
        self.label.setText(str(v))
        self.valueChanged.emit(v)


class TitleBar(QFrame):
    """自绘标题栏：拖动窗口、双击最大化"""

    def __init__(self, window):
        super().__init__()
        self.win = window
        self._drag_offset = None
        lay = QHBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 0)
        lay.setSpacing(8)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(34, 34)
        self.icon_label.setScaledContents(True)
        lay.addWidget(self.icon_label)

        box = QVBoxLayout()
        box.setSpacing(0)
        t = QLabel("批量打印")
        t.setObjectName("title")
        self.sub_label = QLabel("Excel 按工作表选打 · Word / PDF 整篇打印 · 全程静默")
        self.sub_label.setObjectName("sub")
        box.addWidget(t)
        box.addWidget(self.sub_label)
        lay.addLayout(box)
        lay.addStretch(1)

        self.badge = QLabel("正在检测 WPS…")
        self.badge.setObjectName("badge")
        lay.addWidget(self.badge)

        self.min_btn = WinButton("—")
        self.max_btn = WinButton("□")
        self.close_btn = WinButton("✕", close=True)
        for b, slot in ((self.min_btn, self.win.showMinimized),
                        (self.max_btn, self._toggle_max),
                        (self.close_btn, self.win.close)):
            b.clicked.connect(slot)
            lay.addWidget(b)

    def set_icon(self, pixmap):
        self.icon_label.setPixmap(pixmap)

    def _toggle_max(self):
        if self.win.isMaximized():
            self.win.showNormal()
            self.max_btn.setText("□")
        else:
            self.win.showMaximized()
            self.max_btn.setText("❐")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_offset = event.globalPosition().toPoint() - self.win.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self._drag_offset is not None and event.buttons() & Qt.LeftButton:
            if self.win.isMaximized():
                self.win.showNormal()
            self.win.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_offset = None

    def mouseDoubleClickEvent(self, event):
        self._toggle_max()


# ------------------------------------------------------------------ 后台线程
class InspectWorker(QThread):
    fileReady = Signal(dict)
    error = Signal(str)
    finishedAll = Signal()

    def __init__(self, pairs):
        super().__init__()
        self.pairs = pairs  # [(fid, path, kind)]

    def run(self):
        session = None
        writer = None
        try:
            for fid, path, kind in self.pairs:
                if kind == "excel":
                    if session is None:
                        session = wps_print.WpsSession()
                    data = wps_print.inspect_workbook(session, path)
                elif kind == "word":
                    if writer is None:
                        writer = doc_print.WpsWriterSession()
                    data = doc_print.inspect_word(writer, path)
                else:
                    data = doc_print.inspect_pdf(path)
                self.fileReady.emit({"id": fid, "data": data})
        except Exception as e:  # noqa: BLE001
            self.error.emit(str(e))
        finally:
            if session:
                session.quit()
            if writer:
                writer.quit()
            self.finishedAll.emit()


class PrintWorker(QThread):
    event = Signal(dict)

    def __init__(self, printer, items):
        super().__init__()
        self.printer = printer
        self.items = items
        self.stop_flag = threading.Event()

    def run(self):
        try:
            wps_print.print_batch(self.printer, self.items,
                                  lambda ev: self.event.emit(ev),
                                  self.stop_flag, verify_seconds=8)
        except Exception as e:  # noqa: BLE001
            self.event.emit({"type": "fatal", "error": str(e)})


# ------------------------------------------------------------------ 任务卡片
class TaskChipDelegate(QStyledItemDelegate):
    """把每条任务画成一张扁平卡片：独占一行，行高 24px"""

    ROW_H = 24

    def sizeHint(self, option, index):
        view = self.parent()
        w = view.viewport().width() if view is not None else 800
        return QSize(max(240, w - 6), self.ROW_H)

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = option.rect.adjusted(2, 2, -10, -2)
        selected = bool(option.state & QStyle.State_Selected)
        # 底色
        if selected:
            painter.setBrush(QBrush(QColor(183, 157, 240, 60)))
            painter.setPen(QColor("#b79df0"))
        else:
            painter.setBrush(QBrush(QColor(255, 255, 255, 210)))
            painter.setPen(QColor(210, 200, 235, 140))
        painter.drawRoundedRect(rect, 7, 7)

        # 文字
        font = QFont(option.font)
        font.setPixelSize(13)
        fm = QFontMetrics(font)
        text_rect = rect.adjusted(10, 0, -10, 0)
        text = fm.elidedText(str(index.data(Qt.DisplayRole) or ""), Qt.ElideRight,
                             text_rect.width())
        painter.setPen(QColor("#2c2a3a"))
        painter.setFont(font)
        painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, text)
        painter.restore()


# ------------------------------------------------------------------ 可接收拖拽的文件列表
class FileList(QListWidget):
    filesDropped = Signal(list)
    deleteRequested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.deleteRequested.emit()
            return
        super().keyPressEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
            if paths:
                self.filesDropped.emit(paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)


# ------------------------------------------------------------------ 主窗口
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量打印")
        self.resize(1240, 900)
        self.setMinimumSize(1020, 700)
        self.setAcceptDrops(True)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowMinMaxButtonsHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self.files = []          # [{id,name,path,sheets,inspected,error}]
        self.settings = {}       # {fid: {sheet: {print,copies,orient}}}
        self.seq = 0
        self.print_worker = None
        self.inspect_worker = None
        self.row_map = {}        # row -> (fid, sheet)
        self.status_rows = {}    # (fid, sheet) -> row
        self.print_seq_keys = {}  # 打印序号 -> (fid, sheet)
        self.print_state = {}    # (fid, sheet) -> (状态, 说明)
        self.split_manual = False

        self._build_ui()
        self._load_printers()
        self._load_wps_badge()

    # ---------- UI 构建
    def _build_ui(self):
        holder = QWidget()
        holder.setObjectName("holder")
        self.setCentralWidget(holder)
        hb = QVBoxLayout(holder)
        hb.setContentsMargins(14, 14, 14, 14)
        hb.setSpacing(0)

        self.shell = AuroraShell()
        shadow = QGraphicsDropShadowEffect(self.shell)
        shadow.setBlurRadius(48)
        shadow.setOffset(0, 12)
        shadow.setColor(QColor(120, 100, 170, 90))
        self.shell.setGraphicsEffect(shadow)
        hb.addWidget(self.shell)

        outer = QVBoxLayout(self.shell)
        outer.setContentsMargins(18, 10, 18, 14)
        outer.setSpacing(10)

        # 标题栏
        self.titlebar = TitleBar(self)
        icon = QPixmap(_res_path("图标.png"))
        if not icon.isNull():
            icon = icon.scaled(34, 34, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.titlebar.set_icon(icon)
            self.setWindowIcon(QIcon(icon))
        self.wps_badge = self.titlebar.badge
        outer.addWidget(self.titlebar)

        # 打印机条
        card = GlassCard()
        row = QHBoxLayout(card)
        row.setContentsMargins(14, 8, 14, 8)
        row.setSpacing(8)
        lab = QLabel("打 印 机")
        lab.setObjectName("section")
        row.addWidget(lab)
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(340)
        self.printer_combo.currentIndexChanged.connect(self._on_printer_change)
        row.addWidget(self.printer_combo)
        btn = QPushButton("刷新")
        btn.setObjectName("ghost")
        btn.clicked.connect(self._load_printers)
        row.addWidget(btn)
        self.printer_state = QLabel("")
        self.printer_state.setObjectName("muted")
        row.addWidget(self.printer_state, 1)
        outer.addWidget(card)

        # ============ 区域① 导入文件（可上下自由拖高）
        card1 = GlassCard()
        v1 = QVBoxLayout(card1)
        v1.setContentsMargins(14, 10, 14, 10)
        v1.setSpacing(8)
        h1 = QHBoxLayout()
        h1.setSpacing(8)
        lab1 = QLabel("① 导入文件（Excel / Word / PDF）")
        lab1.setObjectName("section")
        h1.addWidget(lab1)
        self.file_count_chip = QLabel("0 个任务")
        self.file_count_chip.setObjectName("chip")
        h1.addWidget(self.file_count_chip)
        h1.addStretch(1)
        for text, slot in (("添加文件", self._add_files),
                           ("添加文件夹", self._add_folder),
                           ("删除选中", self._delete_selected)):
            b = QPushButton(text)
            b.setObjectName("ghost")
            b.clicked.connect(slot)
            h1.addWidget(b)
        v1.addLayout(h1)

        self.file_list = FileList()
        self.file_list.setObjectName("chips")
        self.file_list.setDefaultDropAction(Qt.MoveAction)
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.setViewMode(QListWidget.ListMode)
        self.file_list.setFlow(QListWidget.TopToBottom)
        self.file_list.setWrapping(False)
        self.file_list.setSpacing(0)
        self.file_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.file_list.setMinimumHeight(48)
        self.chip_delegate = TaskChipDelegate(self.file_list)
        self.file_list.setItemDelegate(self.chip_delegate)
        self.file_list.deleteRequested.connect(self._delete_selected)
        self.file_list.filesDropped.connect(self._add_paths)
        self.file_list.model().rowsMoved.connect(self._on_rows_moved)
        v1.addWidget(self.file_list, 1)

        self.drop_hint = QLabel("拖文件或文件夹进来；一行一个任务，Ctrl 多选后点「删除选中」，Ctrl+A 全选")
        self.drop_hint.setObjectName("muted")
        v1.addWidget(self.drop_hint)
        card1.setMinimumHeight(110)

        # ============ 区域② 确认内容 + 实时状态（合并原②③）
        card2 = GlassCard()
        card2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        v2 = QVBoxLayout(card2)
        v2.setContentsMargins(14, 10, 14, 10)
        v2.setSpacing(8)

        h2 = QHBoxLayout()
        h2.setSpacing(8)
        lab2 = QLabel("② 确认要打印的内容 · 状态实时更新")
        lab2.setObjectName("section")
        h2.addWidget(lab2)
        self.stat_line = QLabel("已选 0 · 份数 0 · 页数 —")
        self.stat_line.setObjectName("chip")
        h2.addWidget(self.stat_line)
        h2.addStretch(1)
        for text, val in (("全选", True), ("全不选", False)):
            b3 = QPushButton(text)
            b3.setObjectName("ghost")
            b3.clicked.connect(lambda _=False, v=val: self._toggle_all(v))
            h2.addWidget(b3)
        self.stop_btn = QPushButton("中止")
        self.stop_btn.setObjectName("danger")
        self.stop_btn.setVisible(False)
        self.stop_btn.clicked.connect(self._stop_print)
        h2.addWidget(self.stop_btn)
        self.print_btn = QPushButton("开始打印")
        self.print_btn.setObjectName("primary")
        self.print_btn.setCursor(Qt.PointingHandCursor)
        self.print_btn.setMinimumWidth(126)
        self.print_btn.clicked.connect(self._start_print)
        h2.addWidget(self.print_btn)
        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("muted")
        v2.addLayout(h2)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["打印", "文件", "内容", "页数", "状态", "份数", "方向"])
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setShowGrid(False)
        self.table.setIconSize(QSize(14, 14))
        hdr = self.table.horizontalHeader()
        for i in (0, 3, 4, 5, 6):
            hdr.setSectionResizeMode(i, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setDefaultAlignment(Qt.AlignCenter)
        for c in (0, 3, 4, 5, 6):
            it = self.table.horizontalHeaderItem(c)
            if it is not None:
                it.setTextAlignment(Qt.AlignCenter)
        self.table.setColumnWidth(0, 48)
        self.table.setColumnWidth(1, 230)
        self.table.setColumnWidth(3, 56)
        self.table.setColumnWidth(4, 150)
        self.table.setColumnWidth(5, 74)
        self.table.setColumnWidth(6, 116)
        self.table.itemChanged.connect(self._on_item_changed)
        v2.addWidget(self.table, 1)

        # 日志（默认收起）
        self.log_toggle = QPushButton("显示日志 ▾")
        self.log_toggle.setObjectName("ghost")
        self.log_toggle.setCheckable(True)
        self.log_toggle.clicked.connect(self._toggle_log)
        foot = QHBoxLayout()
        foot.setSpacing(8)
        foot.addWidget(self.status_label, 0, Qt.AlignLeft | Qt.AlignVCenter)
        foot.addStretch(1)
        foot.addWidget(self.log_toggle, 0, Qt.AlignRight)
        v2.addLayout(foot)
        self.log = QTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(78)
        self.log.setVisible(False)
        v2.addWidget(self.log)

        # ①② 可上下自由拖拽高度
        split = QSplitter(Qt.Vertical)
        split.setObjectName("mainSplit")
        split.setChildrenCollapsible(False)
        split.addWidget(card1)
        split.addWidget(card2)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([150, 640])
        split.splitterMoved.connect(lambda *_: setattr(self, "split_manual", True))
        self.main_split = split
        outer.addWidget(split, 1)

        # 右下角缩放把手
        bottom = QHBoxLayout()
        bottom.setContentsMargins(0, 0, 0, 0)
        bottom.addStretch(1)
        grip = QSizeGrip(self.shell)
        grip.setFixedSize(14, 14)
        bottom.addWidget(grip, 0, Qt.AlignBottom | Qt.AlignRight)
        outer.addLayout(bottom)

    def _card(self, deep=False):
        return GlassCard(deep)

    def _muted(self, text):
        lab = QLabel(text)
        lab.setObjectName("muted")
        return lab

    def closeEvent(self, event):
        """关窗前把后台线程收干净，避免退出时崩"""
        try:
            if self.print_worker is not None and self.print_worker.isRunning():
                self.print_worker.stop_flag.set()
                self.print_worker.wait(3000)
            if self.inspect_worker is not None and self.inspect_worker.isRunning():
                self.inspect_worker.wait(3000)
        except Exception:  # noqa: BLE001
            pass
        super().closeEvent(event)

    def _toggle_log(self):
        on = self.log_toggle.isChecked()
        self.log.setVisible(on)
        self.log_toggle.setText("隐藏日志 ▴" if on else "显示日志 ▾")

    def _make_logo(self):
        pm = QPixmap(40, 40)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, 40, 40)
        grad.setColorAt(0, QColor("#f7a8c4"))
        grad.setColorAt(0.55, QColor("#b79df0"))
        grad.setColorAt(1, QColor("#8fc7f2"))
        p.setBrush(QBrush(grad))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, 0, 40, 40, 13, 13)
        p.setPen(QColor("white"))
        f = QFont("Microsoft YaHei UI")
        f.setPointSize(17)
        f.setBold(True)
        p.setFont(f)
        p.drawText(pm.rect(), Qt.AlignCenter, "印")
        p.end()
        return pm

    # ---------- 打印机 / WPS
    def _load_printers(self):
        printers = wps_print.list_printers()
        self.printers = printers
        self.printer_combo.blockSignals(True)
        self.printer_combo.clear()
        for p in printers:
            label = p["name"]
            if p["virtual"]:
                label += "  · 虚拟"
            if p["status_text"] and p["status_text"] != "就绪":
                label += "  · " + p["status_text"]
            self.printer_combo.addItem(label, p["name"])
            if p["is_default"]:
                self.printer_combo.setCurrentIndex(self.printer_combo.count() - 1)
        self.printer_combo.blockSignals(False)
        self._on_printer_change()

    def _on_printer_change(self):
        name = self.printer_combo.currentData()
        p = next((x for x in getattr(self, "printers", []) if x["name"] == name), None)
        if not p:
            self.printer_state.setText("未选择打印机")
            return
        if p["virtual"]:
            self.printer_state.setText("虚拟打印机：会弹「另存为」，不适合批量")
        else:
            self.printer_state.setText("端口 %s · 状态 %s" % (p["port"] or "?", p["status_text"]))

    def _load_wps_badge(self):
        info = None
        try:
            import win32api
            info = _wps_version_info(win32api)
        except Exception:
            info = None
        if info:
            self.wps_badge.setText("本机 WPS %s" % info)
        else:
            self.wps_badge.setText("未检测到 WPS")

    # ---------- 拖拽导入
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
            if paths:
                self._add_paths(paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)

    # ---------- 导入
    def _collect_files(self, paths, depth=0):
        """把文件夹展开成支持的文件清单（最多 3 层）"""
        out = []
        for p in paths:
            if os.path.isdir(p):
                if depth >= 3:
                    continue
                try:
                    names = sorted(os.listdir(p))
                except Exception:
                    continue
                out.extend(self._collect_files([os.path.join(p, n) for n in names], depth + 1))
            elif wps_print.file_kind(p):
                out.append(p)
        return out

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择要打印的文件", "",
            "支持的文件 (*.xlsx *.xls *.xlsm *.xlsb *.doc *.docx *.docm *.rtf *.wps *.pdf);;"
            "Excel 表格 (*.xlsx *.xls *.xlsm *.xlsb);;"
            "Word 文档 (*.doc *.docx *.docm *.rtf *.wps);;"
            "PDF (*.pdf);;所有文件 (*.*)")
        self._add_paths(paths)

    def _add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹（会导入里面的表格/文档/PDF）")
        if d:
            self._add_paths([d])

    def _add_paths(self, paths):
        candidates = self._collect_files(paths)
        new_ids = []
        skipped = 0
        for p in candidates:
            p = os.path.abspath(p)
            if not os.path.exists(p):
                continue
            kind = wps_print.file_kind(p)
            if not kind:
                skipped += 1
                continue
            if any(os.path.normcase(f["path"]) == os.path.normcase(p) for f in self.files):
                skipped += 1
                continue
            self.seq += 1
            rec = {"id": "f%d" % self.seq, "name": os.path.basename(p), "path": p,
                   "kind": kind, "sheets": [], "inspected": False, "error": None}
            self.files.append(rec)
            new_ids.append(rec["id"])
        self._refresh_file_list()
        if paths and not candidates:
            self.status_label.setText("这些文件不支持，只支持 Excel / Word / PDF")
        elif skipped and not new_ids:
            self.status_label.setText("没有新文件（重复或不支持的类型）")
        if new_ids:
            self._inspect(new_ids)

    def _delete_selected(self):
        """删除任务栏里选中的任务（支持 Ctrl 多选 / Ctrl+A 全选）"""
        ids = [it.data(Qt.UserRole) for it in self.file_list.selectedItems()]
        if not ids:
            self.status_label.setText("先在任务栏里选中要删的任务：Ctrl 点选多个，Ctrl+A 全选")
            return
        names = [f["name"] for f in self.files if f["id"] in ids]
        self.files = [f for f in self.files if f["id"] not in ids]
        for i in ids:
            self.settings.pop(i, None)
        self._refresh_file_list()
        self._rebuild_table()
        self._update_summary()
        self.status_label.setText("已删除 %d 个任务：%s" % (
            len(names), "、".join(names[:3]) + ("…" if len(names) > 3 else "")))

    def _clear_files(self):
        self.files = []
        self.settings = {}
        self.row_map = {}
        self.table.setRowCount(0)
        self._refresh_file_list()
        self._update_summary()

    # ---------- 读工作表
    def _inspect(self, ids):
        pairs = [(fid, next(f for f in self.files if f["id"] == fid)["path"],
                  next(f for f in self.files if f["id"] == fid)["kind"]) for fid in ids]
        self.status_label.setText("正在读取文件内容…")
        self.print_btn.setEnabled(False)
        self.inspect_worker = InspectWorker(pairs)
        self.inspect_worker.fileReady.connect(self._on_file_inspected)
        self.inspect_worker.error.connect(lambda msg: self._append_log("读取失败：%s" % msg))
        self.inspect_worker.finishedAll.connect(self._on_inspect_finished)
        self.inspect_worker.start()

    def _on_file_inspected(self, payload):
        rec = next((f for f in self.files if f["id"] == payload["id"]), None)
        if not rec:
            return
        data = payload["data"]
        rec["sheets"] = data["sheets"]
        rec["error"] = data.get("error")
        rec["chart_sheets"] = data.get("chart_sheets", 0)
        rec["inspected"] = True
        self._refresh_file_list()
        self._rebuild_table()

    def _on_inspect_finished(self):
        self.status_label.setText("就绪")
        self.print_btn.setEnabled(True)
        self._rebuild_table()
        self._update_summary()

    # ---------- 文件列表
    def _refresh_file_list(self):
        self.file_list.blockSignals(True)
        self.file_list.clear()
        for idx, f in enumerate(self.files, 1):
            chosen = sum(1 for s in f["sheets"]
                         if self.settings.get(f["id"], {}).get(s["name"], {}).get("print"))
            if not f["inspected"]:
                meta = "读取中…"
            elif f["error"]:
                meta = "读取失败"
            elif f.get("kind") == "excel":
                meta = "%d表 · 已选%d" % (len(f["sheets"]), chosen)
            elif f.get("kind") == "word":
                meta = "Word · 已选%d" % chosen
            else:
                meta = "PDF · 已选%d" % chosen
            text = "%d. %s  ·  %s" % (idx, f["name"], meta)
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, f["id"])
            item.setToolTip("%s\n%s" % (f["path"], meta))
            self.file_list.addItem(item)
        self.file_list.blockSignals(False)
        self.file_count_chip.setText("%d 个任务" % len(self.files))
        self._fit_file_area()

    def _fit_file_area(self):
        """任务多了就把上面这块自动撑高；用户自己拖过之后就不再自动改"""
        if not hasattr(self, "main_split") or self.split_manual:
            return
        need = min(96 + len(self.files) * TaskChipDelegate.ROW_H, 420)
        sizes = self.main_split.sizes()
        if len(sizes) == 2 and sizes[0] < need:
            total = max(sum(sizes), 700)
            self.main_split.setSizes([need, max(150, total - need)])

    def _on_rows_moved(self, *_args):
        order = []
        for i in range(self.file_list.count()):
            fid = self.file_list.item(i).data(Qt.UserRole)
            order.append(fid)
        by_id = {f["id"]: f for f in self.files}
        self.files = [by_id[i] for i in order if i in by_id] + \
                     [f for f in self.files if f["id"] not in order]
        self._refresh_file_list()
        self._rebuild_table()

    # ---------- 工作表表
    def _rebuild_table(self):
        rows = []
        for f in self.files:
            if not f["inspected"] or f["error"]:
                continue
            for s in f["sheets"]:
                default_on = s.get("visible") not in ("隐藏", "深度隐藏")
                st = self.settings.setdefault(f["id"], {}).setdefault(
                    s["name"], {"print": default_on, "copies": 1, "orient": "keep"})
                rows.append((f, s, st))
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        self.row_map = {}
        self.status_rows = {}
        for i, (f, s, st) in enumerate(rows):
            key = (f["id"], s["name"])
            self.row_map[i] = key
            self.status_rows[key] = i
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked if st["print"] else Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            name_item = QTableWidgetItem(f["name"])
            name_item.setToolTip(f["path"])
            self.table.setItem(i, 1, name_item)
            content = s["name"]
            if s.get("visible") in ("隐藏", "深度隐藏"):
                content += "（隐藏）"
            content_item = QTableWidgetItem(content)
            content_item.setToolTip(s.get("note") or "")
            self.table.setItem(i, 2, content_item)
            pages = "—" if s["pages"] is None else str(s["pages"])
            pages_item = QTableWidgetItem(pages)
            pages_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 3, pages_item)
            self.table.setItem(i, 4, QTableWidgetItem(""))
            spin = GlassStepper(int(st["copies"]))
            spin.setFixedWidth(66)
            spin.valueChanged.connect(lambda v, r=i: self._set_cell(r, "copies", v))
            self.table.setCellWidget(i, 5, spin)
            combo = QComboBox()
            for label, val in (("保持原样", "keep"), ("横向", "landscape"),
                               ("纵向", "portrait"), ("适应一页宽", "fitwidth")):
                combo.addItem(label, val)
            idx = combo.findData(st["orient"])
            combo.setCurrentIndex(max(0, idx))
            combo.setFixedHeight(22)
            combo.currentIndexChanged.connect(
                lambda _i, r=i, c=combo: self._set_cell(r, "orient", c.currentData()))
            self.table.setCellWidget(i, 6, combo)
            if key not in self.print_state:
                self.print_state[key] = ("待打印", "")
            self._paint_status(i, key)
        self.table.blockSignals(False)
        self._update_summary()

    # ---------- 实时打印状态
    STATE_STYLE = {
        "待打印": ("#8b87a3", QColor(255, 255, 255, 0)),
        "正在打印": ("#1f6fa8", QColor(150, 205, 246, 90)),
        "已打印": ("#1f7a5b", QColor(150, 240, 200, 90)),
        "已送出": ("#9a6a16", QColor(255, 224, 168, 120)),
        "失败": ("#a8323b", QColor(255, 190, 200, 110)),
    }

    def _style_of(self, state):
        for k, v in self.STATE_STYLE.items():
            if state.startswith(k):
                return k, v[0], v[1]
        return state, "#8b87a3", QColor(0, 0, 0, 0)

    def _paint_status(self, row, key):
        state, note = self.print_state.get(key, ("待打印", ""))
        item = self.table.item(row, 4)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, 4, item)
        item.setText(state if not note else state)
        item.setTextAlignment(Qt.AlignCenter)
        style_key, color, bg = self._style_of(state)
        item.setForeground(QColor(color))
        item.setBackground(bg)
        item.setToolTip(note or state)
        font = item.font()
        font.setBold(style_key in ("正在打印", "已打印"))
        item.setFont(font)

    def _set_status(self, key, state, note=""):
        self.print_state[key] = (state, note)
        row = self.status_rows.get(key)
        if row is not None:
            self._paint_status(row, key)

    def _on_item_changed(self, item):
        if item.column() != 0:
            return
        fid, sheet = self.row_map.get(item.row(), (None, None))
        if not fid:
            return
        self.settings[fid][sheet]["print"] = item.checkState() == Qt.Checked
        self._update_summary()
        self._refresh_file_list()

    def _set_cell(self, row, key, value):
        fid, sheet = self.row_map.get(row, (None, None))
        if not fid:
            return
        self.settings[fid][sheet][key] = value
        self._update_summary()

    def _apply_all(self, kind, value=None):
        """统一设置（界面上不再放这一排按钮，保留接口给自检/批量操作）"""
        for fid, sheets in self.settings.items():
            for name, st in sheets.items():
                if kind == "copies" and value is not None:
                    st["copies"] = max(1, int(value))
                elif kind == "orient" and value is not None:
                    st["orient"] = value
        self._rebuild_table()

    def _toggle_all(self, on):
        for fid, sheets in self.settings.items():
            for st in sheets.values():
                st["print"] = on
        self._rebuild_table()

    def _update_summary(self):
        sheets = copies = pages = 0
        for fid, sheet_map in self.settings.items():
            f = next((x for x in self.files if x["id"] == fid), None)
            if not f:
                continue
            for s in f["sheets"]:
                st = sheet_map.get(s["name"])
                if not st or not st["print"]:
                    continue
                sheets += 1
                copies += st["copies"]
                pages += (s["pages"] or 0) * st["copies"]
        self.stat_line.setText("已选 %d · 份数 %d · 页数 %s" % (
            sheets, copies, pages if pages else "—"))

    # ---------- 打印
    def _build_items(self):
        items = []
        for f in self.files:
            sheets = [s for s in f["sheets"]
                      if self.settings.get(f["id"], {}).get(s["name"], {}).get("print")]
            if not sheets:
                continue
            items.append({
                "path": f["path"],
                "kind": f.get("kind", "excel"),
                "fid": f["id"],
                "sheets": [{"name": s["name"],
                            "copies": self.settings[f["id"]][s["name"]]["copies"],
                            "orientation": self.settings[f["id"]][s["name"]]["orient"]}
                           for s in sheets],
            })
        return items

    def _start_print(self):
        printer = self.printer_combo.currentData()
        items = self._build_items()
        if not printer:
            self.status_label.setText("请先选择打印机")
            return
        if not items:
            self.status_label.setText("还没有勾选任何要打印的内容")
            return
        total = sum(len(i["sheets"]) for i in items)
        # 建立「打印序号 -> 表格行」映射，让状态列实时变化
        self.print_seq_keys = {}
        seq = 0
        for it in items:
            for sh in it["sheets"]:
                seq += 1
                self.print_seq_keys[seq] = (it.get("fid"), sh["name"])
        for it in items:
            for sh in it["sheets"]:
                key = (it.get("fid"), sh["name"])
                self.print_state[key] = ("待打印", "")
                row = self.status_rows.get(key)
                if row is not None:
                    self._paint_status(row, key)
        self.log.clear()
        self.print_btn.setEnabled(False)
        self.stop_btn.setVisible(True)
        self._printed = 0
        self._total_print = total
        self.status_label.setText("正在打印 · 0/%d" % total)
        self.print_btn.setText("打印中…")
        self._append_log("开始打印，共 %d 项" % total)
        self.print_worker = PrintWorker(printer, items)
        self.print_worker.event.connect(self._on_print_event)
        self.print_worker.start()

    def _stop_print(self):
        if self.print_worker:
            self.print_worker.stop_flag.set()
            self.status_label.setText("正在中止…")

    def _on_print_event(self, ev):
        t = ev.get("type")
        if t == "printer":
            self._append_log("打印机：" + ev.get("message", ""))
        elif t == "file_begin":
            self._append_log("处理 %s（%d/%d）" % (ev["file"], ev["index"], ev["total"]))
        elif t == "sheet_begin":
            key = self.print_seq_keys.get(ev.get("seq"))
            if key:
                self._set_status(key, "正在打印", "%d 份 · 已提交给打印机" % ev.get("copies", 1))
            self.status_label.setText("正在打印 · 第 %d/%d 项：%s %s" % (
                ev.get("seq", 0), getattr(self, "_total_print", 0),
                ev.get("file", ""), ev.get("sheet", "")))
        elif t == "sheet_tick":
            key = self.print_seq_keys.get(ev.get("seq"))
            if key:
                p, tot = ev.get("printed", 0), ev.get("pages", 0)
                text = "正在打印 %d/%d页" % (p, tot) if tot else "正在打印"
                self._set_status(key, text, "已出纸 %d / %d 页" % (p, tot))
            self.status_label.setText("正在打印 · 第 %d/%d 项：%s %s" % (
                ev.get("seq", 0), getattr(self, "_total_print", 0),
                ev.get("file", ""), ev.get("sheet", "")))
        elif t in ("sheet_ok", "sheet_error"):
            key = self.print_seq_keys.get(ev.get("seq"))
            if key:
                if t == "sheet_ok":
                    state = "已打印" if ev.get("confirmed", True) else "已送出"
                    self._set_status(key, state, ev.get("note", ""))
                    self._printed = getattr(self, "_printed", 0) + 1
                else:
                    self._set_status(key, "失败", ev.get("error", ""))
            if t == "sheet_error":
                self._append_log("失败：" + ev.get("error", ""))
            else:
                self._append_log("%s %s · %s" % (ev.get("file", ""),
                                                key[1] if key else ev.get("seq", ""),
                                                ev.get("note", "")))
            self.status_label.setText("正在打印 · 已完成 %d/%d" % (
                getattr(self, "_printed", 0), getattr(self, "_total_print", 0)))
        elif t == "file_error":
            self._append_log("文件打不开：%s —— %s" % (ev["file"], ev["error"]))
        elif t == "fatal":
            self._append_log("中断：" + ev.get("error", ""))
        elif t == "done":
            self._append_log("全部执行完毕")
        elif t == "stopped":
            self._append_log("已手动中止")

        if t in ("done", "fatal", "stopped"):
            self.print_btn.setEnabled(True)
            self.print_btn.setText("开始打印")
            self.stop_btn.setVisible(False)
            ok = getattr(self, "_printed", 0)
            total = getattr(self, "_total_print", 0)
            if t == "done":
                self.status_label.setText("打印结束 · 成功 %d/%d" % (ok, total))
            elif t == "stopped":
                self.status_label.setText("已中止 · 成功 %d/%d" % (ok, total))
            else:
                self.status_label.setText("已中断 · 成功 %d/%d" % (ok, total))

    def _append_log(self, msg):
        self.log.append("[%s] %s" % (time.strftime("%H:%M:%S"), msg))


def _wps_version_info(win32api):
    roots = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Kingsoft", "WPS Office"),
        r"C:\Program Files\Kingsoft\WPS Office",
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
    ]
    found = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        for d in os.listdir(root):
            exe = os.path.join(root, d, "office6", "et.exe")
            if os.path.exists(exe):
                found.append((d, exe))
    if not found:
        return None
    found.sort(key=lambda x: [int(n) if n.isdigit() else 0
                              for n in __import__("re").findall(r"\d+", x[0])])
    try:
        return found[-1][0]
    except Exception:
        return None


def _selftest(argv):
    """无界面自检：python 批量打印.exe --selftest 文件1 文件2 --printer 打印机名"""
    import tempfile
    out_path = os.path.join(tempfile.gettempdir(), "batchprint_selftest.txt")
    files = [a for a in argv if a.lower().endswith(SUPPORTED_EXT)]
    printer = ""
    if "--printer" in argv:
        i = argv.index("--printer")
        if i + 1 < len(argv):
            printer = argv[i + 1]
    copies = 1
    if "--copies" in argv:
        i = argv.index("--copies")
        if i + 1 < len(argv):
            try:
                copies = max(1, int(argv[i + 1]))
            except Exception:  # noqa: BLE001
                copies = 1
    buf = []

    def flush():
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(str(x) for x in buf))

    app = QApplication(sys.argv[:1])
    app.setStyleSheet(QSS)
    w = MainWindow()
    w.show()

    def add():
        w._add_paths(files)
        buf.append("add: %d files" % len(files))
        flush()

    def do_print():
        buf.append("files=%d rows=%d summary=%s" % (
            len(w.files), w.table.rowCount(), w.stat_line.text()))
        for f in w.files:
            buf.append("F %s inspected=%s sheets=%d err=%s" % (
                f["name"], f["inspected"], len(f["sheets"]), f["error"]))
        if copies > 1:
            w._apply_all("copies", copies)
            buf.append("copies applied=%s (summary %s)" % (copies, w.stat_line.text()))
        if printer:
            idx = w.printer_combo.findData(printer)
            if idx >= 0:
                w.printer_combo.setCurrentIndex(idx)
            w._start_print()
            buf.append("print started with " + printer)
        else:
            buf.append("no printer, skip print")
        flush()

    def finish():
        states = {}
        for r in range(w.table.rowCount()):
            it = w.table.item(r, 4)
            s = it.text() if it else ""
            states[s] = states.get(s, 0) + 1
        buf.append("rows=%d states=%s status=%s" % (
            w.table.rowCount(), states, w.status_label.text()))
        buf.append("log=" + w.log.toPlainText().replace("\n", " || "))
        flush()
        w.close()
        app.quit()

    QTimer.singleShot(200, add)
    QTimer.singleShot(12000, do_print)
    QTimer.singleShot(34000, finish)
    app.exec()
    return out_path


def main():
    if "--selftest" in sys.argv:
        path = _selftest([a for a in sys.argv[1:] if a != "--selftest"])
        print("selftest result:", path)
        return
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
