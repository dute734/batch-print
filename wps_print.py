# -*- coding: utf-8 -*-
"""
WPS 表格（ET）COM 打印核心
针对 WPS Office 12.x（ProgID: KET.Application / Ket.Application）
- 读取工作簿内每个工作表的信息（名称、是否隐藏、方向、页数、打印区域）
- 按工作表逐个设置份数与打印方向，然后调用 WPS 静默打印
- 打印机名称必须使用 WPS/Office 的本地化格式："打印机名 在 NeXX:"，传纯名字会报 0x80020009
"""
import os
import re
import time
import threading
import winreg
import win32print
import win32gui
import pythoncom
import win32com.client as win32

# WPS 12.x 优先，兼容其他
PROGIDS = ("KET.Application", "Ket.Application", "Excel.Application")

VIRTUAL_KEYWORDS = (
    "print to pdf", "pdf", "xps document writer", "onenote", "fax",
    "导出为wps pdf", "wps pdf", "pdffactory",
)

XL_PORTRAIT = 1
XL_LANDSCAPE = 2
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2

# WPS 打印时自带的进度窗标题关键字（用户看到的就是这个「打印进程」窗口）
PRINT_DIALOG_TITLES = ("打印进程", "打印进度", "正在打印", "打印中")


def _window_process_name(hwnd):
    try:
        import win32process
        import win32api
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        h = win32api.OpenProcess(0x0410, False, pid)   # QUERY_INFORMATION | VM_READ
        try:
            return os.path.basename(win32process.GetModuleFileNameEx(h, 0)).lower()
        finally:
            win32api.CloseHandle(h)
    except Exception:  # noqa: BLE001
        return ""


WPS_PROCESS_NAMES = ("et.exe", "wps.exe", "wpsoffice.exe", "wpspdf.exe", "ksolaunch.exe")


def hide_print_dialogs():
    """把 WPS 自己弹的「打印进程」进度窗隐藏掉。

    只隐藏、不关闭 —— 关闭会被 WPS 当成『取消打印』，隐藏则打印继续进行。
    只处理 WPS 相关进程的窗口，避免误伤别的软件。
    返回被隐藏的窗口标题列表（便于日志核对）。
    """
    hidden = []

    def _cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            if not any(k in title for k in PRINT_DIALOG_TITLES):
                return True
            if _window_process_name(hwnd) not in WPS_PROCESS_NAMES:
                return True
            win32gui.ShowWindow(hwnd, 0)   # SW_HIDE
            hidden.append(title)
        except Exception:  # noqa: BLE001
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:  # noqa: BLE001
        pass
    return hidden


class PrintDialogGuard(threading.Thread):
    """打印期间后台盯着，一出现 WPS 打印进度窗就藏起来"""

    def __init__(self, interval=0.15):
        super().__init__(daemon=True)
        self.interval = interval
        self._stop = threading.Event()
        self.hidden_count = 0

    def run(self):
        while not self._stop.is_set():
            got = hide_print_dialogs()
            if got:
                self.hidden_count += len(got)
            self._stop.wait(self.interval)

    def stop(self):
        self._stop.set()


# ---------------------------------------------------------------- 打印机相关
def list_printers():
    """返回 [{'name','port','virtual','is_default','status_text'}]"""
    result = []
    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        default = ""
    try:
        raw = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 1)
    except Exception:
        raw = []
    seen = set()
    for item in raw:
        name = item[2] if len(item) > 2 else ""
        if not name or name in seen:
            continue
        seen.add(name)
        low = name.lower()
        is_virtual = any(k in low for k in VIRTUAL_KEYWORDS)
        result.append({
            "name": name,
            "port": port_of(name),
            "virtual": is_virtual,
            "is_default": name == default,
            "status_text": printer_status(name),
        })
    result.sort(key=lambda x: (not x["is_default"], x["virtual"], x["name"]))
    return result


def port_of(name):
    """从 HKCU\\...\\Devices 读取打印机端口，如 Ne03:"""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows NT\CurrentVersion\Devices")
        val, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return val.split(",")[-1].strip()
    except Exception:
        return ""


def printer_status(name):
    """返回人类可读状态，如 就绪 / 脱机 / 暂停 / 缺纸"""
    try:
        h = win32print.OpenPrinter(name)
        info = win32print.GetPrinter(h, 2)
        win32print.ClosePrinter(h)
    except Exception:
        return "未知"
    flags = {
        0x00000001: "暂停",
        0x00000002: "出错",
        0x00000004: "正在删除",
        0x00000008: "缺纸",
        0x00000010: "手动进纸",
        0x00000020: "卡纸",
        0x00000040: "脱机",
        0x00000080: "需要用户干预",
        0x00000100: "等待",
        0x00000200: "预热中",
        0x00000400: "正在打印",
        0x00000800: "处理中",
        0x00001000: "忙",
        0x00002000: "初始化",
        0x00004000: "有输出箱",
        0x00008000: "输入箱满",
        0x00010000: "输出箱满",
    }
    st = info.get("Status", 0) if isinstance(info, dict) else 0
    hits = [v for k, v in flags.items() if st & k]
    if not hits:
        return "就绪"
    return "/".join(hits)


def build_active_printer(name):
    """构造 WPS 需要的 ActivePrinter 字符串（中文 Office/WPS 用「在」，英文用「on」）"""
    port = port_of(name)
    if not port:
        return name
    return "%s 在 %s" % (name, port)


# ---------------------------------------------------------------- WPS 会话
class WpsSession:
    """一次批量任务复用一个 WPS 实例"""

    def __init__(self):
        self.app = None
        self.progid = None
        self.last_error = None
        self._com_ready = False
        try:
            pythoncom.CoInitialize()
            self._com_ready = True
        except Exception:
            self._com_ready = False
        for pid in PROGIDS:
            try:
                app = None
                try:
                    app = win32.DispatchEx(pid)
                except Exception:
                    app = win32.Dispatch(pid)
                app.DisplayAlerts = False
                try:
                    app.Visible = False
                except Exception:
                    pass
                try:
                    app.ScreenUpdating = False
                except Exception:
                    pass
                self.app = app
                self.progid = pid
                break
            except Exception as e:  # noqa: BLE001
                self.last_error = "%s: %s" % (pid, e)
        if self.app is None:
            if self._com_ready:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self._com_ready = False
            raise RuntimeError("无法启动 WPS 表格 COM 接口 —— %s" % self.last_error)

    @property
    def info(self):
        app = self.app
        def safe(attr):
            try:
                return str(getattr(app, attr))
            except Exception:
                return ""
        name = safe("Name")
        ver = safe("Version")
        build = safe("Build")
        # WPS 的 Name 会返回 "Microsoft Excel"，Build 是真实版本号
        return {
            "progid": self.progid,
            "name": name,
            "version": ver,
            "build": build,
            "is_wps": name.lower() != "microsoft excel" or bool(build),
            "label": "WPS 表格 %s（内部版本 %s）" % (ver, build) if build else "WPS 表格 %s" % ver,
        }

    def set_printer(self, printer_name):
        if not printer_name:
            return True, "默认打印机"
        # WPS 表格（KET）要「打印机名 在 NeXX:」；少数环境只认纯名字，
        # 所以两种都试一遍，以回读结果为准。
        target = build_active_printer(printer_name)
        candidates = [target]
        if target != printer_name:
            candidates.append(printer_name)
        errors = []
        for cand in candidates:
            try:
                self.app.ActivePrinter = cand
            except Exception as e:  # noqa: BLE001
                errors.append("%s → 报错 %s" % (cand, e))
                continue
            try:
                now = self.app.ActivePrinter
            except Exception:
                now = cand
            if printer_name.lower() in str(now).lower():
                return True, str(now)
            errors.append("%s → 仍是 %s" % (cand, now))
        return False, "打印机未生效（%s）" % "；".join(errors[:2])

    def open(self, path):
        """打开工作簿；不用只读模式，否则无法修改打印方向"""
        errors = []
        candidates = [
            dict(FileName=path, UpdateLinks=0, ReadOnly=False,
                 IgnoreReadOnlyRecommended=True, AddToMru=False),
            dict(FileName=path, ReadOnly=False),
            None,
        ]
        for kw in candidates:
            try:
                if kw is None:
                    wb = self.app.Workbooks.Open(path)
                else:
                    wb = self.app.Workbooks.Open(**kw)
                try:
                    wb.DisplayAlerts = False
                except Exception:
                    pass
                return wb
            except Exception as e:  # noqa: BLE001
                errors.append(str(e))
        raise RuntimeError("打开失败：%s" % " | ".join(errors[:2]))

    def close_all(self):
        try:
            for _ in range(10):
                if self.app.Workbooks.Count == 0:
                    break
                wb = self.app.Workbooks(1)
                wb.Close(SaveChanges=False)
        except Exception:
            pass

    def quit(self):
        try:
            if self.app is not None:
                self.app.Quit()
        except Exception:
            pass
        self.app = None
        if self._com_ready:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_ready = False


# ---------------------------------------------------------------- 读取工作表信息
ORIENT_NAME = {XL_PORTRAIT: "纵向", XL_LANDSCAPE: "横向"}


def _visible_label(val):
    try:
        v = int(val)
    except Exception:
        return "可见"
    if v == XL_SHEET_HIDDEN:
        return "隐藏"
    if v == XL_SHEET_VERY_HIDDEN:
        return "深度隐藏"
    return "可见"


def inspect_workbook(session, path):
    """读取一个 Excel 文件的工作表清单"""
    data = {
        "path": path,
        "name": os.path.basename(path),
        "sheets": [],
        "chart_sheets": 0,
        "error": None,
    }
    wb = None
    try:
        wb = session.open(path)
        total = wb.Worksheets.Count
        for i in range(1, total + 1):
            ws = wb.Worksheets(i)
            item = {"index": i, "name": str(ws.Name), "visible": "可见",
                    "orientation": "纵向", "pages": None, "print_area": "",
                    "rows": None, "cols": None, "note": ""}
            try:
                item["visible"] = _visible_label(ws.Visible)
            except Exception:
                pass
            try:
                ps = ws.PageSetup
                item["orientation"] = ORIENT_NAME.get(int(ps.Orientation), "纵向")
                item["print_area"] = str(ps.PrintArea or "")
                try:
                    item["pages"] = int(ps.Pages.Count)
                except Exception:
                    item["pages"] = None
            except Exception as e:  # noqa: BLE001
                item["note"] = "读取页面设置失败：%s" % e
            try:
                used = ws.UsedRange
                item["rows"] = int(used.Rows.Count)
                item["cols"] = int(used.Columns.Count)
                if item["rows"] and item["rows"] <= 1 and item["cols"] and item["cols"] <= 1:
                    rng = used.Cells(1, 1)
                    if rng.Value in (None, ""):
                        item["note"] = "空白表"
            except Exception:
                pass
            if item["visible"] != "可见":
                item["note"] = (item["note"] + " " if item["note"] else "") + "该表已隐藏，打印前会自动临时显示"
            data["sheets"].append(item)
        try:
            data["chart_sheets"] = wb.Charts.Count
        except Exception:
            data["chart_sheets"] = 0
        wb.Close(SaveChanges=False)
        wb = None
    except Exception as e:  # noqa: BLE001
        data["error"] = str(e)
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    return data


# ---------------------------------------------------------------- 打印执行
def _jobs_on(printer_name):
    """枚举队列中的任务（取不到就返回空）"""
    try:
        h = win32print.OpenPrinter(printer_name)
    except Exception:
        return []
    try:
        out = []
        for level in (2, 1):
            try:
                raw = win32print.EnumJobs(h, 0, 99, level)
            except Exception:
                continue
            for j in raw:
                if isinstance(j, dict):
                    out.append({"id": j.get("JobId"), "doc": j.get("pDocument", ""),
                                "status": j.get("pStatus", ""), "pages": j.get("TotalPages", 0)})
                elif len(j) > 5:
                    out.append({"id": j[0], "doc": j[4], "status": j[6] if len(j) > 6 else "",
                                "pages": j[7] if len(j) > 7 else 0})
            if out:
                break
        return out
    finally:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass


EXCEL_EXT = (".xlsx", ".xls", ".xlsm", ".xlsb", ".et")


def job_state(printer_name):
    """读取队列里每个任务的真实进度：{job_id: {'pages':总页数, 'printed':已打页数, 'status':状态}}"""
    try:
        h = win32print.OpenPrinter(printer_name)
    except Exception:
        return {}
    out = {}
    try:
        for level in (1, 2):
            jobs = None
            try:
                jobs = win32print.EnumJobs(h, 0, 99, level)
            except Exception:
                continue
            if not jobs:
                if jobs == []:
                    return {}
                continue
            for j in jobs:
                jid = None
                if isinstance(j, dict):
                    jid = j.get("JobId")
                    total = j.get("TotalPages") or 0
                    printed = j.get("PagesPrinted") or 0
                    st = j.get("pStatus") or ""
                else:
                    # JOB_INFO_1 元组：(0 JobId … 10 iTotalPages, 11 iPagesPrinted)
                    jid = j[0]
                    total = j[10] if len(j) > 10 else 0
                    printed = j[11] if len(j) > 11 else 0
                    st = j[6] if len(j) > 6 else ""
                if jid is None:
                    continue
                out[int(jid)] = {"pages": int(total or 0),
                                 "printed": int(printed or 0),
                                 "status": st or ""}
            if out:
                break
        return out
    finally:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass


def _pages_text(printed, total):
    if total:
        return "%d/%d 页" % (printed, total)
    return "%d 页" % printed


def wait_spool_done(printer_name, base_ids, stop_flag=None, on_tick=None,
                    hard_cap=420.0):
    """
    等到「本次提交的任务」全部离开打印队列 —— 这才是真的打完了。
    返回 (state, note)：
      ok      队列已清空，确认出纸完成
      offline 任务还挂在队列里且打印机没在动（脱机/需干预），只算已送出
      timeout 等太久仍未清空
      stopped 用户中止
    """
    base = set(base_ids or ())
    started = time.time()
    started_all = started
    stall_since = None
    max_pages = 0
    last_printed = -1
    limit = 60.0
    while True:
        if stop_flag is not None and stop_flag.is_set():
            return "stopped", "已中止"
        jobs = job_state(printer_name) or {}
        mine = {k: v for k, v in jobs.items() if k not in base}
        if not mine:
            return "ok", "打印完成（用时 %.0f 秒）" % (time.time() - started_all)
        total = sum(v["pages"] for v in mine.values())
        printed = sum(v["printed"] for v in mine.values())
        max_pages = max(max_pages, total)
        last_total_pages = total
        last_max_printed = printed
        if on_tick:
            on_tick(printed, total)
        st = printer_status(printer_name)
        moving = printed > 0
        if st not in ("就绪", "") and not moving:
            if stall_since is None:
                stall_since = time.time()
            if time.time() - stall_since > 15:
                return "offline", "已送入打印队列，打印机未取纸（%s）" % st
        elif moving:
            stall_since = None
        # 只要还在出纸就续命：按剩余页数重新算时限（每页给 4 秒，最少 45 秒）
        if printed > last_printed:
            last_printed = printed
            remain = max(total - printed, 0)
            limit = min(hard_cap, max(45.0, remain * 4.0 + 20.0))
            started = time.time()
        if time.time() - started > limit:
            return "timeout", "队列长时间未清空（%s，仍 %d 个任务在队列）" % (
                _pages_text(printed, total), len(mine))
        time.sleep(0.6)
WORD_EXT = (".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".wps", ".wpt")
PDF_EXT = (".pdf",)


def file_kind(path):
    """按扩展名判断文件类型：excel / word / pdf / None"""
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXT:
        return "excel"
    if ext in WORD_EXT:
        return "word"
    if ext in PDF_EXT:
        return "pdf"
    return None


def _print_excel_file(session, printer_name, item, seq, on_event, stop_flag, verify_seconds):
    """打印一个 Excel 文件里的多个工作表，返回最后一个序号"""
    path = item["path"]
    fname = os.path.basename(path)
    wb = None
    try:
        wb = session.open(path)
    except Exception as e:  # noqa: BLE001
        on_event({"type": "file_error", "file": fname, "error": str(e)})
        return seq

    try:
        for sheet in item["sheets"]:
            if stop_flag is not None and stop_flag.is_set():
                on_event({"type": "stopped"})
                return seq
            seq += 1
            name = sheet["name"]
            copies = max(1, int(sheet.get("copies") or 1))
            mode = sheet.get("orientation") or "keep"
            on_event({"type": "sheet_begin", "seq": seq, "file": fname,
                      "sheet": name, "copies": copies, "orientation": mode})
            try:
                ws = wb.Worksheets(name)
            except Exception as e:  # noqa: BLE001
                on_event({"type": "sheet_error", "seq": seq, "error": "找不到工作表：%s" % e})
                continue
            need_show = False
            orig_visible = None
            try:
                orig_visible = int(ws.Visible)
            except Exception:
                orig_visible = None
            try:
                if orig_visible is not None and orig_visible != XL_SHEET_VISIBLE:
                    need_show = True
                    ws.Visible = XL_SHEET_VISIBLE
                ps = ws.PageSetup
                if mode == "landscape":
                    ps.Orientation = XL_LANDSCAPE
                elif mode == "portrait":
                    ps.Orientation = XL_PORTRAIT
                elif mode == "fitwidth":
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                before = {j["id"] for j in _jobs_on(printer_name)} if verify_seconds else set()
                pre_ids = set(before)          # 本次提交前队列里已有的任务（等待时排除它们）
                new_ids = []
                for _c in range(copies):
                    hide_print_dialogs()
                    ws.PrintOut()
                    hide_print_dialogs()
                    if verify_seconds:
                        waited = 0.0
                        while waited < 3.0:
                            cur = _jobs_on(printer_name)
                            fresh = [j["id"] for j in cur if j["id"] not in before]
                            if fresh:
                                new_ids.append(max(fresh))
                                before.update(fresh)
                                break
                            hide_print_dialogs()
                            time.sleep(0.4)
                            waited += 0.4
                    else:
                        # 没有队列核验时也要给 WPS 一点喘息时间，避免连续 PrintOut 被合并成一份
                        time.sleep(1.0)
                # ---- 真正等出纸：本次提交的任务从队列消失才算打完
                state, wnote = "ok", ""
                if verify_seconds and new_ids:
                    hide_print_dialogs()
                    state, wnote = wait_spool_done(
                        printer_name, pre_ids, stop_flag,
                        on_tick=lambda p, t: on_event({"type": "sheet_tick", "seq": seq,
                                                       "printed": p, "pages": t}))
                elif verify_seconds:
                    # 没抓到任务号（空表等瞬间完成的情况）
                    state, wnote = "quick", ""
                note = "已提交 %d 份" % copies
                if verify_seconds:
                    if state == "ok":
                        note = "已打印完成（%d 份%s）" % (
                            copies, " · " + wnote if wnote and "用时" in wnote else "")
                    elif state == "quick":
                        note = "已提交 %d 份（队列未捕获任务，可能瞬间完成）" % copies
                    else:
                        note = "%d 份 · %s" % (copies, wnote)
                if stop_flag is not None and stop_flag.is_set():
                    on_event({"type": "stopped"})
                    return seq
                on_event({"type": "sheet_ok", "seq": seq, "note": note,
                          "confirmed": state == "ok"})
            except Exception as e:  # noqa: BLE001
                on_event({"type": "sheet_error", "seq": seq, "error": str(e)})
                continue
            finally:
                if need_show and orig_visible is not None:
                    try:
                        ws.Visible = orig_visible
                    except Exception:
                        pass
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    return seq


def print_batch(printer_name, items, on_event, stop_flag=None, verify_seconds=0):
    """
    items: [{'path':..., 'kind':'excel|word|pdf',
             'sheets':[{'name':..., 'copies':int, 'orientation': 'keep|portrait|landscape|fitwidth'}]}]
    on_event: callable(dict) —— 每次状态变化回调
    verify_seconds: >0 时，提交后检查打印队列确认任务确实入队
    """
    seq = 0
    session = None   # WPS 表格
    writer = None    # WPS 文字
    printer_ok = False
    guard = PrintDialogGuard()
    guard.start()
    try:
        for idx, item in enumerate(items, 1):
            if stop_flag is not None and stop_flag.is_set():
                on_event({"type": "stopped"})
                return
            kind = item.get("kind") or "excel"
            path = item["path"]
            fname = os.path.basename(path)
            on_event({"type": "file_begin", "file": fname, "index": idx, "total": len(items)})

            # ---- Excel
            if kind == "excel":
                if session is None:
                    session = WpsSession()
                    on_event({"type": "session", "info": session.info})
                if not printer_ok:
                    ok, msg = session.set_printer(printer_name)
                    on_event({"type": "printer", "ok": ok, "message": msg})
                    if not ok:
                        raise RuntimeError(msg)
                    printer_ok = True
                seq = _print_excel_file(session, printer_name, item, seq,
                                        on_event, stop_flag, verify_seconds)
                continue

            # ---- Word
            if kind == "word":
                import doc_print
                if writer is None:
                    writer = doc_print.WpsWriterSession()
                    on_event({"type": "session", "info": writer.info})
                if not printer_ok:
                    ok, msg = writer.set_printer(printer_name)
                    on_event({"type": "printer", "ok": ok, "message": msg})
                    if not ok:
                        raise RuntimeError(msg)
                    printer_ok = True
                for sheet in item["sheets"]:
                    if stop_flag is not None and stop_flag.is_set():
                        on_event({"type": "stopped"})
                        return
                    seq += 1
                    copies = max(1, int(sheet.get("copies") or 1))
                    mode = sheet.get("orientation") or "keep"
                    label = sheet.get("name") or "整篇文档"
                    on_event({"type": "sheet_begin", "seq": seq, "file": fname,
                              "sheet": label, "copies": copies, "orientation": mode})
                    try:
                        base_ids = {j["id"] for j in _jobs_on(printer_name)} if verify_seconds else set()
                        note = doc_print.print_word(writer, path, copies, mode)
                        state = "ok"
                        if verify_seconds:
                            hide_print_dialogs()
                            state, wnote = wait_spool_done(
                                printer_name, base_ids, stop_flag,
                                on_tick=lambda p, t: on_event({"type": "sheet_tick", "seq": seq,
                                                               "printed": p, "pages": t}))
                            note = ("已打印完成（%d 份%s）" % (copies, " · " + wnote if "用时" in (wnote or "") else "")
                                    if state == "ok"
                                    else "%d 份 · %s" % (copies, wnote))
                        if stop_flag is not None and stop_flag.is_set():
                            on_event({"type": "stopped"})
                            return
                        on_event({"type": "sheet_ok", "seq": seq, "note": note,
                                  "confirmed": state == "ok"})
                    except Exception as e:  # noqa: BLE001
                        on_event({"type": "sheet_error", "seq": seq, "error": str(e)})
                continue

            # ---- PDF
            if kind == "pdf":
                import doc_print
                if not printer_ok:
                    try:
                        doc_print.test_pdf_printer(printer_name)
                        ok, msg = True, printer_name
                    except Exception as e:  # noqa: BLE001
                        ok, msg = False, "打印机不可用：%s" % e
                    on_event({"type": "printer", "ok": ok, "message": msg})
                    if not ok:
                        raise RuntimeError(msg)
                    printer_ok = True
                for sheet in item["sheets"]:
                    if stop_flag is not None and stop_flag.is_set():
                        on_event({"type": "stopped"})
                        return
                    seq += 1
                    copies = max(1, int(sheet.get("copies") or 1))
                    mode = sheet.get("orientation") or "keep"
                    label = sheet.get("name") or "PDF 全文"
                    on_event({"type": "sheet_begin", "seq": seq, "file": fname,
                              "sheet": label, "copies": copies, "orientation": mode})
                    try:
                        base_ids = {j["id"] for j in _jobs_on(printer_name)} if verify_seconds else set()
                        note = doc_print.print_pdf(path, copies, mode, printer_name)
                        state = "ok"
                        if verify_seconds:
                            state, wnote = wait_spool_done(
                                printer_name, base_ids, stop_flag,
                                on_tick=lambda p, t: on_event({"type": "sheet_tick", "seq": seq,
                                                               "printed": p, "pages": t}))
                            note = ("已打印完成（%d 份%s）" % (copies, " · " + wnote if "用时" in (wnote or "") else "")
                                    if state == "ok"
                                    else "%d 份 · %s" % (copies, wnote))
                        if stop_flag is not None and stop_flag.is_set():
                            on_event({"type": "stopped"})
                            return
                        on_event({"type": "sheet_ok", "seq": seq, "note": note,
                                  "confirmed": state == "ok"})
                    except Exception as e:  # noqa: BLE001
                        on_event({"type": "sheet_error", "seq": seq, "error": str(e)})
                continue

            on_event({"type": "file_error", "file": fname, "error": "不支持的文件类型"})

        on_event({"type": "done"})
    except Exception as e:  # noqa: BLE001
        on_event({"type": "fatal", "error": str(e)})
    finally:
        guard.stop()
        if session is not None:
            session.quit()
        if writer is not None:
            writer.quit()
