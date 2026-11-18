# -*- coding: utf-8 -*-
"""
Excel 批量打印工具（接入本机 WPS 表格 COM）

启动后自动用浏览器打开界面，全程本机处理，不上传任何文件。
"""
import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from email import policy
from email.parser import BytesParser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

import win32api

import wps_print

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "ExcelBatchPrint", "uploads")
EXCEL_EXT = (".xlsx", ".xls", ".xlsm", ".xlsb")

LOCK = threading.RLock()
STATE = {
    "files": [],          # [{id,name,path,size,sheets,error,inspected}]
    "seq": 0,
    "running": False,
    "progress": [],       # 打印明细行
    "logs": [],
    "summary": {"total": 0, "done": 0, "failed": 0},
    "wps_info": None,
    "printer_note": "",
}
STOP_FLAG = threading.Event()


# ---------------------------------------------------------------- 工具函数
def log(msg):
    with LOCK:
        STATE["logs"].append({"t": time.strftime("%H:%M:%S"), "msg": str(msg)})
        STATE["logs"] = STATE["logs"][-300:]


def wps_version():
    """从已安装的 WPS 目录读取真实版本号（不启动 WPS，速度快）"""
    roots = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Kingsoft", "WPS Office"),
        r"C:\Program Files\Kingsoft\WPS Office",
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
    ]
    found = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        try:
            for d in os.listdir(root):
                exe = os.path.join(root, d, "office6", "et.exe")
                if os.path.exists(exe):
                    found.append((d, exe))
        except Exception:
            pass
    if not found:
        return None
    found.sort(key=lambda x: [int(n) if n.isdigit() else 0 for n in re.findall(r"\d+", x[0])])
    folder, exe = found[-1]
    info = {"folder": folder, "path": exe, "file_version": "", "product_version": ""}
    try:
        vi = win32api.GetFileVersionInfo(exe, "\\")
        ms, ls = vi["FileVersionMS"], vi["FileVersionLS"]
        info["file_version"] = "%d.%d.%d.%d" % (ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF)
        pms, pls = vi["ProductVersionMS"], vi["ProductVersionLS"]
        info["product_version"] = "%d.%d.%d.%d" % (pms >> 16, pms & 0xFFFF, pls >> 16, pls & 0xFFFF)
    except Exception:
        pass
    info["all_versions"] = [f[0] for f in found]
    return info


def next_id():
    STATE["seq"] += 1
    return "f%d" % STATE["seq"]


def add_path(path):
    """把一个 Excel 文件加入任务列表，返回文件记录"""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        return None, "文件不存在"
    if os.path.splitext(path)[1].lower() not in EXCEL_EXT:
        return None, "不是 Excel 文件：%s" % os.path.basename(path)
    with LOCK:
        for f in STATE["files"]:
            if os.path.normcase(f["path"]) == os.path.normcase(path):
                return f, "已在列表中"
        rec = {
            "id": next_id(),
            "name": os.path.basename(path),
            "path": path,
            "size": os.path.getsize(path),
            "sheets": [],
            "error": None,
            "inspected": False,
        }
        STATE["files"].append(rec)
        return rec, ""


def scan_folder(folder):
    added = []
    try:
        names = sorted(os.listdir(folder))
    except Exception as e:  # noqa: BLE001
        return added, str(e)
    for n in names:
        p = os.path.join(folder, n)
        if os.path.isfile(p) and os.path.splitext(n)[1].lower() in EXCEL_EXT:
            rec, msg = add_path(p)
            if rec and not msg:
                added.append(rec["id"])
    return added, ""


def inspect_ids(ids):
    """用同一个 WPS 会话依次读取这些文件的工作表信息"""
    result = []
    session = None
    try:
        session = wps_print.WpsSession()
        with LOCK:
            STATE["wps_info"] = session.info
        for fid in ids:
            with LOCK:
                rec = next((f for f in STATE["files"] if f["id"] == fid), None)
            if not rec:
                continue
            data = wps_print.inspect_workbook(session, rec["path"])
            with LOCK:
                rec["sheets"] = data["sheets"]
                rec["chart_sheets"] = data.get("chart_sheets", 0)
                rec["error"] = data["error"]
                rec["inspected"] = True
            result.append(rec)
    except Exception as e:  # noqa: BLE001
        log("读取工作表失败：%s" % e)
        for fid in ids:
            with LOCK:
                rec = next((f for f in STATE["files"] if f["id"] == fid), None)
            if rec:
                rec["error"] = str(e)
                rec["inspected"] = True
                result.append(rec)
    finally:
        if session:
            session.quit()
    return result


def public_files():
    with LOCK:
        return [{
            "id": f["id"], "name": f["name"], "path": f["path"], "size": f["size"],
            "sheets": f.get("sheets", []), "error": f.get("error"),
            "inspected": f.get("inspected", False),
            "chart_sheets": f.get("chart_sheets", 0),
        } for f in STATE["files"]]


# ---------------------------------------------------------------- 打印线程
def print_worker(printer, items):
    STOP_FLAG.clear()

    def on_event(ev):
        t = ev.get("type")
        with LOCK:
            if t == "session":
                STATE["wps_info"] = ev.get("info")
            elif t == "printer":
                STATE["printer_note"] = ev.get("message", "")
                log("打印机：%s" % ev.get("message", ""))
            elif t == "file_begin":
                log("开始处理 %s（%d/%d）" % (ev["file"], ev["index"], ev["total"]))
            elif t == "sheet_begin":
                STATE["progress"].append({
                    "seq": ev["seq"], "file": ev["file"], "sheet": ev["sheet"],
                    "copies": ev["copies"], "orientation": ev["orientation"],
                    "status": "printing", "note": "",
                })
            elif t == "sheet_ok":
                for r in STATE["progress"]:
                    if r["seq"] == ev["seq"]:
                        r["status"] = "ok"
                        r["note"] = ev.get("note", "")
                STATE["summary"]["done"] += 1
            elif t == "sheet_error":
                for r in STATE["progress"]:
                    if r["seq"] == ev["seq"]:
                        r["status"] = "fail"
                        r["note"] = ev.get("error", "")
                STATE["summary"]["failed"] += 1
                log("失败：%s" % ev.get("error"))
            elif t == "file_error":
                log("文件打不开：%s —— %s" % (ev["file"], ev["error"]))
                STATE["summary"]["failed"] += 1
            elif t == "fatal":
                log("批量打印中断：%s" % ev["error"])
                STATE["summary"]["failed"] += 1
            elif t == "done":
                log("全部任务执行完毕")
            elif t == "stopped":
                log("已手动中止")

    try:
        wps_print.print_batch(printer, items, on_event, STOP_FLAG, verify_seconds=8)
    except Exception as e:  # noqa: BLE001
        log("异常：%s" % e)
        log(traceback.format_exc())
    finally:
        with LOCK:
            STATE["running"] = False
            STOP_FLAG.clear()
        # 收尾：仍在“打印中”的行标注未知
        with LOCK:
            for r in STATE["progress"]:
                if r["status"] == "printing":
                    r["status"] = "unknown"
                    r["note"] = r["note"] or "状态未知"


# ---------------------------------------------------------------- HTTP 服务
class Handler(BaseHTTPRequestHandler):
    server_version = "ExcelBatchPrint/1.0"

    def log_message(self, fmt, *args):  # 去掉控制台噪音
        pass

    # ---------- 基础响应
    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _html(self, text):
        body = text.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    def do_GET(self):  # noqa: N802
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            p = os.path.join(TEMPLATE_DIR, "index.html")
            with open(p, "r", encoding="utf-8") as fh:
                self._html(fh.read())
            return
        if path == "/api/init":
            if STATE["wps_info"] is None:
                with LOCK:
                    STATE["wps_info_file"] = STATE.get("wps_info_file") or wps_version()
            self._json({
                "files": public_files(),
                "printers": wps_print.list_printers(),
                "wps": STATE.get("wps_info") or STATE.get("wps_info_file"),
                "running": STATE["running"],
                "version": "1.0",
            })
            return
        if path == "/api/state":
            with LOCK:
                self._json({
                    "running": STATE["running"],
                    "progress": STATE["progress"],
                    "logs": STATE["logs"],
                    "summary": STATE["summary"],
                    "printer_note": STATE["printer_note"],
                    "files": public_files(),
                })
            return
        if path == "/api/printers":
            self._json({"printers": wps_print.list_printers()})
            return
        self._json({"error": "not found"}, 404)

    def do_POST(self):  # noqa: N802
        path = self.path.split("?")[0]
        try:
            if path == "/api/upload":
                self._handle_upload()
                return
            payload = json.loads(self._body() or b"{}")
            route = {
                "/api/add": self._api_add,
                "/api/pick": self._api_pick,
                "/api/inspect": self._api_inspect,
                "/api/order": self._api_order,
                "/api/remove": self._api_remove,
                "/api/print": self._api_print,
                "/api/stop": self._api_stop,
                "/api/clear": self._api_clear,
            }.get(path)
            if route is None:
                self._json({"error": "not found"}, 404)
                return
            route(payload)
        except Exception as e:  # noqa: BLE001
            log(traceback.format_exc())
            self._json({"error": str(e)}, 500)

    # ---------- 各接口
    def _api_add(self, payload):
        paths = payload.get("paths") or []
        added, skipped = [], []
        for p in paths:
            if os.path.isdir(p):
                ids, err = scan_folder(p)
                added.extend(ids)
                if err:
                    skipped.append(err)
                continue
            rec, msg = add_path(p)
            if rec:
                added.append(rec["id"])
                if msg:
                    skipped.append("%s：%s" % (os.path.basename(p), msg))
            else:
                skipped.append("%s：%s" % (os.path.basename(p), msg))
        self._json({"added": added, "skipped": skipped, "files": public_files()})

    def _api_pick(self, payload):
        mode = payload.get("mode", "files")
        try:
            import pick_files
            paths = pick_files.pick(mode)
        except Exception as e:  # noqa: BLE001
            self._json({"error": "选择框启动失败：%s" % e}, 500)
            return
        if not paths:
            self._json({"paths": [], "files": public_files()})
            return
        added, skipped = [], []
        for p in paths:
            if os.path.isdir(p):
                ids, err = scan_folder(p)
                added.extend(ids)
            else:
                rec, msg = add_path(p)
                if rec:
                    added.append(rec["id"])
                    if msg:
                        skipped.append(msg)
        self._json({"paths": paths, "added": added, "skipped": skipped, "files": public_files()})

    def _api_inspect(self, payload):
        ids = payload.get("ids") or []
        files = inspect_ids(ids)
        self._json({"files": public_files(), "result": files})

    def _api_order(self, payload):
        ids = payload.get("ids") or []
        with LOCK:
            by_id = {f["id"]: f for f in STATE["files"]}
            ordered = [by_id[i] for i in ids if i in by_id]
            for f in STATE["files"]:
                if f["id"] not in ids:
                    ordered.append(f)
            STATE["files"] = ordered
        self._json({"files": public_files()})

    def _api_remove(self, payload):
        ids = set(payload.get("ids") or [])
        with LOCK:
            STATE["files"] = [f for f in STATE["files"] if f["id"] not in ids]
        self._json({"files": public_files()})

    def _api_clear(self, payload):
        with LOCK:
            STATE["files"] = []
            STATE["progress"] = []
            STATE["logs"] = []
            STATE["summary"] = {"total": 0, "done": 0, "failed": 0}
            STATE["printer_note"] = ""
        self._json({"files": []})

    def _api_print(self, payload):
        with LOCK:
            if STATE["running"]:
                self._json({"error": "正在打印中，请先中止"}, 409)
                return
            printer = payload.get("printer") or ""
            items = payload.get("items") or []
            if not items:
                self._json({"error": "没有要打印的工作表"}, 400)
                return
            total = sum(len(i["sheets"]) for i in items)
            STATE["progress"] = []
            STATE["logs"] = []
            STATE["summary"] = {"total": total, "done": 0, "failed": 0}
            STATE["running"] = True
            STATE["printer_note"] = ""
        th = threading.Thread(target=print_worker, args=(printer, items), daemon=True)
        th.start()
        self._json({"ok": True, "total": total})

    def _api_stop(self, payload):
        STOP_FLAG.set()
        log("收到中止指令")
        self._json({"ok": True})

    def _handle_upload(self):
        try:
            ctype = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ctype:
                self._json({"error": "需要 multipart 上传"}, 400)
                return
            body = self._body()
            msg = BytesParser(policy=policy.default).parsebytes(
                b"Content-Type: " + ctype.encode("utf-8") +
                b"\r\nMIME-Version: 1.0\r\n\r\n" + body)
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            added = []
            for part in msg.iter_parts():
                filename = part.get_filename()
                if not filename:
                    continue
                name = os.path.basename(filename)
                dest = os.path.join(UPLOAD_DIR, name)
                if os.path.exists(dest):
                    base, ext = os.path.splitext(name)
                    dest = os.path.join(
                        UPLOAD_DIR,
                        "%s_%d%s" % (base, int(time.time() * 1000) % 100000, ext))
                payload = part.get_payload(decode=True) or b""
                with open(dest, "wb") as fh:
                    fh.write(payload)
                rec, _msg = add_path(dest)
                if rec:
                    added.append(rec["id"])
            self._json({"added": added, "files": public_files()})
        except Exception as e:  # noqa: BLE001
            log(traceback.format_exc())
            self._json({"error": str(e)}, 500)


def find_free_port(start=8765, end=8790):
    for p in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    return start


def main():
    port = find_free_port()
    with LOCK:
        STATE["wps_info_file"] = wps_version()
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError:
        port = find_free_port(8791, 8820)
        httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = "http://127.0.0.1:%d" % port
    print("Excel 批量打印工具已启动：%s" % url)
    print("关闭窗口即停止服务。")
    if os.environ.get("EXCELPRINT_NO_BROWSER") != "1":
        threading.Timer(1.0, lambda: _open_browser(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


def _open_browser(url):
    try:
        os.startfile(url)
    except Exception:
        try:
            subprocess.Popen(["cmd", "/c", "start", "", url])
        except Exception:
            pass


if __name__ == "__main__":
    main()
