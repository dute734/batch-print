# -*- coding: utf-8 -*-
"""批量打印 · Word / PDF 打印核心

- Word：走 WPS 文字 COM（KWps.Application，兼容 Word.Application），
  可切方向、可按份数打印，整篇文档作为一条"内容"。
- PDF：用 PyMuPDF 把页面渲染成图，再用 Windows GDI 直接送到打印机，
  可横向/纵向切纸张方向，可"适应一页宽"（超出部分自动续页）。

所有接口都不弹窗，报错以异常形式返回给上层写日志。
"""
import os

WORD_PROGIDS = ("KWps.Application", "Word.Application")
WORD_EXT = (".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".wps", ".wpt")
PDF_EXT = (".pdf",)

WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1
WD_STAT_PAGES = 2

PDF_RENDER_DPI = 200


# ------------------------------------------------------------------ Word
class WpsWriterSession:
    """一次批量任务复用一个 WPS 文字实例（和表格实例分开）"""

    def __init__(self):
        import pythoncom
        import win32com.client as win32

        self.app = None
        self.progid = None
        self.last_error = None
        self._com_ready = False
        try:
            pythoncom.CoInitialize()
            self._com_ready = True
        except Exception:
            self._com_ready = False
        for pid in WORD_PROGIDS:
            try:
                try:
                    app = win32.DispatchEx(pid)
                except Exception:
                    app = win32.Dispatch(pid)
                app.Visible = False
                try:
                    app.DisplayAlerts = 0
                except Exception:
                    pass
                self.app = app
                self.progid = pid
                break
            except Exception as e:  # noqa: BLE001
                self.last_error = "%s: %s" % (pid, e)
        if self.app is None:
            self._uninit()
            raise RuntimeError("无法启动 WPS 文字 COM 接口 —— %s" % self.last_error)

    def _uninit(self):
        if self._com_ready:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_ready = False

    @property
    def info(self):
        app = self.app
        try:
            name = str(app.Name)
        except Exception:
            name = ""
        try:
            ver = str(app.Version)
        except Exception:
            ver = ""
        return {"progid": self.progid, "name": name, "version": ver}

    def set_printer(self, printer_name):
        from wps_print import build_active_printer, port_of

        if not printer_name:
            return True, "默认打印机"
        # WPS 文字（KWps）只认**纯打印机名**：带「名 在 NeXX:」的写法会被静默忽略，
        # 但保险起见两种写法都试，谁真正生效就用谁。
        candidates = [printer_name]
        port = port_of(printer_name)
        if port:
            candidates.append(build_active_printer(printer_name))
        errors = []
        for target in candidates:
            try:
                self.app.ActivePrinter = target
            except Exception as e:  # noqa: BLE001
                errors.append("%s → 报错 %s" % (target, e))
                continue
            try:
                now = str(self.app.ActivePrinter)
            except Exception:
                now = target
            if printer_name.lower() in now.lower():
                return True, now
            errors.append("%s → 仍是 %s" % (target, now))
        return False, "打印机未生效（%s）" % "；".join(errors[:2])

    def open(self, path):
        errors = []
        candidates = [
            dict(FileName=path, ReadOnly=True, AddToRecentFiles=False,
                 ConfirmConversions=False, Revert=False),
            dict(FileName=path, ReadOnly=True),
            None,
        ]
        for kw in candidates:
            try:
                if kw is None:
                    doc = self.app.Documents.Open(path)
                else:
                    doc = self.app.Documents.Open(**kw)
                return doc
            except Exception as e:  # noqa: BLE001
                errors.append(str(e))
        raise RuntimeError("打开失败：%s" % " | ".join(errors[:2]))

    def quit(self):
        try:
            if self.app is not None:
                self.app.Quit()
        except Exception:
            pass
        self.app = None
        self._uninit()


def _word_blank_pages(doc, pages):
    """找出没有文字的页（空白页）；返回 [页码,...]（1 起）

    做法：逐段落读它所在页码（Range.Information(3) = wdActiveEndPageNumber），
    把该段的可见字符累加到页码上；再叠加锚定在该页的图片/图形，都没内容的页算空白。
    """
    text_len = {}
    try:
        for para in doc.Paragraphs:
            try:
                rng = para.Range
                pg = int(rng.Information(3))
            except Exception:
                continue
            try:
                t = (rng.Text or "").strip()
            except Exception:
                t = ""
            if t:
                text_len[pg] = text_len.get(pg, 0) + len(t)
    except Exception:
        return []

    shape_pages = set()
    for coll in ("InlineShapes", "Shapes"):
        try:
            items = getattr(doc, coll)
            for k in range(1, int(items.Count) + 1):
                try:
                    anchor = items(k).Anchor if coll == "Shapes" else items(k).Range
                    shape_pages.add(int(anchor.Information(3)))
                except Exception:
                    continue
        except Exception:
            continue

    blanks = []
    for p in range(1, int(pages) + 1):
        if text_len.get(p, 0) == 0 and p not in shape_pages:
            blanks.append(p)
    # 整篇都空就别报了，交给调用方判断
    if len(blanks) >= int(pages):
        return []
    return blanks


def _pages_range_text(pages, blanks, total):
    """把「要打印的页」转成 PrintOut 的 PageRanges 字符串（如 1,3-5）；全要打就返回空串"""
    keep = [p for p in range(1, int(total) + 1) if p not in set(blanks or ())]
    if not keep or len(keep) == int(total):
        return ""
    parts = []
    start = prev = keep[0]
    for p in keep[1:]:
        if p == prev + 1:
            prev = p
            continue
        parts.append("%d" % start if start == prev else "%d-%d" % (start, prev))
        start = prev = p
    parts.append("%d" % start if start == prev else "%d-%d" % (start, prev))
    return ",".join(parts)


def inspect_word(session, path):
    """Word 文档没有工作表概念，整篇作为一条内容"""
    data = {"path": path, "name": os.path.basename(path), "sheets": [], "error": None,
            "blank_pages": []}
    doc = None
    try:
        doc = session.open(path)
        pages = None
        try:
            pages = int(doc.ComputeStatistics(WD_STAT_PAGES))
        except Exception:
            pages = None
        try:
            orient = int(doc.PageSetup.Orientation)
        except Exception:
            orient = WD_ORIENT_PORTRAIT
        blanks = _word_blank_pages(doc, pages) if pages else []
        data["blank_pages"] = blanks
        note = "Word 文档按整篇打印"
        if blanks:
            note = "共 %d 页，第 %s 页是空白页（自动跳过）" % (
                pages, "、".join(str(p) for p in blanks))
        data["sheets"].append({
            "index": 1,
            "name": "整篇文档",
            "visible": "文档",
            "orientation": "横向" if orient == WD_ORIENT_LANDSCAPE else "纵向",
            "pages": pages,
            "print_area": "",
            "rows": None,
            "cols": None,
            "note": note,
        })
    except Exception as e:  # noqa: BLE001
        data["error"] = str(e)
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
    return data


def print_word(session, path, copies, orient, blank_pages=None, printer_name=None):
    """打印整篇 Word；返回说明文字

    注意：WPS 文字的 PrintOut(Copies=n) 参数不可靠（实测只出 1 份），
    所以这里统一按「每次 1 份、循环 n 次」提交，和表格的份数处理保持一致。
    blank_pages 里的页码不打印：WPS 文字实测不认 PageRanges / From-To，
    所以有空白页时先导出成临时 PDF，再按 PDF 的页面逐页送打印机（真跳页）。
    """
    import time
    from wps_print import hide_print_dialogs

    n = max(1, int(copies))
    doc = None
    blanks = list(blank_pages or [])
    total = 0
    try:
        doc = session.open(path)
        if orient == "landscape":
            doc.PageSetup.Orientation = WD_ORIENT_LANDSCAPE
        elif orient == "portrait":
            doc.PageSetup.Orientation = WD_ORIENT_PORTRAIT
        try:
            total = int(doc.ComputeStatistics(WD_STAT_PAGES))
        except Exception:
            total = 0
        if blanks and total and printer_name:
            # 转成 PDF，逐页渲染，空白页直接不送 → 这是唯一可靠的跳页办法
            import tempfile
            fd, tmp = tempfile.mkstemp(suffix=".pdf", prefix="bp_w_")
            os.close(fd)
            exported = False
            try:
                try:
                    doc.ExportAsFixedFormat(tmp, 17)      # 17 = PDF
                    exported = True
                except Exception:
                    try:
                        doc.SaveAs2(tmp, 17)
                        exported = True
                    except Exception:
                        exported = False
                if not exported or not os.path.exists(tmp) or os.path.getsize(tmp) == 0:
                    raise RuntimeError("导出 PDF 失败")
                try:
                    doc.Close(0)
                except Exception:
                    pass
                doc = None
                return print_pdf(tmp, n, orient, printer_name, blank_pages=blanks)
            finally:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
        page_ranges = _pages_range_text(None, blanks, total) if (blanks and total) else ""
        kwargs = {"Background": False}
        if orient == "fitwidth":
            kwargs["PrintZoomColumn"] = 1
            kwargs["PrintZoomRow"] = 99
        if page_ranges:
            kwargs["PageRanges"] = page_ranges
        for _ in range(n):
            hide_print_dialogs()
            try:
                doc.PrintOut(**kwargs)
            except Exception:
                # 个别版本不认 PageRanges，退回整篇打印，别把任务弄丢
                kwargs.pop("PageRanges", None)
                doc.PrintOut(**kwargs)
            hide_print_dialogs()
            if n > 1:
                time.sleep(1.0)
        note = "已发送 %d 份" % n
        if orient == "fitwidth":
            note += "（按一页宽缩放）"
        if blanks:
            note += "（跳过 %d 页空白：%s）" % (len(blanks), "、".join(str(p) for p in blanks))
        return note
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass


# ------------------------------------------------------------------ PDF
class _PdfDoc:
    """PDF 打开适配层：优先 pypdfium2（体积小），没有就退回 PyMuPDF"""

    def __init__(self, path):
        self.engine = None
        self._doc = None
        try:  # 轻量引擎
            import pypdfium2 as pdfium
            self._doc = pdfium.PdfDocument(path)
            self.engine = "pdfium"
            self.page_count = len(self._doc)
        except Exception:
            import pymupdf
            self._doc = pymupdf.open(path)
            self.engine = "mupdf"
            self.page_count = self._doc.page_count

    def page_size(self, index):
        """返回页面尺寸（点，1 点 = 1/72 英寸）"""
        if self.engine == "pdfium":
            w, h = self._doc[index].get_size()
            return float(w), float(h)
        rect = self._doc[index].rect
        return float(rect.width), float(rect.height)

    def render(self, index, dpi=PDF_RENDER_DPI):
        """渲染成 PIL 图像（白底 RGB）"""
        from PIL import Image
        if self.engine == "pdfium":
            bmp = self._doc[index].render(scale=dpi / 72.0)
            img = bmp.to_pil()
            if img.mode in ("RGBA", "LA", "P"):
                img = img.convert("RGBA")
                bg = Image.new("RGB", img.size, "white")
                bg.paste(img, mask=img.split()[-1])
                return bg
            return img.convert("RGB")
        page = self._doc[index]
        pix = page.get_pixmap(dpi=dpi)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    def close(self):
        try:
            self._doc.close()
        except Exception:
            pass


def pdf_page_count(path):
    doc = _PdfDoc(path)
    try:
        return doc.page_count
    finally:
        doc.close()


def _page_is_blank(doc, index, dpi=24):
    """把该页按极低 DPI 渲染出来看是不是全白（空白页）；渲染失败就当非空白"""
    try:
        from PIL import Image
        img = doc.render(index, dpi)
        g = img.convert("L")
        lo, hi = g.getextrema()
        return lo >= 245          # 最暗的点都接近白 → 整页空白
    except Exception:
        return False


def _pdf_blank_pages(doc, limit=400):
    """返回空白页列表（1 起）；页数过多的文件只查前 limit 页，避免拖慢导入"""
    out = []
    n = doc.page_count
    if n > limit:
        return []
    for i in range(n):
        if _page_is_blank(doc, i):
            out.append(i + 1)
    if len(out) >= n:
        return []                 # 整份都空白，交给调用方按“全空白”处理
    return out


def inspect_pdf(path):
    data = {"path": path, "name": os.path.basename(path), "sheets": [], "error": None,
            "blank_pages": []}
    doc = None
    try:
        doc = _PdfDoc(path)
        page_count = doc.page_count
        landscape = False
        if page_count:
            w, h = doc.page_size(0)
            landscape = w > h
        blanks = _pdf_blank_pages(doc)
        data["blank_pages"] = blanks
        note = "PDF 按页面渲染后送打印机"
        if blanks:
            note = "共 %d 页，第 %s 页是空白页（自动跳过）" % (
                page_count, "、".join(str(p) for p in blanks))
        data["sheets"].append({
            "index": 1,
            "name": "PDF 全文（%d 页）" % page_count,
            "visible": "文档",
            "orientation": "横向" if landscape else "纵向",
            "pages": page_count,
            "print_area": "",
            "rows": None,
            "cols": None,
            "note": note,
        })
    except Exception as e:  # noqa: BLE001
        data["error"] = str(e)
    finally:
        if doc is not None:
            doc.close()
    return data


def _make_dc(printer_name, landscape):
    """按指定纸张方向创建打印机 DC（不影响系统默认设置）"""
    import win32print
    import win32gui
    import win32ui

    h = win32print.OpenPrinter(printer_name)
    try:
        dm = win32print.GetPrinter(h, 2)["pDevMode"]
    finally:
        win32print.ClosePrinter(h)
    dm.Orientation = 2 if landscape else 1
    hdc = win32gui.CreateDC("WINSPOOL", printer_name, dm)
    return win32ui.CreateDCFromHandle(hdc)


def print_pdf(path, copies, orient, printer_name, dpi=PDF_RENDER_DPI, blank_pages=None):
    """打印 PDF；orient: keep / landscape / portrait / fitwidth；返回说明文字

    blank_pages 里的页码（1 起）会被跳过，不送打印。
    """
    from PIL import ImageWin
    import win32con

    doc = _PdfDoc(path)
    try:
        if not doc.page_count:
            raise RuntimeError("这个 PDF 没有页面")
        skip = {int(p) for p in (blank_pages or ())}
        keep = [i for i in range(doc.page_count) if (i + 1) not in skip]
        if not keep:
            raise RuntimeError("整份 PDF 都是空白页，已跳过")
        w0, h0 = doc.page_size(0)
        base_landscape = w0 > h0
        if orient == "landscape":
            landscape = True
        elif orient == "portrait":
            landscape = False
        else:
            landscape = base_landscape
        for _c in range(max(1, int(copies))):
            dc = _make_dc(printer_name, landscape)
            try:
                pw = dc.GetDeviceCaps(win32con.HORZRES)
                ph = dc.GetDeviceCaps(win32con.VERTRES)
                dc.StartDoc("批量打印 - %s" % os.path.basename(path))
                try:
                    for i in keep:
                        img = doc.render(i, dpi)
                        pw_pt, ph_pt = doc.page_size(i)
                        page_landscape = pw_pt > ph_pt
                        if page_landscape != landscape:
                            img = img.rotate(90, expand=True)
                        if orient == "fitwidth":
                            scale = pw / float(img.width)
                            nw = pw
                            nh = max(1, int(img.height * scale))
                            img = img.resize((nw, nh), Image.LANCZOS)
                            slices = (nh + ph - 1) // ph
                            for s in range(slices):
                                top = s * ph
                                bottom = min(nh, top + ph)
                                piece = img.crop((0, top, nw, bottom))
                                dc.StartPage()
                                ImageWin.Dib(piece).draw(
                                    dc.GetHandleOutput(), (0, 0, nw, bottom - top))
                                dc.EndPage()
                        else:
                            scale = min(pw / float(img.width), ph / float(img.height))
                            nw = max(1, int(img.width * scale))
                            nh = max(1, int(img.height * scale))
                            x = (pw - nw) // 2
                            y = (ph - nh) // 2
                            dc.StartPage()
                            ImageWin.Dib(img).draw(
                                dc.GetHandleOutput(), (x, y, x + nw, y + nh))
                            dc.EndPage()
                finally:
                    dc.EndDoc()
            finally:
                dc.DeleteDC()
        note = "已发送 %d 份（%d 页）" % (max(1, int(copies)), len(keep))
        if skip:
            note += "，跳过 %d 页空白" % len(skip)
        return note
    finally:
        doc.close()


def test_pdf_printer(printer_name):
    """PDF 打印前确认打印机可用"""
    dc = _make_dc(printer_name, False)
    try:
        return True, printer_name
    finally:
        dc.DeleteDC()
