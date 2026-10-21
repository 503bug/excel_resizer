import os
import sys
import re
import zipfile
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import traceback
from openpyxl.utils import column_index_from_string, get_column_letter

try:
    from lxml import etree
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

try:
    import win32com.client as win32  # noqa: F401
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

XMLNS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
SHEET_RE = re.compile(r'^xl/worksheets/sheet\d+\.xml$')


# ============================================================
#  XML 直改（不依赖 Office，保留所有嵌入资源）
# ============================================================
def _set_cols_width(root, col_start, col_end, width):
    sd = root.find(f'{{{XMLNS_MAIN}}}sheetData')
    cols = root.find(f'{{{XMLNS_MAIN}}}cols')
    target = set(range(col_start, col_end + 1))
    if cols is not None:
        for col in list(cols):
            cmin = int(col.get('min', '1'))
            cmax = int(col.get('max', '1'))
            if set(range(cmin, cmax + 1)) & target:
                cols.remove(col)
        newcol = etree.SubElement(cols, f'{{{XMLNS_MAIN}}}col')
        newcol.set('min', str(col_start))
        newcol.set('max', str(col_end))
        newcol.set('width', str(width))
        newcol.set('customWidth', '1')
    else:
        newcols = etree.Element(f'{{{XMLNS_MAIN}}}cols')
        newcol = etree.SubElement(newcols, f'{{{XMLNS_MAIN}}}col')
        newcol.set('min', str(col_start))
        newcol.set('max', str(col_end))
        newcol.set('width', str(width))
        newcol.set('customWidth', '1')
        if sd is not None:
            sd.addprevious(newcols)
        else:
            root.append(newcols)


def _set_rows_height(root, row_start, row_end, height):
    sd = root.find(f'{{{XMLNS_MAIN}}}sheetData')
    if sd is None:
        return
    for row in sd.findall(f'{{{XMLNS_MAIN}}}row'):
        r_str = row.get('r')
        if r_str is None:
            continue
        r = int(r_str)
        if row_start is not None:
            if row_start <= r <= row_end:
                row.set('ht', str(height))
                row.set('customHeight', '1')
        else:
            row.set('ht', str(height))
            row.set('customHeight', '1')


def modify_sheet_xml(data, col_start, col_end, col_width,
                     row_start, row_end, row_height):
    root = etree.fromstring(data)
    if col_start is not None:
        _set_cols_width(root, col_start, col_end, col_width)
    _set_rows_height(root, row_start, row_end, row_height)
    return etree.tostring(root, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


def process_xlsx_file_xml(fp, col_start, col_end, col_width,
                          row_start, row_end, row_height):
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(fp))
    os.close(fd)
    try:
        with zipfile.ZipFile(fp, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if SHEET_RE.match(item.filename):
                        data = modify_sheet_xml(
                            data, col_start, col_end, col_width,
                            row_start, row_end, row_height)
                    zout.writestr(item, data)
        os.replace(tmp, fp)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


# ============================================================
#  GUI
# ============================================================
class ExcelResizerApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel/WPS 批量调整列宽行高（保留所有内容）")
        root.geometry("800x740")
        self.stop_event = threading.Event()
        self._build_ui()
        self._print_env()

    def _build_ui(self):
        f1 = ttk.Frame(self.root, padding=6)
        f1.pack(fill=tk.X)
        ttk.Label(f1, text="文件夹：").pack(side=tk.LEFT)
        self.folder_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self.folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f1, text="浏览...", command=self.browse_folder).pack(side=tk.LEFT)

        f2 = ttk.LabelFrame(self.root, text="参数", padding=8)
        f2.pack(fill=tk.X, padx=6, pady=4)

        ttk.Label(f2, text="列宽(字符)：").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.col_width_var = tk.StringVar(value="10")
        ttk.Entry(f2, textvariable=self.col_width_var, width=12).grid(row=0, column=1, padx=5)

        ttk.Label(f2, text="行高(磅)：").grid(row=0, column=2, sticky=tk.W, padx=(20, 0), pady=3)
        self.row_height_var = tk.StringVar(value="15")
        ttk.Entry(f2, textvariable=self.row_height_var, width=12).grid(row=0, column=3, padx=5)

        ttk.Label(f2, text="列范围：").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.col_range_var = tk.StringVar(value="A:Z")
        ttk.Entry(f2, textvariable=self.col_range_var, width=12).grid(row=1, column=1, padx=5)
        ttk.Label(f2, text="例 A:Z", foreground="gray").grid(
            row=1, column=2, columnspan=2, sticky=tk.W, padx=(20, 0))

        ttk.Label(f2, text="行范围：").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.row_range_var = tk.StringVar(value="")
        ttk.Entry(f2, textvariable=self.row_range_var, width=12).grid(row=2, column=1, padx=5)
        ttk.Label(f2, text="例 1:100；留空=所有已有行", foreground="gray").grid(
            row=2, column=2, columnspan=2, sticky=tk.W, padx=(20, 0))

        f3 = ttk.LabelFrame(self.root, text="处理引擎", padding=8)
        f3.pack(fill=tk.X, padx=6, pady=4)
        self.engine_var = tk.StringVar(value="xml")
        ttk.Radiobutton(
            f3, text="XML 直改（推荐 · 不依赖 Office · 保留所有内容）",
            variable=self.engine_var, value="xml").pack(anchor=tk.W)
        ttk.Radiobutton(
            f3, text="COM（WPS 优先 / Excel 备选，需 pywin32 + 对应软件）",
            variable=self.engine_var, value="com").pack(anchor=tk.W)
        ttk.Radiobutton(
            f3, text="openpyxl（不推荐，可能丢失嵌入图片/图表）",
            variable=self.engine_var, value="openpyxl").pack(anchor=tk.W)

        f4 = ttk.Frame(self.root, padding=6)
        f4.pack(fill=tk.X)
        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f4, text="包含子文件夹", variable=self.recursive_var).pack(side=tk.LEFT)
        ttk.Button(f4, text="重新检测环境", command=self._print_env).pack(side=tk.RIGHT)

        f5 = ttk.Frame(self.root, padding=6)
        f5.pack(fill=tk.X)
        self.start_btn = ttk.Button(f5, text="开始处理", command=self.start)
        self.start_btn.pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(f5, text="停止", command=self.stop,
                                   state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(6, 0))
        self.progress = ttk.Progressbar(f5, mode="determinate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        self.log = scrolledtext.ScrolledText(self.root, height=18, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

    # ---------- 环境信息 ----------
    def _print_env(self):
        self.log.delete(1.0, tk.END)
        self.log_msg(f"Python  : {sys.version.split()[0]}")
        self.log_msg(f"解释器  : {sys.executable}")
        self.log_msg(f"lxml    : {'✅ 已安装' if HAS_LXML else '❌ 未安装'}")
        self.log_msg(f"pywin32 : {'✅ 已安装' if HAS_PYWIN32 else '❌ 未安装'}")
        if not HAS_LXML:
            self.log_msg(f'   装 lxml   : "{sys.executable}" -m pip install lxml')
        if not HAS_PYWIN32:
            self.log_msg(f'   装 pywin32: "{sys.executable}" -m pip install pywin32')
        self.log_msg("")
        self.log_msg("提示：XML 直改引擎只需 lxml，且能 100% 保留 WPS 嵌入图片。")
        self.log_msg("      若此窗口显示的路径与 `pip install` 用的 Python 不同，请用该路径启动脚本。")
        self.log_msg("")

    # ---------- 工具 ----------
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    # ---------- 启动 / 停止 ----------
    def start(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self.log_msg("请选择有效的文件夹")
            return

        try:
            col_width = float(self.col_width_var.get())
            row_height = float(self.row_height_var.get())
        except ValueError:
            self.log_msg("列宽和行高必须是数字")
            return

        col_start = col_end = None
        col_range = self.col_range_var.get().strip().upper()
        if col_range:
            try:
                if ":" in col_range:
                    a, b = col_range.split(":")
                    col_start = column_index_from_string(a.strip())
                    col_end = column_index_from_string(b.strip())
                else:
                    col_start = col_end = column_index_from_string(col_range)
            except Exception:
                self.log_msg("列范围格式错误，例如 A:Z")
                return

        row_start = row_end = None
        row_range = self.row_range_var.get().strip()
        if row_range:
            try:
                if ":" in row_range:
                    a, b = row_range.split(":")
                    row_start, row_end = int(a), int(b)
                else:
                    row_start = row_end = int(row_range)
            except Exception:
                self.log_msg("行范围格式错误，例如 1:100")
                return

        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log.delete(1.0, tk.END)

        args = (folder, col_width, row_height,
                col_start, col_end, row_start, row_end,
                self.recursive_var.get(), self.engine_var.get())
        threading.Thread(target=self.process, args=args, daemon=True).start()

    def stop(self):
        if self.stop_event.is_set():
            return
        self.stop_event.set()
        self.stop_btn.config(state=tk.DISABLED)
        self.log_msg("⏹ 已发出停止指令，等待当前文件处理完成...")

    # ---------- 主流程 ----------
    def process(self, folder, col_width, row_height,
                col_start, col_end, row_start, row_end, recursive, engine):
        try:
            files = []
            if recursive:
                for dp, _, fs in os.walk(folder):
                    for f in fs:
                        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                            files.append(os.path.join(dp, f))
            else:
                for f in os.listdir(folder):
                    full = os.path.join(folder, f)
                    if (os.path.isfile(full) and f.lower().endswith(".xlsx")
                            and not f.startswith("~$")):
                        files.append(full)

            if not files:
                self.log_msg("未找到 .xlsx 文件")
                return

            self.log_msg(f"共找到 {len(files)} 个 .xlsx 文件\n")
            self.progress["maximum"] = len(files)
            self.progress["value"] = 0

            if self.stop_event.is_set():
                self.log_msg("⏹ 已停止（未开始处理）")
                return

            if engine == "xml":
                if not HAS_LXML:
                    self.log_msg("✗ XML 直改需要 lxml，请先安装：")
                    self.log_msg(f'   "{sys.executable}" -m pip install lxml')
                    return
                self.log_msg("▶ 使用 XML 直改引擎（内容零改动）\n")
                self._process_xml(files, col_width, row_height,
                                  col_start, col_end, row_start, row_end)

            elif engine == "com":
                if not HAS_PYWIN32:
                    self.log_msg("✗ 未安装 pywin32：")
                    self.log_msg(f'   "{sys.executable}" -m pip install pywin32')
                    return
                self._process_com_entry(files, col_width, row_height,
                                        col_start, col_end, row_start, row_end)

            else:
                self.log_msg("⚠ 使用 openpyxl 引擎（可能丢失嵌入图片/图表）\n")
                self._process_openpyxl(files, col_width, row_height,
                                       col_start, col_end, row_start, row_end)
        except Exception:
            self.log_msg("发生异常：\n" + traceback.format_exc())
        finally:
            self.stop_event.clear()
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    # ---------- XML ----------
    def _process_xml(self, files, col_width, row_height,
                     col_start, col_end, row_start, row_end):
        ok = bad = 0
        total = len(files)
        stopped = False
        for i, fp in enumerate(files, 1):
            if self.stop_event.is_set():
                stopped = True
                self.log_msg(f"\n⏹ 已停止，剩余 {total - i + 1} 个文件未处理")
                break
            try:
                process_xlsx_file_xml(fp, col_start, col_end, col_width,
                                      row_start, row_end, row_height)
                ok += 1
                self.log_msg(f"[成功] {os.path.basename(fp)}")
            except Exception as e:
                bad += 1
                self.log_msg(f"[失败] {os.path.basename(fp)} → {e}")
            self.progress["value"] = i
        prefix = "已停止" if stopped else "完成"
        self.log_msg(f"\n===== {prefix}：成功 {ok} 个，失败 {bad} 个 =====")

    # ---------- COM ----------
    def _process_com_entry(self, files, col_width, row_height,
                           col_start, col_end, row_start, row_end):
        if self.stop_event.is_set():
            self.log_msg("⏹ 已停止（未开始处理）")
            return
        app, kind = self._get_com_app()
        if app is None:
            self.log_msg("✗ 未找到可用的 WPS/Excel COM 接口（本机可能未安装）")
            self.log_msg("  建议：切换到「XML 直改」引擎")
            return
        self.log_msg(f"✓ 使用 {kind.upper()} 引擎\n")
        self._process_com(app, kind, files, col_width, row_height,
                          col_start, col_end, row_start, row_end)

    def _get_com_app(self):
        if not HAS_PYWIN32:
            self.log_msg("✗ 未安装 pywin32")
            return None, None
        for pid in ("KET.Application", "ET.Application",
                    "WPS.Application", "Excel.Application"):
            try:
                return win32.Dispatch(pid), \
                    ("wps" if pid.startswith(("KET", "ET", "WPS")) else "excel")
            except Exception:
                continue
        return None, None

    def _process_com(self, app, kind, files, col_width, row_height,
                     col_start, col_end, row_start, row_end):
        ok = bad = 0
        total = len(files)
        stopped = False
        for attr, val in (("Visible", False), ("DisplayAlerts", False),
                          ("ScreenUpdating", False), ("EnableEvents", False),
                          ("AskToUpdateLinks", False)):
            try:
                setattr(app, attr, val)
            except Exception:
                pass
        try:
            for i, fp in enumerate(files, 1):
                if self.stop_event.is_set():
                    stopped = True
                    self.log_msg(f"\n⏹ 已停止，剩余 {total - i + 1} 个文件未处理")
                    break
                abspath = os.path.abspath(fp)
                wb = None
                try:
                    try:
                        wb = app.Workbooks.Open(abspath, 0, False)
                    except Exception:
                        wb = app.Workbooks.Open(abspath)
                    for ws in wb.Worksheets:
                        self._com_apply(ws, col_width, row_height,
                                        col_start, col_end, row_start, row_end)
                    wb.Save()
                    ok += 1
                    self.log_msg(f"[成功] {os.path.basename(fp)}")
                except Exception as e:
                    bad += 1
                    self.log_msg(f"[失败] {os.path.basename(fp)} → {e}")
                finally:
                    if wb is not None:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass
                self.progress["value"] = i
        finally:
            try:
                app.Quit()
            except Exception:
                pass
        prefix = "已停止" if stopped else "完成"
        self.log_msg(f"\n===== {prefix}：成功 {ok} 个，失败 {bad} 个 =====")

    def _com_apply(self, ws, col_width, row_height,
                   col_start, col_end, row_start, row_end):
        try:
            if col_start is not None:
                c1 = get_column_letter(col_start)
                c2 = get_column_letter(col_end)
                ws.Range(f"{c1}:{c2}").ColumnWidth = col_width
            else:
                used = ws.UsedRange
                sc = int(used.Column)
                ec = sc + int(used.Columns.Count) - 1
                ws.Range(
                    f"{get_column_letter(max(1, sc))}:{get_column_letter(max(1, ec))}"
                ).ColumnWidth = col_width
        except Exception as e:
            raise RuntimeError(f"设列宽失败: {e}")

        try:
            if row_start is not None:
                ws.Range(f"{row_start}:{row_end}").RowHeight = row_height
            else:
                used = ws.UsedRange
                sr = int(used.Row)
                er = sr + int(used.Rows.Count) - 1
                ws.Range(f"{max(1, sr)}:{max(1, er)}").RowHeight = row_height
        except Exception as e:
            raise RuntimeError(f"设行高失败: {e}")

    # ---------- openpyxl ----------
    def _process_openpyxl(self, files, col_width, row_height,
                          col_start, col_end, row_start, row_end):
        from openpyxl import load_workbook
        ok = bad = 0
        total = len(files)
        stopped = False
        for i, fp in enumerate(files, 1):
            if self.stop_event.is_set():
                stopped = True
                self.log_msg(f"\n⏹ 已停止，剩余 {total - i + 1} 个文件未处理")
                break
            try:
                wb = load_workbook(fp)
                for ws in wb.worksheets:
                    if col_start is not None:
                        for c in range(col_start, col_end + 1):
                            ws.column_dimensions[get_column_letter(c)].width = col_width
                    elif ws.max_column >= 1:
                        for c in range(1, ws.max_column + 1):
                            ws.column_dimensions[get_column_letter(c)].width = col_width
                    if row_start is not None:
                        for r in range(row_start, row_end + 1):
                            ws.row_dimensions[r].height = row_height
                    else:
                        for r in range(1, ws.max_row + 1):
                            ws.row_dimensions[r].height = row_height
                wb.save(fp)
                ok += 1
                self.log_msg(f"[成功] {os.path.basename(fp)}")
            except Exception as e:
                bad += 1
                self.log_msg(f"[失败] {os.path.basename(fp)} → {e}")
            self.progress["value"] = i
        prefix = "已停止" if stopped else "完成"
        self.log_msg(f"\n===== {prefix}：成功 {ok} 个，失败 {bad} 个 =====")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelResizerApp(root)
    root.mainloop()
