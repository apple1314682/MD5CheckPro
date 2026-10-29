# -*- coding: utf-8 -*-
"""
文件 MD5 校验工具 Pro —— macOS 风格界面
- 拖入文件 / 文件夹（递归）/ 选择文件，自动计算 MD5
- 校验比对：粘贴已知 MD5，自动显示一致 / 不一致
- 每算完一个文件，自动把「文件名 + MD5」等追加写入同目录的 MD5校验记录.xlsx
- 可复制全部结果 / 导出 TXT 报告

运行环境：Python 3.10+（tkinter 不需要），依赖 pywebview、openpyxl
    pip install pywebview openpyxl
"""
import os
import sys
import json
import time
import queue
import hashlib
import threading
import ctypes
import datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _fatal(msg):
    """无控制台（pythonw）时的兑底：写日志 + 弹窗，绝不静默退出"""
    try:
        with open(os.path.join(APP_DIR, '启动错误.log'), 'w', encoding='utf-8') as f:
            f.write(msg)
    except Exception:
        pass
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, 'MD5 校验工具', 0x10)
    except Exception:
        pass
    sys.exit(1)


try:
    import webview
except Exception:
    import traceback
    _fatal('启动失败：缺少依赖 pywebview。\n\n请执行：\n  D:\\python\\python.exe -m pip install pywebview openpyxl\n\n详情：\n' + traceback.format_exc())


def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE，渲染更清晰
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_enable_dpi_awareness()

UI_PATH = os.path.join(APP_DIR, 'ui.html')
ICON_PATH = os.path.join(APP_DIR, 'app.ico')
APP_TITLE = 'MD5 校验工具 Pro'
APP_VERSION = '2.0'

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except Exception:
    EXCEL_OK = False


def human_size(n):
    if n < 1024:
        return '%s B' % n
    x = float(n)
    for u in ('KB', 'MB', 'GB', 'TB'):
        x /= 1024.0
        if x < 1024:
            return '%.1f %s' % (x, u)
    return '%.1f PB' % x


def fmt_ts(ts):
    try:
        return datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return '-'


def set_clipboard(text):
    """写剪贴板：优先 .NET Clipboard（STA 线程，内部自带重试），失败再用 Win32 API 兜底"""
    if not text:
        return False
    try:
        from System.Threading import Thread, ThreadStart, ApartmentState
        box = []

        def _run():
            try:
                from System.Windows.Forms import Clipboard
                Clipboard.SetText(text)
                box.append(True)
            except Exception:
                box.append(False)

        t = Thread(ThreadStart(_run))
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        t.Join(3000)
        if box and box[0]:
            return True
    except Exception:
        pass

    CF_UNICODETEXT, GMEM_MOVEABLE = 13, 0x42
    u32 = ctypes.windll.user32
    k32 = ctypes.windll.kernel32
    for _ in range(20):
        if u32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        return False
    try:
        u32.EmptyClipboard()
        buf = ctypes.create_unicode_buffer(text)
        h = k32.GlobalAlloc(GMEM_MOVEABLE, ctypes.sizeof(buf))
        if not h:
            return False
        p = k32.GlobalLock(h)
        if not p:
            k32.GlobalFree(h)
            return False
        ctypes.memmove(p, buf, ctypes.sizeof(buf))
        k32.GlobalUnlock(h)
        if u32.SetClipboardData(CF_UNICODETEXT, h):
            return True
        k32.GlobalFree(h)
        return False
    finally:
        u32.CloseClipboard()


def hash_file(path, on_progress, abort):
    h = hashlib.md5()
    size = os.path.getsize(path)
    read = 0
    with open(path, 'rb') as f:
        while True:
            if abort.is_set():
                return None
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
            read += len(chunk)
            if size:
                on_progress(read / size)
    return h.hexdigest().upper()


class ExcelLog:
    """固定总表：MD5校验记录.xlsx，每完成一个文件追加一行"""
    HEAD = ['记录时间', '文件名', 'MD5', '大小（字节）', '大小（可读）',
            '创建时间', '修改时间', '耗时（秒）', '校验结果', '文件路径']
    WIDTH = [19, 42, 38, 13, 11, 20, 20, 10, 10, 55]

    def __init__(self):
        self.lock = threading.Lock()
        self.path = os.path.join(APP_DIR, 'MD5校验记录.xlsx')
        self.count = 0
        self.err = None
        self.available = EXCEL_OK
        self._wb = None
        self._ws = None
        if EXCEL_OK:
            self._load()

    def _new(self):
        self._wb = openpyxl.Workbook()
        self._ws = self._wb.active
        self._ws.title = 'MD5记录'
        self._ws.append(self.HEAD)
        for i, w in enumerate(self.WIDTH, 1):
            self._ws.column_dimensions[get_column_letter(i)].width = w
        for c in self._ws[1]:
            c.font = Font(bold=True, color='FFFFFF', size=10, name='微软雅黑')
            c.fill = PatternFill('solid', fgColor='0A84FF')
            c.alignment = Alignment(horizontal='center', vertical='center')
        self._ws.row_dimensions[1].height = 22
        self._ws.freeze_panes = 'A2'
        self.count = 0

    def _load(self):
        if os.path.exists(self.path):
            try:
                self._wb = openpyxl.load_workbook(self.path)
                self._ws = self._wb.active
                self.count = max(self._ws.max_row - 1, 0)
                return
            except Exception as e:
                self.err = '无法读取记录表：%s' % e
        self._new()

    def append(self, row):
        with self.lock:
            if not self.available:
                return {'ok': False, 'err': self.err or '未安装 openpyxl'}
            try:
                self._ws.append(row)
                self._ws.auto_filter.ref = 'A1:%s%d' % (
                    get_column_letter(len(self.HEAD)), self._ws.max_row)
                self._wb.save(self.path)
                self.count += 1
                self.err = None
                return {'ok': True, 'count': self.count,
                        'name': os.path.basename(self.path), 'path': self.path}
            except PermissionError:
                # 记录表正被 Excel/WPS 打开占用 → 另存新表
                alt = os.path.join(APP_DIR, 'MD5校验记录_%s.xlsx' %
                                   datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
                try:
                    self.path = alt
                    self._new()
                    self._ws.append(row)
                    self._wb.save(alt)
                    self.count = 1
                    return {'ok': True, 'count': 1, 'name': os.path.basename(alt),
                            'path': alt, 'note': '原记录表被占用，已另存新表'}
                except Exception as e2:
                    self.err = str(e2)
                    return {'ok': False, 'err': self.err}
            except Exception as e:
                self.err = str(e)
                return {'ok': False, 'err': self.err}

    def open(self):
        try:
            if os.path.exists(self.path):
                os.startfile(self.path)
            else:
                os.startfile(APP_DIR)
            return True
        except Exception:
            return False


class Api:
    def __init__(self):
        self.win = None
        self.q = queue.Queue()
        self.lock = threading.Lock()
        self.gen = 1            # 代数：清空后 +1，旧任务作废
        self.seq = 0
        self.known = set()      # 会话内去重
        self.done = []          # 已完成结果（复制 / 导出用）
        self.total = 0
        self.finished = 0
        self.abort = threading.Event()
        self.excel = ExcelLog()
        self.expected = None    # 校验比对的期望 MD5
        self._maxed = False
        self._ready = False     # 页面就绪前，事件先缓存
        self._pending = []
        self._demo = None

    # ---------- 内部 ----------
    def attach(self, win):
        self.win = win

    def _emit(self, ev):
        if not self._ready or not self.win:
            self._pending.append(ev)
            return
        try:
            self.win.evaluate_js('window.__onEvent(%s)' % json.dumps(ev, ensure_ascii=False))
        except Exception:
            pass

    def _status(self, text=None, busy=None):
        ev = {'type': 'status', 'done': self.finished, 'total': self.total}
        if text is not None:
            ev['text'] = text
        if busy is not None:
            ev['busy'] = busy
        self._emit(ev)

    def __dir__(self):
        # 只把公开方法暴露给前端，避免 pywebview 递归扫描内部属性（.NET 对象会卡死）
        return ['get_init', 'set_expected', 'add_paths', 'stop', 'clear_all',
                'copy_md5', 'copy_all', 'export_txt', 'open_excel',
                'pick_files', 'pick_folder', 'ui_ready',
                'minimize_window', 'toggle_max', 'close_window']

    # ---------- JS API ----------
    def ui_ready(self):
        self._ready = True
        pending, self._pending = self._pending, []
        for ev in pending:
            self._emit(ev)
        self._register_drop()
        if self._demo:
            demo, self._demo = self._demo, None
            self.add_paths(demo)
        return True

    def _register_drop(self):
        """系统级拖放：必须用 pywebview 的 DOM API 注册，JS 端拿不到真实文件路径"""
        def _on_drop(event):
            try:
                files = (event.get('dataTransfer') or {}).get('files') or []
                paths = [f.get('pywebviewFullPath') or f.get('path')
                         for f in files if isinstance(f, dict)]
                paths = [p for p in paths if p]
                if paths:
                    self.add_paths(paths)
            except Exception:
                pass
        try:
            self.win.dom.document.events.drop += _on_drop
        except Exception as e:
            print('DROP-REG-FAIL', repr(e), flush=True)

    def get_init(self):
        return {'version': APP_VERSION,
                'excel': {'ok': self.excel.available,
                          'name': os.path.basename(self.excel.path),
                          'count': self.excel.count if self.excel.available else 0,
                          'err': self.excel.err, 'path': self.excel.path}}

    def set_expected(self, md5):
        v = (md5 or '').strip().lower()
        if len(v) == 32 and all(c in '0123456789abcdef' for c in v):
            self.expected = v
        else:
            self.expected = None
        return True

    def add_paths(self, paths):
        if isinstance(paths, str):
            paths = [paths]
        files = []
        for p in (paths or []):
            if not isinstance(p, str) or not p.strip():
                continue
            p = p.strip().strip('"')
            if not os.path.isabs(p):
                continue
            if os.path.isdir(p):
                for root, dirs, fs in os.walk(p):
                    dirs.sort()
                    for f in sorted(fs):
                        files.append(os.path.join(root, f))
            elif os.path.isfile(p):
                files.append(p)
        added = skipped = 0
        with self.lock:
            g = self.gen
            for p in files:
                p = os.path.abspath(p)
                k = os.path.normcase(p)
                if k in self.known:
                    skipped += 1
                    continue
                self.known.add(k)
                self.seq += 1
                self.total += 1
                added += 1
                self.q.put({'gen': g, 'id': self.seq, 'path': p})
                self._emit({'type': 'queued', 'id': self.seq,
                            'name': os.path.basename(p), 'path': p})
        if added:
            self._status('已加入 %d 个文件' % added, self.total > self.finished)
        return {'added': added, 'skipped': skipped}

    def stop(self):
        self.abort.set()
        while True:
            try:
                job = self.q.get_nowait()
            except queue.Empty:
                break
            self._emit({'type': 'stopped', 'id': job['id']})
        self._status('已停止', False)
        return True

    def clear_all(self):
        with self.lock:
            self.gen += 1
            self.known.clear()
            self.done.clear()
            self.total = 0
            self.finished = 0
        self.abort.set()
        while True:
            try:
                self.q.get_nowait()
            except queue.Empty:
                break
        self.abort.clear()
        self._emit({'type': 'cleared'})
        self._status('就绪', False)
        return True

    def copy_md5(self, md5):
        return set_clipboard(md5 or '')

    def copy_all(self):
        if not self.done:
            return False
        parts = []
        for r in self.done:
            if r.get('error'):
                parts.append('文件：%s\n读取失败：%s' % (r['name'], r['error']))
            else:
                parts.append('文件：%s\n大小：%s（%s 字节）\n创建：%s\n修改：%s\nMD5：%s\n耗时：%.3f 秒' %
                             (r['name'], r['size_h'], r['size'], r['ctime'], r['mtime'],
                              r['md5'], r['elapsed']))
        return set_clipboard('\n------------------------\n'.join(parts))

    def export_txt(self):
        if not self.done:
            return None
        default = 'MD5校验报告_%s.txt' % datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        try:
            p = self.win.create_file_dialog(webview.SAVE_DIALOG, save_filename=default)
        except Exception:
            return None
        if not p:
            return None
        if isinstance(p, (list, tuple)):
            p = p[0]
        lines = []
        for r in self.done:
            if r.get('error'):
                lines.append('文件：%s\n读取失败：%s\n----------------------------------------' %
                             (r['name'], r['error']))
            else:
                lines.append('文件：%s\n大小：%s（%s 字节）\n创建：%s\n修改：%s\nMD5：%s\n耗时：%.3f 秒\n----------------------------------------' %
                             (r['name'], r['size_h'], r['size'], r['ctime'], r['mtime'],
                              r['md5'], r['elapsed']))
        with open(p, 'w', encoding='utf-8-sig') as f:
            f.write('\n'.join(lines))
        return p

    def open_excel(self):
        return self.excel.open()

    def pick_files(self):
        try:
            res = self.win.create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=True, file_types=('所有文件 (*.*)',))
        except Exception:
            res = None
        if res:
            self.add_paths(list(res))
        return True

    def pick_folder(self):
        try:
            res = self.win.create_file_dialog(webview.FOLDER_DIALOG)
        except Exception:
            res = None
        if res:
            if not isinstance(res, (list, tuple)):
                res = [res]
            self.add_paths(list(res))
        return True

    # ---------- 窗口控制（macOS 红绿灯） ----------
    def minimize_window(self):
        try:
            self.win.minimize()
        except Exception:
            pass
        return True

    def toggle_max(self):
        try:
            if self._maxed:
                self.win.restore()
                self._maxed = False
            else:
                self.win.maximize()
                self._maxed = True
        except Exception:
            pass
        return True

    def close_window(self):
        try:
            self.win.destroy()
        except Exception:
            pass
        return True

    # ---------- 工作线程 ----------
    def worker_loop(self):
        while True:
            job = self.q.get()
            with self.lock:
                stale = job['gen'] != self.gen
            if stale:
                self._emit({'type': 'stopped', 'id': job['id']})
                continue
            path, jid = job['path'], job['id']
            name = os.path.basename(path)
            self.abort.clear()
            self._emit({'type': 'hashing', 'id': jid})
            self._status('正在校验 %d/%d：%s' % (self.finished + 1, self.total, name), True)
            t0 = time.perf_counter()
            last = [0.0]

            def cb(frac, jid=jid, last=last):
                now = time.time()
                if now - last[0] > 0.06 or frac >= 1:
                    last[0] = now
                    self._emit({'type': 'progress', 'id': jid, 'frac': round(frac, 4)})

            try:
                md5 = hash_file(path, cb, self.abort)
            except Exception as e:
                elapsed = time.perf_counter() - t0
                rec = {'id': jid, 'name': name, 'path': path, 'error': str(e),
                       'elapsed': elapsed, 'md5': '', 'size': 0, 'size_h': '-',
                       'ctime': '-', 'mtime': '-'}
                with self.lock:
                    self.finished += 1
                    self.done.append(rec)
                rec['excel'] = self.excel.append([fmt_ts(time.time()), name, '读取失败：%s' % e,
                                                  0, '-', '-', '-', round(elapsed, 3), '—', path])
                self._emit(dict(type='done', **rec))
                self._status(None, self.total > self.finished)
                continue
            if md5 is None:
                with self.lock:
                    if job['gen'] == self.gen:
                        self.total -= 1
                self._emit({'type': 'stopped', 'id': jid})
                self._status('已停止', False)
                continue

            elapsed = time.perf_counter() - t0
            try:
                size = os.path.getsize(path)
            except Exception:
                size = 0
            now_s = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            verify = '—'
            if self.expected:
                verify = '一致' if md5.lower() == self.expected else '不一致'
            rec = {'id': jid, 'name': name, 'path': path, 'md5': md5, 'size': size,
                   'size_h': human_size(size),
                   'ctime': fmt_ts(_safe_ctime(path)), 'mtime': fmt_ts(os.path.getmtime(path)),
                   'elapsed': elapsed, 'verify': verify}
            ex = self.excel.append([now_s, name, md5, size, human_size(size),
                                    rec['ctime'], rec['mtime'], round(elapsed, 3), verify, path])
            rec['excel'] = ex
            with self.lock:
                self.finished += 1
                self.done.append(rec)
            self._emit(dict(type='done', **rec))
            if self.finished >= self.total:
                self._status('全部完成，共 %d 个文件' % self.total, False)
            else:
                self._status(None, True)


def _safe_ctime(path):
    try:
        return os.path.getctime(path)
    except Exception:
        return time.time()


def _demo_paths():
    if '--demo' in sys.argv:
        i = sys.argv.index('--demo')
        return sys.argv[i + 1:]
    return None


def main():
    api = Api()
    with open(UI_PATH, 'r', encoding='utf-8') as f:
        ui_html = f.read()
    win = webview.create_window(APP_TITLE, html=ui_html, js_api=api,
                                width=980, height=680, frameless=True,
                                background_color='#141416')
    api.attach(win)

    def on_loaded():
        def later():
            time.sleep(0.4)
            try:
                hwnd = int(win.native.Handle.ToInt64())
                val = ctypes.c_int(2)   # DWMWCP_ROUND，Win11 圆角窗口
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(val), 4)
            except Exception:
                pass
            try:
                from System.Windows.Forms import MethodInvoker
                from System.Drawing import Icon as WinIcon
                if os.path.exists(ICON_PATH):
                    icon = WinIcon(ICON_PATH)

                    def _set_icon():
                        try:
                            win.native.Icon = icon
                        except Exception:
                            pass

                    win.native.BeginInvoke(MethodInvoker(_set_icon))
            except Exception:
                pass

        threading.Thread(target=later, daemon=True).start()

    win.events.loaded += on_loaded

    def bootstrap():
        threading.Thread(target=api.worker_loop, daemon=True).start()
        api._demo = _demo_paths()

    webview.start(bootstrap)


if __name__ == '__main__':
    main()
