# -*- coding: utf-8 -*-
"""
文件 MD5 校验工具 Pro —— macOS 风格界面
- 拖入文件 / 文件夹（递归）/ 选择文件，自动计算 MD5
- 校验比对：粘贴已知 MD5，自动显示一致 / 不一致
- 每算完一个文件，自动把「文件名 + MD5」等追加写入同目录的 MD5校验记录.xlsx
- 可复制全部结果 / 导出 TXT 报告

源码运行：Python 3.10+，依赖 pywebview、openpyxl
    pip install pywebview openpyxl
    python MD5校验工具.py
打包发布（单文件 exe，免 Python 环境）：
    pip install pyinstaller
    pyinstaller --onefile --noconsole --icon app.ico --name "MD5校验工具" ^
        --add-data "ui.html;." --add-data "app.ico;." ^
        --collect-all webview --collect-all pythonnet --collect-all clr_loader ^
        --splash icon.png MD5校验工具.py
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

if getattr(sys, 'frozen', False):
    # PyInstaller 打包运行：资源在解包目录 _MEIPASS，记录表/日志放在 exe 旁边
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
    BUNDLE_DIR = getattr(sys, '_MEIPASS', APP_DIR)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = APP_DIR


WEBVIEW2_REG_KEY = r'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'


def _find_webview2_runtime():
    """从注册表查 WebView2 运行时版本，找不到返回 None（Evergreen 常青版/固定版都会写这里）"""
    try:
        import winreg
    except Exception:
        return 'ok'
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY, 0):
            try:
                k = winreg.OpenKey(root, WEBVIEW2_REG_KEY, 0, winreg.KEY_READ | view)
            except OSError:
                continue
            try:
                pv = winreg.QueryValueEx(k, 'pv')[0]
                if pv and pv != '0.0.0.0':
                    return pv
            except OSError:
                pass
            finally:
                try:
                    k.Close()
                except Exception:
                    pass
    return None


def _fatal(msg, hard=False):
    """无控制台（pythonw/exe）时的兜底：写日志 + 弹窗，绝不静默退出。
    hard=True 用 os._exit（供非主线程调用，sys.exit 在子线程杀不掉进程）"""
    try:
        with open(os.path.join(APP_DIR, '启动错误.log'), 'w', encoding='utf-8') as f:
            f.write(msg)
    except Exception:
        pass
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, 'MD5 校验工具', 0x10)
    except Exception:
        pass
    if hard:
        os._exit(1)
    sys.exit(1)


_BENCH_T0 = time.perf_counter()


def _bench(tag):
    """--bench 启动打点：各阶段耗时（含绝对时间）写入 启动耗时.log，仅测试用"""
    if '--bench' not in sys.argv:
        return
    try:
        with open(os.path.join(APP_DIR, '启动耗时.log'), 'a', encoding='utf-8') as f:
            f.write('%s | %6.0f ms | %s\n' % (
                datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3],
                (time.perf_counter() - _BENCH_T0) * 1000, tag))
    except Exception:
        pass


try:
    import webview
except Exception:
    import traceback
    _fatal('启动失败：缺少依赖 pywebview。\n\n请执行：\n  python -m pip install pywebview openpyxl\n\n详情：\n' + traceback.format_exc())

_bench('imports-done')


def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE，渲染更清晰
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_enable_dpi_awareness()

UI_PATH = os.path.join(BUNDLE_DIR, 'ui.html')
ICON_PATH = os.path.join(BUNDLE_DIR, 'app.ico')
APP_TITLE = 'MD5 校验工具 Pro'
APP_VERSION = '2.6.1'


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
    """固定总表：MD5校验记录.xlsx，每完成一个文件追加一行。
    openpyxl 的导入与记录表加载延迟到首次使用，不拖慢首屏。"""
    HEAD = ['记录时间', '文件名', 'MD5', '大小（字节）', '大小（可读）',
            '创建时间', '修改时间', '耗时（秒）', '校验结果', '文件路径']
    WIDTH = [19, 42, 38, 13, 11, 20, 20, 10, 10, 55]

    def __init__(self):
        self.lock = threading.Lock()
        self.path = os.path.join(APP_DIR, 'MD5校验记录.xlsx')
        self.count = 0
        self.err = None
        self.available = False
        self._tried = False
        self.rows = {}          # 本会话写入的记录 id → 工作表行号

    def status(self):
        return {'ok': self.available, 'loading': not self._tried,
                'name': os.path.basename(self.path),
                'count': self.count if self.available else 0,
                'err': self.err, 'path': self.path}

    def ensure_loaded(self):
        with self.lock:
            if not self._tried:
                self._tried = True
                self._init_engine()
            return self.status()

    def _init_engine(self):
        """持有 self.lock 调用：导入 openpyxl 并加载（或新建）记录表"""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            self._oxl = openpyxl
            self._Font, self._PatternFill, self._Alignment = Font, PatternFill, Alignment
            self._gcl = get_column_letter
        except Exception as e:
            self.available = False
            self.err = '未安装 openpyxl：%s' % e
            return
        if os.path.exists(self.path):
            try:
                self._wb = openpyxl.load_workbook(self.path)
                self._ws = self._wb.active
                self.count = max(self._ws.max_row - 1, 0)
                self.available = True
                return
            except Exception as e:
                self.err = '无法读取记录表：%s' % e
        self._new()
        self.available = True

    def _new(self):
        """持有 self.lock 调用"""
        self._wb = self._oxl.Workbook()
        self._ws = self._wb.active
        self._ws.title = 'MD5记录'
        self._ws.append(self.HEAD)
        for i, w in enumerate(self.WIDTH, 1):
            self._ws.column_dimensions[self._gcl(i)].width = w
        for c in self._ws[1]:
            c.font = self._Font(bold=True, color='FFFFFF', size=10, name='微软雅黑')
            c.fill = self._PatternFill('solid', fgColor='0A84FF')
            c.alignment = self._Alignment(horizontal='center', vertical='center')
        self._ws.row_dimensions[1].height = 22
        self._ws.freeze_panes = 'A2'
        self.count = 0

    def append(self, row, rid=None):
        with self.lock:
            if not self._tried:
                self._tried = True
                self._init_engine()
            if not self.available:
                return {'ok': False, 'err': self.err or '未安装 openpyxl'}
            try:
                self._ws.append(row)
                self._ws.auto_filter.ref = 'A1:%s%d' % (
                    self._gcl(len(self.HEAD)), self._ws.max_row)
                self._wb.save(self.path)
                if rid is not None:
                    self.rows[rid] = self._ws.max_row
                self.count = max(self._ws.max_row - 1, 0)
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
                    self.rows = {}
                    if rid is not None:
                        self.rows[rid] = self._ws.max_row
                    self.count = 1
                    return {'ok': True, 'count': 1, 'name': os.path.basename(alt),
                            'path': alt, 'note': '原记录表被占用，已另存新表'}
                except Exception as e2:
                    self.err = str(e2)
                    return {'ok': False, 'err': self.err}
            except Exception as e:
                self.err = str(e)
                return {'ok': False, 'err': self.err}

    def delete_row(self, rid):
        """删除某条记录对应的工作表行；返回是否删除成功"""
        with self.lock:
            if not self.available or rid not in self.rows:
                return False
            r = self.rows.pop(rid)
            try:
                self._ws.delete_rows(r)
                for k, v in list(self.rows.items()):
                    if v > r:
                        self.rows[k] = v - 1
                self._wb.save(self.path)
                self.count = max(self._ws.max_row - 1, 0)
                self.err = None
                return True
            except Exception as e:
                self.err = str(e)
                return False

    def clear_data(self):
        """清空记录表全部数据行（保留表头，含历史记录）；返回删除的行数，失败返回 -1"""
        with self.lock:
            if not self._tried:
                self._tried = True
                self._init_engine()
            if not self.available:
                return -1
            try:
                removed = max(self._ws.max_row - 1, 0)
                if removed:
                    self._ws.delete_rows(2, removed)
                    self._wb.save(self.path)
                self.rows = {}
                self.count = 0
                self.err = None
                return removed
            except Exception as e:
                self.err = str(e)
                return -1

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
        self.cancelled = set()  # 被删除但工作线程还没处理完的任务 id

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
                'delete_records', 'reorder', 'copy_md5', 'copy_all', 'export_txt', 'open_excel',
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
        return {'version': APP_VERSION, 'excel': self.excel.status()}

    def _load_excel_async(self):
        """后台线程加载 Excel 组件与记录表，完成后推送给界面刷新底栏"""
        def job():
            st = self.excel.ensure_loaded()
            self._emit({'type': 'excel', **st})
        threading.Thread(target=job, daemon=True).start()

    def _watchdog(self, sec=20):
        """页面迟迟没就绪（WebView2 缺失/过旧回退到坏引擎）时给出明确提示"""
        def job():
            time.sleep(sec)
            if not self._ready:
                _fatal('界面加载超时：WebView2 运行时缺失、过旧或不可用。\n\n'
                       '请到下面地址安装最新「Evergreen 常青版」后重试：\n'
                       'https://developer.microsoft.com/microsoft-edge/webview2/',
                       hard=True)
        threading.Thread(target=job, daemon=True).start()

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
            if job['id'] in self.cancelled:
                self._emit({'type': 'removed', 'id': job['id']})
            else:
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
        self.cancelled.clear()
        self.excel.clear_data()   # 清空列表同时清空 Excel 全部记录（含历史）
        self._emit({'type': 'cleared'})
        self._emit({'type': 'excel', **self.excel.status()})
        self._status('就绪', False)
        return True

    def delete_records(self, ids):
        """删除选中的记录：移除列表卡片，并删除 Excel 中对应的行"""
        if not isinstance(ids, (list, tuple)):
            ids = [ids]
        removed = excel_deleted = 0
        emits = []
        with self.lock:
            for rid in ids:
                try:
                    rid = int(rid)
                except (TypeError, ValueError):
                    continue
                rec = next((r for r in self.done if r['id'] == rid), None)
                if rec is not None:
                    if self.excel.delete_row(rid):
                        excel_deleted += 1
                    self.done.remove(rec)
                    self.known.discard(os.path.normcase(rec['path']))
                    if self.total > 0:
                        self.total -= 1
                    if self.finished > 0:
                        self.finished -= 1
                    removed += 1
                    emits.append({'type': 'removed', 'id': rid})
                else:
                    # 还在排队/计算中：标记取消，由工作线程发出 removed
                    self.cancelled.add(rid)
                    removed += 1
                    emits.append({'type': 'removed', 'id': rid})
        for ev in emits:
            self._emit(ev)
        if removed:
            self._status(None, self.total > self.finished)
            self._emit({'type': 'excel', **self.excel.status()})
        return {'removed': removed, 'excel_deleted': excel_deleted}

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

    def reorder(self, ids):
        """按界面顺序重排已完成记录（复制 / 导出 TXT 顺序跟随列表）"""
        try:
            order = [int(x) for x in (ids or [])]
        except (TypeError, ValueError):
            return False
        idset = set(order)
        with self.lock:
            byid = {r['id']: r for r in self.done}
            self.done = ([byid[i] for i in order if i in byid] +
                         [r for r in self.done if r['id'] not in idset])
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
                if job['id'] in self.cancelled:
                    self._emit({'type': 'removed', 'id': job['id']})
                else:
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
                                                  0, '-', '-', '-', round(elapsed, 3), '—', path], rid=jid)
                self._emit(dict(type='done', **rec))
                self._status(None, self.total > self.finished)
                continue
            if md5 is None:
                with self.lock:
                    if job['gen'] == self.gen:
                        self.total -= 1
                if jid in self.cancelled:
                    self._emit({'type': 'removed', 'id': jid})
                    self._status(None, self.total > self.finished)
                else:
                    self._emit({'type': 'stopped', 'id': jid})
                    self._status('已停止', False)
                continue

            if jid in self.cancelled:
                with self.lock:
                    self.total -= 1
                self._emit({'type': 'removed', 'id': jid})
                self._status(None, self.total > self.finished)
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
                                    rec['ctime'], rec['mtime'], round(elapsed, 3), verify, path], rid=jid)
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


def _autoquit_seconds():
    """--autoquit N：N 秒后自动退出（自动化测试用）"""
    if '--autoquit' in sys.argv:
        i = sys.argv.index('--autoquit')
        try:
            return float(sys.argv[i + 1])
        except (IndexError, ValueError):
            return None
    return None


def _ensure_single_instance():
    """已有实例运行时，把旧窗口调到前台并退出（避免多实例并发写同一个 Excel 记录表）"""
    try:
        k32 = ctypes.WinDLL('kernel32', use_last_error=True)
        k32.CreateMutexW(None, False, 'Local\\MD5校验工具Pro单实例')
        if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
            u32 = ctypes.windll.user32
            hwnd = u32.FindWindowW(None, APP_TITLE)
            if hwnd:
                u32.ShowWindow(hwnd, 9)  # SW_RESTORE
                u32.SetForegroundWindow(hwnd)
                u32.FlashWindow(hwnd, True)
            if getattr(sys, 'frozen', False):
                try:
                    import pyi_splash
                    pyi_splash.close()
                except Exception:
                    pass
            return False
    except Exception:
        pass
    return True


def main():
    _bench('main-start')
    if not _ensure_single_instance():
        return
    pv = _find_webview2_runtime()
    if pv is None:
        _fatal('未检测到 WebView2 运行时（本工具的界面依赖它）。\n\n'
               'Win10/11 一般自带；如缺失请到下面地址安装「Evergreen 常青版」：\n'
               'https://developer.microsoft.com/microsoft-edge/webview2/')

    api = Api()
    with open(UI_PATH, 'r', encoding='utf-8') as f:
        ui_html = f.read()
    # easy_drag=False：关闭 pywebview 默认的「任意位置按住拖动窗口」，
    # 否则拖动文件卡片时整个窗口会跟着跑（窗口拖拽与卡片排序冲突）。标题栏拖动走区域判断，不受影响。
    win = webview.create_window(APP_TITLE, html=ui_html, js_api=api,
                                width=980, height=680, frameless=True,
                                background_color='#141416', hidden=True,
                                easy_drag=False)
    api.attach(win)
    _bench('window-created')

    def on_loaded():
        _bench('page-loaded')
        # 关闭 PyInstaller 启动画面（如有），窗口就绪后才显示，避免黑屏等待
        if getattr(sys, 'frozen', False):
            try:
                import pyi_splash
                pyi_splash.close()
            except Exception:
                pass
        try:
            win.show()
            _bench('window-shown')
        except Exception:
            pass

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
        _bench('gui-start')
        threading.Thread(target=api.worker_loop, daemon=True).start()
        api._load_excel_async()
        api._watchdog(20)
        api._demo = _demo_paths()
        autoquit = _autoquit_seconds()
        if autoquit:
            t = threading.Timer(autoquit, api.close_window)
            t.daemon = True
            t.start()

    # 持久化 WebView2 用户目录：二次启动热启动更快；配合单实例互斥避免目录锁冲突。
    # 目录优先级：LOCALAPPDATA → USERPROFILE\AppData\Local（兑底，绝不落在 exe 目录里攒缓存）
    base = (os.environ.get('LOCALAPPDATA')
            or (os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local')
                if os.environ.get('USERPROFILE') else None)
            or APP_DIR)
    storage = os.path.join(base, 'MD5校验工具')
    try:
        os.makedirs(storage, exist_ok=True)
    except Exception:
        storage = None
    webview.start(bootstrap, private_mode=False, storage_path=storage)


if __name__ == '__main__':
    main()
