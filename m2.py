#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys, fnmatch, argparse, shutil, ctypes, math, subprocess, time
from contextlib import contextmanager
from pathlib import Path

from PyQt5 import QtCore
from PyQt5.QtCore import (
    Qt, QDir, QUrl, QDateTime, QSortFilterProxyModel,
    pyqtSignal, QSettings, QEvent, QTimer, QSize, QAbstractTableModel,
    QIdentityProxyModel, QElapsedTimer
)
from PyQt5.QtGui import (
    QDesktopServices, QPalette, QColor, QKeySequence, QIcon,
    QStandardItemModel, QStandardItem, QPainter, QPixmap, QPen, QBrush,
    QCursor, QPolygonF, QGuiApplication, QFont
)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTreeView, QFileSystemModel,
    QLineEdit, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout,
    QAction, QInputDialog, QMessageBox, QAbstractItemView,
    QMenu, QStyle, QHeaderView, QScrollArea, QFrame, QLabel, QShortcut,
    QToolButton, QDialog, QDialogButtonBox, QTableWidget, QTableWidgetItem,
    QCheckBox, QFileDialog, QProgressDialog, QToolTip, QSizePolicy, QFileIconProvider,
    QComboBox, QSpacerItem
)

# -------------------- Perf / Debug --------------------
DEBUG = True
def dlog(msg):
    if DEBUG:
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")

@contextmanager
def perf(name):
    t0 = time.perf_counter()
    try:
        yield
    finally:
        dt = (time.perf_counter() - t0) * 1000
        dlog(f"{name} took {dt:.1f} ms")

ORG_NAME = "MultiPane"
APP_NAME = "Multi-Pane File Explorer"

# -------------------- Global UI Tuning --------------------
FONT_PT = 9.5
UI_H = max(22, int(round(FONT_PT * 2.6)))
GRID_GAPS    = {2: 5, 3: 3, 4: 2}
GRID_MARG_LR = {2: 8, 3: 6, 4: 6}
PANE_MARGIN = (4, 1, 4, 1)
ROW_SPACING = 4

ITEM_VPAD   = 1
TREE_PAD    = 1
HEADER_VPAD = 1
HEADER_HPAD = 5
CONTROL_VPAD= 1
CONTROL_HPAD= 6

CRUMB_MAX_SEG_W = 180
ALWAYS_GENERIC_ICONS = False

# -------------------- Optional: pywin32 --------------------
HAS_PYWIN32 = True
try:
    import pythoncom
    import win32con, win32gui, win32api
    from win32com.shell import shell, shellcon
except Exception:
    HAS_PYWIN32 = False

# -------------------- Optional: send2trash --------------------
try:
    from send2trash import send2trash
    HAS_SEND2TRASH = True
except Exception:
    HAS_SEND2TRASH = False

# -------------------- HiDPI --------------------
def _enable_win_per_monitor_v2():
    if sys.platform != "win32": return
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)); return
    except Exception: pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try: ctypes.windll.user32.SetProcessDPIAware()
        except Exception: pass

os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

# -------------------- Helpers --------------------
def _normalize_fs_path(p: str) -> str:
    try: p = os.path.normpath(p)
    except Exception: pass
    if os.name == "nt" and len(p) == 2 and p[1] == ":": p = p + os.sep
    return p

def nice_path(p: str) -> str:
    try: return str(Path(p).resolve())
    except Exception: return _normalize_fs_path(p)

def human_size(n: int) -> str:
    if n is None: return ""
    size = float(n); units = ["B", "KB", "MB", "GB", "TB", "PB"]; i = 0
    while size >= 1024 and i < len(units)-1: size /= 1024.0; i += 1
    if i == 0: return f"{int(size)} B"
    return f"{size:.1f} {units[i]}" if size < 10 else f"{size:.0f} {units[i]}"

def unique_dest_path(dst_dir: str, name: str) -> str:
    base, ext = os.path.splitext(name); candidate = name; i = 1
    while os.path.exists(os.path.join(dst_dir, candidate)):
        suffix = " - Copy" if i == 1 else f" - Copy ({i})"
        candidate = f"{base}{suffix}{ext}"; i += 1
    return os.path.join(dst_dir, candidate)

def remove_any(path: str):
    if not os.path.exists(path): return
    if os.path.isdir(path) and not os.path.islink(path): shutil.rmtree(path)
    else: os.remove(path)

def move_with_collision(src: str, dst_dir: str) -> str:
    name = os.path.basename(src); dst = os.path.join(dst_dir, name)
    if os.path.exists(dst): dst = unique_dest_path(dst_dir, name)
    return shutil.move(src, dst)

def recycle_to_trash(paths: list, hwnd: int = 0) -> bool:
    if not paths: return True
    if HAS_SEND2TRASH:
        try:
            for p in paths: send2trash(p)
            return True
        except Exception as e:
            if DEBUG: print("[delete] send2trash failed:", e)
    if HAS_PYWIN32:
        try:
            pFrom = ("\0".join(_normalize_fs_path(p) for p in paths) + "\0\0")
            flags = (shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION |
                     shellcon.FOF_NOERRORUI | shellcon.FOF_SILENT)
            res, aborted = shell.SHFileOperation((int(hwnd), shellcon.FO_DELETE, pFrom, None, flags, False, None, None))
            return (res == 0) and (not aborted)
        except Exception as e:
            if DEBUG: print("[delete] SHFileOperation failed:", e)
    try:
        for p in paths: remove_any(p)
        return True
    except Exception as e:
        if DEBUG: print("[delete] fallback remove failed:", e)
        return False

# -------------------- Background File Ops (with per-file conflicts) --------------------
class FileOpWorker(QtCore.QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished_ok = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, op: str, srcs: list, dst_dir: str, conflict_map: dict | None = None):
        super().__init__()
        self.op = op  # "copy" or "move"
        self.srcs = list(srcs)
        self.dst_dir = dst_dir
        self.conflict_map = dict(conflict_map or {})  # src -> "overwrite"|"skip"|"copy"
        self._cancel = False
        self._total = 0
        self._done = 0

    def cancel(self): self._cancel = True

    def _iter_files(self, path):
        if os.path.isdir(path) and not os.path.islink(path):
            for root, dirs, files in os.walk(path):
                for f in files:
                    fp = os.path.join(root, f)
                    try: size = os.path.getsize(fp)
                    except Exception: size = 0
                    yield fp, size
        else:
            try: size = os.path.getsize(path)
            except Exception: size = 0
            yield path, size

    def _size_of(self, path) -> int:
        return sum(sz for _, sz in self._iter_files(path))

    def _calc_total(self):
        self._total = max(1, sum(self._size_of(s) for s in self.srcs))

    def _tick_progress(self, delta_bytes):
        self._done += max(0, int(delta_bytes))
        self.progress.emit(min(100, int(self._done * 100 / self._total)))

    def _copy_file(self, src, dst):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
            while True:
                if self._cancel: return
                buf = fsrc.read(1024 * 1024)
                if not buf: break
                fdst.write(buf)
                self._tick_progress(len(buf))
        try: shutil.copystat(src, dst, follow_symlinks=True)
        except Exception: pass

    def _copy_dir_recursive(self, src_dir, dst_dir):
        for root, dirs, files in os.walk(src_dir):
            if self._cancel: return
            rel = os.path.relpath(root, src_dir)
            target_root = os.path.join(dst_dir, "" if rel == "." else rel)
            os.makedirs(target_root, exist_ok=True)
            for f in files:
                if self._cancel: return
                sfile = os.path.join(root, f)
                dfile = os.path.join(target_root, f)
                self._copy_file(sfile, dfile)

    def run(self):
        try:
            self._calc_total()
            self.status.emit(f"Preparing {self.op} …")

            for src in self.srcs:
                if self._cancel: break
                base = os.path.basename(src)
                dst = os.path.join(self.dst_dir, base)
                exists = os.path.exists(dst)
                action = self.conflict_map.get(src) if exists else None  # None for non-conflict

                # ---- COPY ----
                if self.op == "copy":
                    if os.path.isdir(src) and not os.path.islink(src):
                        if exists:
                            if action == "skip":
                                self._tick_progress(self._size_of(src)); continue
                            elif action == "copy":
                                dst = unique_dest_path(self.dst_dir, base)
                            elif action == "overwrite":
                                try: shutil.rmtree(dst)
                                except Exception: pass
                        os.makedirs(dst, exist_ok=True)
                        self._copy_dir_recursive(src, dst)
                    else:
                        if exists:
                            if action == "skip":
                                self._tick_progress(self._size_of(src)); continue
                            elif action == "copy":
                                dst = unique_dest_path(self.dst_dir, base)
                            # overwrite → 덮어씀
                        self._copy_file(src, dst)

                # ---- MOVE ----
                else:
                    if exists:
                        if action == "skip":
                            self._tick_progress(self._size_of(src)); continue
                        elif action == "copy":  # keep both → move with new name
                            dst = unique_dest_path(self.dst_dir, base)
                        elif action == "overwrite":
                            try:
                                if os.path.isdir(dst) and not os.path.islink(dst): shutil.rmtree(dst)
                                else: os.remove(dst)
                            except Exception: pass
                    try:
                        final = shutil.move(src, dst if exists and action=="copy" else self.dst_dir)
                        self._tick_progress(self._size_of(final if os.path.exists(final) else src))
                    except Exception:
                        # 폴백: copy 후 삭제
                        if os.path.isdir(src) and not os.path.islink(src):
                            if os.path.exists(dst) and action == "overwrite":
                                try: shutil.rmtree(dst)
                                except Exception: pass
                            if os.path.exists(dst) and action == "copy":
                                dst = unique_dest_path(self.dst_dir, base)
                            os.makedirs(dst, exist_ok=True)
                            self._copy_dir_recursive(src, dst)
                            if not self._cancel:
                                shutil.rmtree(src)
                        else:
                            if os.path.exists(dst) and action == "overwrite":
                                try: os.remove(dst)
                                except Exception: pass
                            if os.path.exists(dst) and action == "copy":
                                dst = unique_dest_path(self.dst_dir, base)
                            self._copy_file(src, dst)
                            if not self._cancel:
                                try: os.remove(src)
                                except Exception: pass

                self.progress.emit(min(100, int(self._done * 100 / self._total)))

            if self._cancel:
                self.error.emit("Operation cancelled."); return
            self.progress.emit(100); self.finished_ok.emit()
        except Exception as e:
            self.error.emit(str(e))

# -------------------- Styles --------------------
def _common_css():
    return f"""
    QWidget {{ font-family: Segoe UI, Pretendard, "Noto Sans", sans-serif; font-size: {FONT_PT}pt; }}
    QScrollArea, QAbstractScrollArea {{ padding: 0; margin: 0; border: 0; }}
    QAbstractScrollArea::viewport {{ margin: 0; padding: 0; }}
    QLineEdit[clearButtonEnabled="true"] {{ padding-right: 22px; }}
    QToolTip {{ border: 1px solid rgba(0,0,0,0.25); }}

    QTreeView {{ padding: {TREE_PAD}px; }}
    QTreeView::item {{ padding: {ITEM_VPAD}px 6px; }}
    QHeaderView::section {{ padding: {HEADER_VPAD}px {HEADER_HPAD}px; }}
    QLineEdit, QPushButton, QToolButton {{ padding: {CONTROL_VPAD}px {CONTROL_HPAD}px; }}

    QLabel#crumbSep {{ padding: 0 2px; }}
    """

def apply_dark_style(app: QApplication):
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(28, 30, 34))
    pal.setColor(QPalette.Base, QColor(22, 24, 28))
    pal.setColor(QPalette.AlternateBase, QColor(30, 32, 36))
    pal.setColor(QPalette.Text, QColor(230, 233, 238))
    pal.setColor(QPalette.ButtonText, QColor(230, 233, 238))
    pal.setColor(QPalette.WindowText, QColor(230, 233, 238))
    pal.setColor(QPalette.ToolTipBase, QColor(255,255,255))
    pal.setColor(QPalette.ToolTipText, QColor(30,30,30))
    pal.setColor(QPalette.Highlight, QColor(64, 128, 255))
    pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    pal.setColor(QPalette.Button, QColor(38, 40, 46))
    app.setPalette(pal)
    app.setStyleSheet(_common_css() + """
        QMainWindow { background: #1C1E22; }
        QLineEdit, QPushButton, QToolButton {
            background: #26282E; border: 1px solid #33363D; border-radius: 8px;
            color: #E6E9EE; min-height: 0px;
        }
        QLineEdit:focus { border: 1px solid #5E9BFF; }
        QTreeView {
            background: #16181C; alternate-background-color: #1E2026;
            border: 1px solid #2B2E34; border-radius: 10px;
        }
        QTreeView::item { color: #E6E9EE; }
        QTreeView::item:selected { background: #4068FF; color: white; }
        QTreeView::item:hover { background: rgba(160,190,255,0.22); }

        QHeaderView::section {
            background: #20232A; color: #D6DAE2; border: 0; border-right: 1px solid #2E3138;
        }
        QMenu { background-color: #20232A; color: #E6E9EE; border: 1px solid #2B2E34; border-radius: 8px; }
        QMenu::item { padding: 6px 12px; }
        QMenu::item:selected { background: #2D3550; }

        QPushButton#crumb {
            background: rgba(255,255,255,0.05);
            border: 1px solid #2B2E34;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #E6E9EE;
        }
        QPushButton#crumb:hover { background: rgba(255,255,255,0.09); }
        QLabel#crumbSep { color: #7F8796; }

        /* 메시지 박스: 흰 배경/검정 글씨 */
        QMessageBox { background: #FFFFFF; color: #000000; }
        QMessageBox QLabel { color: #000000; }
        QMessageBox QPushButton {
            color: #000000;
            background: #F2F4F8;
            border: 1px solid #D0D5DD;
            border-radius: 6px;
            padding: 4px 10px;
        }
        QMessageBox QPushButton:hover { background: #EAEFFF; }
    """)

def apply_light_style(app: QApplication):
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(248, 249, 251))
    pal.setColor(QPalette.Base, QColor(255, 255, 255))
    pal.setColor(QPalette.AlternateBase, QColor(246, 248, 250))
    pal.setColor(QPalette.Text, QColor(28, 28, 30))
    pal.setColor(QPalette.ButtonText, QColor(28, 28, 30))
    pal.setColor(QPalette.WindowText, QColor(28, 28, 30))
    pal.setColor(QPalette.ToolTipBase, QColor(255,255,255))
    pal.setColor(QPalette.ToolTipText, QColor(28,28,30))
    pal.setColor(QPalette.Highlight, QColor(64, 128, 255))
    pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    pal.setColor(QPalette.Button, QColor(242, 244, 248))
    app.setPalette(pal)
    app.setStyleSheet(_common_css() + """
        QMainWindow { background: #F8F9FB; }
        QLineEdit, QPushButton, QToolButton {
            background: #FFFFFF; border: 1px solid #DDE1E6; border-radius: 8px;
            color: #1C1C1E; min-height: 0px;
        }
        QLineEdit:focus { border: 1px solid #5E9BFF; }
        QTreeView {
            background: #FFFFFF; alternate-background-color: #F6F8FA;
            border: 1px solid #DDE1E6; border-radius: 10px;
        }
        QTreeView::item { color: #1A1A1A; }
        QTreeView::item:selected { background: #2A63FF; color: #FFFFFF; }
        QTreeView::item:hover { background: rgba(64,104,255,0.14); color: #1A1A1A; }

        QHeaderView::section {
            background: #F1F3F7; color: #333; border: 0; border-right: 1px solid #E5E8EE;
        }
        QMenu { background-color: #FFFFFF; color: #1C1C1E; border: 1px solid #CED3DB; border-radius: 8px; }
        QMenu::item { padding: 6px 12px; }
        QMenu::item:selected { background: #EAEFFF; }

        QPushButton#crumb {
            background: rgba(0,0,0,0.04);
            border: 1px solid #E5E8EE;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #1C1C1E;
        }
        QPushButton#crumb:hover { background: rgba(0,0,0,0.07); }
        QLabel#crumbSep { color: #7A7F89; }
    """)

# -------------------- Vector Icons --------------------
def _make_icon(w, h, painter_fn):
    pm = QPixmap(w, h); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing, True)
    try: painter_fn(p, w, h)
    finally: p.end()
    return QIcon(pm)

def icon_grid_layout(state: int, theme: str):
    def paint(p: QPainter, w, h):
        pen = QPen(QColor(180,180,190) if theme=="dark" else QColor(90,90,100), 1.6)
        p.setPen(pen); p.setBrush(Qt.NoBrush)
        cols = 2 if state == 4 else (3 if state == 6 else 4); rows = 2
        margin = 3; cellw = (w - margin*2) / cols; cellh = (h - margin*2) / rows
        for r in range(rows):
            for c in range(cols):
                x = int(margin + c * cellw + 0.5); y = int(margin + r * cellh + 0.5)
                p.drawRect(x, y, int(cellw-3), int(cellh-3))
    return _make_icon(22, 22, paint)

def icon_theme_toggle(theme: str):
    def paint(p: QPainter, w, h):
        cx, cy, r = w/2, h/2, min(w,h)/3
        if theme == "dark":
            p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(255,195,70)))
            p.drawEllipse(QtCore.QPointF(cx, cy), r, r)
            p.setPen(QPen(QColor(255,195,70), 2))
            for i in range(8):
                a = i * (math.pi/4.0)
                p.drawLine(QtCore.QPointF(cx + math.cos(a)*r*1.5, cy + math.sin(a)*r*1.5),
                           QtCore.QPointF(cx + math.cos(a)*r*2.0, cy + math.sin(a)*r*2.0))
        else:
            p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(60,60,80)))
            p.drawEllipse(QtCore.QPointF(cx, cy), r, r)
            p.setBrush(QBrush(QColor(240,240,250)))
            p.drawEllipse(QtCore.QPointF(cx + r*0.45, cy - r*0.2), r*0.9, r*0.9)
    return _make_icon(22, 22, paint)

def icon_star(checked: bool, theme: str):
    def paint(p: QPainter, w, h):
        cx, cy, r = w/2, h/2, min(w,h)/2.6; pts = []
        for i in range(10):
            angle = -math.pi/2 + i * math.pi/5; rad = r if i % 2 == 0 else r*0.45
            pts.append(QtCore.QPointF(cx + math.cos(angle)*rad, cy + math.sin(angle)*rad))
        poly = QPolygonF(pts)
        if checked:
            p.setBrush(QBrush(QColor(255, 200, 0))); p.setPen(QPen(QColor(160,120,0), 1.2))
        else:
            p.setBrush(Qt.NoBrush); p.setPen(QPen(QColor(200,200,210) if theme=="dark" else QColor(90,90,100), 1.8))
        p.drawPolygon(poly)
    return _make_icon(20, 20, paint)

def icon_edit(theme: str):
    def paint(p: QPainter, w, h):
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(100, 180, 255) if theme=="dark" else QColor(40,120,220)))
        p.drawRect(5, 13, 12, 4)
        tri = QPolygonF([QtCore.QPointF(5,12), QtCore.QPointF(5,7), QtCore.QPointF(9,11)])
        p.setBrush(QBrush(QColor(240, 200, 80))); p.drawPolygon(tri)
    return _make_icon(22, 22, paint)

def icon_info(theme: str):
    def paint(p: QPainter, w, h):
        c = QColor(160,190,255) if theme=="dark" else QColor(60,90,200)
        p.setPen(QPen(c, 2)); p.setBrush(Qt.NoBrush)
        p.drawEllipse(3,3,w-6,h-6)
        p.drawPoint(w//2, h//2-4); p.drawLine(w//2, h//2-2, w//2, h//2+6)
    return _make_icon(22, 22, paint)

def icon_cmd(theme: str):
    def paint(p: QPainter, w, h):
        border = QColor(190, 195, 210) if theme=="dark" else QColor(90, 100, 120)
        textc = QColor(230, 233, 238) if theme=="dark" else QColor(30, 30, 35)
        bg = QColor(38, 42, 50) if theme=="dark" else QColor(245, 247, 250)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(QPen(border, 1.6)); p.setBrush(QBrush(bg))
        p.drawRoundedRect(2, 4, w-4, h-6, 4, 4)
        p.setPen(QPen(textc, 2))
        p.drawLine(6, h//2, 10, h//2-3); p.drawLine(6, h//2, 10, h//2+3); p.drawLine(12, h//2+5, w-6, h//2+5)
    return _make_icon(22, 22, paint)

class GenericIconProvider(QFileIconProvider):
    def __init__(self, style): super().__init__(); self._file=style.standardIcon(QStyle.SP_FileIcon); self._dir=style.standardIcon(QStyle.SP_DirIcon)
    def icon(self, arg):
        try:
            if isinstance(arg, QFileIconProvider.IconType):
                return self._dir if arg == QFileIconProvider.Folder else self._file
            return self._dir if arg.isDir() else self._file
        except Exception:
            return self._file

# -------------------- Native Context Menu (pywin32) --------------------
class _MSG(ctypes.Structure):
    _fields_=[("hwnd",ctypes.c_void_p),("message",ctypes.c_uint),("wParam",ctypes.c_size_t),("lParam",ctypes.c_size_t),("time",ctypes.c_uint),("pt_x",ctypes.c_long),("pt_y",ctypes.c_long)]

class WinCtxMenuEventFilter(QtCore.QAbstractNativeEventFilter):
    def __init__(self): super().__init__(); self._cm2=None; self._cm3=None
    def set_context(self, cm_iface):
        self.clear()
        if not HAS_PYWIN32 or cm_iface is None: return
        try: self._cm3 = cm_iface.QueryInterface(shell.IID_IContextMenu3)
        except Exception: self._cm3 = None
        if not self._cm3:
            try: self._cm2 = cm_iface.QueryInterface(shell.IID_IContextMenu2)
            except Exception: self._cm2 = None
    def clear(self): self._cm2=None; self._cm3=None
    def nativeEventFilter(self, eventType, message):
        if eventType != 'windows_generic_MSG': return False, 0
        if not (self._cm2 or self._cm3): return False, 0
        msg = _MSG.from_address(int(message)); m = msg.message
        if HAS_PYWIN32 and m in (win32con.WM_INITMENU, win32con.WM_INITMENUPOPUP, win32con.WM_DRAWITEM, win32con.WM_MEASUREITEM, win32con.WM_MENUCHAR):
            try:
                if self._cm3 and m == win32con.WM_MENUCHAR:
                    handled, result = self._cm3.HandleMenuMsg2(int(m), int(msg.wParam), int(msg.lParam))
                    return bool(handled), int(result or 0)
                cm = self._cm3 or self._cm2
                if cm: cm.HandleMenuMsg(int(m), int(msg.wParam), int(msg.lParam))
            except Exception: pass
        return False, 0

def _ensure_event_filter(app: QApplication) -> WinCtxMenuEventFilter:
    f = app.property('_win_ctx_filter')
    if f is None:
        f = WinCtxMenuEventFilter(); app.installNativeEventFilter(f); app.setProperty('_win_ctx_filter', f)
    return f

def _as_interface(obj):
    if obj is None: return None
    if isinstance(obj,(list,tuple)): obj = obj[0] if obj else None
    if isinstance(obj,int): return None
    return obj

def _abs_pidl(path_str):
    pidl, _attrs = shell.SHParseDisplayName(_normalize_fs_path(path_str), 0)
    return pidl

def _bind_folder(path_str):
    desktop = shell.SHGetDesktopFolder()
    pidl = _abs_pidl(path_str)
    return desktop.BindToObject(pidl, None, shell.IID_IShellFolder)

def _icm_via_shellitems(paths):
    try:
        if len(paths) == 1:
            it = shell.SHCreateItemFromParsingName(_normalize_fs_path(paths[0]), None, shell.IID_IShellItem)
            return _as_interface(it.BindToHandler(None, shell.BHID_SFUIObject, shell.IID_IContextMenu))
        pidls = tuple(_abs_pidl(p) for p in paths)
        try: sia = shell.SHCreateShellItemArrayFromIDLists(pidls)
        except Exception: return None
        return _as_interface(sia.BindToHandler(None, shell.BHID_SFUIObject, shell.IID_IContextMenu))
    except Exception as e:
        if DEBUG: print("[ctx] ShellItems route failed:", e); return None

def _post_null(hwnd):
    try: win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
    except Exception: pass

def _get_canonical_verb(cm, idx):
    verb=None
    try:
        GCS_VERBW=getattr(shellcon,"GCS_VERBW",4); v=cm.GetCommandString(idx, GCS_VERBW)
        if v:
            if isinstance(v,(bytes,bytearray)):
                try: verb=v.decode("utf-16le",errors="ignore").strip("\x00")
                except Exception: verb=v.decode(errors="ignore")
            else: verb=str(v)
    except Exception: pass
    if not verb:
        try:
            GCS_VERBA=getattr(shellcon,"GCS_VERBA",2); v=cm.GetCommandString(idx, GCS_VERBA)
            if v: verb=v.decode(errors="ignore") if isinstance(v,(bytes,bytearray)) else str(v)
        except Exception: pass
    return (verb or "").strip().lower()

def _invoke_menu(owner_hwnd, cm, hmenu, screen_pt, work_dir, paths=None, id_first=1):
    """
    True  = 명령 실행/취소(폴백 불필요), False = 실패/팝업헤더 선택(폴백 메뉴 제공)
    """
    shown=False
    try:
        win32gui.SetForegroundWindow(owner_hwnd)
    except Exception: pass
    try:
        cmd_id = win32gui.TrackPopupMenu(hmenu, win32con.TPM_LEFTALIGN|win32con.TPM_RETURNCMD|win32con.TPM_RIGHTBUTTON,
                                         int(screen_pt[0]), int(screen_pt[1]), 0, int(owner_hwnd), None)
        shown=True
    except Exception as e:
        if DEBUG: print("[ctx] TrackPopupMenu failed:", e)
        return False
    if not cmd_id:  # user cancelled
        return True

    # 서브메뉴 헤더 클릭 → 폴백 메뉴로 처리
    try:
        state = win32gui.GetMenuState(hmenu, cmd_id, win32con.MF_BYCOMMAND)
        if state & win32con.MF_POPUP:
            if DEBUG: print("[ctx] popup header selected; will fallback")
            return False
    except Exception:
        pass

    idx=int(cmd_id)-int(id_first); verb=None
    try:
        verb = _get_canonical_verb(cm, idx) or None
    except Exception: verb=None
    if DEBUG: print(f"[ctx] chosen cmd_id={cmd_id} id_first={id_first} -> idx={idx}, verb='{verb or ''}'")

    if verb and verb.lower() in ("properties","prop","property"):
        try:
            target = paths[0] if (paths and len(paths)>0) else work_dir
            shell.SHObjectProperties(owner_hwnd, 0, _normalize_fs_path(target), None)
            _post_null(owner_hwnd); return True
        except Exception as e:
            if DEBUG: print("[ctx] SHObjectProperties failed:", e)
            return False

    try:
        pici_int=(0,int(owner_hwnd),int(idx),None,None,win32con.SW_SHOWNORMAL,0,0)
        cm.InvokeCommand(pici_int); _post_null(owner_hwnd); return True
    except Exception as e:
        if DEBUG: print("[ctx] InvokeCommand(int) failed:", e)

    if verb:
        try:
            pici_str=(0,int(owner_hwnd),str(verb),None,None,win32con.SW_SHOWNORMAL,0,0)
            cm.InvokeCommand(pici_str); _post_null(owner_hwnd); return True
        except Exception as e:
            if DEBUG: print(f"[ctx] InvokeCommand verb='{verb}' failed:", e)

    return False

def show_explorer_context_menu(owner_hwnd, paths, screen_pt):
    if not HAS_PYWIN32 or not paths: return False
    pythoncom.CoInitialize()
    try:
        parents={str(Path(p).parent) for p in paths}
        if len(parents)!=1: paths=[paths[0]]; parents={str(Path(paths[0]).parent)}
        parent_dir=_normalize_fs_path(next(iter(parents)))
        app=QApplication.instance(); evf=_ensure_event_filter(app)

        cm=_icm_via_shellitems(paths)
        if cm:
            evf.set_context(cm); hMenu=win32gui.CreatePopupMenu()
            flags=shellcon.CMF_NORMAL|shellcon.CMF_EXPLORE|shellcon.CMF_INCLUDESTATIC
            if win32api.GetKeyState(win32con.VK_SHIFT)<0: flags|=shellcon.CMF_EXTENDEDVERBS
            id_first=1
            try:
                cm.QueryContextMenu(hMenu,0,id_first,0x7FFF,flags)
                ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,parent_dir,paths=paths,id_first=id_first)
                evf.clear()
                if ok: return True
            except Exception as e:
                if DEBUG: print("[ctx] ShellItems QueryContextMenu failed:", e)
                evf.clear()

        try:
            desktop=shell.SHGetDesktopFolder()
            abs_pidls=tuple(_abs_pidl(p) for p in paths)
            cm=desktop.GetUIObjectOf(0,abs_pidls,shell.IID_IContextMenu,0); cm=_as_interface(cm)
        except Exception as e:
            if DEBUG: print("[ctx] desktop GetUIObjectOf failed:", e); cm=None
        if not cm: return False

        evf.set_context(cm); hMenu=win32gui.CreatePopupMenu()
        flags=shellcon.CMF_NORMAL|shellcon.CMF_EXPLORE|shellcon.CMF_INCLUDESTATIC
        if win32api.GetKeyState(win32con.VK_SHIFT)<0: flags|=shellcon.CMF_EXTENDEDVERBS
        id_first=1; cm.QueryContextMenu(hMenu,0,id_first,0x7FFF,flags)
        ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,parent_dir,paths=paths,id_first=id_first)
        evf.clear(); return ok
    finally:
        pythoncom.CoUninitialize()

def show_explorer_background_menu(owner_hwnd, folder_path, screen_pt):
    if not HAS_PYWIN32: return False
    pythoncom.CoInitialize()
    try:
        sf=_bind_folder(folder_path)
        try: cm=sf.CreateViewObject(0, shell.IID_IContextMenu); cm=_as_interface(cm)
        except Exception: cm=None
        if not cm: return False
        app=QApplication.instance(); evf=_ensure_event_filter(app); evf.set_context(cm)
        hMenu=win32gui.CreatePopupMenu()
        flags=shellcon.CMF_NORMAL|shellcon.CMF_EXPLORE|shellcon.CMF_INCLUDESTATIC
        if win32api.GetKeyState(win32con.VK_SHIFT)<0: flags|=shellcon.CMF_EXTENDEDVERBS
        id_first=1; cm.QueryContextMenu(hMenu,0,id_first,0x7FFF,flags)
        ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,folder_path,paths=[folder_path],id_first=id_first)
        evf.clear(); return ok
    finally:
        pythoncom.CoUninitialize()

# --- ShellNew 기반 새 파일 생성 유틸리티 ------------------------------------
try:
    import winreg
    HAS_WINREG = True
except Exception:
    HAS_WINREG = False

def _shellnew_template_for_ext(ext_with_dot: str) -> tuple[bool, str | None]:
    """
    반환: (is_nullfile, template_path_or_None)
    """
    if not HAS_WINREG:
        return (False, None)
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ext_with_dot) as k:
            progid, _ = winreg.QueryValueEx(k, None)
        if not progid:
            return (False, None)
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\ShellNew") as ks:
            try:
                fname, _ = winreg.QueryValueEx(ks, "FileName")
                if fname:
                    candidates = [fname]
                    if not os.path.isabs(fname):
                        windir = os.environ.get("WINDIR", r"C:\Windows")
                        candidates.insert(0, os.path.join(windir, "ShellNew", fname))
                    for p in candidates:
                        if os.path.exists(p):
                            return (False, p)
            except Exception:
                pass
            try:
                _null, _ = winreg.QueryValueEx(ks, "NullFile")
                return (True, None)
            except Exception:
                pass
    except Exception:
        pass
    return (False, None)

def _create_new_file_with_template(dst_dir: str, filename: str, ext_with_dot: str) -> str:
    os.makedirs(dst_dir, exist_ok=True)
    target = unique_dest_path(dst_dir, filename)
    is_null, templ = _shellnew_template_for_ext(ext_with_dot)
    try:
        if templ:
            shutil.copyfile(templ, target)
        else:
            with open(target, "wb"):
                pass
    except Exception:
        with open(target, "wb"):
            pass
    return target

# -------------------- Bookmarks Store --------------------
def load_named_bookmarks() -> list:
    s = QSettings(ORG_NAME, APP_NAME)
    val = s.value("bookmarks/named_items", [])
    if isinstance(val, list):
        out=[]
        for it in val:
            try:
                out.append({"enabled": bool(it.get("enabled", False)),
                            "name": str(it.get("name","")),
                            "path": str(it.get("path",""))})
            except Exception: pass
        return out[:10]
    return []

def save_named_bookmarks(items: list):
    s = QSettings(ORG_NAME, APP_NAME)
    s.setValue("bookmarks/named_items", items[:10]); s.sync()

def _derive_name_from_path(p: str) -> str:
    try:
        p = nice_path(p)
        if p.endswith(os.sep) and len(p) <= 3: return p
        base = os.path.basename(p.rstrip("\\/")); return base or p
    except Exception: return p

def migrate_legacy_favorites_into_named(items: list) -> list:
    try:
        s = QSettings(ORG_NAME, APP_NAME)
        favs = s.value("favorites/paths", [])
        if not favs: return items[:10]
        existing = {os.path.normcase(x.get("path","")) for x in items}
        for p in favs:
            np = nice_path(str(p))
            if os.path.normcase(np) in existing: continue
            if len(items) >= 10: break
            items.append({"enabled": True, "name": _derive_name_from_path(np), "path": np})
        s.remove("favorites/paths"); return items[:10]
    except Exception:
        return items[:10]

# -------------------- Roles & Proxies --------------------
IS_DIR_ROLE = Qt.UserRole + 99
SIZE_BYTES_ROLE = Qt.UserRole + 100

class FsSortProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDynamicSortFilter(True)
        self.setSortCaseSensitivity(Qt.CaseInsensitive)
        self.setSortRole(Qt.EditRole)
        self.setSortLocaleAware(True)
    def filterAcceptsRow(self, source_row, source_parent): return True
    def sort(self, column, order=Qt.AscendingOrder):
        # 현재 정렬 방향을 기억해 두었다가 lessThan에서 폴더 우선 규칙을 유지합니다.
        self._sort_order = order
        super().sort(column, order)
    def lessThan(self, left, right):
        col = left.column(); src = self.sourceModel()

        # 폴더가 항상 파일보다 먼저 오도록(정렬 방향과 무관)
        try:
            ldir = bool(src.isDir(left)) if hasattr(src, "isDir") else bool(src.data(left, IS_DIR_ROLE))
            rdir = bool(src.isDir(right)) if hasattr(src, "isDir") else bool(src.data(right, IS_DIR_ROLE))
            if ldir != rdir:
                order = getattr(self, "_sort_order", Qt.AscendingOrder)
                if order == Qt.AscendingOrder:
                    # 오름차순: 폴더 < 파일
                    return ldir and not rdir
                else:
                    # 내림차순: 오름차순 기준을 뒤집어도 폴더가 먼저 보이도록 파일 < 폴더로 판단
                    return (not ldir) and rdir
        except Exception:
            pass

        if col == 1:
            try:
                lv = int(src.data(left, SIZE_BYTES_ROLE) or src.data(left, Qt.EditRole) or 0)
                rv = int(src.data(right, SIZE_BYTES_ROLE) or src.data(right, Qt.EditRole) or 0)
                return lv < rv
            except Exception:
                pass

        if col in (2, 3):
            lv = src.data(left, Qt.EditRole); rv = src.data(right, Qt.EditRole)
            if isinstance(lv, QDateTime) and isinstance(rv, QDateTime):
                return lv < rv

        try:
            lv = src.data(left, Qt.EditRole); rv = src.data(right, Qt.EditRole)
            return str(lv).lower() < str(rv).lower()
        except Exception:
            return super().lessThan(left, right)


class FastDirModel(QAbstractTableModel):
    HEADERS = ["Name", "Size", "Type", "Date Modified"]
    def __init__(self, parent=None):
        super().__init__(parent); self._root=""; self._rows=[]
        # 아이콘 캐시
        self._icon_cache = {}
        self._icon_file = None
        self._icon_dir  = None
    def rootPath(self): return self._root
    def reset_dir(self, path:str):
        self.beginResetModel(); self._root=path; self._rows=[]; self._icon_cache.clear(); self.endResetModel()
    @QtCore.pyqtSlot(list)
    def append_rows(self, rows:list):
        if not rows: return
        start=len(self._rows); self.beginInsertRows(QtCore.QModelIndex(), start, start+len(rows)-1)
        self._rows.extend(rows); self.endInsertRows()
    def row_path(self, row:int)->str: return self._rows[row]["path"] if 0<=row<len(self._rows) else ""
    def has_stat(self, row:int)->bool:
        if 0<=row<len(self._rows):
            return (self._rows[row]["size"] is not None) and (self._rows[row]["mtime"] is not None)
        return False
    def has_icon(self, row:int)->bool:
        return row in self._icon_cache
    @QtCore.pyqtSlot(int, object, object)
    def apply_stat(self, row:int, size_val, mtime_val):
        if not (0<=row<len(self._rows)): return
        changed=[]
        if self._rows[row]["size"] is None and size_val is not None:
            self._rows[row]["size"]=int(size_val); changed.append(1)
        if self._rows[row]["mtime"] is None and mtime_val is not None:
            self._rows[row]["mtime"]=float(mtime_val); changed.append(3)
        if changed:
            for col in changed:
                ix=self.index(row,col); self.dataChanged.emit(ix,ix,[Qt.DisplayRole,Qt.EditRole,SIZE_BYTES_ROLE])
    def apply_icon(self, row:int, icon:QIcon):
        if 0 <= row < len(self._rows):
            self._icon_cache[row] = icon
            ix = self.index(row, 0)
            self.dataChanged.emit(ix, ix, [Qt.DecorationRole])
    def rowCount(self, parent=QtCore.QModelIndex()): return 0 if parent.isValid() else len(self._rows)
    def columnCount(self, parent=QtCore.QModelIndex()): return 4
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        return self.HEADERS[section] if role==Qt.DisplayRole and orientation==Qt.Horizontal else None
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
        r=self._rows[index.row()]; c=index.column()

        # 아이콘
        if role == Qt.DecorationRole and c == 0:
            ic = self._icon_cache.get(index.row())
            if ic is not None:
                return ic
            try:
                if self._icon_file is None or self._icon_dir is None:
                    st = QApplication.instance().style()
                    self._icon_file = st.standardIcon(QStyle.SP_FileIcon) if st else QIcon()
                    self._icon_dir  = st.standardIcon(QStyle.SP_DirIcon)  if st else QIcon()
            except Exception:
                return None
            return self._icon_dir if r["is_dir"] else self._icon_file

        if role==Qt.DisplayRole:
            if c==0: return r["name"]
            if c==1: return "" if r["size"] is None else human_size(int(r["size"]))
            if c==2: return r["type"]
            if c==3:
                if r["mtime"] is None: return ""
                dt=QDateTime.fromSecsSinceEpoch(int(r["mtime"])); return dt.toString("yyyy-MM-dd HH:mm:ss")
        elif role==Qt.EditRole:
            if c==0: return r["name"]
            if c==1: return 0 if r["size"] is None else int(r["size"])
            if c==3: return QDateTime.fromSecsSinceEpoch(int(r["mtime"])) if r["mtime"] else QDateTime()
            else: return r["type"]
        elif role==Qt.ToolTipRole: return r["path"]
        elif role==Qt.UserRole: return r["path"]
        elif role==IS_DIR_ROLE: return r["is_dir"]
        elif role==SIZE_BYTES_ROLE: return 0 if r["size"] is None else int(r["size"])
        return None

class FastStatWorker(QtCore.QThread):
    statReady=pyqtSignal(int, object, object); finishedCycle=pyqtSignal()
    def __init__(self, model:FastDirModel, root:str, rows:list[int], parent=None):
        super().__init__(parent); self._model=model; self._root=root; self._rows=list(rows); self._cancel=False
    def cancel(self): self._cancel=True
    def run(self):
        try:
            for row in self._rows:
                if self._cancel: break
                if self._model.rootPath()!=self._root: break
                if self._model.has_stat(row): continue
                p=self._model.row_path(row)
                try:
                    st=os.stat(p, follow_symlinks=False)
                    size_val=0 if os.path.isdir(p) else int(st.st_size)
                    mtime_val=float(st.st_mtime)
                except Exception:
                    size_val=0; mtime_val=None
                self.statReady.emit(row,size_val,mtime_val)
        finally:
            self.finishedCycle.emit()

class DirEnumWorker(QtCore.QThread):
    batchReady=QtCore.pyqtSignal(list); finished=QtCore.pyqtSignal(); error=QtCore.pyqtSignal(str)
    def __init__(self, root:str): super().__init__(); self.root=root; self._cancel=False
    def cancel(self): self._cancel=True
    def run(self):
        batch, BATCH=[], 400
        try:
            with os.scandir(self.root) as it:
                for entry in it:
                    if self._cancel: break
                    name=entry.name; p=os.path.join(self.root,name)
                    try: is_dir=entry.is_dir(follow_symlinks=False)
                    except Exception: is_dir=os.path.isdir(p)
                    ext=os.path.splitext(name)[1]
                    typ="Folder" if is_dir else (ext.lstrip(".").upper()+" file" if ext else "File")
                    batch.append({"name":name,"path":p,"is_dir":is_dir,"size":None,"mtime":None,"type":typ})
                    if len(batch)>=BATCH: self.batchReady.emit(batch); batch=[]
                if batch: self.batchReady.emit(batch)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()

class NormalStatWorker(QtCore.QThread):
    statReady=pyqtSignal(str, object, object); finishedCycle=pyqtSignal()
    def __init__(self, paths:list[str], parent=None):
        super().__init__(parent); self._paths=list(paths); self._cancel=False
    def cancel(self): self._cancel=True
    def run(self):
        try:
            for p in self._paths:
                if self._cancel: break
                try:
                    st=os.stat(p, follow_symlinks=False)
                    size_val=0 if os.path.isdir(p) else int(st.st_size)
                    mtime_val=float(st.st_mtime)
                except Exception:
                    size_val=0; mtime_val=None
                self.statReady.emit(p,size_val,mtime_val)
        finally:
            self.finishedCycle.emit()

class SearchWorker(QtCore.QThread):
    batchReady = pyqtSignal(str, list)   # (base_path, [ {name,path,is_dir,folder} ... ])
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, base_path: str, pattern_str: str, parent=None):
        super().__init__(parent)
        self.base = base_path
        self._cancel = False
        # 패턴 분해: 콤마/세미콜론/공백 구분 → 소문자
        raw = (pattern_str or "").replace(",", " ").replace(";", " ").split()
        self._patterns = [p.lower() for p in raw] if raw else ["*"]
        # 빠른 확장자 체크 함수들 구성 (*.txt 처럼 단순 확장자 패턴은 endswith로 최적화)
        self._tests = []
        for p in self._patterns:
            simple_ext = (p.startswith("*.") and ("*" not in p[2:]) and ("?" not in p) and ("[" not in p) and ("]" not in p))
            if simple_ext:
                ext = p[1:]  # ".txt"
                self._tests.append(lambda name, ext=ext: name.endswith(ext))
            else:
                # fnmatch의 대소문자 민감도 회피를 위해 lower-case 고정
                self._tests.append(lambda name, pat=p: fnmatch.fnmatchcase(name, pat))

    def cancel(self): self._cancel = True

    def _match(self, name_lower: str) -> bool:
        for t in self._tests:
            if t(name_lower):
                return True
        return False

    def run(self):
        try:
            base = self.base
            stack = [base]
            BATCH = 600
            batch = []
            while stack and not self._cancel:
                d = stack.pop()
                try:
                    with os.scandir(d) as it:
                        for entry in it:
                            if self._cancel:
                                break
                            try:
                                is_dir = entry.is_dir(follow_symlinks=False)
                            except Exception:
                                is_dir = os.path.isdir(entry.path)

                            name_l = entry.name.lower()
                            if self._match(name_l):
                                rel = os.path.relpath(d, base)
                                if rel == ".":
                                    rel = ""
                                batch.append({
                                    "name": entry.name,
                                    "path": entry.path,
                                    "is_dir": is_dir,
                                    "folder": rel
                                })
                                if len(batch) >= BATCH:
                                    self.batchReady.emit(base, batch)
                                    batch = []

                            # 하위 디렉토리 탐색 (심볼릭 링크/재분석 지점은 피함)
                            if is_dir:
                                try:
                                    if entry.is_symlink():
                                        continue
                                except Exception:
                                    pass
                                stack.append(entry.path)
                except Exception:
                    # 접근 거부 등은 무시하고 계속
                    continue

            if batch:
                self.batchReady.emit(base, batch)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()

class StatOverlayProxy(QIdentityProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent); self._cache={}; self._pending=set(); self._worker=None
    def filePath(self, index):
        src=self.sourceModel(); s=self.mapToSource(index); return src.filePath(s)
    def isDir(self, index):
        src=self.sourceModel(); s=self.mapToSource(index)
        try: return src.isDir(s)
        except Exception:
            try: return os.path.isdir(src.filePath(s))
            except Exception: return False
    def clear_cache(self):
        self._cache.clear(); self._pending.clear(); self._cancel_worker()
    def _cancel_worker(self):
        w=self._worker
        if w and w.isRunning(): w.cancel(); w.wait(100)
        self._worker=None
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        col = index.column()
        # Size(1), Date Modified(3)만 오버레이 처리
        if col not in (1, 3):
            return super().data(index, role)

        src = self.sourceModel()
        sidx = self.mapToSource(index)

        # 파일 경로 / 폴더 여부
        try:
            p = src.filePath(sidx)
        except Exception:
            p = None

        try:
            is_dir = src.isDir(sidx)
        except Exception:
            try:
                is_dir = os.path.isdir(p) if p else False
            except Exception:
                is_dir = False

        rec = self._cache.get(p) if p else None  # (size, mtime) or None

        # ---------------- Size Column ----------------
        if col == 1:
            # 폴더는 항상 0/빈 문자열
            if is_dir:
                if role == SIZE_BYTES_ROLE:
                    return 0
                if role == Qt.EditRole:
                    return 0
                if role == Qt.DisplayRole:
                    return ""
                return super().data(index, role)

            # 캐시에 있으면 즉시 사용
            if rec is not None:
                size_val = int(rec[0] or 0)
                if role == SIZE_BYTES_ROLE:
                    return size_val
                if role == Qt.EditRole:
                    return size_val
                if role == Qt.DisplayRole:
                    return human_size(size_val)

            # 캐시가 없으면 즉시 stat으로 계산해서 표시 (초기 진입 직후 F5 없이도 보이게)
            if p:
                try:
                    st = os.stat(p, follow_symlinks=False)
                    size_val = int(st.st_size)
                    # 캐시에 적재하여 이후 정렬/표시에 활용
                    old_mtime = rec[1] if (rec is not None and len(rec) > 1) else None
                    self._cache[p] = (size_val, old_mtime)

                    if role == SIZE_BYTES_ROLE:
                        return size_val
                    if role == Qt.EditRole:
                        return size_val
                    if role == Qt.DisplayRole:
                        return human_size(size_val)
                except Exception:
                    pass

            # 마지막으로 원본 모델 값 시도
            try:
                if role == SIZE_BYTES_ROLE:
                    v = src.data(sidx, Qt.EditRole)
                    return int(v) if isinstance(v, int) else 0
                if role == Qt.EditRole:
                    v = src.data(sidx, Qt.EditRole)
                    return int(v) if isinstance(v, int) else 0
                if role == Qt.DisplayRole:
                    v = src.data(sidx, Qt.EditRole)
                    return human_size(int(v)) if isinstance(v, int) else ""
            except Exception:
                return 0 if role in (Qt.EditRole, SIZE_BYTES_ROLE) else ""

            return super().data(index, role)

        # ---------------- Date Modified Column ----------------
        if col == 3:
            # 캐시에 mtime 있으면 사용
            if rec and rec[1] is not None:
                if role == Qt.DisplayRole:
                    dt = QDateTime.fromSecsSinceEpoch(int(rec[1]))
                    return dt.toString("yyyy-MM-dd HH:mm:ss")
                if role == Qt.EditRole:
                    return QDateTime.fromSecsSinceEpoch(int(rec[1]))

            # 캐시가 없으면 파일 시스템에서 즉시 조회 (표시 용도)
            if p:
                try:
                    mtime = os.path.getmtime(p)
                    # 캐시에 병합 저장 (size는 유지)
                    old_size = rec[0] if rec else None
                    self._cache[p] = (int(old_size or 0), float(mtime))

                    if role == Qt.DisplayRole:
                        dt = QDateTime.fromSecsSinceEpoch(int(mtime))
                        return dt.toString("yyyy-MM-dd HH:mm:ss")
                    if role == Qt.EditRole:
                        return QDateTime.fromSecsSinceEpoch(int(mtime))
                except Exception:
                    pass

            # 원본 모델 값으로 폴백
            try:
                if role == Qt.DisplayRole:
                    v = src.data(sidx, Qt.DisplayRole)
                    return v
                if role == Qt.EditRole:
                    v = src.data(sidx, Qt.EditRole)
                    return v if isinstance(v, QDateTime) else QDateTime()
            except Exception:
                if role == Qt.DisplayRole:
                    return ""
                if role == Qt.EditRole:
                    return QDateTime()

            return super().data(index, role)

    def request_paths(self, paths:list[str], batch_limit:int=256):
        todo=[]
        for p in paths:
            if not p: continue
            if p in self._cache or p in self._pending: continue
            self._pending.add(p); todo.append(p)
            if len(todo)>=batch_limit: break
        if not todo: return
        self._cancel_worker()
        w=NormalStatWorker(todo, self)
        w.statReady.connect(self._apply_stat, Qt.QueuedConnection)
        w.finishedCycle.connect(lambda: self._on_cycle_finished(todo))
        self._worker=w; w.start()
    @QtCore.pyqtSlot(str, object, object)
    def _apply_stat(self, path:str, size_val, mtime_val):
        self._cache[path]=(int(size_val or 0), float(mtime_val) if mtime_val is not None else None)
        try:
            src=self.sourceModel(); sidx0=src.index(path)
            if sidx0.isValid():
                for col in (1,3):
                    sidx=sidx0.sibling(sidx0.row(), col)
                    pidx=self.mapFromSource(sidx)
                    self.dataChanged.emit(pidx,pidx,[Qt.DisplayRole,Qt.EditRole,SIZE_BYTES_ROLE])
        except Exception: pass
    def _on_cycle_finished(self, batch):
        for p in batch: self._pending.discard(p)

# -------------------- Breadcrumb --------------------
class PathBar(QWidget):
    pathSubmitted=pyqtSignal(str)
    def __init__(self, parent=None):
        super().__init__(parent); self._current_path=QDir.homePath()
        self._host=QWidget(self); self._hlay=QHBoxLayout(self._host)
        self._hlay.setContentsMargins(4,0,4,0); self._hlay.setSpacing(max(0, ROW_SPACING-2))
        self._scroll=QScrollArea(self); self._scroll.setWidget(self._host)
        self._scroll.setWidgetResizable(True); self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff); self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setViewportMargins(0,0,0,0); self._scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._scroll.setFixedHeight(UI_H); self.setFixedHeight(UI_H)
        self._edit=QLineEdit(self); self._edit.hide(); self._edit.setClearButtonEnabled(True); self._edit.setFixedHeight(UI_H)
        self._edit.returnPressed.connect(self._on_edit_return)
        wrap=QHBoxLayout(self); wrap.setContentsMargins(0,0,0,0); wrap.setSpacing(0); wrap.addWidget(self._scroll,1); wrap.addWidget(self._edit,1)
        self._host.installEventFilter(self); self._edit.installEventFilter(self)
        self.set_path(self._current_path)
    def sizeHint(self): return QSize(200, UI_H)
    def minimumSizeHint(self): return QSize(100, UI_H)
    def eventFilter(self, obj, ev):
        if obj is self._host and ev.type()==QEvent.MouseButtonDblClick: self.start_edit(); return True
        if obj is self._edit and ev.type()==QEvent.FocusOut: self.cancel_edit()
        return super().eventFilter(obj, ev)
    def start_edit(self):
        self._edit.setText(self._current_path); self._scroll.hide(); self._edit.show()
        self._edit.setFocus(); self._edit.selectAll()
    def cancel_edit(self): self._edit.hide(); self._scroll.show()
    def _on_edit_return(self):
        t=self._edit.text().strip(); self.cancel_edit()
        if t: self.pathSubmitted.emit(t)
    def set_path(self, path:str):
        self._current_path = nice_path(path); self._rebuild()
    def _rebuild(self):
        while self._hlay.count():
            it=self._hlay.takeAt(0); w=it.widget()
            if w: w.deleteLater()
        p=self._current_path; parts=[]
        if p.startswith("\\\\"):
            comps=[c for c in p.split(os.sep) if c]
            if len(comps)>=2:
                root=f"\\\\{comps[0]}\\{comps[1]}\\"
                parts.append((root,root)); acc=root
                for c in comps[2:]:
                    acc=os.path.join(acc,c); parts.append((c,acc))
            else: parts.append((p,p))
        else:
            drive,_=os.path.splitdrive(p); root=(drive+os.sep) if drive else os.sep
            parts.append((root,root)); sub=p[len(root):].strip("\\/")
            for seg in [s for s in sub.split(os.sep) if s]:
                curr=os.path.join(parts[-1][1], seg); parts.append((seg,curr))
        fm=self.fontMetrics()
        for i,(label,target) in enumerate(parts):
            btn=QPushButton(self); btn.setObjectName("crumb"); btn.setFlat(True); btn.setCursor(Qt.PointingHandCursor)
            elided=fm.elidedText(label, Qt.ElideMiddle, CRUMB_MAX_SEG_W)
            btn.setText(elided); btn.setToolTip(label); btn.setMinimumHeight(UI_H)
            btn.clicked.connect(lambda _,t=target: self.pathSubmitted.emit(t))
            self._hlay.addWidget(btn)
            if i < len(parts)-1:
                s=QLabel("›", self); s.setObjectName("crumbSep"); s.setContentsMargins(0,0,0,0); self._hlay.addWidget(s)
        self._hlay.addStretch(1)
        QTimer.singleShot(0, lambda: self._scroll.horizontalScrollBar().setValue(self._scroll.horizontalScrollBar().maximum()))

# -------------------- SearchResultModel --------------------
class SearchResultModel(QStandardItemModel):
    def data(self, index, role=Qt.DisplayRole):
        if role==Qt.DisplayRole and index.column()==1:
            b=super().data(index, SIZE_BYTES_ROLE)
            if b is None: b=super().data(index, Qt.EditRole)
            if isinstance(b,(int,float)) and b: return human_size(int(b))
            return ""
        return super().data(index, role)

# -------------------- Conflict Resolution Dialog --------------------
class ConflictResolutionDialog(QDialog):
    """
    Per-file conflict chooser:
      - Table with columns: Name, Destination, Action (Overwrite/Skip/Copy)
      - Top buttons: Overwrite All / Skip All / Copy All
      - OK / Cancel
    """
    def __init__(self, parent, conflicts:list[tuple[str,str]], dst_dir:str):
        super().__init__(parent)
        self.setWindowTitle("Resolve name conflicts")
        self.resize(720, 420)
        self._conflicts = conflicts
        self._dst_dir = dst_dir

        lay = QVBoxLayout(self)

        # Top "apply to all" row
        top = QHBoxLayout()
        lbl = QLabel("Apply to all:", self)
        btn_over = QPushButton("Overwrite All", self)
        btn_skip = QPushButton("Skip All", self)
        btn_copy = QPushButton("Copy All", self)
        top.addWidget(lbl)
        top.addSpacing(8)
        top.addWidget(btn_over)
        top.addWidget(btn_skip)
        top.addWidget(btn_copy)
        top.addItem(QSpacerItem(10,10, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lay.addLayout(top)

        # Table
        self.tbl = QTableWidget(self)
        self.tbl.setColumnCount(3)
        self.tbl.setHorizontalHeaderLabels(["Name", "Destination", "Action"])
        self.tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl.setRowCount(len(conflicts))
        self._combos = []

        for r, (src, dst) in enumerate(conflicts):
            name = os.path.basename(src)
            self.tbl.setItem(r, 0, QTableWidgetItem(name))
            self.tbl.setItem(r, 1, QTableWidgetItem(dst))
            combo = QComboBox(self.tbl)
            combo.addItems(["Overwrite", "Skip", "Copy"])
            self.tbl.setCellWidget(r, 2, combo)
            self._combos.append(combo)

        lay.addWidget(self.tbl, 1)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        lay.addWidget(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        # Apply-to-all handlers
        btn_over.clicked.connect(lambda: self._apply_all("Overwrite"))
        btn_skip.clicked.connect(lambda: self._apply_all("Skip"))
        btn_copy.clicked.connect(lambda: self._apply_all("Copy"))

        # 다크 모드에서 가독성 개선
        theme = getattr(getattr(parent, "host", None), "theme", "dark")
        if theme == "dark":
            pal = self.palette()
            pal.setColor(QPalette.Window, QColor(255, 255, 255))
            pal.setColor(QPalette.Base, QColor(255, 255, 255))
            pal.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
            pal.setColor(QPalette.Text, QColor(0, 0, 0))
            pal.setColor(QPalette.ButtonText, QColor(0, 0, 0))
            pal.setColor(QPalette.WindowText, QColor(0, 0, 0))
            self.setPalette(pal)
            self.setStyleSheet("""
                QDialog, QLabel, QTableWidget, QLineEdit { color: #000000; background: #FFFFFF; }
                QHeaderView::section { color: #000000; background: #F1F3F7; border: 0; border-right: 1px solid #E5E8EE; }
                QComboBox { color: #000000; background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 2px 6px; }
                QComboBox:hover { border: 1px solid #5E9BFF; }
                QComboBox QAbstractItemView { color: #000000; background: #FFFFFF; }
                QTableWidget QTableCornerButton::section { background: #FFFFFF; }
            """)

    def _apply_all(self, which:str):
        for c in self._combos:
            idx = c.findText(which)
            if idx >= 0: c.setCurrentIndex(idx)

    def result_map(self) -> dict:
        """
        Returns: dict[src_path] = "overwrite"|"skip"|"copy"
        """
        out = {}
        for (src, _dst), combo in zip(self._conflicts, self._combos):
            choice = combo.currentText().strip().lower()
            if choice not in ("overwrite","skip","copy"): choice="overwrite"
            out[src] = choice
        return out

# -------------------- Custom View (robust D&D) --------------------
class ExplorerView(QTreeView):
    def __init__(self, pane):
        super().__init__(pane); self.pane=pane
        self.setDragEnabled(True); self.setAcceptDrops(True)
        self.setDropIndicatorShown(True); self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        # ▶ 단축키가 항상 이 뷰까지 도달하도록 포커스 정책 강화
        self.setFocusPolicy(Qt.StrongFocus)
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
        else: super().dragEnterEvent(e)
    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            if e.keyboardModifiers() & Qt.ControlModifier: e.setDropAction(Qt.CopyAction)
            else: e.setDropAction(Qt.MoveAction)
            e.accept()
        else: super().dragMoveEvent(e)
    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            urls=e.mimeData().urls(); srcs=[u.toLocalFile() for u in urls if u.isLocalFile()]
            if srcs:
                op="copy" if (e.dropAction()==Qt.CopyAction or (e.keyboardModifiers() & Qt.ControlModifier)) else "move"
                self.pane._start_bg_op(op, srcs, self.pane.current_path()); e.acceptProposedAction(); return
        super().dropEvent(e)

    # F5 = 하드 리프레시
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F5:
            try:
                self.pane.hard_refresh()
            finally:
                e.accept()
            return
        super().keyPressEvent(e)

    # 단일 선택 항목을 재클릭해도 선택 해제되지 않게
    def mousePressEvent(self, e):
        # ▶ 어떤 버튼이든 클릭 시 먼저 포커스를 이 뷰로 강제 이동
        try:
            if not self.hasFocus():
                self.setFocus(Qt.MouseFocusReason)
        except Exception:
            pass

        # 단일 선택 항목을 재클릭해도 선택 해제되지 않게 (기존 동작 유지)
        if e.button() == Qt.LeftButton and e.modifiers() == Qt.NoModifier:
            clicked = self.indexAt(e.pos())
            sm = self.selectionModel()
            prev_rows = sm.selectedRows(0)
            same_single = (clicked.isValid() and len(prev_rows)==1 and
                           prev_rows[0].row()==clicked.row() and prev_rows[0].parent()==clicked.parent())
            super().mousePressEvent(e)
            if same_single and len(sm.selectedRows(0)) == 0:
                sm.select(clicked, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                sm.setCurrentIndex(clicked, QtCore.QItemSelectionModel.NoUpdate)
                e.accept()
                return
            return

        super().mousePressEvent(e)


# -------------------- Explorer Pane --------------------
class ExplorerPane(QWidget):
    requestBackgroundOp=pyqtSignal(str, list, str)
    def __init__(self, _unused, start_path: str, pane_id: int, host_main):
        super().__init__()
        self.pane_id=pane_id; self.host=host_main
        self._search_mode=False; self._search_model=None; self._search_proxy=None
        self._back_stack=[]; self._fwd_stack=[]; self._undo_stack=[]
        self._last_hover_index=QtCore.QModelIndex(); self._tooltip_last_ms=0.0; self._tooltip_interval_ms=180; self._tooltip_last_text=""
        self._fast_model=FastDirModel(self); self._fast_proxy=FsSortProxy(self); self._fast_proxy.setSourceModel(self._fast_model)
        self._using_fast=False; self._fast_stat_worker=None; self._enum_worker=None; self._pending_normal_root=None

        # Toolbar
        self.btn_star=QToolButton(self); self.btn_star.setCheckable(True)
        self.btn_star.setIcon(icon_star(False, getattr(self.host,"theme","dark"))); self.btn_star.setToolTip("Add bookmark for this folder"); self.btn_star.setFixedHeight(UI_H)
        self._bm_btn_container=QWidget(self); self._bm_btn_layout=QHBoxLayout(self._bm_btn_container)
        self._bm_btn_layout.setContentsMargins(0,0,0,0); self._bm_btn_layout.setSpacing(ROW_SPACING)

        self.btn_cmd=QToolButton(self); self.btn_cmd.setIcon(icon_cmd(self.host.theme)); self.btn_cmd.setToolTip("Open Command Prompt here"); self.btn_cmd.setFixedHeight(UI_H)
        self.btn_up=QToolButton(self); self.btn_up.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp)); self.btn_up.setToolTip("Up"); self.btn_up.setFixedHeight(UI_H)
        self.btn_new=QToolButton(self); self.btn_new.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder)); self.btn_new.setToolTip("New Folder"); self.btn_new.setFixedHeight(UI_H)

        # 새 문서(.txt) 버튼
        self.btn_new_file=QToolButton(self)
        self.btn_new_file.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        self.btn_new_file.setToolTip("New Text File (.txt)")
        self.btn_new_file.setFixedHeight(UI_H)

        self.btn_refresh=QToolButton(self); self.btn_refresh.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload)); self.btn_refresh.setToolTip("Refresh"); self.btn_refresh.setFixedHeight(UI_H)

        row_toolbar=QHBoxLayout()
        row_toolbar.setContentsMargins(0,0,0,0)
        # ▶ 우상단 아이콘 사이 간격 축소
        row_toolbar.setSpacing(max(0, ROW_SPACING-2))
        row_toolbar.addWidget(self.btn_star)
        row_toolbar.addWidget(self._bm_btn_container,1)
        row_toolbar.addWidget(self.btn_cmd)
        row_toolbar.addWidget(self.btn_up)
        row_toolbar.addWidget(self.btn_new)
        row_toolbar.addWidget(self.btn_new_file)
        row_toolbar.addWidget(self.btn_refresh)
        self._row_toolbar=row_toolbar

        # ▶ 우상단 아이콘 좌우 패딩 축소(해당 버튼에만 적용)
        _tight_css = "QToolButton{padding-left:4px;padding-right:4px;}"
        for b in (self.btn_cmd, self.btn_up, self.btn_new, self.btn_new_file, self.btn_refresh):
            b.setStyleSheet(_tight_css)
            b.setAutoRaise(True)  # 테두리/여백 느낌 최소화

        # Path
        self.path_bar=PathBar(self); self.path_bar.setToolTip("Breadcrumb — Double-click or F4/Ctrl+L to enter path")
        row_path=QHBoxLayout(); row_path.setContentsMargins(0,0,0,0); row_path.setSpacing(0); row_path.addWidget(self.path_bar,1)

        # Filter
        self.filter_label=QLabel("Filter:", self)
        self.filter_edit=QLineEdit(self); self.filter_edit.setPlaceholderText("Filter (*.pdf, *file*.xls*, …)"); self.filter_edit.setClearButtonEnabled(True); self.filter_edit.setFixedHeight(UI_H)
        self.filter_label.setFixedHeight(UI_H); self.filter_label.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.btn_search=QToolButton(self); self.btn_search.setText("Search"); self.btn_search.setToolTip("Run recursive search"); self.btn_search.setFixedHeight(UI_H)
        row_filter=QHBoxLayout(); row_filter.setContentsMargins(0,0,0,0); row_filter.setSpacing(ROW_SPACING)
        row_filter.addWidget(self.filter_label); row_filter.addWidget(self.filter_edit,1); row_filter.addWidget(self.btn_search,0)

        # Models / View
        self.source_model=QFileSystemModel(self); self.source_model.setReadOnly(False)
        try: self.source_model.setResolveSymlinks(False)
        except Exception: pass
        self.source_model.setFilter(QDir.AllEntries|QDir.NoDotAndDotDot|QDir.Hidden|QDir.System|QDir.Drives|QDir.AllDirs)
        self._generic_icons=GenericIconProvider(self.style())
        if ALWAYS_GENERIC_ICONS: self.source_model.setIconProvider(self._generic_icons)

        self.stat_proxy=StatOverlayProxy(self); self.stat_proxy.setSourceModel(self.source_model)
        self.proxy=FsSortProxy(self); self.proxy.setSourceModel(self.stat_proxy)
        self.source_model.directoryLoaded.connect(self._on_directory_loaded)

        self.view=ExplorerView(self); self.view.setModel(self.proxy); self.view.setSortingEnabled(True)
        self.view.setAlternatingRowColors(True); self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._on_context_menu)
        self.view.setMouseTracking(True)
        self.view.setUniformRowHeights(True); self.view.setAnimated(False); self.view.setExpandsOnDoubleClick(False); self.view.setRootIsDecorated(False)
        header=self.view.header()
        header.setStretchLastSection(False)
        for i in range(4): header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, 90); header.resizeSection(3, 150)
        header.sectionClicked.connect(self._on_header_clicked)
        self.view.setColumnHidden(2, True)

        # Pane status
        self.lbl_sel=QLabel("", self); self.lbl_free=QLabel("", self)
        row_status=QHBoxLayout(); row_status.setContentsMargins(0,0,0,0); row_status.setSpacing(ROW_SPACING)
        row_status.addWidget(self.lbl_sel,0); row_status.addStretch(1); row_status.addWidget(self.lbl_free,0)
        self._row_status=row_status

        # Layout
        root_layout=QVBoxLayout(self); root_layout.setContentsMargins(*PANE_MARGIN); root_layout.setSpacing(max(1, ROW_SPACING//2))
        root_layout.addLayout(row_toolbar); root_layout.addLayout(row_path); root_layout.addLayout(row_filter)
        root_layout.addWidget(self.view,1); root_layout.addLayout(row_status)

        # Bookmarks, start
        self.host.namedBookmarksChanged.connect(self._on_bookmarks_changed)
        self._dirload_timer={}
        self.set_path(start_path or QDir.homePath(), push_history=False)
        self._update_star_button(); self._rebuild_quick_bookmark_buttons()

        # Signals
        self.path_bar.pathSubmitted.connect(lambda p: self.set_path(p, push_history=True))
        self.btn_star.clicked.connect(self._on_star_toggle)
        self.btn_cmd.clicked.connect(self._open_cmd_here)
        self.btn_up.clicked.connect(self.go_up)
        self.btn_new.clicked.connect(self.create_folder)
        self.btn_new_file.clicked.connect(self.create_text_file)  # 새 문서 생성
        self.btn_refresh.clicked.connect(self.hard_refresh)  # 하드 리프레시
        self.view.activated.connect(self._on_double_click)
        self.view.viewport().installEventFilter(self)
        self.view.selectionModel().selectionChanged.connect(self._on_selection_changed)
        self.filter_edit.returnPressed.connect(self._apply_filter); self.btn_search.clicked.connect(self._apply_filter)
        try: self.view.verticalScrollBar().valueChanged.connect(lambda _v: self._schedule_visible_stats())
        except Exception: pass

        # Shortcuts
        def add_sc(seq, slot):
            sc=QShortcut(QKeySequence(seq), self.view)
            sc.setContext(Qt.WidgetWithChildrenShortcut); sc.activated.connect(slot); return sc
        add_sc("Backspace", self.go_back); add_sc("Alt+Left", self.go_back); add_sc("Alt+Right", self.go_forward)
        add_sc("Ctrl+L", self.path_bar.start_edit); add_sc("F4", self.path_bar.start_edit); add_sc("F3", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+C", self.copy_selection); add_sc("Ctrl+X", self.cut_selection); add_sc("Ctrl+V", self.paste_into_current); add_sc("Ctrl+Z", self.undo_last)
        add_sc("Delete", self.delete_selection); add_sc("Shift+Delete", lambda: self.delete_selection(permanent=True)); add_sc("F2", self.rename_selection)
        add_sc(Qt.Key_Return, self._open_current); add_sc(Qt.Key_Enter, self._open_current); add_sc("Ctrl+O", self._open_current)

        # 경로 복사 단축키
        add_sc("Ctrl+Shift+C", lambda: self._copy_path_shortcut(False))  # 전체 경로
        add_sc("Alt+Shift+C",  lambda: self._copy_path_shortcut(True))   # 폴더까지만

        # 기본 정렬: 크기 내림차순 + 날짜 폭 확장
        self._apply_default_sort_size()

        self._update_pane_status()


    def _copy_path_shortcut(self, folder_only: bool = False):
        """
        Ctrl+Shift+C : 선택한 항목의 전체 경로 복사
        Alt+Shift+C  : 파일은 폴더 경로만, 폴더는 그대로 복사
        """
        sel = self._selected_paths()
        if len(sel) != 1:
            try:
                self.host.statusBar().showMessage("하나의 파일/폴더를 선택하세요.", 2000)
            except Exception:
                pass
            return

        p = sel[0]
        try:
            # 폴더-only 모드이면, 파일일 때만 dirname 적용 (폴더는 동일)
            if folder_only and os.path.isfile(p):
                p = os.path.dirname(p)
        except Exception:
            # 경로 판별 실패 시 그대로 둠
            pass

        try:
            QApplication.clipboard().setText(p)
            if folder_only:
                self.host.flash_status("폴더 경로가 클립보드에 복사되었습니다.")
            else:
                self.host.flash_status("전체 경로가 클립보드에 복사되었습니다.")
        except Exception:
            try:
                self.host.statusBar().showMessage("클립보드 복사에 실패했습니다.", 2000)
            except Exception:
                pass

    def _cancel_search_worker(self):
        # 검색/가시 영역 stat 워커 모두 중단
        try:
            w = getattr(self, "_search_worker", None)
            if w and w.isRunning():
                w.cancel()
                w.wait(100)
        except Exception:
            pass
        self._search_worker = None
        try:
            sw = getattr(self, "_search_stat_worker", None)
            if sw and sw.isRunning():
                sw.cancel()
                sw.wait(50)
        except Exception:
            pass
        self._search_stat_worker = None

    @QtCore.pyqtSlot(str, list)
    def _on_search_batch(self, base_path: str, rows: list):
        if not self._search_mode or not self._search_model:
            return
        root_item = self._search_model.invisibleRootItem()

        for rec in rows:
            name = rec.get("name", "")
            full = rec.get("path", "")
            isdir = bool(rec.get("is_dir", False))
            rel_folder = rec.get("folder", "")

            item_name = QStandardItem(name)
            item_name.setData(full, Qt.UserRole)
            item_name.setData(isdir, IS_DIR_ROLE)
            item_name.setData(full, Qt.ToolTipRole)
            # 우선 기본 아이콘(빠름)
            item_name.setIcon(self._default_icon(isdir))

            # 크기/날짜는 나중에 가시 영역만 채움
            item_size = QStandardItem()
            item_size.setData(0, Qt.EditRole)
            item_size.setData(0, SIZE_BYTES_ROLE)

            item_date = QStandardItem("")
            item_date.setData(QDateTime(), Qt.EditRole)

            item_folder = QStandardItem(rel_folder)

            root_item.appendRow([item_name, item_size, item_date, item_folder])

            # 경로 → (size_item, date_item) 매핑 저장
            if full:
                d = getattr(self, "_search_item_by_path", None)
                if isinstance(d, dict):
                    d[full] = (item_size, item_date)

        # 배치가 들어올 때마다 가시 영역의 아이콘/크기 갱신 예약
        QTimer.singleShot(0, self._fill_search_visible_icons)

    @QtCore.pyqtSlot()
    def _on_search_finished(self):
        QApplication.restoreOverrideCursor()
        self._search_worker = None
        # 마무리로 한 번 더 가시 영역 갱신
        QTimer.singleShot(0, self._fill_search_visible_icons)

    @QtCore.pyqtSlot(str, object, object)
    def _apply_search_stat(self, path: str, size_val, mtime_val):
        # 가시 영역 stat 워커가 알려준 값을 검색 모델에 반영
        d = getattr(self, "_search_item_by_path", None)
        if not isinstance(d, dict):
            return
        pair = d.get(path)
        if not pair:
            return
        item_size, item_date = pair
        try:
            sv = int(size_val or 0)
        except Exception:
            sv = 0
        item_size.setData(sv, Qt.EditRole)
        item_size.setData(sv, SIZE_BYTES_ROLE)

        if mtime_val is not None:
            try:
                dt = QDateTime.fromSecsSinceEpoch(int(mtime_val))
                item_date.setData(dt.toString("yyyy-MM-dd HH:mm:ss"), Qt.DisplayRole)
                item_date.setData(dt, Qt.EditRole)
            except Exception:
                item_date.setData("", Qt.DisplayRole)
                item_date.setData(QDateTime(), Qt.EditRole)


    def create_text_file(self):
        """
        현재 폴더에 임의(시간기반)의 새 .txt 파일을 생성합니다.
        ShellNew 템플릿(.txt)이 등록되어 있으면 그것을 사용합니다.
        """
        base_dir = self.current_path()
        try:
            # 시각 기반 고유 이름 (동일 초에 여러 번 눌러도 unique_dest_path로 충돌 방지)
            name = f"New Document {time.strftime('%Y%m%d-%H%M%S')}.txt"
            _create_new_file_with_template(base_dir, name, ".txt")
            self.hard_refresh()
            self.host.flash_status("Text file created")
        except Exception as e:
            QMessageBox.critical(self, "Create failed", str(e))


        # Shortcuts
        def add_sc(seq, slot):
            sc=QShortcut(QKeySequence(seq), self.view)
            sc.setContext(Qt.WidgetWithChildrenShortcut); sc.activated.connect(slot); return sc
        add_sc("Backspace", self.go_back); add_sc("Alt+Left", self.go_back); add_sc("Alt+Right", self.go_forward)
        add_sc("Ctrl+L", self.path_bar.start_edit); add_sc("F4", self.path_bar.start_edit); add_sc("F3", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+C", self.copy_selection); add_sc("Ctrl+X", self.cut_selection); add_sc("Ctrl+V", self.paste_into_current); add_sc("Ctrl+Z", self.undo_last)
        add_sc("Delete", self.delete_selection); add_sc("Shift+Delete", lambda: self.delete_selection(permanent=True)); add_sc("F2", self.rename_selection)
        add_sc(Qt.Key_Return, self._open_current); add_sc(Qt.Key_Enter, self._open_current); add_sc("Ctrl+O", self._open_current)

        # 기본 정렬: 크기 내림차순 + 날짜 폭 확장
        self._apply_default_sort_size()

        self._update_pane_status()

    def _apply_default_sort_size(self):
        v = self.view
        if not v.isSortingEnabled():
            v.setSortingEnabled(True)
        v.header().setSortIndicator(1, Qt.DescendingOrder)
        v.sortByColumn(1, Qt.DescendingOrder)
        try:
            v.header().resizeSection(3, 190)
        except Exception:
            pass

    def _default_icon(self, is_dir: bool) -> QIcon:
        try:
            if ALWAYS_GENERIC_ICONS:
                return self._generic_icons.icon(QFileIconProvider.Folder if is_dir else QFileIconProvider.File)
            return self.style().standardIcon(QStyle.SP_DirIcon if is_dir else QStyle.SP_FileIcon)
        except Exception: return QIcon()

    def _cancel_fast_stat_worker(self):
        w=self._fast_stat_worker
        if w and w.isRunning(): w.cancel(); w.wait(100)
        self._fast_stat_worker=None

    def _schedule_visible_stats(self):
        # 검색 모드에서는 먼저 아이콘만 채움
        if self._search_mode:
            self._fill_search_visible_icons()

        vp = self.view.viewport()

        # -------- Fast 모델(빠른 나열)일 때: 자체 워커로 size/mtime 채움 --------
        if self._using_fast:
            rc = self._fast_model.rowCount()
            if rc <= 0:
                return

            top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
            bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))

            if top_ix.isValid() and bot_ix.isValid():
                proxy_start = max(0, top_ix.row() - 30)
                proxy_end   = min(rc - 1, bot_ix.row() + 50)
            else:
                # 초기 진입 직후: 아직 화면 배치 전이면 선두 N개를 선제적으로 수집
                proxy_start = 0
                proxy_end   = min(rc - 1, 199)

            to_rows = []
            for r in range(proxy_start, proxy_end + 1):
                prx_ix = self._fast_proxy.index(r, 0)
                src_ix = self._fast_proxy.mapToSource(prx_ix)
                row    = src_ix.row()
                if row is None or row < 0:
                    continue

                # 파일 크기/시간 미수집이면 수집 대상으로 추가
                if not self._fast_model.has_stat(row):
                    to_rows.append(row)

                # 아이콘은 즉시 표시
                if not self._fast_model.has_icon(row):
                    p = self._fast_model.row_path(row)
                    if p:
                        idx = self.source_model.index(p)
                        if idx.isValid():
                            icon = self.source_model.fileIcon(idx)
                            if icon and not icon.isNull():
                                self._fast_model.apply_icon(row, icon)

            if not to_rows:
                return

            self._cancel_fast_stat_worker()
            root = self._fast_model.rootPath()
            w = FastStatWorker(self._fast_model, root, to_rows, self)
            w.statReady.connect(self._fast_model.apply_stat, Qt.QueuedConnection)
            w.finishedCycle.connect(lambda: None)
            self._fast_stat_worker = w
            w.start()
            return

        # -------- Normal 모델(QFileSystemModel + StatOverlayProxy)일 때 --------
        model = self.proxy
        stat_proxy = self.stat_proxy
        src_model = self.source_model

        mrc = model.rowCount()
        if mrc <= 0:
            return

        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))

        if top_ix.isValid() and bot_ix.isValid():
            start = max(0, top_ix.row() - 40)
            end   = min(mrc - 1, bot_ix.row() + 80)
        else:
            # 초기 진입 직후: 화면 배치 전이면 선두 N개 경로를 선제적으로 요청
            start = 0
            end   = min(mrc - 1, 199)

        if end < start:
            end = start

        paths = []
        for r in range(start, end + 1):
            prx_ix = model.index(r, 0)
            st_ix  = model.mapToSource(prx_ix)
            src_ix = stat_proxy.mapToSource(st_ix)
            try:
                p = src_model.filePath(src_ix)
            except Exception:
                p = None
            if not p:
                continue
            paths.append(p)

        if paths:
            stat_proxy.request_paths(paths)


    def _on_header_clicked(self, col:int):
        v=self.view
        if not v.isSortingEnabled(): v.setSortingEnabled(True)
        if col == 1:
            v.header().setSortIndicator(1, Qt.DescendingOrder)
            v.sortByColumn(1, Qt.DescendingOrder)
            try: v.header().resizeSection(3, 190)
            except Exception: pass
            return
        v.header().setSortIndicator(col, Qt.AscendingOrder); v.sortByColumn(col, Qt.AscendingOrder)

    def eventFilter(self, obj, ev):
        if obj is self.view.viewport():
            if ev.type()==QEvent.MouseButtonPress:
                if ev.button()==Qt.XButton1: self.go_back(); return True
                if ev.button()==Qt.XButton2: self.go_forward(); return True
            if ev.type()==QEvent.MouseMove:
                ix=self.view.indexAt(ev.pos())
                if ix!=self._last_hover_index: self._last_hover_index=ix
                if ix.isValid():
                    now_ms=time.perf_counter()*1000.0
                    if (now_ms-self._tooltip_last_ms)>=self._tooltip_interval_ms:
                        name=ix.sibling(ix.row(),0).data(Qt.DisplayRole); full=self._index_to_full_path(ix)
                        tip=full if full else name
                        if tip!=self._tooltip_last_text:
                            QToolTip.showText(QCursor.pos(), tip, self.view); self._tooltip_last_text=tip; self._tooltip_last_ms=now_ms
            if ev.type() in (QEvent.Resize, QEvent.Show):
                QTimer.singleShot(0, self._schedule_visible_stats)
            return False
        return super().eventFilter(obj, ev)

    def _open_cmd_here(self):
        path=self.current_path()
        try:
            flags=getattr(subprocess,"CREATE_NEW_CONSOLE",0)
            subprocess.Popen(["cmd.exe","/K"], cwd=path, creationflags=flags)
        except Exception:
            try: subprocess.Popen('start "" cmd.exe', shell=True, cwd=path)
            except Exception as e: QMessageBox.critical(self,"Command Prompt",f"Failed to launch cmd.exe:\n{e}")

    def _on_bookmarks_changed(self,*_): self._update_star_button(); self._rebuild_quick_bookmark_buttons()
    def _on_star_toggle(self): self.host.toggle_bookmark(self.current_path())
    def _update_star_button(self):
        idx,it=self.host.is_path_bookmarked(self.current_path()); checked=bool(it and it.get("enabled"))
        self.btn_star.setChecked(checked); self.btn_star.setIcon(icon_star(checked, getattr(self.host,"theme","dark")))
        self.btn_star.setToolTip("Remove bookmark for this folder" if checked else "Add bookmark for this folder")
    def _rebuild_quick_bookmark_buttons(self):
        while self._bm_btn_layout.count():
            it=self._bm_btn_layout.takeAt(0); w=it.widget()
            if w: w.deleteLater()
        for it in self.host.get_enabled_bookmarks():
            name=it.get("name") or _derive_name_from_path(it.get("path","")); p=it.get("path","")
            btn=QToolButton(self._bm_btn_container); btn.setText(name); btn.setToolTip(p); btn.setFixedHeight(UI_H)
            btn.clicked.connect(lambda _=False, path=p: self.set_path(path, push_history=True))
            self._bm_btn_layout.addWidget(btn)

    def _selected_paths(self):
        paths=[]; sel=self.view.selectionModel().selectedRows(0)
        for ix in sel:
            p=self._index_to_full_path(ix)
            if p: paths.append(p)
        seen=set(); out=[]
        for p in paths:
            np=os.path.normpath(p)
            if np not in seen: seen.add(np); out.append(np)
        return out

    def _index_to_full_path(self, index):
        if not index.isValid(): return None
        if self._using_fast or self._search_mode:
            return index.sibling(index.row(),0).data(Qt.UserRole)
        else:
            st_ix=self.proxy.mapToSource(index); src_ix=self.stat_proxy.mapToSource(st_ix)
            return self.source_model.filePath(src_ix)

    def _use_fast_model(self, path:str):
        self._cancel_fast_stat_worker()
        if self._enum_worker and self._enum_worker.isRunning(): self._enum_worker.cancel(); self._enum_worker.wait(100)
        self.stat_proxy.clear_cache(); self._using_fast=True
        self._fast_model.reset_dir(path); self.view.setModel(self._fast_proxy)
        header=self.view.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.resizeSection(1, 90); header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)
        if not self.view.isSortingEnabled(): self.view.setSortingEnabled(True)
        header.setSortIndicator(0, Qt.AscendingOrder); self.view.sortByColumn(0, Qt.AscendingOrder)
        self._apply_default_sort_size()
        self._enum_worker=DirEnumWorker(path)
        self._enum_worker.batchReady.connect(self._fast_model.append_rows, QtCore.Qt.QueuedConnection)
        self._enum_worker.batchReady.connect(lambda _rows: self._schedule_visible_stats(), QtCore.Qt.QueuedConnection)
        self._enum_worker.error.connect(lambda msg: self.host.statusBar().showMessage(f"List error: {msg}", 4000))
        self._enum_worker.finished.connect(lambda: self._schedule_visible_stats())
        self._enum_worker.start()
        QTimer.singleShot(0, self._schedule_visible_stats)

    def _start_normal_model_loading(self, path:str):
        self._pending_normal_root=path; t=QElapsedTimer(); t.start(); self._dirload_timer[path.lower()]=t
        if ALWAYS_GENERIC_ICONS: self.source_model.setIconProvider(self._generic_icons)
        _ = self.source_model.setRootPath(path)

    def set_path(self, path:str, push_history:bool=True):
        with perf(f"set_path begin -> {path}"):
            path=nice_path(path)
            if not os.path.exists(path): QMessageBox.warning(self,"Path not found",path); return
            cur=getattr(self.path_bar,"_current_path",None)
            if push_history and cur and os.path.normcase(cur)!=os.path.normcase(path):
                self._back_stack.append(cur); self._fwd_stack.clear()
            self.path_bar.set_path(path); self._update_star_button()
            self._use_fast_model(path)
            QTimer.singleShot(0, lambda: self._start_normal_model_loading(path))
            QTimer.singleShot(50, self._update_pane_status); self._update_statusbar_selection()

    @QtCore.pyqtSlot(str)
    def _on_directory_loaded(self, loaded_path:str):
        key=loaded_path.lower()
        if key in self._dirload_timer:
            ms=self._dirload_timer[key].elapsed(); dlog(f"directoryLoaded: '{loaded_path}' in {ms} ms")
            self._dirload_timer.pop(key,None)
        if (os.path.normcase(loaded_path)==os.path.normcase(self.current_path())
            and self._using_fast and self._pending_normal_root
            and os.path.normcase(self._pending_normal_root)==os.path.normcase(loaded_path)
            and not self._search_mode):
            try:
                src_idx=self.source_model.index(loaded_path)
                self.view.setModel(self.proxy)
                self.view.setRootIndex(self.proxy.mapFromSource(self.stat_proxy.mapFromSource(src_idx)))
                self._using_fast=False; self._pending_normal_root=None
                header=self.view.header()
                for i in range(4): header.setSectionResizeMode(i, QHeaderView.Interactive)
                header.resizeSection(1, 90); header.resizeSection(3, 150)
                self.view.setColumnHidden(2, True)
                if not self.view.isSortingEnabled(): self.view.setSortingEnabled(True)
                header.setSortIndicator(0, Qt.AscendingOrder); self.view.sortByColumn(0, Qt.AscendingOrder)
                self._apply_default_sort_size()
                QTimer.singleShot(0, self._schedule_visible_stats)
            except Exception: pass

    def current_path(self)->str: return self.path_bar._current_path or QDir.homePath()
    def go_back(self):
        if not self._back_stack: return
        dst=self._back_stack.pop(); self._fwd_stack.append(self.current_path()); self.set_path(dst, push_history=False)
    def go_forward(self):
        if not self._fwd_stack: return
        dst=self._fwd_stack.pop(); self._back_stack.append(self.current_path()); self.set_path(dst, push_history=False)
    def go_up(self):
        parent=Path(self.current_path()).parent; self.set_path(str(parent), push_history=True)
    def refresh(self):  # 버튼/단축키 → 하드 리프레시 동일
        self.hard_refresh()
    def hard_refresh(self):
        if self._search_mode:
            self._apply_filter(); return
        try: self._cancel_fast_stat_worker()
        except Exception: pass
        try:
            if self._enum_worker and self._enum_worker.isRunning():
                self._enum_worker.cancel(); self._enum_worker.wait(100)
        except Exception: pass
        try: self.stat_proxy.clear_cache()
        except Exception: pass
        self.set_path(self.current_path(), push_history=False)
        self.host.flash_status("Hard refresh")

    # ---- Open helpers ----
    def _open_file_with_cwd(self, path:str):
        folder=os.path.dirname(path) or self.current_path()
        try:
            if HAS_PYWIN32:
                win32api.ShellExecute(int(self.window().winId()) if self.window() else 0, None, path, None, folder, win32con.SW_SHOWNORMAL); return
        except Exception: pass
        try:
            if path.lower().endswith((".bat",".cmd")):
                flags=getattr(subprocess,"CREATE_NEW_CONSOLE",0)
                subprocess.Popen(["cmd.exe","/C", path], cwd=folder, creationflags=flags)
            else:
                subprocess.Popen(f'start "" "{path}"', shell=True, cwd=folder)
        except Exception:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _open_many(self, paths:list[str]):
        files=[p for p in paths if os.path.isfile(p)]
        for p in files: self._open_file_with_cwd(p)
        dirs=[p for p in paths if os.path.isdir(p)]
        if not files and len(dirs)==1: self.set_path(dirs[0], push_history=True)

    def _on_double_click(self, index):
        if not index.isValid(): return
        path=self._index_to_full_path(index)
        if not path: return
        if os.path.isdir(path): self.set_path(path, push_history=True)
        else: self._open_file_with_cwd(path)

    def _open_current(self):
        sel=self._selected_paths()
        if len(sel)>=2: self._open_many(sel); return
        ix=self.view.currentIndex()
        if not ix.isValid():
            if sel:
                p=sel[0]; self.set_path(p, True) if os.path.isdir(p) else self._open_file_with_cwd(p)
            return
        p=self._index_to_full_path(ix)
        if not p: return
        self.set_path(p,True) if os.path.isdir(p) else self._open_file_with_cwd(p)

    def create_folder(self):
        base=self.current_path()
        name,ok=QInputDialog.getText(self,"New Folder","Name:", text="New Folder")
        if ok and name:
            target=os.path.join(base,name)
            try:
                os.makedirs(target, exist_ok=False)
                self._undo_stack.append({"type":"mkdir","path":target}); self.refresh()
            except FileExistsError:
                QMessageBox.information(self,"Exists","Folder already exists:\n"+target)
            except Exception as e:
                QMessageBox.critical(self,"Error",str(e))

    # ---- Clipboard ops (w/ conflicts dialog) ----
    def copy_selection(self):
        paths=self._selected_paths()
        if not paths: return
        self.host.set_clipboard({"op":"copy","paths":paths}); self.host.flash_status(f"Copied {len(paths)} item(s)")
    def cut_selection(self):
        paths=self._selected_paths()
        if not paths: return
        self.host.set_clipboard({"op":"cut","paths":paths}); self.host.flash_status(f"Cut {len(paths)} item(s)")
    def paste_into_current(self):
        clip=self.host.get_clipboard()
        if not clip: return
        dst_dir=self.current_path(); op=clip.get("op"); srcs=clip.get("paths") or []
        if not srcs: return
        self._start_bg_op("copy" if op=="copy" else "move", srcs, dst_dir)
        if op=="cut": self.host.clear_clipboard()

    def _start_bg_op(self, op, srcs, dst_dir):
        # 1) Build conflicts
        conflicts=[]
        for src in srcs:
            base=os.path.basename(src); dst=os.path.join(dst_dir, base)
            if os.path.exists(dst): conflicts.append((src,dst))

        conflict_map=None
        if conflicts:
            dlg=ConflictResolutionDialog(self, conflicts, dst_dir)
            if dlg.exec_()!=QDialog.Accepted: return
            # src -> "overwrite"|"skip"|"copy"
            conflict_map=dlg.result_map()

        # 2) Run worker
        worker=FileOpWorker(op, srcs, dst_dir, conflict_map=conflict_map)
        dlgp=QProgressDialog(f"{op.title()} in progress…","Cancel",0,100,self)
        dlgp.setWindowTitle(f"{op.title()} files"); dlgp.setWindowModality(Qt.ApplicationModal)
        dlgp.setAutoClose(True); dlgp.setAutoReset(True)

        worker.progress.connect(dlgp.setValue)
        worker.status.connect(lambda s: self.host.statusBar().showMessage(s,2000))
        worker.error.connect(lambda msg:(dlgp.close(), QMessageBox.critical(self,f"{op.title()} failed",msg)))

        def _finish_ok():
            try: dlgp.setValue(100); dlgp.close()
            except Exception: pass
            if not self._using_fast and not self._search_mode: self.stat_proxy.clear_cache()
            QTimer.singleShot(0, self._schedule_visible_stats); self._update_pane_status()
            self.host.flash_status(f"{op.title()} complete")

        worker.finished_ok.connect(_finish_ok); dlgp.canceled.connect(worker.cancel)
        worker.start(); dlgp.exec_()

    # ---- Delete / Rename / Undo ----
    def delete_selection(self, permanent:bool=False):
        paths=self._selected_paths()
        if not paths: return
        title="Delete permanently" if permanent else "Delete"
        action="permanently delete" if permanent else "move to Recycle Bin"
        msg=f"{len(paths)} item(s) will be {action}.\n\nAre you sure?"
        btn=QMessageBox.question(self,title,msg,QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
        if btn!=QMessageBox.Yes: return
        if permanent:
            errors=[]
            for p in paths:
                try: remove_any(p)
                except Exception as e: errors.append(f"{p}: {e}")
            if errors: QMessageBox.critical(self,"Delete failed","\n".join(errors)[:2000])
            else: self.host.flash_status(f"Deleted {len(paths)} item(s)")
            self.refresh(); return
        hwnd=int(self.window().winId()) if HAS_PYWIN32 else 0
        ok=recycle_to_trash(paths, hwnd)
        if ok: self.host.flash_status(f"Sent {len(paths)} item(s) to Recycle Bin"); self.refresh()
        else: QMessageBox.critical(self,"Delete failed","Could not move items to Recycle Bin.")

    def rename_selection(self):
        paths=self._selected_paths()
        if len(paths)!=1: return
        src=paths[0]; base=os.path.basename(src)
        new_name,ok=QInputDialog.getText(self,"Rename","New name:", text=base)
        if not ok or not new_name or new_name==base: return
        dst=os.path.join(os.path.dirname(src), new_name)
        if os.path.exists(dst):
            QMessageBox.warning(self,"Rename","A file or folder with that name already exists."); return
        try:
            os.rename(src, dst)
            self._undo_stack.append({"type":"move_back","pairs":[(dst,src)]}); self.refresh(); self.host.flash_status("Renamed")
        except Exception as e:
            QMessageBox.critical(self,"Rename failed",str(e))

    def undo_last(self):
        if not self._undo_stack: self.host.flash_status("Nothing to undo"); return
        act=self._undo_stack.pop(); t=act.get("type")
        try:
            if t=="mkdir":
                path=act["path"]
                try: os.rmdir(path)
                except OSError:
                    QMessageBox.information(self,"Undo New Folder","Folder is not empty; cannot undo safely."); return
            elif t=="delete":
                for p in act.get("paths",[]): remove_any(p)
            elif t=="move_back":
                for dst,src in act.get("pairs",[]):
                    target_dir=os.path.dirname(src)
                    if os.path.exists(src):
                        base=os.path.basename(src); src=unique_dest_path(target_dir, base)
                    shutil.move(dst, src)
            else: return
            self.refresh(); self.host.flash_status("Undone")
        except Exception as e:
            QMessageBox.critical(self,"Undo failed",str(e))

    # ---- Search ----
    def _enter_browse_mode(self):
        if not self._search_mode: return
        self._search_mode=False; self._search_model=None; self._search_proxy=None
        if self._using_fast: self.view.setModel(self._fast_proxy)
        else: self.view.setModel(self.proxy)
        path=self.current_path()
        if not self._using_fast:
            src_idx=self.source_model.index(path)
            self.view.setRootIndex(self.proxy.mapFromSource(self.stat_proxy.mapFromSource(src_idx)))
        header=self.view.header(); header.setStretchLastSection(False)
        for i in range(4): header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, 90); header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)
        if not self.view.isSortingEnabled(): self.view.setSortingEnabled(True)
        header.setSortIndicator(0, Qt.AscendingOrder); self.view.sortByColumn(0, Qt.AscendingOrder)
        self._apply_default_sort_size()
        QTimer.singleShot(0, self._schedule_visible_stats)

    def _enter_search_mode(self, model:QStandardItemModel):
        self._cancel_fast_stat_worker(); self._using_fast=False; self._search_mode=True
        self._search_model=model; self._search_proxy=FsSortProxy(self); self._search_proxy.setSourceModel(self._search_model)
        self.view.setModel(self._search_proxy); self.view.setRootIndex(QtCore.QModelIndex())
        header=self.view.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.resizeSection(1,90); header.resizeSection(2,150)
        if not self.view.isSortingEnabled(): self.view.setSortingEnabled(True)
        header.setSortIndicator(0, Qt.AscendingOrder); self.view.sortByColumn(0, Qt.AscendingOrder)

    def _apply_filter(self):
        pattern = self.filter_edit.text().strip()
        if not pattern:
            # 빈 패턴 → 탐색 모드 복귀
            self._enter_browse_mode()
            return

        # 이전 검색 중단
        self._cancel_search_worker()

        base = self.current_path()

        # 검색 결과 모델 준비 (가벼운 데이터만 먼저 채우고, stat은 나중에 화면 가시 영역만)
        model = SearchResultModel(self)
        model.setHorizontalHeaderLabels(["Name", "Size", "Date Modified", "Folder"])
        self._enter_search_mode(model)

        # 경로→아이템 매핑/상태 초기화 (가시 영역 stat 채우기용)
        self._search_item_by_path = {}
        self._search_stats_done = set()
        self._search_stat_worker = None

        # 워커 시작 (비동기 스트리밍)
        w = SearchWorker(base, pattern, self)
        self._search_worker = w
        w.batchReady.connect(self._on_search_batch, Qt.QueuedConnection)
        w.error.connect(lambda msg: self.host.statusBar().showMessage(f"Search error: {msg}", 4000))
        w.finished.connect(self._on_search_finished, Qt.QueuedConnection)

        QApplication.setOverrideCursor(Qt.WaitCursor)
        w.start()


    def _fill_search_visible_icons(self):
        if not self._search_mode or not self._search_proxy or not self._search_model:
            return

        vp = self.view.viewport()
        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        end = bot_ix.row() if bot_ix.isValid() else min(start + 200, self._search_proxy.rowCount() - 1)
        start = max(0, start - 40)
        end = min(self._search_proxy.rowCount() - 1, end + 100)
        if end < start:
            end = start

        paths_need_stat = []

        for r in range(start, end + 1):
            prx_ix = self._search_proxy.index(r, 0)
            src_ix = self._search_proxy.mapToSource(prx_ix)
            item_name = self._search_model.item(src_ix.row(), 0)
            item_size = self._search_model.item(src_ix.row(), 1)
            item_date = self._search_model.item(src_ix.row(), 2)
            if not item_name:
                continue

            p = item_name.data(Qt.UserRole)
            isdir = bool(item_name.data(IS_DIR_ROLE))

            # 아이콘: 실제 OS 아이콘으로 교체
            if p:
                idx = self.source_model.index(p)
                if idx.isValid():
                    icon = self.source_model.fileIcon(idx)
                    if icon and not icon.isNull():
                        item_name.setIcon(icon)

            # 가시 영역 stat: 폴더 제외, 아직 처리하지 않은 경로만
            if p and not isdir and p not in getattr(self, "_search_stats_done", set()):
                # size/date 가 비어 있는 경우에만 요청 (0-byte 파일의 재요청 방지 위해 set으로 관리)
                paths_need_stat.append(p)
                self._search_stats_done.add(p)

        if paths_need_stat:
            # 이전 워커가 돌고 있으면 취소
            try:
                if self._search_stat_worker and self._search_stat_worker.isRunning():
                    self._search_stat_worker.cancel()
                    self._search_stat_worker.wait(50)
            except Exception:
                pass
            w = NormalStatWorker(paths_need_stat, self)
            w.statReady.connect(self._apply_search_stat, Qt.QueuedConnection)
            self._search_stat_worker = w
            w.start()


    def _on_context_menu(self, pos):
        owner_hwnd=int(self.window().winId()) if HAS_PYWIN32 else 0; paths=self._selected_paths()
        handled = False
        if HAS_PYWIN32:
            try: cx,cy=win32api.GetCursorPos(); screen_pt=(int(cx),int(cy))
            except Exception:
                g=self.view.viewport().mapToGlobal(pos); screen_pt=(g.x(),g.y())
            if paths:
                handled = show_explorer_context_menu(owner_hwnd, paths, screen_pt)
            else:
                handled = show_explorer_background_menu(owner_hwnd, self.current_path(), screen_pt)
            if handled:
                return

        # Fallback menu (includes New)
        dst_dir = self.current_path()
        global_pt=QCursor.pos(); menu=QMenu(self)
        act_open=menu.addAction("Open"); act_reveal=menu.addAction("Open in Explorer"); act_copy_path=menu.addAction("Copy Path")
        menu.addSeparator()
        m_new = menu.addMenu("New")
        a_new_folder = m_new.addAction("Folder")
        a_new_txt    = m_new.addAction("Text Document (.txt)")
        a_new_docx   = m_new.addAction("Word Document (.docx)")
        a_new_xlsx   = m_new.addAction("Excel Workbook (.xlsx)")
        a_new_pptx   = m_new.addAction("PowerPoint Presentation (.pptx)")

        action=menu.exec_(global_pt); target=paths[0] if paths else dst_dir
        if action==act_open and paths: self._open_many(paths); return
        elif action==act_open and os.path.exists(target):
            self.set_path(target, True) if os.path.isdir(target) else self._open_file_with_cwd(target); return
        elif action==act_reveal:
            if os.path.isdir(target): os.startfile(target)
            else: os.system(f'explorer /select,\"{target}\"'); return
        elif action==act_copy_path: QApplication.clipboard().setText(target); return

        try:
            if action == a_new_folder:
                newp = unique_dest_path(dst_dir, "New Folder"); os.makedirs(newp, exist_ok=False)
                self.hard_refresh(); self.host.flash_status("Folder created"); return
            if action == a_new_txt:
                _create_new_file_with_template(dst_dir, "New Text Document.txt", ".txt")
                self.hard_refresh(); self.host.flash_status("Text file created"); return
            if action == a_new_docx:
                _create_new_file_with_template(dst_dir, "New Word Document.docx", ".docx")
                self.hard_refresh(); self.host.flash_status("Word document created"); return
            if action == a_new_xlsx:
                _create_new_file_with_template(dst_dir, "New Excel Workbook.xlsx", ".xlsx")
                self.hard_refresh(); self.host.flash_status("Excel workbook created"); return
            if action == a_new_pptx:
                _create_new_file_with_template(dst_dir, "New PowerPoint Presentation.pptx", ".pptx")
                self.hard_refresh(); self.host.flash_status("PowerPoint presentation created"); return
        except Exception as e:
            QMessageBox.critical(self, "Create failed", str(e))

    def _on_selection_changed(self,*_): self._update_statusbar_selection(); self._update_pane_status()
    def _update_statusbar_selection(self):
        sel=self._selected_paths(); count=len(sel); size=0
        for p in sel:
            if os.path.isdir(p):
                for root,dirs,files in os.walk(p):
                    for f in files:
                        try: size+=os.path.getsize(os.path.join(root,f))
                        except Exception: pass
            else:
                try: size+=os.path.getsize(p)
                except Exception: pass
        txt=f"Pane {self.pane_id} — selected {count} item(s)"; 
        if count: txt+=f" / {human_size(size)}"
        self.host.statusBar().showMessage(txt, 2000)
    def _drive_label(self, path:str)->str:
        if path.startswith("\\\\"):
            comps=[c for c in path.split("\\") if c]
            return f"\\\\{comps[0]}\\{comps[1]}" if len(comps)>=2 else "\\\\"
        drv,_=os.path.splitdrive(path); return drv if drv else os.sep
    def _update_pane_status(self):
        cnt=len(self._selected_paths()); self.lbl_sel.setText(f"{cnt} selected" if cnt else "")
        path=self.current_path()
        try:
            total,used,free=shutil.disk_usage(path)
            self.lbl_free.setText(f"{self._drive_label(path)} free {human_size(free)}")
        except Exception:
            self.lbl_free.setText("")

# -------------------- Main Window --------------------
class MultiExplorer(QMainWindow):
    namedBookmarksChanged=pyqtSignal(list)
    def __init__(self, pane_count:int=6, start_paths=None, initial_theme:str="dark"):
        super().__init__()
        self.theme=initial_theme if initial_theme in ("dark","light") else "dark"
        self._layout_states=[4,6,8]; self._layout_idx=self._layout_states.index(pane_count) if pane_count in self._layout_states else 1
        self.setWindowTitle(f"Multi-Pane File Explorer — {pane_count} panes"); self.resize(1500,900)

        # Top bar
        top=QWidget(self); top_lay=QHBoxLayout(top); top_lay.setContentsMargins(6,2,6,2); top_lay.setSpacing(ROW_SPACING)
        self.btn_layout=QToolButton(top); self.btn_layout.setToolTip("Toggle layout (4 ↔ 6 ↔ 8)"); self.btn_layout.setFixedHeight(UI_H)
        self.btn_theme=QToolButton(top); self.btn_theme.setToolTip("Toggle Light/Dark"); self.btn_theme.setFixedHeight(UI_H)
        self.btn_bm_edit=QToolButton(top); self.btn_bm_edit.setToolTip("Edit Bookmarks"); self.btn_bm_edit.setFixedHeight(UI_H)
        self.btn_about=QToolButton(top); self.btn_about.setToolTip("About"); self.btn_about.setFixedHeight(UI_H)
        top_lay.addWidget(self.btn_layout,0); top_lay.addWidget(self.btn_theme,0); top_lay.addWidget(self.btn_bm_edit,0); top_lay.addWidget(self.btn_about,0); top_lay.addStretch(1)

        self.central=QWidget(self); self.setCentralWidget(self.central)
        vmain=QVBoxLayout(self.central); vmain.setContentsMargins(0,0,0,0); vmain.setSpacing(ROW_SPACING)
        vmain.addWidget(top,0); self.grid=QGridLayout(); vmain.addLayout(self.grid,1)

        self.named_bookmarks=migrate_legacy_favorites_into_named(load_named_bookmarks()); save_named_bookmarks(self.named_bookmarks)
        self._clipboard=None; self._bm_dlg=None

        self._update_layout_icon(); self._update_theme_icon()
        self.btn_layout.clicked.connect(self._cycle_layout); self.btn_theme.clicked.connect(self._toggle_theme)
        self.btn_bm_edit.clicked.connect(self._open_bookmark_editor); self.btn_about.clicked.connect(self._show_about)

        self.panes=[]; self.build_panes(pane_count, start_paths or []); self._update_theme_dependent_icons()
        self.statusBar().showMessage("Ready", 1500)

        self._wd_timer=QTimer(self); self._wd_timer.setInterval(50); self._wd_last=time.perf_counter()
        def _wd_tick():
            now=time.perf_counter(); gap=(now-self._wd_last)*1000
            if gap>200: dlog(f"[STALL] UI event loop blocked ~{gap:.0f} ms")
            self._wd_last=now
        self._wd_timer.timeout.connect(_wd_tick); self._wd_timer.start()

        settings=QSettings(ORG_NAME, APP_NAME); geo=settings.value("window/geometry")
        if isinstance(geo, QtCore.QByteArray): self.restoreGeometry(geo)

    def _update_layout_icon(self):
        states=getattr(self,"_layout_states",[4,6,8]); idx=getattr(self,"_layout_idx",0)
        if not states: states=[4,6,8]
        if idx>=len(states) or idx<0: idx=0; self._layout_idx=0
        state=states[idx]; self.btn_layout.setIcon(icon_grid_layout(state, self.theme))
    def _update_theme_icon(self):
        if hasattr(self,"btn_theme") and self.btn_theme: self.btn_theme.setIcon(icon_theme_toggle(self.theme))
        if hasattr(self,"btn_bm_edit") and self.btn_bm_edit: self.btn_bm_edit.setIcon(icon_edit(self.theme))
        if hasattr(self,"btn_about") and self.btn_about: self.btn_about.setIcon(icon_info(self.theme))
    def _update_theme_dependent_icons(self):
        self._update_layout_icon(); self._update_theme_icon()
        for p in getattr(self,"panes",[]):
            try: p.btn_star.setIcon(icon_star(p.btn_star.isChecked(), self.theme)); p.btn_cmd.setIcon(icon_cmd(self.theme))
            except Exception: pass
    def _cycle_layout(self):
        self._layout_idx=(self._layout_idx+1)%len(self._layout_states); n=self._layout_states[self._layout_idx]; self.build_panes(n, self._current_paths())
    def _toggle_theme(self):
        self.theme="light" if self.theme=="dark" else "dark"
        app=QApplication.instance()
        apply_dark_style(app) if self.theme=="dark" else apply_light_style(app)
        self._update_theme_dependent_icons()
        s=QSettings(ORG_NAME, APP_NAME); s.setValue("ui/theme", self.theme); s.sync()

    def build_panes(self, n:int, start_paths):
        cols={4:2,6:3,8:4}.get(n,3); gap=GRID_GAPS.get(cols,3); margin_lr=GRID_MARG_LR.get(cols,6)
        self.grid.setSpacing(gap); self.grid.setContentsMargins(margin_lr,2,margin_lr,4)
        while self.grid.count():
            it=self.grid.takeAt(0); w=it.widget()
            if w and isinstance(w, ExplorerPane): w.setParent(None)
        self.panes.clear()
        for i in range(n):
            spath=start_paths[i] if i<len(start_paths) else None
            pane=ExplorerPane(None, start_path=spath, pane_id=i+1, host_main=self)
            self.panes.append(pane); r=i//cols; c=i%cols; self.grid.addWidget(pane, r, c)
        self.setWindowTitle(f"Multi-Pane File Explorer — {n} panes")
        self._update_theme_dependent_icons()

    def _current_paths(self): return [p.current_path() for p in self.panes]
    def set_clipboard(self,payload:dict): self._clipboard=payload
    def get_clipboard(self): return self._clipboard
    def clear_clipboard(self): self._clipboard=None
    def flash_status(self,text:str):
        try: self.statusBar().showMessage(text,2000)
        except Exception: pass

    # Bookmarks
    def _find_bookmark_index_by_path(self, path:str):
        np=os.path.normcase(nice_path(path))
        for i,it in enumerate(self.named_bookmarks):
            if os.path.normcase(it.get("path",""))==np: return i
        return -1
    def is_path_bookmarked(self, path:str):
        i=self._find_bookmark_index_by_path(path)
        return (i, self.named_bookmarks[i]) if i>=0 else (-1,None)
    def get_enabled_bookmarks(self): return [it for it in self.named_bookmarks if it.get("enabled") and it.get("path")]
    def toggle_bookmark(self, path:str):
        np=nice_path(path); idx=self._find_bookmark_index_by_path(np)
        if idx>=0:
            it=dict(self.named_bookmarks[idx]); it["enabled"]=not bool(it.get("enabled"))
            if not it.get("name"): it["name"]=_derive_name_from_path(np)
            self.named_bookmarks[idx]=it
        else:
            if len(self.named_bookmarks)>=10 and all(x.get("enabled") for x in self.named_bookmarks):
                QMessageBox.information(self,"Bookmarks","Bookmark limit reached (10). Please edit bookmarks to free a slot.")
                self._open_bookmark_editor(); return
            reused=False
            for i,it in enumerate(self.named_bookmarks):
                if not it.get("enabled") and not it.get("name") and not it.get("path"):
                    self.named_bookmarks[i]={"enabled":True,"name":_derive_name_from_path(np),"path":np}; reused=True; break
            if not reused: self.named_bookmarks.append({"enabled":True,"name":_derive_name_from_path(np),"path":np})
            if len(self.named_bookmarks)>10: self.named_bookmarks=self.named_bookmarks[:10]
        save_named_bookmarks(self.named_bookmarks); self.namedBookmarksChanged.emit(self.named_bookmarks); self.flash_status("Bookmarks updated")

    def _open_bookmark_editor(self):
        if getattr(self,"_bm_dlg",None) and self._bm_dlg.isVisible():
            self._bm_dlg.raise_(); self._bm_dlg.activateWindow(); return
        dlg=BookmarkEditDialog(self, items=self.named_bookmarks); self._bm_dlg=dlg
        try: self.namedBookmarksChanged.connect(dlg.set_items)
        except Exception: pass
        dlg.finished.connect(self._on_bmdlg_closed)
        if dlg.exec_()==QDialog.Accepted:
            new_items=dlg.values(); cleaned=[]
            for it in new_items[:10]:
                cleaned.append({"name":it.get("name","").strip(),"path":it.get("path","").strip(),"enabled":bool(it.get("enabled",False))})
            self.named_bookmarks=cleaned[:10]; save_named_bookmarks(self.named_bookmarks); self.namedBookmarksChanged.emit(self.named_bookmarks)

    def _on_bmdlg_closed(self,*_):
        try:
            if getattr(self,"_bm_dlg",None): self.namedBookmarksChanged.disconnect(self._bm_dlg.set_items)
        except Exception: pass
        self._bm_dlg=None

    def _show_about(self):
        dlg=QDialog(self); dlg.setWindowTitle("About")
        lay=QVBoxLayout(dlg)
        lbl=QLabel(dlg); lbl.setTextFormat(Qt.RichText)
        lbl.setText(
            "<div style='color:#000; font-size:12pt;'><b>Multi-Pane File Explorer</b></div>"
            "<div style='color:#111; margin-top:6px;'>A compact multi-pane file explorer for Windows (PyQt5).</div>"
            "<div style='color:#111; margin-top:6px;'>For feedback, contact <b>kkongt2.kang</b>.</div>"
            "<div style='color:#333; margin-top:6px; font-size:10pt;'>© 2025</div>"
        )
        lay.addWidget(lbl)
        btns=QDialogButtonBox(QDialogButtonBox.Ok, dlg); lay.addWidget(btns); btns.accepted.connect(dlg.accept)
        pal=dlg.palette(); pal.setColor(QPalette.Window, QColor(255,255,255)); pal.setColor(QPalette.WindowText, QColor(0,0,0))
        dlg.setPalette(pal); dlg.setStyleSheet("QLabel { color: #000; } QDialog { background: #FFF; }")
        dlg.resize(380,180); dlg.exec_()

    def closeEvent(self,e):
        settings=QSettings(ORG_NAME, APP_NAME)
        settings.setValue("window/geometry", self.saveGeometry())
        settings.setValue("layout/pane_count", len(self.panes))
        for i,p in enumerate(self.panes): settings.setValue(f"layout/pane_{i}_path", p.current_path())
        settings.sync(); super().closeEvent(e)

# -------------------- Bookmark Editor --------------------
class BookmarkEditDialog(QDialog):
    def __init__(self, parent=None, items=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Bookmarks (max 10)")
        self.resize(760, 420)

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Enabled", "Name", "Path"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.table.setRowCount(10)
        self._rows = []

        lay = QVBoxLayout(self); lay.addWidget(self.table, 1)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        lay.addWidget(btns)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)

        items = list(items or [])
        for i in range(10):
            it = items[i] if i < len(items) else {"enabled": False, "name": "", "path": ""}
            self._add_row(i, it)

    def _add_row(self, row: int, data: dict):
        chk = QCheckBox(self.table); chk.setChecked(bool(data.get("enabled", False)))
        self.table.setCellWidget(row, 0, chk)

        name_edit = QLineEdit(self.table); name_edit.setText(str(data.get("name", "")))
        name_edit.setPlaceholderText("Bookmark name"); name_edit.setClearButtonEnabled(True); name_edit.setFixedHeight(UI_H)
        self.table.setCellWidget(row, 1, name_edit)

        path_wrap = QWidget(self.table); h = QHBoxLayout(path_wrap); h.setContentsMargins(0,0,0,0); h.setSpacing(ROW_SPACING)
        path_edit = QLineEdit(path_wrap); path_edit.setText(str(data.get("path", ""))); path_edit.setPlaceholderText("Folder path"); path_edit.setClearButtonEnabled(True); path_edit.setFixedHeight(UI_H)
        btn = QToolButton(path_wrap); btn.setText("…"); btn.setFixedHeight(UI_H)
        def browse():
            start = path_edit.text().strip() or QDir.homePath()
            d = QFileDialog.getExistingDirectory(self, "Select Folder", start)
            if d: path_edit.setText(d)
        btn.clicked.connect(browse)
        h.addWidget(path_edit, 1); h.addWidget(btn, 0)
        self.table.setCellWidget(row, 2, path_wrap)
        self._rows.append((chk, name_edit, path_edit))

    def values(self) -> list:
        out=[]
        for chk, name_edit, path_edit in self._rows:
            enabled = chk.isChecked()
            name = name_edit.text().strip()
            path = path_edit.text().strip()
            if name or path or enabled:
                out.append({"enabled": enabled, "name": name, "path": path})
        return out

    def set_items(self, items: list):
        items = list(items or [])
        for r in range(10):
            it = items[r] if r < len(items) else {"enabled": False, "name": "", "path": ""}
            chk, name_edit, path_edit = self._rows[r]
            chk.setChecked(bool(it.get("enabled", False)))
            name_edit.setText(str(it.get("name", "")))
            path_edit.setText(str(it.get("path", "")))
        self.table.resizeColumnsToContents()

# -------------------- Boot --------------------
def _load_start_paths(desired_panes:int, cli_paths):
    s=QSettings(ORG_NAME, APP_NAME); saved_n=s.value("layout/pane_count", type=int); paths=[]
    if not cli_paths and saved_n:
        for i in range(desired_panes):
            p=s.value(f"layout/pane_{i}_path", QDir.homePath(), type=str)
            paths.append(p if p and os.path.exists(p) else QDir.homePath())
    else:
        for i in range(desired_panes):
            if i<len(cli_paths) and os.path.exists(cli_paths[i]): paths.append(cli_paths[i])
            else:
                p=s.value(f"layout/pane_{i}_path", QDir.homePath(), type=str)
                paths.append(p if p and os.path.exists(p) else QDir.homePath())
    return paths

def parse_args():
    ap=argparse.ArgumentParser(description="Multi-Pane File Explorer (PyQt5)")
    ap.add_argument("paths", nargs="*", help="Optional start paths per pane")
    ap.add_argument("--panes", type=int, choices=[4,6,8], default=6, help="Number of panes: 4, 6 or 8")
    return ap.parse_args()

def main():
    args=parse_args()
    _enable_win_per_monitor_v2()
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app=QApplication(sys.argv)
    base_font=QFont("Segoe UI"); base_font.setPointSizeF(FONT_PT); app.setFont(base_font)
    try:
        if hasattr(QGuiApplication,"setHighDpiScaleFactorRoundingPolicy"):
            QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception: pass
    app.setOrganizationName(ORG_NAME); app.setApplicationName(APP_NAME)
    settings=QSettings(ORG_NAME, APP_NAME); theme=settings.value("ui/theme","dark")
    if theme not in ("dark","light"): theme="dark"
    apply_dark_style(app) if theme=="dark" else apply_light_style(app)
    start_paths=_load_start_paths(args.panes, args.paths)
    w=MultiExplorer(pane_count=args.panes, start_paths=start_paths, initial_theme=theme); w.show()
    sys.exit(app.exec_())

if __name__=="__main__":
    main()
