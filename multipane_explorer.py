

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


def _env_flag(name: str) -> bool:
    v = os.environ.get(name, "")
    return str(v).strip().lower() in {"1", "true", "yes", "on", "y"}

DEBUG = _env_flag("MULTIPANE_DEBUG")
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
SEARCH_RESULT_LIMIT = 50000
FILEOP_SIZE_SCAN_FILE_LIMIT = 6000
FILEOP_SIZE_SCAN_TIME_MS = 1200


HAS_PYWIN32 = True
try:
    import pythoncom
    import win32con, win32gui, win32api, win32clipboard
    from win32com.shell import shell, shellcon
except Exception:
    HAS_PYWIN32 = False


try:
    from send2trash import send2trash
    HAS_SEND2TRASH = True
except Exception:
    HAS_SEND2TRASH = False


def _enable_win_per_monitor_v2():
    if sys.platform != "win32": return
    # Windows only: try per-monitor DPI awareness with fallbacks.
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)); return
    except Exception: pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try: ctypes.windll.user32.SetProcessDPIAware()
        except Exception: pass

os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")


def _normalize_fs_path(p: str) -> str:
    try: p = os.path.normpath(p)
    except Exception: pass
    if os.name == "nt" and len(p) == 2 and p[1] == ":":
        # Normalize drive-root paths like "C:" to "C:\\".
        p = p + os.sep
    return p

def nice_path(p: str) -> str:
    try: return str(Path(p).resolve())
    except Exception: return _normalize_fs_path(p)

def _path_key(p: str) -> str:
    try:
        p = os.path.abspath(_normalize_fs_path(p))
    except Exception:
        p = _normalize_fs_path(p)
    return os.path.normcase(p)

def _paths_same(a: str, b: str) -> bool:
    try:
        return os.path.samefile(a, b)
    except Exception:
        return _path_key(a) == _path_key(b)

def _is_subpath(child: str, parent: str) -> bool:
    child_key = _path_key(child)
    parent_key = _path_key(parent)
    try:
        return os.path.commonpath([child_key, parent_key]) == parent_key
    except Exception:
        # Different drives on Windows can raise ValueError here.
        return False

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
    # Prefer trash providers; fall back to permanent delete.
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

def icon_bookmark_edit(theme: str):
    def paint(p: QPainter, w, h):
        p.setRenderHint(QPainter.Antialiasing, True)


        cx, cy = w/2 - 2, h/2 - 1
        r_outer = min(w, h) * 0.40
        r_inner = r_outer * 0.44
        pts = []
        for i in range(10):
            ang = -math.pi/2 + i * (math.pi / 5.0)
            r = r_outer if (i % 2 == 0) else r_inner
            pts.append(QtCore.QPointF(cx + math.cos(ang) * r, cy + math.sin(ang) * r))
        star = QPolygonF(pts)

        fill = QColor(255, 210, 60) if theme == "dark" else QColor(255, 190, 0)
        stroke = QColor(160, 120, 0) if theme == "dark" else QColor(150, 110, 0)
        p.setPen(QPen(stroke, 1.6))
        p.setBrush(QBrush(fill))
        p.drawPolygon(star)


        p.save()
        body = QColor(100, 180, 255) if theme == "dark" else QColor(40, 120, 220)
        tip  = QColor(240, 200, 80)


        p.translate(w * 0.64, h * 0.68)
        p.rotate(-25)


        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(body))
        p.drawRect(-5, -2, 12, 4)


        p.setBrush(QBrush(tip))
        tri = QPolygonF([
            QtCore.QPointF(7, -2),
            QtCore.QPointF(7,  2),
            QtCore.QPointF(10, 0)
        ])
        p.drawPolygon(tri)


        eraser = QColor(230, 230, 240) if theme == "dark" else QColor(250, 250, 255)
        p.setBrush(QBrush(eraser))
        p.drawRect(QtCore.QRectF(-6.5, -2.2, 2.6, 4.4))

        p.restore()

    return _make_icon(22, 22, paint)



class FileOpWorker(QtCore.QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished_ok = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, op: str, srcs: list, dst_dir: str, conflict_map: dict | None = None, parent=None):
        super().__init__(parent)
        self.op = op
        self.srcs = list(srcs)
        self.dst_dir = dst_dir
        self.conflict_map = dict(conflict_map or {})
        self._cancel = False
        self._total = 0
        self._done = 0
        self._count_progress = False
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._src_size_cache = {}

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
        scanned = 0
        total = 0
        self._src_size_cache = {}
        deadline = time.perf_counter() + (FILEOP_SIZE_SCAN_TIME_MS / 1000.0)
        for s in self.srcs:
            if self._cancel:
                break
            skey = _path_key(s)
            if os.path.isdir(s) and not os.path.islink(s):
                src_total = 0
                for _fp, sz in self._iter_files(s):
                    cur = max(0, int(sz or 0))
                    src_total += cur
                    total += cur
                    scanned += 1
                    if scanned >= FILEOP_SIZE_SCAN_FILE_LIMIT or time.perf_counter() >= deadline:
                        # Switch to count-based progress when size scan is too large/slow.
                        self._src_size_cache[skey] = src_total
                        self._count_progress = True
                        self._total = max(1, scanned, len(self.srcs))
                        self._done = 0
                        return
                self._src_size_cache[skey] = src_total
            else:
                src_total = 0
                try:
                    src_total = max(0, int(os.path.getsize(s)))
                    total += src_total
                except Exception:
                    pass
                self._src_size_cache[skey] = src_total
                scanned += 1
                if scanned >= FILEOP_SIZE_SCAN_FILE_LIMIT or time.perf_counter() >= deadline:
                    self._count_progress = True
                    self._total = max(1, scanned, len(self.srcs))
                    self._done = 0
                    return
        self._count_progress = False
        self._total = max(1, total)
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0

    def _emit_progress(self, force: bool = False):
        total = max(1, int(self._total or 1))
        pct = min(100, int(self._done * 100 / total))
        now = time.perf_counter()
        should_emit = (
            force
            or pct >= 100
            or self._last_progress_pct < 0
            or pct > self._last_progress_pct
            and (
                (pct - self._last_progress_pct) >= 1
                or (now - self._last_progress_emit_ts) >= 0.05
            )
        )
        if not should_emit:
            return
        self._last_progress_pct = pct
        self._last_progress_emit_ts = now
        self.progress.emit(pct)

    def _tick_progress(self, delta_bytes):
        if self._count_progress:
            return
        self._done += max(0, int(delta_bytes))
        self._emit_progress()

    def _tick_count_unit(self, units: int = 1):
        if not self._count_progress:
            return
        self._done += max(0, int(units))
        self._emit_progress()

    def _skip_source_progress(self, src):
        if self._count_progress:
            return
        try:
            delta = self._src_size_cache.get(_path_key(src))
        except Exception:
            delta = None
        if delta is None:
            delta = self._size_of(src)
        self._tick_progress(delta)

    def _emit_source_done(self):
        if self._count_progress:
            self._done += 1
            self._emit_progress()
            return
        self._emit_progress()

    def _copy_file(self, src, dst):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
            while True:
                if self._cancel: return
                buf = fsrc.read(1024 * 1024)
                if not buf: break
                fdst.write(buf)
                self._tick_progress(len(buf))
        self._tick_count_unit(1)
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
            if self._count_progress:
                self.status.emit(f"Preparing {self.op} (quick estimate) ...")
            else:
                self.status.emit(f"Preparing {self.op} ...")

            for src in self.srcs:
                if self._cancel: break
                if not os.path.exists(src):
                    self._emit_source_done()
                    continue

                base = os.path.basename(src.rstrip("\\/")) or os.path.basename(src)
                dst = os.path.join(self.dst_dir, base)


                if _paths_same(src, dst):
                    if self.op == "copy":
                        dst = unique_dest_path(self.dst_dir, base)
                    else:
                        self.status.emit(f"Skipped same path: {base}")
                        self._skip_source_progress(src)
                        self._emit_source_done()
                        continue


                if os.path.isdir(src) and not os.path.islink(src) and _is_subpath(dst, src):
                    # Prevent copying/moving a folder into its own subtree.
                    self.status.emit(f"Skipped nested destination: {base}")
                    self._skip_source_progress(src)
                    self._emit_source_done()
                    continue

                exists = os.path.exists(dst)
                action = self.conflict_map.get(src) if exists else None


                if self.op == "copy":
                    if os.path.isdir(src) and not os.path.islink(src):
                        if exists:
                            if action == "skip":
                                self._skip_source_progress(src); self._emit_source_done(); continue
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
                                self._skip_source_progress(src); self._emit_source_done(); continue
                            elif action == "copy":
                                dst = unique_dest_path(self.dst_dir, base)

                        self._copy_file(src, dst)


                else:
                    keep_both = (action == "copy")
                    src_progress_size = self._src_size_cache.get(_path_key(src))
                    if exists:
                        if action == "skip":
                            self._skip_source_progress(src); self._emit_source_done(); continue
                        elif keep_both:
                            dst = unique_dest_path(self.dst_dir, base)
                        elif action == "overwrite":
                            try:
                                if os.path.isdir(dst) and not os.path.islink(dst): shutil.rmtree(dst)
                                else: os.remove(dst)
                            except Exception: pass
                    try:
                        final = shutil.move(src, dst if keep_both else self.dst_dir)
                        if src_progress_size is None:
                            src_progress_size = self._size_of(final if os.path.exists(final) else src)
                        self._tick_progress(src_progress_size)
                    except Exception:
                        # Cross-device move or permission failures: fall back to copy.
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

                self._emit_source_done()

            if self._cancel:
                self.error.emit("Operation cancelled."); return
            self._done = max(self._done, self._total)
            self._emit_progress(force=True)
            self.finished_ok.emit()
        except Exception as e:
            self.error.emit(str(e))


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
    QToolButton#quickBookmarkBtn {{ text-align: left; }}

    QLabel#crumbSep {{ padding: 0 0px; margin: 0; }}
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

        /* Default crumb button */
        QPushButton#crumb {
            background: rgba(255,255,255,0.05);
            border: 1px solid #2B2E34;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #E6E9EE;
        }
        QPushButton#crumb:hover { background: rgba(255,255,255,0.09); }
        QLabel#crumbSep { color: #7F8796; }

        /* Remove border/background from breadcrumb scroll area */
        QScrollArea#crumbScroll { border: 0px solid transparent; }
        QScrollArea#crumbScroll[active="true"] { border: 0px solid transparent; }
        QScrollArea#crumbScroll > QWidget#crumbViewport { background: transparent; }

        /* Active pane: highlight only crumb buttons */
        QWidget#paneRoot[active="true"] QPushButton#crumb {
            background: rgba(94,155,255,0.16);
            border-color: rgba(94,155,255,0.40);
        }
        QWidget#paneRoot[active="true"] QPushButton#crumb:hover {
            background: rgba(94,155,255,0.22);
        }

        /* Pane-level active highlight */
        QWidget#paneRoot {
            border: 1px solid transparent;
            border-radius: 10px;
        }
        QWidget#paneRoot[active="true"] {
            border: 1px solid #5E9BFF;
            background: rgba(94, 155, 255, 0.06);
        }
        QWidget#paneRoot[active="true"] QTreeView { border-color: rgba(94,155,255,0.45); }
        QWidget#paneRoot[active="true"] QLineEdit { border: 1px solid rgba(94,155,255,0.35); }
        QWidget#paneRoot[active="true"] QLineEdit:focus { border: 1px solid #5E9BFF; }
        QWidget#paneRoot[active="true"] QToolButton, QWidget#paneRoot[active="true"] QPushButton {
            border-color: rgba(94,155,255,0.25);
        }

        /* Message boxes: white background, black text */
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

        /* Default crumb button */
        QPushButton#crumb {
            background: rgba(0,0,0,0.04);
            border: 1px solid #E5E8EE;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #1C1C1E;
        }
        QPushButton#crumb:hover { background: rgba(0,0,0,0.07); }
        QLabel#crumbSep { color: #7A7F89; }

        /* Remove border/background from breadcrumb scroll area */
        QScrollArea#crumbScroll { border: 0px solid transparent; }
        QScrollArea#crumbScroll[active="true"] { border: 0px solid transparent; }
        QScrollArea#crumbScroll > QWidget#crumbViewport { background: transparent; }

        /* Active pane: highlight only crumb buttons */
        QWidget#paneRoot[active="true"] QPushButton#crumb {
            background: rgba(64,128,255,0.12);
            border-color: rgba(64,128,255,0.40);
        }
        QWidget#paneRoot[active="true"] QPushButton#crumb:hover {
            background: rgba(64,128,255,0.18);
        }

        /* Pane-level active highlight */
        QWidget#paneRoot {
            border: 1px solid transparent;
            border-radius: 10px;
        }
        QWidget#paneRoot[active="true"] {
            border: 1px solid #5E9BFF;
            background: rgba(64, 128, 255, 0.06);
        }
        QWidget#paneRoot[active="true"] QTreeView { border-color: rgba(64,128,255,0.40); }
        QWidget#paneRoot[active="true"] QLineEdit { border: 1px solid rgba(64,128,255,0.35); }
        QWidget#paneRoot[active="true"] QLineEdit:focus { border: 1px solid #5E9BFF; }
        QWidget#paneRoot[active="true"] QToolButton, QWidget#paneRoot[active="true"] QPushButton {
            border-color: rgba(64,128,255,0.25);
        }
    """)



def icon_copy_squares(theme: str):
    def paint(p: QPainter, w, h):

        stroke = QColor(210, 214, 225) if theme == "dark" else QColor(85, 95, 115)
        fill   = QColor(255, 255, 255)

        p.setRenderHint(QPainter.Antialiasing, True)
        pen = QPen(stroke, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        p.setPen(pen)
        p.setBrush(QBrush(fill))


        front_rect = QtCore.QRect(6, 3, 11, 11)
        back_rect  = QtCore.QRect(3, 6, 11, 11)
        radius = 3


        p.drawRoundedRect(front_rect, radius, radius)


        p.drawRoundedRect(back_rect, radius, radius)

    return _make_icon(20, 20, paint)


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

def icon_session(theme: str):
    def paint(p: QPainter, w, h):
        p.setRenderHint(QPainter.Antialiasing, True)


        line = QColor(190, 195, 210) if theme == "dark" else QColor(90, 100, 120)
        fill = QColor(60, 66, 80) if theme == "dark" else QColor(245, 247, 250)
        tab  = QColor(100, 150, 255) if theme == "dark" else QColor(80, 120, 230)
        star_fill  = QColor(255, 210, 60)
        star_edge  = QColor(160, 120, 0)


        p.setPen(QPen(line, 1.3))
        p.setBrush(QBrush(fill))
        p.drawRoundedRect(QtCore.QRectF(3.0, 6.5, 12.5, 9.0), 2.5, 2.5)
        p.drawRoundedRect(QtCore.QRectF(5.0, 5.0, 12.5, 9.0), 2.5, 2.5)
        p.setBrush(QBrush(tab))
        p.drawRoundedRect(QtCore.QRectF(7.0, 3.5, 12.5, 9.0), 2.5, 2.5)


        cx, cy, r = w - 6.0, h - 6.0, 3.2
        pts = []
        for i in range(10):
            ang = -math.pi/2 + i * (math.pi/5.0)
            rad = r if (i % 2 == 0) else r * 0.44
            pts.append(QtCore.QPointF(cx + math.cos(ang)*rad, cy + math.sin(ang)*rad))
        p.setPen(QPen(star_edge, 1.0))
        p.setBrush(QBrush(star_fill))
        p.drawPolygon(QPolygonF(pts))

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
    if isinstance(obj, (list, tuple)):


        cand = None
        for v in obj:
            if isinstance(v, int):
                continue
            cand = v
            break
        obj = cand
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

def _query_ctx_menu_id_last(cm, hmenu, id_first: int, flags: int, id_limit: int = 0x7FFF) -> int:
    qret = cm.QueryContextMenu(hmenu, 0, int(id_first), int(id_limit), int(flags))
    count = int(qret) & 0xFFFF
    return int(id_first) + int(count) - 1 if count > 0 else int(id_first) - 1

def _menu_item_text(hmenu, cmd_id: int) -> str:
    try:
        txt = win32gui.GetMenuString(hmenu, int(cmd_id), win32con.MF_BYCOMMAND)
        return (txt or "").strip().lower()
    except Exception:
        return ""

def _context_target_dir(work_dir: str, paths=None) -> str:
    target = work_dir
    if paths:
        try:
            p0 = str(paths[0])
            target = p0 if os.path.isdir(p0) else (os.path.dirname(p0) or work_dir)
        except Exception:
            pass
    target = _normalize_fs_path(target or os.getcwd())
    try:
        target = os.path.abspath(target)
    except Exception:
        pass
    return target

def _launch_powershell_here(owner_hwnd, work_dir: str, paths=None) -> bool:
    target = _context_target_dir(work_dir, paths=paths)


    ps_literal = target.replace("'", "''")
    args = f"-NoExit -Command Set-Location -LiteralPath '{ps_literal}'"
    try:
        win32api.ShellExecute(int(owner_hwnd) if owner_hwnd else 0,
                              "open",
                              "powershell.exe",
                              args,
                              target,
                              win32con.SW_SHOWNORMAL)
        return True
    except Exception as e:
        if DEBUG: print("[ctx] direct PowerShell launch failed:", e)
        return False

def _is_git_bash_action(verb: str | None, menu_text: str | None) -> bool:
    v = (verb or "").strip().lower()
    t = (menu_text or "").strip().lower()
    if not v and not t:
        return False
    if ("git bash" in v) or ("git bash" in t):
        return True
    if ("git-bash" in v) or ("git_bash" in v) or ("gitbash" in v):
        return True
    if "git_shell" in v:
        return True
    return ("git" in v and "bash" in v) or ("git" in t and "bash" in t)

def _first_existing_path(candidates: list[str]) -> str | None:
    seen = set()
    for p in candidates:
        if not p:
            continue
        try:
            pp = os.path.abspath(_normalize_fs_path(str(p)))
        except Exception:
            pp = _normalize_fs_path(str(p))
        key = os.path.normcase(pp)
        if key in seen:
            continue
        seen.add(key)
        if os.path.isfile(pp):
            return pp
    return None

def _discover_git_for_windows_tools() -> tuple[str | None, str | None]:
    git_bash_candidates = []
    bash_candidates = []

    wb = shutil.which("git-bash.exe")
    if wb:
        git_bash_candidates.append(wb)

    wgit = shutil.which("git.exe")
    if wgit:
        try:
            gdir = os.path.dirname(os.path.abspath(wgit))
            roots = [os.path.dirname(gdir), os.path.dirname(os.path.dirname(gdir))]
            for root in roots:
                if not root:
                    continue
                git_bash_candidates.append(os.path.join(root, "git-bash.exe"))
                bash_candidates.append(os.path.join(root, "bin", "bash.exe"))
                bash_candidates.append(os.path.join(root, "usr", "bin", "bash.exe"))
        except Exception:
            pass

    for env_name in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_name)
        if not base:
            continue
        root = os.path.join(base, "Git")
        git_bash_candidates.append(os.path.join(root, "git-bash.exe"))
        bash_candidates.append(os.path.join(root, "bin", "bash.exe"))
        bash_candidates.append(os.path.join(root, "usr", "bin", "bash.exe"))

    lad = os.environ.get("LocalAppData")
    if lad:
        for root in (os.path.join(lad, "Programs", "Git"), os.path.join(lad, "Git")):
            git_bash_candidates.append(os.path.join(root, "git-bash.exe"))
            bash_candidates.append(os.path.join(root, "bin", "bash.exe"))
            bash_candidates.append(os.path.join(root, "usr", "bin", "bash.exe"))

    wbash = shutil.which("bash.exe")
    if wbash:
        bash_candidates.append(wbash)

    git_bash = _first_existing_path(git_bash_candidates)
    bash_exe = _first_existing_path(bash_candidates)
    return git_bash, bash_exe

def _notify_git_bash_not_found():
    msg = "Git Bash executable was not found. Install Git for Windows or add it to PATH."
    try:
        QTimer.singleShot(0, lambda m=msg: QMessageBox.warning(None, "Git Bash", m))
    except Exception:
        pass

def _launch_git_bash_here(owner_hwnd, work_dir: str, paths=None) -> bool:
    target = _context_target_dir(work_dir, paths=paths)
    git_bash, bash_exe = _discover_git_for_windows_tools()

    if git_bash:
        args = f'--cd="{target}"'
        try:
            win32api.ShellExecute(int(owner_hwnd) if owner_hwnd else 0,
                                  "open",
                                  git_bash,
                                  args,
                                  target,
                                  win32con.SW_SHOWNORMAL)
            return True
        except Exception as e:
            if DEBUG: print("[ctx] direct Git Bash launch failed:", e)
        try:
            subprocess.Popen([git_bash], cwd=target)
            return True
        except Exception as e:
            if DEBUG: print("[ctx] Git Bash fallback launch failed:", e)

    if bash_exe:
        try:
            flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
            subprocess.Popen([bash_exe, "--login", "-i"], cwd=target, creationflags=flags)
            return True
        except Exception as e:
            if DEBUG: print("[ctx] bash.exe launch failed:", e)

    _notify_git_bash_not_found()
    return False

def _invoke_menu(owner_hwnd, cm, hmenu, screen_pt, work_dir, paths=None, id_first=1, id_last=None):
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
    if not cmd_id:
        return True


    try:
        state = win32gui.GetMenuState(hmenu, cmd_id, win32con.MF_BYCOMMAND)
        if state & win32con.MF_POPUP:
            if DEBUG: print("[ctx] popup header selected; will fallback")
            return False
    except Exception:
        pass

    if id_last is not None and (int(cmd_id) < int(id_first) or int(cmd_id) > int(id_last)):
        if DEBUG: print(f"[ctx] cmd_id={cmd_id} out of range [{id_first}, {id_last}]")
        return False

    idx=int(cmd_id)-int(id_first); verb=None
    try:
        verb = _get_canonical_verb(cm, idx) or None
    except Exception: verb=None
    menu_text = _menu_item_text(hmenu, cmd_id)
    if DEBUG: print(f"[ctx] chosen cmd_id={cmd_id} id_first={id_first} -> idx={idx}, verb='{verb or ''}', text='{menu_text}'")



    if (verb and "powershell" in verb.lower()) or ("powershell" in menu_text):
        if _launch_powershell_here(owner_hwnd, work_dir, paths=paths):
            _post_null(owner_hwnd)
            return True



    if _is_git_bash_action(verb, menu_text):
        _launch_git_bash_here(owner_hwnd, work_dir, paths=paths)
        _post_null(owner_hwnd)
        return True


    if verb and verb.lower() in ("properties","prop","property"):
        try:
            target = paths[0] if (paths and len(paths)>0) else work_dir
            target = _normalize_fs_path(target)

            ok = False

            try:
                fn = getattr(shell, "SHObjectProperties", None)
                if callable(fn):

                    fn(int(owner_hwnd), 0x00000002, target, None)
                    ok = True
            except Exception as e:
                if DEBUG: print("[ctx] SHObjectProperties (pywin32) failed:", e)


            if not ok:
                try:
                    shell.ShellExecuteEx(
                        hwnd=int(owner_hwnd),
                        fMask=shellcon.SEE_MASK_INVOKEIDLIST,
                        lpVerb='properties',
                        lpFile=target,
                        nShow=win32con.SW_SHOW
                    )
                    ok = True
                except Exception as e:
                    if DEBUG: print("[ctx] ShellExecuteEx(properties) failed:", e)


            if not ok:
                try:
                    SHOP_FILEPATH = 0x00000002
                    shell32 = ctypes.windll.shell32
                    res = shell32.SHObjectProperties(ctypes.c_void_p(int(owner_hwnd)),
                                                     ctypes.c_uint(SHOP_FILEPATH),
                                                     ctypes.c_wchar_p(target),
                                                     ctypes.c_wchar_p(None))
                    ok = bool(res)
                except Exception as e:
                    if DEBUG: print("[ctx] SHObjectProperties(ctypes) failed:", e)

            if ok:
                _post_null(owner_hwnd)
                return True
        except Exception as e:
            if DEBUG: print("[ctx] open properties failed:", e)
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

        norm_paths = []
        seen = set()
        for p in paths:
            if not p:
                continue
            np = _normalize_fs_path(p)
            key = os.path.normcase(os.path.normpath(np))
            if key in seen:
                continue
            seen.add(key)
            norm_paths.append(np)
        if not norm_paths:
            return False

        parent_dir = _normalize_fs_path(os.path.dirname(norm_paths[0]) or os.getcwd())
        app=QApplication.instance(); evf=_ensure_event_filter(app)

        cm=_icm_via_shellitems(norm_paths)
        if cm:
            evf.set_context(cm); hMenu=win32gui.CreatePopupMenu()
            flags=shellcon.CMF_NORMAL|shellcon.CMF_EXPLORE|shellcon.CMF_INCLUDESTATIC
            if win32api.GetKeyState(win32con.VK_SHIFT)<0: flags|=shellcon.CMF_EXTENDEDVERBS
            id_first=1
            try:
                id_last = _query_ctx_menu_id_last(cm, hMenu, id_first, flags)
                ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,parent_dir,paths=norm_paths,id_first=id_first,id_last=id_last)
                evf.clear()
                if ok: return True
            except Exception as e:
                if DEBUG: print("[ctx] ShellItems QueryContextMenu failed:", e)
                evf.clear()

        try:
            desktop=shell.SHGetDesktopFolder()
            abs_pidls=tuple(_abs_pidl(p) for p in norm_paths)
            cm=desktop.GetUIObjectOf(0,abs_pidls,shell.IID_IContextMenu,0); cm=_as_interface(cm)
        except Exception as e:
            if DEBUG: print("[ctx] desktop GetUIObjectOf failed:", e); cm=None
        if not cm: return False

        evf.set_context(cm); hMenu=win32gui.CreatePopupMenu()
        flags=shellcon.CMF_NORMAL|shellcon.CMF_EXPLORE|shellcon.CMF_INCLUDESTATIC
        if win32api.GetKeyState(win32con.VK_SHIFT)<0: flags|=shellcon.CMF_EXTENDEDVERBS
        id_first=1
        id_last = _query_ctx_menu_id_last(cm, hMenu, id_first, flags)
        ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,parent_dir,paths=norm_paths,id_first=id_first,id_last=id_last)
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
        id_first=1
        id_last = _query_ctx_menu_id_last(cm, hMenu, id_first, flags)
        ok = _invoke_menu(owner_hwnd,cm,hMenu,screen_pt,folder_path,paths=[folder_path],id_first=id_first,id_last=id_last)
        evf.clear(); return ok
    finally:
        pythoncom.CoUninitialize()


try:
    import winreg
    HAS_WINREG = True
except Exception:
    HAS_WINREG = False

def _shellnew_template_for_ext(ext_with_dot: str) -> tuple[bool, str | None]:
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


IS_DIR_ROLE = Qt.UserRole + 99
SIZE_BYTES_ROLE = Qt.UserRole + 100
SEARCH_ICON_READY_ROLE = Qt.UserRole + 101
NAME_FOLD_ROLE = Qt.UserRole + 102

class FsSortProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDynamicSortFilter(True)
        self.setSortCaseSensitivity(Qt.CaseInsensitive)
        self.setSortRole(Qt.EditRole)
        self.setSortLocaleAware(False)
    def _same_model(self, a, b) -> bool:
        if a is None or b is None:
            return False
        if a is b:
            return True
        try:
            return bool(a == b)
        except Exception:
            return False
    def mapToSource(self, proxyIndex):
        if not proxyIndex.isValid():
            return QtCore.QModelIndex()
        try:
            if not self._same_model(proxyIndex.model(), self):
                return QtCore.QModelIndex()
        except Exception:
            return QtCore.QModelIndex()
        return super().mapToSource(proxyIndex)
    def mapFromSource(self, sourceIndex):
        if not sourceIndex.isValid():
            return QtCore.QModelIndex()
        src = self.sourceModel()
        if src is None:
            return QtCore.QModelIndex()
        try:
            if not self._same_model(sourceIndex.model(), src):
                return QtCore.QModelIndex()
        except Exception:
            return QtCore.QModelIndex()
        return super().mapFromSource(sourceIndex)
    def filterAcceptsRow(self, source_row, source_parent): return True
    def sort(self, column, order=Qt.AscendingOrder):

        self._sort_order = order
        super().sort(column, order)
    def lessThan(self, left, right):
        col = left.column(); src = self.sourceModel()


        try:
            ldir = bool(src.isDir(left)) if hasattr(src, "isDir") else bool(src.data(left, IS_DIR_ROLE))
            rdir = bool(src.isDir(right)) if hasattr(src, "isDir") else bool(src.data(right, IS_DIR_ROLE))
            if ldir != rdir:
                order = getattr(self, "_sort_order", Qt.AscendingOrder)
                if order == Qt.AscendingOrder:

                    return ldir and not rdir
                else:

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

        if col == 0:
            try:
                lv = src.data(left, NAME_FOLD_ROLE)
                rv = src.data(right, NAME_FOLD_ROLE)
                if lv is not None and rv is not None:
                    return str(lv) < str(rv)
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
            if c==0:
                return r["name"]
            if c==1:

                if r.get("is_dir", False): return ""
                if r["size"] is None: return ""
                return human_size(int(r["size"]))
            if c==2:
                return r["type"]
            if c==3:
                if r["mtime"] is None: return ""
                dt=QDateTime.fromSecsSinceEpoch(int(r["mtime"])); return dt.toString("yyyy-MM-dd HH:mm:ss")

        elif role==Qt.EditRole:
            if c==0: return r["name"]
            if c==1:

                if r.get("is_dir", False): return 0
                return 0 if r["size"] is None else int(r["size"])
            if c==3:
                return QDateTime.fromSecsSinceEpoch(int(r["mtime"])) if r["mtime"] else QDateTime()
            else:
                return r["type"]

        elif role==Qt.ToolTipRole:
            return r["path"]
        elif role==Qt.UserRole:
            return r["path"]
        elif role==IS_DIR_ROLE:
            return r["is_dir"]
        elif role==SIZE_BYTES_ROLE:

            if r.get("is_dir", False): return 0
            return 0 if r["size"] is None else int(r["size"])
        elif role==NAME_FOLD_ROLE and c==0:
            return r.get("name_l") or str(r.get("name", "")).lower()

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
    def __init__(self, root:str, parent=None): super().__init__(parent); self.root=root; self._cancel=False
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
                    batch.append({
                        "name": name,
                        "name_l": name.lower(),
                        "path": p,
                        "is_dir": is_dir,
                        "size": None,
                        "mtime": None,
                        "type": typ,
                    })
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
    batchReady = pyqtSignal(str, list)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    truncated = pyqtSignal(int)

    def __init__(self, base_path: str, pattern_str: str, parent=None, max_results: int = SEARCH_RESULT_LIMIT):
        super().__init__(parent)
        self.base = base_path
        self._cancel = False
        self._max_results = max(1, int(max_results))
        self._matches = 0
        self._truncated = False

        raw = (pattern_str or "").replace(",", " ").replace(";", " ").split()
        self._patterns = [p.lower() for p in raw] if raw else ["*"]

        self._tests = []
        for p in self._patterns:
            simple_ext = (p.startswith("*.") and ("*" not in p[2:]) and ("?" not in p) and ("[" not in p) and ("]" not in p))
            if simple_ext:
                ext = p[1:]
                self._tests.append(lambda name, ext=ext: name.endswith(ext))
            else:

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
                                if self._matches >= self._max_results:
                                    self._truncated = True
                                    self._cancel = True
                                    break
                                self._matches += 1
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


                            if is_dir:
                                try:
                                    if entry.is_symlink():
                                        continue
                                except Exception:
                                    pass
                                stack.append(entry.path)
                except Exception:

                    continue

            if batch:
                self.batchReady.emit(base, batch)
            if self._truncated:
                self.truncated.emit(self._matches)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()

class StatOverlayProxy(QIdentityProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._cache = {}
        self._pending = set()
        self._queue = []
        self._worker = None
        self._batch_limit = 256

    def filePath(self, index):
        src = self.sourceModel()
        s = self.mapToSource(index)
        return src.filePath(s)

    def isDir(self, index):
        src = self.sourceModel()
        s = self.mapToSource(index)
        try:
            return src.isDir(s)
        except Exception:
            try:
                return os.path.isdir(src.filePath(s))
            except Exception:
                return False

    def clear_cache(self):
        self._cache.clear()
        self._pending.clear()
        self._queue.clear()
        self._cancel_worker()

    def _cancel_worker(self):
        w = self._worker
        if w and w.isRunning():
            w.cancel()
            w.wait(100)
        self._worker = None

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        col = index.column()
        if col not in (1, 3):
            return super().data(index, role)

        src = self.sourceModel()
        sidx = self.mapToSource(index)

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

        rec = self._cache.get(p) if p else None

        if col == 1:
            if is_dir:
                if role == SIZE_BYTES_ROLE:
                    return 0
                if role == Qt.EditRole:
                    return 0
                if role == Qt.DisplayRole:
                    return ""
                return super().data(index, role)

            if rec is not None:
                size_val = int(rec[0] or 0)
                if role == SIZE_BYTES_ROLE:
                    return size_val
                if role == Qt.EditRole:
                    return size_val
                if role == Qt.DisplayRole:
                    return human_size(size_val)


            if role in (SIZE_BYTES_ROLE, Qt.EditRole):
                return 0
            if role == Qt.DisplayRole:
                return ""
            return super().data(index, role)

        if col == 3:
            if rec and rec[1] is not None:
                dt = QDateTime.fromSecsSinceEpoch(int(rec[1]))
                if role == Qt.DisplayRole:
                    return dt.toString("yyyy-MM-dd HH:mm:ss")
                if role == Qt.EditRole:
                    return dt


            if role == Qt.DisplayRole:
                return ""
            if role == Qt.EditRole:
                return QDateTime()
            return super().data(index, role)

    def request_paths(self, paths: list[str], batch_limit: int = 256):
        try:
            self._batch_limit = max(1, int(batch_limit))
        except Exception:
            self._batch_limit = 256

        added = False
        for p in paths:
            if not p:
                continue
            if p in self._cache or p in self._pending:
                continue
            self._pending.add(p)
            self._queue.append(p)
            added = True

        if added:
            self._start_next_batch()

    def _start_next_batch(self):
        if self._worker and self._worker.isRunning():
            return
        if not self._queue:
            self._worker = None
            return

        batch_size = max(1, int(self._batch_limit))
        batch = self._queue[:batch_size]
        del self._queue[:batch_size]

        w = NormalStatWorker(batch, self)
        w.statReady.connect(self._apply_stat, Qt.QueuedConnection)
        w.finishedCycle.connect(lambda b=batch: self._on_cycle_finished(b), Qt.QueuedConnection)
        self._worker = w
        w.start()

    @QtCore.pyqtSlot(str, object, object)
    def _apply_stat(self, path: str, size_val, mtime_val):
        self._cache[path] = (int(size_val or 0), float(mtime_val) if mtime_val is not None else None)
        try:
            src = self.sourceModel()
            sidx0 = src.index(path)
            if sidx0.isValid():
                for col in (1, 3):
                    sidx = sidx0.sibling(sidx0.row(), col)
                    pidx = self.mapFromSource(sidx)
                    self.dataChanged.emit(pidx, pidx, [Qt.DisplayRole, Qt.EditRole, SIZE_BYTES_ROLE])
        except Exception:
            pass

    def _on_cycle_finished(self, batch):
        for p in batch:
            self._pending.discard(p)
        self._worker = None
        self._start_next_batch()

class PathBar(QWidget):
    pathSubmitted=pyqtSignal(str)
    def __init__(self, parent=None):
        super().__init__(parent); self._current_path=QDir.homePath()
        self.setObjectName("pathbar")

        self._host=QWidget(); self._hlay=QHBoxLayout(self._host)
        self._host.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self._host.setMinimumHeight(UI_H)
        self._hlay.setContentsMargins(4,0,4,0); self._hlay.setSpacing(max(0, ROW_SPACING-2))

        self._scroll=QScrollArea(self); self._scroll.setObjectName("crumbScroll")
        self._scroll.setWidget(self._host)
        self._scroll.setWidgetResizable(False); self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff); self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setViewportMargins(0,0,0,0); self._scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._scroll.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self._hbar = self._scroll.horizontalScrollBar()
        self._scroll.setFixedHeight(UI_H); self.setFixedHeight(UI_H)


        try:
            vp = self._scroll.viewport()
            vp.setObjectName("crumbViewport")
            vp.setAttribute(Qt.WA_StyledBackground, True)
            vp.installEventFilter(self)
        except Exception:
            pass


        self._scroll.setProperty("active", False)

        self._edit=QLineEdit(self); self._edit.hide(); self._edit.setClearButtonEnabled(True); self._edit.setFixedHeight(UI_H)
        self._edit.returnPressed.connect(self._on_edit_return)


        self._btn_copy = QToolButton(self)
        self._btn_copy.setToolTip("Copy current path")
        self._btn_copy.setFixedHeight(UI_H)

        theme = getattr(getattr(parent, "host", None), "theme", "dark")
        try:
            self._btn_copy.setIcon(icon_copy_squares(theme))
        except Exception:

            self._btn_copy.setText("Copy")
        self._btn_copy.clicked.connect(self._copy_current_path)


        wrap=QHBoxLayout(self); wrap.setContentsMargins(0,0,0,0); wrap.setSpacing(0)
        wrap.addWidget(self._scroll, 1)
        wrap.addWidget(self._edit, 1)
        wrap.addWidget(self._btn_copy, 0)

        self._host.installEventFilter(self); self._edit.installEventFilter(self)
        self.set_path(self._current_path)

    def _copy_current_path(self):

        t = self._edit.text().strip() if self._edit.isVisible() else self._current_path
        if not t:
            t = self._current_path
        QApplication.clipboard().setText(t)
        QToolTip.showText(QCursor.pos(), f"Copied: {t}", self)


    def set_active(self, active: bool):
        try:
            self._scroll.setProperty("active", bool(active))

            vp = self._scroll.viewport()
            for w in (self._scroll, vp):
                w.style().unpolish(w)
                w.style().polish(w)
                w.update()
        except Exception:
            pass

    def sizeHint(self): return QSize(200, UI_H)
    def minimumSizeHint(self): return QSize(100, UI_H)
    def eventFilter(self, obj, ev):
        if obj is self._host and ev.type()==QEvent.MouseButtonDblClick: self.start_edit(); return True
        if obj is self._scroll.viewport() and ev.type()==QEvent.MouseButtonDblClick:
            try:
                if self._edit.isVisible():
                    return True
                vp = self._scroll.viewport()
                if vp is None:
                    return False
                content_w = max(self._host.sizeHint().width(), self._host.width())
                pos = ev.pos()
                # Enter edit mode only when user double-clicks the right-side blank area.
                if content_w < vp.width() and pos.x() >= content_w:
                    self.start_edit()
                    return True
            except Exception:
                return False
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
            if w:
                w.setParent(None)
                w.deleteLater()
        p=self._current_path; parts=[]
        p_unc = p.replace("/", "\\")
        if p_unc.startswith("\\\\"):
            comps=[c for c in p_unc.split("\\") if c]
            if len(comps)>=2:
                server = f"\\\\{comps[0]}"
                share = comps[1]
                server_root = server + "\\"
                share_root = server_root + share + "\\"
                server_target = server_root if os.path.exists(server_root) else share_root
                parts.append((server, server_target))
                parts.append((share, share_root))
                acc = share_root.rstrip("\\")
                for c in comps[2:]:
                    acc=os.path.join(acc,c); parts.append((c,acc))
            elif len(comps)==1:
                server = f"\\\\{comps[0]}"
                parts.append((server, server + "\\"))
            else:
                parts.append((p,p))
        else:
            drive,_=os.path.splitdrive(p); root=(drive+os.sep) if drive else os.sep
            parts.append((root,root)); sub=p[len(root):].strip("\\/")
            for seg in [s for s in sub.split(os.sep) if s]:
                curr=os.path.join(parts[-1][1], seg); parts.append((seg,curr))
        fm=self.fontMetrics()
        for i,(label,target) in enumerate(parts):
            btn=QPushButton(self._host); btn.setObjectName("crumb"); btn.setFlat(True); btn.setCursor(Qt.PointingHandCursor)
            elided=fm.elidedText(label, Qt.ElideMiddle, CRUMB_MAX_SEG_W)
            btn.setText(elided); btn.setToolTip(label); btn.setMinimumHeight(UI_H)
            btn.clicked.connect(lambda _,t=target: self.pathSubmitted.emit(t))
            self._hlay.addWidget(btn)
            if i < len(parts)-1:
                s=QLabel(">", self._host); s.setObjectName("crumbSep"); s.setContentsMargins(0,0,0,0); self._hlay.addWidget(s)
        self._hlay.activate()
        m = self._hlay.contentsMargins()
        item_w = 0
        item_n = 0
        for i in range(self._hlay.count()):
            it = self._hlay.itemAt(i)
            w = it.widget() if it else None
            if w is None:
                continue
            item_w += max(0, w.sizeHint().width())
            item_n += 1
        total_w = m.left() + m.right() + item_w + (max(0, item_n - 1) * self._hlay.spacing())
        total_h = max(UI_H, self._hlay.sizeHint().height(), 1)
        self._host.setFixedSize(max(1, total_w), total_h)
        self._host.updateGeometry()


        self._pin_to_right()
        QTimer.singleShot(0, self._pin_to_right)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)

        QTimer.singleShot(0, self._pin_to_right)

    def _pin_to_right(self):
        try:
            if not (hasattr(self, "_hbar") and self._hbar):
                return
            vp = self._scroll.viewport()
            if vp is None:
                return
            viewport_w = vp.width()
            if viewport_w <= 0:
                return
            content_w = max(self._host.sizeHint().width(), self._host.width())
            if content_w > (viewport_w + 1):

                self._hbar.setValue(self._hbar.maximum())
            else:

                self._hbar.setValue(self._hbar.minimum())
        except Exception:
            pass


class SearchResultModel(QStandardItemModel):
    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and index.column() == 1:
            b = super().data(index, SIZE_BYTES_ROLE)
            if b is None:
                b = super().data(index, Qt.EditRole)
            if isinstance(b, (int, float)) and b:
                return human_size(int(b))
            return ""
        return super().data(index, role)


    def mimeTypes(self):

        return ["text/uri-list"]

    def mimeData(self, indexes):
        md = QtCore.QMimeData()

        rows = set()
        for ix in indexes:
            if ix.isValid():
                rows.add(ix.row())

        from PyQt5.QtCore import QUrl
        urls = []
        for r in rows:
            it = self.item(r, 0)
            if not it:
                continue
            path = it.data(Qt.UserRole)
            if path:
                urls.append(QUrl.fromLocalFile(path))

        md.setUrls(urls)
        return md

    def flags(self, index):

        f = super().flags(index)
        if index.isValid():
            f |= Qt.ItemIsDragEnabled
        return f

    def supportedDragActions(self):

        return Qt.CopyAction

    def startDrag(self, supportedActions):

        from PyQt5.QtGui import QDrag
        md = QtCore.QMimeData()


        paths = [p for p in self.pane._selected_paths() if p and os.path.exists(p)]
        if not paths:
            return


        md.setUrls([QUrl.fromLocalFile(p) for p in paths])
        md.setText("\r\n".join(paths))

        drag = QDrag(self)
        drag.setMimeData(md)


        drag.exec_(Qt.CopyAction | Qt.MoveAction, Qt.CopyAction)



class ConflictResolutionDialog(QDialog):
    def __init__(self, parent, conflicts:list[tuple[str,str]], dst_dir:str):
        super().__init__(parent)
        self.setWindowTitle("Resolve name conflicts")
        self.resize(720, 420)
        self._conflicts = conflicts
        self._dst_dir = dst_dir

        lay = QVBoxLayout(self)


        top = QHBoxLayout()
        lbl = QLabel("Apply to all:", self)
        btn_over = QPushButton("Overwrite All", self)
        btn_skip = QPushButton("Skip All", self)
        btn_copy = QPushButton("Copy All", self)
        top.addStretch(1)
        top.addWidget(lbl)
        top.addSpacing(8)
        top.addWidget(btn_over)
        top.addWidget(btn_skip)
        top.addWidget(btn_copy)
        lay.addLayout(top)


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


        btn_over.clicked.connect(lambda: self._apply_all("Overwrite"))
        btn_skip.clicked.connect(lambda: self._apply_all("Skip"))
        btn_copy.clicked.connect(lambda: self._apply_all("Copy"))


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
        out = {}
        for (src, _dst), combo in zip(self._conflicts, self._combos):
            choice = combo.currentText().strip().lower()
            if choice not in ("overwrite","skip","copy"): choice="overwrite"
            out[src] = choice
        return out


class ExplorerView(QTreeView):
    def __init__(self, pane):
        super().__init__(pane); self.pane=pane
        self.setDragEnabled(True); self.setAcceptDrops(True)
        self.setDropIndicatorShown(True); self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self._drag_start_pos = None
        self._drag_start_index = QtCore.QModelIndex()
        self._drag_start_modifiers = Qt.NoModifier
        self._drag_start_was_selected = False
        self._drag_ready = False

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


    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F5:
            try:
                self.pane.hard_refresh()
            finally:
                e.accept()
            return
        super().keyPressEvent(e)


    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_start_pos = e.pos()
            self._drag_start_index = self.indexAt(e.pos())
            self._drag_start_modifiers = e.modifiers()
            sm = self.selectionModel()
            self._drag_start_was_selected = bool(self._drag_start_index.isValid() and sm and sm.isSelected(self._drag_start_index))
            self._drag_ready = False


        try:
            if not self.hasFocus():
                self.setFocus(Qt.MouseFocusReason)
        except Exception:
            pass


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

    def mouseMoveEvent(self, e):
        if (e.buttons() & Qt.LeftButton) and self._drag_start_pos is not None:

            if (e.pos() - self._drag_start_pos).manhattanLength() >= QApplication.startDragDistance():
                if not self._drag_ready:
                    ix = self._drag_start_index
                    sm = self.selectionModel()



                    if (ix.isValid() and sm
                        and bool(self._drag_start_modifiers & Qt.ControlModifier)
                        and self._drag_start_was_selected
                        and (not sm.isSelected(ix))):
                        try:
                            sm.select(ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                        except Exception:
                            pass
                    if ix.isValid() and sm and sm.isSelected(ix):


                        has_shift = bool(self._drag_start_modifiers & Qt.ShiftModifier)
                        has_ctrl = bool(self._drag_start_modifiers & Qt.ControlModifier)
                        allow_with_ctrl = (not has_ctrl) or self._drag_start_was_selected
                        self._drag_ready = (not has_shift) and allow_with_ctrl
                if self._drag_ready:
                    self._clear_drag_state()
                    self.startDrag(Qt.CopyAction | Qt.MoveAction)
                    return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        self._clear_drag_state()
        super().mouseReleaseEvent(e)

    def _clear_drag_state(self):
        self._drag_start_pos = None
        self._drag_start_index = QtCore.QModelIndex()
        self._drag_start_modifiers = Qt.NoModifier
        self._drag_start_was_selected = False
        self._drag_ready = False



class ExplorerPane(QWidget):
    requestBackgroundOp=pyqtSignal(str, list, str)
    _FALLBACK_NEW_ACTION_SPECS = (
        ("New Folder", "folder", "New Folder", None, "Folder created"),
        ("New Text File (.txt)", "file", "New Text Document.txt", ".txt", "Text file created"),
        ("New Word Document (.docx)", "file", "New Word Document.docx", ".docx", "Word document created"),
        ("New Excel Workbook (.xlsx)", "file", "New Excel Workbook.xlsx", ".xlsx", "Excel workbook created"),
        ("New PowerPoint Presentation (.pptx)", "file", "New PowerPoint Presentation.pptx", ".pptx", "PowerPoint presentation created"),
    )
    def __init__(self, _unused, start_path: str, pane_id: int, host_main):
        super().__init__()
        self.pane_id=pane_id; self.host=host_main
        self._init_state()

        row_toolbar = self._build_toolbar()
        row_path = self._build_path_row()
        row_filter = self._build_filter_row()

        self._init_models()
        self._setup_view()
        row_status = self._build_status_row()

        self._apply_layout(row_toolbar, row_path, row_filter, row_status)


        self._load_sort_settings()

        self.set_path(start_path or QDir.homePath(), push_history=False)
        self._update_star_button(); self._rebuild_quick_bookmark_buttons()

        self._connect_signals()
        self._register_shortcuts()


        self._apply_saved_sort()
        self._update_pane_status()

    def _init_state(self):
        self._search_mode=False; self._search_model=None; self._search_proxy=None
        self._search_pending_items={}; self._search_stats_done=set(); self._search_stat_worker=None
        self._search_stat_queue=[]; self._search_stat_pending=set()
        self._back_stack=[]; self._fwd_stack=[]; self._undo_stack=[]
        self._last_hover_index=QtCore.QModelIndex(); self._tooltip_last_ms=0.0; self._tooltip_interval_ms=180; self._tooltip_last_text=""
        self._fast_model=FastDirModel(self); self._fast_proxy=FsSortProxy(self); self._fast_proxy.setSourceModel(self._fast_model)
        self._using_fast=False; self._fast_stat_worker=None; self._enum_worker=None; self._pending_normal_root=None
        self._fast_enum_count = 0
        self._fast_enum_root = ""
        self._fast_enum_done = False
        self._deferred_normal_load_path = None
        self._file_worker=None
        self._op_progress_dialog=None
        self._dirload_timer={}
        self._sort_column = 0
        self._sort_order = Qt.AscendingOrder
        self._header_resize_guard = False
        self._browse_name_min_width = 140
        self._visible_stats_interval_ms = 60
        self._visible_stats_timer = None
        self._selection_update_interval_ms = 120
        self._selection_update_timer = None
        self._selection_cache_sig = None
        self._selection_cache_data = (0, False, 0)
        self._selection_cache_ts = 0.0
        self._disk_free_cache_key = None
        self._disk_free_cache_text = ""
        self._disk_free_cache_ts = 0.0
        self._disk_free_ttl_s = 2.0

    def _build_toolbar(self):
        self.btn_star=QToolButton(self); self.btn_star.setCheckable(True)
        self.btn_star.setIcon(icon_star(False, getattr(self.host,"theme","dark"))); self.btn_star.setToolTip("Add bookmark for this folder"); self.btn_star.setFixedHeight(UI_H)
        self._bm_btn_container=QWidget(self); self._bm_btn_layout=QHBoxLayout(self._bm_btn_container)
        self._bm_btn_layout.setContentsMargins(0,0,0,0); self._bm_btn_layout.setSpacing(ROW_SPACING)

        self.btn_cmd=QToolButton(self); self.btn_cmd.setIcon(icon_cmd(self.host.theme)); self.btn_cmd.setToolTip("Open Command Prompt here"); self.btn_cmd.setFixedHeight(UI_H)
        self.btn_up=QToolButton(self); self.btn_up.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp)); self.btn_up.setToolTip("Up"); self.btn_up.setFixedHeight(UI_H)
        self.btn_new=QToolButton(self); self.btn_new.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder)); self.btn_new.setToolTip("New Folder"); self.btn_new.setFixedHeight(UI_H)


        self.btn_new_file=QToolButton(self)
        self.btn_new_file.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        self.btn_new_file.setToolTip("New Text File (.txt)")
        self.btn_new_file.setFixedHeight(UI_H)

        self.btn_refresh=QToolButton(self); self.btn_refresh.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload)); self.btn_refresh.setToolTip("Refresh"); self.btn_refresh.setFixedHeight(UI_H)

        row_toolbar=QHBoxLayout()
        row_toolbar.setContentsMargins(0,0,0,0)

        row_toolbar.setSpacing(max(0, ROW_SPACING-2))
        row_toolbar.addWidget(self.btn_star)
        row_toolbar.addWidget(self._bm_btn_container,1)
        row_toolbar.addWidget(self.btn_cmd)
        row_toolbar.addWidget(self.btn_up)
        row_toolbar.addWidget(self.btn_new)
        row_toolbar.addWidget(self.btn_new_file)
        row_toolbar.addWidget(self.btn_refresh)
        self._row_toolbar=row_toolbar


        _tight_css = "QToolButton{padding-left:4px;padding-right:4px;}"
        for b in (self.btn_cmd, self.btn_up, self.btn_new, self.btn_new_file, self.btn_refresh):
            b.setStyleSheet(_tight_css)
            b.setAutoRaise(True)
        return row_toolbar

    def _build_path_row(self):
        self.path_bar=PathBar(self); self.path_bar.setToolTip("Breadcrumb - Double-click or F4/Ctrl+L to enter path")
        row_path=QHBoxLayout(); row_path.setContentsMargins(0,0,0,0); row_path.setSpacing(0); row_path.addWidget(self.path_bar,1)
        return row_path

    def _build_filter_row(self):
        self.filter_label=QLabel("Filter:", self)
        self.filter_edit=QLineEdit(self); self.filter_edit.setPlaceholderText("Filter (*.pdf, *file*.xls*, *.txt)"); self.filter_edit.setClearButtonEnabled(True); self.filter_edit.setFixedHeight(UI_H)
        self.filter_label.setFixedHeight(UI_H); self.filter_label.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.btn_search=QToolButton(self); self.btn_search.setText("Search"); self.btn_search.setToolTip("Run recursive search"); self.btn_search.setFixedHeight(UI_H)
        row_filter=QHBoxLayout(); row_filter.setContentsMargins(0,0,0,0); row_filter.setSpacing(ROW_SPACING)
        row_filter.addWidget(self.filter_label); row_filter.addWidget(self.filter_edit,1); row_filter.addWidget(self.btn_search,0)
        return row_filter

    def _init_models(self):
        self.source_model=QFileSystemModel(self); self.source_model.setReadOnly(False)
        try: self.source_model.setResolveSymlinks(False)
        except Exception: pass
        self.source_model.setFilter(QDir.AllEntries|QDir.NoDotAndDotDot|QDir.Hidden|QDir.System|QDir.Drives|QDir.AllDirs)
        self._native_icons=QFileIconProvider()
        self._generic_icons=GenericIconProvider(self.style())
        self._icon_provider_mode="native"
        if ALWAYS_GENERIC_ICONS:
            self.source_model.setIconProvider(self._generic_icons)
            self._icon_provider_mode="generic"
        else:
            self.source_model.setIconProvider(self._native_icons)

        self.stat_proxy=StatOverlayProxy(self); self.stat_proxy.setSourceModel(self.source_model)
        self.proxy=FsSortProxy(self); self.proxy.setSourceModel(self.stat_proxy)
        self.source_model.directoryLoaded.connect(self._on_directory_loaded)

    def _setup_view(self):
        self.view=ExplorerView(self); self.view.setModel(self.proxy); self.view.setSortingEnabled(True)
        self.view.setAlternatingRowColors(True); self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._on_context_menu)
        self.view.setMouseTracking(True)
        self.view.setUniformRowHeights(True); self.view.setAnimated(False); self.view.setExpandsOnDoubleClick(False); self.view.setRootIsDecorated(False)
        self._configure_header_browse()
        self.view.header().sectionClicked.connect(self._on_header_clicked)
        self.view.header().sectionResized.connect(self._on_header_section_resized)

    def _configure_header_browse(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, 90)
        header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)
        self._schedule_browse_name_autofit()

    def _configure_header_fast(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, 90)
        header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)
        self._schedule_browse_name_autofit()

    def _configure_header_search(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.resizeSection(1, 90)
        header.resizeSection(2, 150)

    def _schedule_browse_name_autofit(self):
        if self._search_mode:
            return
        QTimer.singleShot(0, self._autofit_browse_name_column)

    def _autofit_browse_name_column(self):
        if self._search_mode:
            return
        v = getattr(self, "view", None)
        if v is None:
            return
        header = v.header()
        if header is None:
            return

        if not v.isColumnHidden(2):
            return
        vp_w = v.viewport().width()
        if vp_w <= 0:
            return
        target = vp_w - header.sectionSize(1) - header.sectionSize(3)
        target = max(self._browse_name_min_width, target)
        if abs(header.sectionSize(0) - target) <= 1:
            return
        self._header_resize_guard = True
        try:
            header.blockSignals(True)
            header.resizeSection(0, target)
        finally:
            header.blockSignals(False)
            self._header_resize_guard = False

    def _on_header_section_resized(self, logical_index:int, _old_size:int, _new_size:int):
        if self._header_resize_guard or self._search_mode:
            return

        if logical_index in (1, 3):
            self._schedule_browse_name_autofit()

    def _build_status_row(self):
        self.lbl_sel=QLabel("", self); self.lbl_free=QLabel("", self)
        row_status=QHBoxLayout(); row_status.setContentsMargins(0,0,0,0); row_status.setSpacing(ROW_SPACING)
        row_status.addWidget(self.lbl_sel,0); row_status.addStretch(1); row_status.addWidget(self.lbl_free,0)
        self._row_status=row_status
        return row_status

    def _apply_layout(self, row_toolbar, row_path, row_filter, row_status):
        root_layout=QVBoxLayout(self); root_layout.setContentsMargins(*PANE_MARGIN); root_layout.setSpacing(max(1, ROW_SPACING//2))
        root_layout.addLayout(row_toolbar); root_layout.addLayout(row_path); root_layout.addLayout(row_filter)
        root_layout.addWidget(self.view,1); root_layout.addLayout(row_status)

    def _connect_signals(self):
        self.host.namedBookmarksChanged.connect(self._on_bookmarks_changed)
        self.path_bar.pathSubmitted.connect(lambda p: self.set_path(p, push_history=True))
        self.btn_star.clicked.connect(self._on_star_toggle)
        self.btn_cmd.clicked.connect(self._open_cmd_here)
        self.btn_up.clicked.connect(self.go_up)
        self.btn_new.clicked.connect(self.create_folder)
        self.btn_new_file.clicked.connect(self.create_text_file)
        self.btn_refresh.clicked.connect(self.hard_refresh)
        self.view.activated.connect(self._on_double_click)
        self.view.viewport().installEventFilter(self)
        self.view.installEventFilter(self)
        self.path_bar.installEventFilter(self)
        self._bm_btn_container.installEventFilter(self)
        self.filter_edit.installEventFilter(self)
        self._sel_model = None
        self._hook_selection_model()
        self.filter_edit.returnPressed.connect(self._apply_filter)
        self.btn_search.clicked.connect(self._apply_filter)
        self.filter_edit.textChanged.connect(self._on_filter_text_changed)
        try: self.view.verticalScrollBar().valueChanged.connect(lambda _v: self._request_visible_stats())
        except Exception: pass
        try: self.proxy.rowsInserted.connect(lambda *_: self._request_visible_stats(0))
        except Exception: pass
        try: self.proxy.modelReset.connect(lambda: self._request_visible_stats(0))
        except Exception: pass
        try: self.proxy.layoutChanged.connect(lambda *_: self._request_visible_stats(0))
        except Exception: pass

    def _register_shortcuts(self):
        def add_sc(seq, slot):
            sc=QShortcut(QKeySequence(seq), self.view)
            sc.setContext(Qt.WidgetWithChildrenShortcut); sc.activated.connect(slot); return sc
        add_sc("Backspace", self.go_back); add_sc("Alt+Left", self.go_back); add_sc("Alt+Right", self.go_forward)
        add_sc("Alt+Up", self.go_up)
        add_sc("Ctrl+L", self.path_bar.start_edit); add_sc("F4", self.path_bar.start_edit)
        add_sc("F3", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+F", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+C", self.copy_selection); add_sc("Ctrl+X", self.cut_selection); add_sc("Ctrl+V", self.paste_into_current);add_sc("Ctrl+Z", self.undo_last)
        add_sc("Delete", self.delete_selection); add_sc("Shift+Delete", lambda: self.delete_selection(permanent=True)); add_sc("F2", self.rename_selection)
        add_sc(Qt.Key_Return, self._open_current); add_sc(Qt.Key_Enter, self._open_current); add_sc("Ctrl+O", self._open_current)


        add_sc("Ctrl+Shift+C", lambda: self._copy_path_shortcut(False))
        add_sc("Alt+Shift+C",  lambda: self._copy_path_shortcut(True))

    def _load_sort_settings(self):
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            self._sort_column = s.value(f"pane_{self.pane_id}/sort_column", 0, type=int)
            order_val = s.value(f"pane_{self.pane_id}/sort_order", Qt.AscendingOrder, type=int)
            self._sort_order = Qt.DescendingOrder if order_val == Qt.DescendingOrder else Qt.AscendingOrder
        except Exception:
            self._sort_column = 0
            self._sort_order = Qt.AscendingOrder

    def _save_sort_settings(self):
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            s.setValue(f"pane_{self.pane_id}/sort_column", self._sort_column)
            s.setValue(f"pane_{self.pane_id}/sort_order", int(self._sort_order))
            s.sync()
        except Exception:
            pass

    def _apply_saved_sort(self):
        try:
            v = self.view
            if not v.isSortingEnabled():
                v.setSortingEnabled(True)
            v.header().setSortIndicator(self._sort_column, self._sort_order)
            v.sortByColumn(self._sort_column, self._sort_order)
        except Exception:
            pass


    def set_active_visual(self, active: bool):
        try:

            if self.objectName() != "paneRoot":
                self.setObjectName("paneRoot")
            self.setProperty("active", bool(active))


            targets = [self,
                       getattr(self, "view", None),
                       getattr(self, "filter_edit", None),
                       getattr(self, "path_bar", None)]
            for w in targets:
                if w:
                    try:
                        st = w.style()
                        st.unpolish(w)
                        st.polish(w)
                        w.update()
                    except Exception:
                        pass


            host = getattr(getattr(self, "path_bar", None), "_host", None)
            if host:
                from PyQt5.QtWidgets import QPushButton
                for btn in host.findChildren(QPushButton, "crumb"):
                    try:
                        st = btn.style()
                        st.unpolish(btn)
                        st.polish(btn)
                        btn.update()
                    except Exception:
                        pass
        except Exception:
            pass

    def _hook_selection_model(self):
        try:
            old = getattr(self, "_sel_model", None)
            if old is not None:
                try:
                    old.selectionChanged.disconnect(self._on_selection_changed)
                except Exception:
                    pass
        except Exception:
            pass

        self._sel_model = self.view.selectionModel()
        try:
            if self._sel_model:
                self._sel_model.selectionChanged.connect(self._on_selection_changed)
        except Exception:
            pass


        self._request_selection_status_update(immediate=True)

    def _copy_path_shortcut(self, folder_only: bool = False):
        sel = self._selected_paths()
        if len(sel) != 1:
            try:
                self.host.statusBar().showMessage("Select exactly one item.", 2000)
            except Exception:
                pass
            return

        p = sel[0]
        try:

            if folder_only and os.path.isfile(p):
                p = os.path.dirname(p)
        except Exception:

            pass

        try:
            QApplication.clipboard().setText(p)
            if folder_only:
                self.host.flash_status("Copied folder path to clipboard")
            else:
                self.host.flash_status("Copied full path to clipboard")
        except Exception:
            try:
                self.host.statusBar().showMessage("Failed to copy path to clipboard.", 2000)
            except Exception:
                pass

    def _stop_worker_thread(self, w, wait_ms: int = 100, label: str = "") -> bool:
        if not w:
            return True
        try:
            if w.isRunning():
                try:
                    if hasattr(w, "cancel"):
                        w.cancel()
                except Exception:
                    pass
                if not w.wait(wait_ms):
                    # Don't block UI; defer deletion after thread finishes.
                    try:
                        w.finished.connect(w.deleteLater, QtCore.Qt.UniqueConnection)
                    except Exception:
                        try:
                            w.finished.connect(w.deleteLater)
                        except Exception:
                            pass
                    if DEBUG and label:
                        dlog(f"[thread] deferred cleanup: {label}")
                    return False
            w.deleteLater()
            return True
        except Exception:
            return False

    def _cancel_search_worker(self):

        self._stop_worker_thread(getattr(self, "_search_worker", None), 120, "search")
        self._search_worker = None
        self._stop_worker_thread(getattr(self, "_search_stat_worker", None), 80, "search-stat")
        self._search_stat_worker = None
        self._search_pending_items = {}
        self._search_stats_done = set()
        self._search_stat_queue = []
        self._search_stat_pending = set()
        try:
            while QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()
        except Exception:
            pass

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
            item_name.setData(str(name).lower(), NAME_FOLD_ROLE)
            item_name.setData(False, SEARCH_ICON_READY_ROLE)
            item_name.setData(full, Qt.ToolTipRole)

            item_name.setIcon(self._default_icon(isdir))


            item_size = QStandardItem()
            item_size.setData(0, Qt.EditRole)
            item_size.setData(0, SIZE_BYTES_ROLE)

            item_date = QStandardItem("")
            item_date.setData(QDateTime(), Qt.EditRole)

            item_folder = QStandardItem(rel_folder)

            root_item.appendRow([item_name, item_size, item_date, item_folder])


        self._request_visible_stats(0)

    @QtCore.pyqtSlot()
    def _on_search_finished(self):
        try:
            if QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()
        except Exception:
            pass
        self._search_worker = None
        try:
            if self._search_mode and self._search_proxy and self.view.model() is self._search_proxy:
                hdr = self.view.header()
                col = hdr.sortIndicatorSection()
                order = hdr.sortIndicatorOrder()
                if not self.view.isSortingEnabled():
                    self.view.setSortingEnabled(True)
                self.view.sortByColumn(col, order)
        except Exception:
            pass

        self._request_visible_stats(0)

    def _start_next_search_stat_worker(self, batch_limit: int = 220):
        cur = getattr(self, "_search_stat_worker", None)
        if cur and cur.isRunning():
            return
        if not self._search_stat_queue:
            self._search_stat_worker = None
            return

        size = max(1, int(batch_limit))
        batch = self._search_stat_queue[:size]
        del self._search_stat_queue[:size]

        w = NormalStatWorker(batch, self)
        w.statReady.connect(self._apply_search_stat, Qt.QueuedConnection)
        w.finishedCycle.connect(lambda b=batch: self._on_search_stat_cycle_finished(b), Qt.QueuedConnection)
        self._search_stat_worker = w
        w.start()

    def _enqueue_search_stat_paths(self, paths: list[str], batch_limit: int = 220):
        added = False
        for p in paths:
            if not p:
                continue
            if p in self._search_stat_pending:
                continue
            self._search_stat_pending.add(p)
            self._search_stat_queue.append(p)
            added = True
        if added:
            self._start_next_search_stat_worker(batch_limit=batch_limit)

    def _on_search_stat_cycle_finished(self, batch):
        for p in batch:
            self._search_stat_pending.discard(p)
        self._search_stat_worker = None
        if self._search_mode:
            self._start_next_search_stat_worker()

    def _on_filter_text_changed(self, text: str):

        if not (text or "").strip():
            self._enter_browse_mode()

    @QtCore.pyqtSlot(str, object, object)
    def _apply_search_stat(self, path: str, size_val, mtime_val):

        d = getattr(self, "_search_pending_items", None)
        if not isinstance(d, dict):
            return
        pair = d.pop(path, None)
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
        base_dir = self.current_path()
        try:
            name = f"New Document {time.strftime('%Y%m%d-%H%M%S')}.txt"
            new_path = _create_new_file_with_template(base_dir, name, ".txt")
        except Exception as e:
            QMessageBox.critical(self, "Create failed", str(e))
            return


        self.hard_refresh()


        def _try_select():

            try:
                if self.view and not self.view.hasFocus():
                    self.view.setFocus(Qt.ShortcutFocusReason)
            except Exception:
                pass


            if self._search_mode:
                return

            try:
                if self._using_fast:

                    rows = self._fast_model.rowCount()
                    for r in range(rows):
                        rp = self._fast_model.row_path(r)
                        if rp and os.path.normcase(rp) == os.path.normcase(new_path):
                            prx_ix = self._fast_proxy.index(r, 0)
                            sm = self.view.selectionModel()
                            sm.clearSelection()
                            sm.select(prx_ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                            self.view.scrollTo(prx_ix, QAbstractItemView.PositionAtCenter)
                            self.view.setCurrentIndex(prx_ix)
                            return
                else:

                    src_ix = self.source_model.index(new_path)
                    if src_ix.isValid():
                        st_ix = self.stat_proxy.mapFromSource(src_ix)
                        prx_ix = self.proxy.mapFromSource(st_ix)
                        sm = self.view.selectionModel()
                        sm.clearSelection()
                        sm.select(prx_ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                        self.view.scrollTo(prx_ix, QAbstractItemView.PositionAtCenter)
                        self.view.setCurrentIndex(prx_ix)
                        return
            except Exception:
                pass


        for delay in (0, 80, 200, 450):
            QTimer.singleShot(delay, _try_select)

        try:
            self.host.flash_status("Text file created")
        except Exception:
            pass

    def _default_icon(self, is_dir: bool) -> QIcon:
        try:
            if ALWAYS_GENERIC_ICONS:
                return self._generic_icons.icon(QFileIconProvider.Folder if is_dir else QFileIconProvider.File)
            return self.style().standardIcon(QStyle.SP_DirIcon if is_dir else QStyle.SP_FileIcon)
        except Exception: return QIcon()

    def _cancel_fast_stat_worker(self):
        self._stop_worker_thread(self._fast_stat_worker, 120, "fast-stat")
        self._fast_stat_worker=None

    def _cancel_enum_worker(self, wait_ms: int = 150):
        self._stop_worker_thread(self._enum_worker, wait_ms, "dir-enum")
        self._enum_worker = None

    def _cancel_file_worker(self, wait_ms: int = 300):
        self._stop_worker_thread(getattr(self, "_file_worker", None), wait_ms, "file-op")
        self._file_worker = None
        dlg = getattr(self, "_op_progress_dialog", None)
        if dlg:
            try:
                dlg.close()
                dlg.deleteLater()
            except Exception:
                pass
            self._op_progress_dialog = None

    def shutdown(self, wait_ms: int = 300):
        try:
            self._cancel_search_worker()
        except Exception:
            pass
        try:
            self._cancel_fast_stat_worker()
        except Exception:
            pass
        try:
            self._cancel_enum_worker(wait_ms)
        except Exception:
            pass
        try:
            self._cancel_file_worker(wait_ms)
        except Exception:
            pass

    def _ensure_visible_stats_timer(self):
        if self._visible_stats_timer is not None:
            return
        t = QTimer(self)
        t.setSingleShot(True)
        t.setInterval(self._visible_stats_interval_ms)
        t.timeout.connect(self._schedule_visible_stats)
        self._visible_stats_timer = t

    def _request_visible_stats(self, delay_ms: int | None = None):
        self._ensure_visible_stats_timer()
        t = self._visible_stats_timer
        delay = self._visible_stats_interval_ms if delay_ms is None else max(0, int(delay_ms))
        if t.isActive():
            remaining = t.remainingTime()
            if remaining >= 0 and remaining <= delay:
                return
            t.stop()
        t.start(delay)

    def _ensure_selection_update_timer(self):
        if self._selection_update_timer is not None:
            return
        t = QTimer(self)
        t.setSingleShot(True)
        t.setInterval(self._selection_update_interval_ms)
        t.timeout.connect(self._flush_selection_status_update)
        self._selection_update_timer = t

    def _request_selection_status_update(self, immediate: bool = False):
        self._ensure_selection_update_timer()
        if immediate:
            if self._selection_update_timer.isActive():
                self._selection_update_timer.stop()
            self._flush_selection_status_update()
            return
        self._selection_update_timer.start(self._selection_update_interval_ms)

    def _selection_summary(self):
        sel = self._selected_paths()
        sig = tuple(sel)
        now = time.perf_counter()
        if sig == self._selection_cache_sig and (now - self._selection_cache_ts) <= 0.2:
            return self._selection_cache_data

        cnt = len(sel)
        only_files = (cnt > 0)
        total = 0
        if only_files:
            for p in sel:
                if not os.path.isfile(p):
                    only_files = False
                    total = 0
                    break
                try:
                    total += os.path.getsize(p)
                except Exception:
                    pass

        data = (cnt, only_files, total)
        self._selection_cache_sig = sig
        self._selection_cache_data = data
        self._selection_cache_ts = now
        return data

    def _update_free_space_label(self, force: bool = False):
        path = self.current_path()
        if self._is_network_path(path):
            self.lbl_free.setText("")
            return

        key = self._drive_label(path)
        now = time.perf_counter()
        if (not force and self._disk_free_cache_key == key
            and (now - self._disk_free_cache_ts) <= self._disk_free_ttl_s):
            self.lbl_free.setText(self._disk_free_cache_text)
            return

        try:
            _total, _used, free = shutil.disk_usage(path)
            text = f"{key} free {human_size(free)}"
        except Exception:
            text = ""

        self._disk_free_cache_key = key
        self._disk_free_cache_text = text
        self._disk_free_cache_ts = now
        self.lbl_free.setText(text)

    def _flush_selection_status_update(self):
        self._render_selection_status(update_statusbar=True, update_label=True, update_free=True)

    def _render_selection_status(self, update_statusbar: bool, update_label: bool, update_free: bool):
        cnt, only_files, total = self._selection_summary()

        if update_statusbar:
            msg = f"Pane {self.pane_id} / selected {cnt} item(s)"
            if cnt and only_files:
                msg += f" / {human_size(total)}"
            try:
                self.host.statusBar().showMessage(msg, 2000)
            except Exception:
                pass

        if update_label:
            text = ""
            if cnt:
                if only_files:
                    text = f"{cnt} selected / {human_size(total)}"
                else:
                    text = f"{cnt} selected"
            self.lbl_sel.setText(text)

        if update_free:
            self._update_free_space_label(force=False)

    def _schedule_visible_stats(self):

        if self._search_mode:
            self._fill_search_visible_icons()
            return


        current_model = self.view.model()
        root_ix = self.view.rootIndex()
        vp = self.view.viewport()


        if self._using_fast:
            if current_model is not self._fast_proxy:
                return
            rc = self._fast_proxy.rowCount(root_ix)
            if rc <= 0:
                return
            top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
            bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
            proxy_start = top_ix.row() if top_ix.isValid() else 0
            proxy_end   = bot_ix.row() if bot_ix.isValid() else min(proxy_start + 80, rc - 1)
            proxy_start = max(0, proxy_start - 30)
            proxy_end   = min(rc - 1, proxy_end + 50)

            to_rows = []
            for r in range(proxy_start, proxy_end + 1):
                prx_ix = self._fast_proxy.index(r, 0, root_ix)
                src_ix = self._fast_proxy.mapToSource(prx_ix)
                row = src_ix.row()
                if row is None or row < 0:
                    continue
                if not self._fast_model.has_stat(row):
                    to_rows.append(row)
                    if len(to_rows) >= 220:
                        break


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
            if self._fast_stat_worker and self._fast_stat_worker.isRunning():
                return
            root = self._fast_model.rootPath()
            w = FastStatWorker(self._fast_model, root, to_rows, self)
            w.statReady.connect(self._fast_model.apply_stat, Qt.QueuedConnection)
            def _on_fast_cycle_finished():
                if self._fast_stat_worker is w:
                    self._fast_stat_worker = None
                self._request_visible_stats(0)
            w.finishedCycle.connect(_on_fast_cycle_finished, QtCore.Qt.QueuedConnection)
            self._fast_stat_worker = w
            w.start()
            return


        if current_model is not self.proxy:
            return

        model = self.proxy
        stat_proxy = self.stat_proxy
        src_model = self.source_model

        paths = []
        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        rc = model.rowCount(root_ix)
        end = bot_ix.row() if bot_ix.isValid() else min(start + 120, rc - 1)
        start = max(0, start - 40)
        end = min(rc - 1, end + 80)
        if end < start:
            end = start

        for r in range(start, end + 1):
            prx_ix = model.index(r, 0, root_ix)
            if not prx_ix.isValid():
                continue
            st_ix  = model.mapToSource(prx_ix)
            if not st_ix.isValid():
                continue
            src_ix = stat_proxy.mapToSource(st_ix)
            if not src_ix.isValid():
                continue
            try:
                p = src_model.filePath(src_ix)
            except Exception:
                p = None
            if p:
                paths.append(p)

        if paths:
            stat_proxy.request_paths(paths)

    def _on_header_clicked(self, col:int):
        v=self.view
        if not v.isSortingEnabled():
            v.setSortingEnabled(True)


        if col == self._sort_column:
            new_order = Qt.DescendingOrder if self._sort_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            new_order = Qt.AscendingOrder

        self._sort_column = col
        self._sort_order = new_order

        v.header().setSortIndicator(col, new_order)
        v.sortByColumn(col, new_order)


        self._save_sort_settings()

    def _mark_self_active(self):
        try:
            if hasattr(self.host, "mark_active_pane"):
                self.host.mark_active_pane(self)
        except Exception:
            pass

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
                self._request_visible_stats(0)
                self._schedule_browse_name_autofit()
            return False

        if obj is getattr(self, "_bm_btn_container", None):
            if ev.type() in (QEvent.Resize, QEvent.Show):
                QTimer.singleShot(0, self._refresh_quick_bookmark_button_texts)
            return False


        if obj is self.filter_edit:
            if ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Escape:
                try:
                    self.filter_edit.clear()
                finally:

                    self._enter_browse_mode()
                ev.accept()
                return True
            return False

        return super().eventFilter(obj, ev)


    def _open_cmd_here(self):
        path = self.current_path() or os.getcwd()
        try:
            path = os.path.abspath(path)
        except Exception:
            pass


        comspec = os.environ.get("ComSpec") or r"C:\Windows\System32\cmd.exe"


        if HAS_PYWIN32:
            try:


                params = f'/K title Multi-Pane File Explorer & cd /d "{path}"'
                win32api.ShellExecute(
                    int(self.window().winId()) if self.window() else 0,
                    "open",
                    comspec,
                    params,
                    path,
                    win32con.SW_SHOWNORMAL
                )
                return
            except Exception:
                pass


        try:
            flags = 0
            flags |= getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
            flags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)


            si = None
            try:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = 1
            except Exception:
                si = None


            subprocess.Popen(
                [comspec, "/K", f'cd /d "{path}"'],
                cwd=path,
                creationflags=flags,
                startupinfo=si
            )
            return
        except Exception:
            pass


        try:

            cmdline = f'start "" /D "{path}" "{comspec}" /K cd /d "{path}"'
            subprocess.Popen(cmdline, shell=True)
            return
        except Exception as e:
            QMessageBox.critical(self, "Command Prompt", f"Failed to launch cmd.exe:\n{e}")


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
            btn=QToolButton(self._bm_btn_container)
            btn.setObjectName("quickBookmarkBtn")
            btn.setText(str(name))
            btn.setProperty("fullText", str(name))
            btn.setToolTip(p)
            btn.setFixedHeight(UI_H)
            btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            btn.clicked.connect(lambda _=False, path=p: self.set_path(path, push_history=True))
            self._bm_btn_layout.addWidget(btn)
        self._bm_btn_layout.addStretch(1)
        QTimer.singleShot(0, self._refresh_quick_bookmark_button_texts)

    def _refresh_quick_bookmark_button_texts(self):
        try:
            for btn in self._bm_btn_container.findChildren(QToolButton, "quickBookmarkBtn"):
                full = str(btn.property("fullText") or btn.text() or "")
                w = max(24, int(btn.width()) - 12)
                elided = btn.fontMetrics().elidedText(full, Qt.ElideRight, w)
                if btn.text() != elided:
                    btn.setText(elided)
        except Exception:
            pass

    def _selected_paths(self):
        paths=[]; sel=self.view.selectionModel().selectedRows(0)
        for ix in sel:
            p=self._index_to_full_path(ix)
            if p: paths.append(p)
        seen=set(); out=[]
        for p in paths:
            np=os.path.normcase(os.path.normpath(p))
            if np not in seen: seen.add(np); out.append(np)
        return out

    def _index_to_full_path(self, index):
        if not index.isValid():
            return None

        model = None
        try:
            model = index.model()
        except Exception:
            model = None

        try:

            if model in (self._fast_proxy, self._fast_model, self._search_proxy, self._search_model):
                return index.sibling(index.row(), 0).data(Qt.UserRole)


            if model is self.proxy:
                st_ix = self.proxy.mapToSource(index)
                if not st_ix.isValid():
                    return None
                src_ix = self.stat_proxy.mapToSource(st_ix)
                return self.source_model.filePath(src_ix) if src_ix.isValid() else None

            if model is self.stat_proxy:
                src_ix = self.stat_proxy.mapToSource(index)
                return self.source_model.filePath(src_ix) if src_ix.isValid() else None

            if model is self.source_model:
                return self.source_model.filePath(index)


            return index.sibling(index.row(), 0).data(Qt.UserRole)
        except Exception:
            return None

    def _use_fast_model(self, path: str):

        self._cancel_fast_stat_worker()
        self._cancel_enum_worker(wait_ms=100)
        try:
            self.stat_proxy.clear_cache()
        except Exception:
            pass


        self._using_fast = True
        self._fast_model.reset_dir(path)
        self.view.setModel(self._fast_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())
        self._configure_header_fast()
        self._fast_enum_count = 0
        self._fast_enum_root = path
        self._fast_enum_done = False




        orig_apply_icon = getattr(self._fast_model, "apply_icon", None)

        def _noop_apply_icon(_row, _icon):

            return None

        try:
            self._fast_model.apply_icon = _noop_apply_icon
        except Exception:
            orig_apply_icon = None


        was_sorting = self.view.isSortingEnabled()
        if was_sorting:
            self.view.setSortingEnabled(False)
        self.view.setUpdatesEnabled(False)

        self._fast_batch_counter = 0
        self._enum_worker = DirEnumWorker(path, self)


        def _on_batch(rows):
            self._fast_model.append_rows(rows)
            self._fast_enum_count += len(rows or [])
            self._fast_batch_counter += 1
            if (self._fast_batch_counter % 6) == 0:
                self._request_visible_stats(0)

        self._enum_worker.batchReady.connect(_on_batch, QtCore.Qt.QueuedConnection)
        self._enum_worker.error.connect(
            lambda msg: self.host.statusBar().showMessage(f"List error: {msg}", 4000)
        )

        def _restore_apply_icon():
            if orig_apply_icon is None:
                return
            try:
                self._fast_model.apply_icon = orig_apply_icon
            except Exception:
                pass


        def _on_finished():
            try:
                self._fast_enum_done = True

                self.view.setUpdatesEnabled(True)

                self._apply_saved_sort()
                self.view.setSortingEnabled(True)


                _restore_apply_icon()


                self._request_visible_stats(0)
                self._request_visible_stats(80)
                try:
                    cur = self.current_path()
                    deferred = getattr(self, "_deferred_normal_load_path", None)
                    if (deferred
                        and not self._search_mode
                        and os.path.normcase(deferred) == os.path.normcase(path)
                        and os.path.normcase(cur) == os.path.normcase(path)):
                        self._deferred_normal_load_path = None
                        QTimer.singleShot(0, lambda p=path, c=self._fast_enum_count: self._start_normal_model_loading(p, known_count=c))
                except Exception:
                    pass
            finally:

                if not was_sorting:
                    self.view.setSortingEnabled(False)

        self._enum_worker.finished.connect(_on_finished)


        self._request_visible_stats(0)


        self._enum_worker.start()


    def _start_normal_model_loading(self, path:str, known_count: int | None = None):
        self._pending_normal_root = path
        t = QElapsedTimer(); t.start()
        self._dirload_timer[path.lower()] = t


        HUGE_THRESHOLD = 3000
        GENERIC_THRESHOLD = 1200
        if known_count is not None:
            try:
                count = max(0, int(known_count))
            except Exception:
                count = 0
            is_huge = count >= HUGE_THRESHOLD
        else:
            count = 0
            is_huge = False
            try:
                with os.scandir(path) as it:
                    for _ in it:
                        count += 1
                        if count >= HUGE_THRESHOLD:
                            is_huge = True
                            break
            except Exception:

                pass

        if is_huge:

            dlog(f"[perf] Skip QFileSystemModel for huge folder (>= {HUGE_THRESHOLD} items): {path}")
            self._pending_normal_root = None
            self._request_visible_stats(0)
            return


        try:
            use_generic = ALWAYS_GENERIC_ICONS or count >= GENERIC_THRESHOLD
            if use_generic and self._icon_provider_mode != "generic":
                self.source_model.setIconProvider(self._generic_icons)
                self._icon_provider_mode = "generic"
            elif (not use_generic) and self._icon_provider_mode != "native":
                self.source_model.setIconProvider(self._native_icons)
                self._icon_provider_mode = "native"
        except Exception:
            pass


        _ = self.source_model.setRootPath(path)

    def _unc_share_root(self, path:str)->str:
        if not path:
            return ""
        p = path.replace("/", "\\")
        if not p.startswith("\\\\"):
            return ""
        comps = [c for c in p.split("\\") if c]
        if len(comps) < 2:
            return ""
        return f"\\\\{comps[0]}\\{comps[1]}"

    def _try_network_auth_prompt(self, path:str)->tuple[bool, bool]:
        target = self._unc_share_root(path)
        prompted = False

        if target:
            try:
                class _NETRESOURCEW(ctypes.Structure):
                    _fields_ = [
                        ("dwScope", ctypes.c_ulong),
                        ("dwType", ctypes.c_ulong),
                        ("dwDisplayType", ctypes.c_ulong),
                        ("dwUsage", ctypes.c_ulong),
                        ("lpLocalName", ctypes.c_wchar_p),
                        ("lpRemoteName", ctypes.c_wchar_p),
                        ("lpComment", ctypes.c_wchar_p),
                        ("lpProvider", ctypes.c_wchar_p),
                    ]

                nr = _NETRESOURCEW()
                nr.dwType = 1
                nr.lpRemoteName = target

                CONNECT_INTERACTIVE = 0x00000008
                CONNECT_PROMPT = 0x00000010
                CONNECT_TEMPORARY = 0x00000004
                flags = CONNECT_INTERACTIVE | CONNECT_PROMPT | CONNECT_TEMPORARY

                hwnd = int(self.window().winId()) if self.window() else 0
                rc = ctypes.windll.mpr.WNetAddConnection3W(
                    ctypes.c_void_p(hwnd),
                    ctypes.byref(nr),
                    None,
                    None,
                    flags,
                )
                prompted = True
                if DEBUG:
                    dlog(f"[net] WNetAddConnection3W rc={rc} target={target}")
            except Exception as e:
                if DEBUG:
                    dlog(f"[net] WNetAddConnection3W failed: {e}")

            if os.path.exists(path) or os.path.exists(target):
                return True, prompted

        open_target = target or path
        try:
            subprocess.Popen(["explorer.exe", open_target])
            prompted = True
        except Exception as e:
            if DEBUG:
                dlog(f"[net] explorer launch failed ({open_target}): {e}")

        return os.path.exists(path), prompted

    def set_path(self, path:str, push_history:bool=True):
        with perf(f"set_path begin -> {path}"):
            path = nice_path(path)
            if (not os.path.exists(path)) and self._is_network_path(path):
                accessible, prompted = self._try_network_auth_prompt(path)
                if accessible:
                    pass
                elif prompted:
                    QMessageBox.information(
                        self,
                        "Network Sign-in",
                        "Network location requires sign-in.\nPlease complete sign-in and try again.",
                    )
                    return
            if not os.path.exists(path):
                QMessageBox.warning(self, "Path not found", path)
                return



            if self._search_mode:
                self._enter_browse_mode()

            cur = getattr(self.path_bar, "_current_path", None)
            if push_history and cur and os.path.normcase(cur) != os.path.normcase(path):
                self._back_stack.append(cur)
                self._fwd_stack.clear()


            self.path_bar.set_path(path)
            self._update_star_button()


            try:
                self._bind_fs_watcher(path)
            except Exception:
                pass


            self._deferred_normal_load_path = path
            self._use_fast_model(path)
            if self._fast_enum_done and os.path.normcase(self._fast_enum_root) == os.path.normcase(path):
                self._deferred_normal_load_path = None
                QTimer.singleShot(0, lambda p=path, c=self._fast_enum_count: self._start_normal_model_loading(p, known_count=c))


            QTimer.singleShot(50, self._update_pane_status)
            self._update_statusbar_selection()

    def _bind_fs_watcher(self, folder_path: str):

        if not hasattr(self, "_fswatch"):
            self._fswatch = QtCore.QFileSystemWatcher(self)
            self._fswatch.directoryChanged.connect(self._on_fs_changed)
            self._fswatch.fileChanged.connect(self._on_fs_changed)

        if not hasattr(self, "_fswatch_debounce"):
            self._fswatch_debounce = QTimer(self)
            self._fswatch_debounce.setSingleShot(True)
            self._fswatch_debounce.setInterval(600)
            self._fswatch_debounce.timeout.connect(self._apply_fs_change)


        try:
            dirs = list(self._fswatch.directories())
            if dirs:
                self._fswatch.removePaths(dirs)
        except Exception:
            pass

        try:

            if os.path.isdir(folder_path):
                self._fswatch.addPath(folder_path)
        except Exception:

            pass

    def _on_fs_changed(self, _path: str):
        try:
            if self._fswatch_debounce.isActive():
                self._fswatch_debounce.stop()
            self._fswatch_debounce.start()
        except Exception:
            pass

    def _apply_fs_change(self):
        try:

            if getattr(self, "_search_mode", False):
                pattern = self.filter_edit.text().strip()
                if pattern:

                    self._apply_filter()
                else:

                    self._enter_browse_mode()
                return


            if getattr(self, "_using_fast", False):
                self._use_fast_model(self.current_path())
                return


            self._request_visible_stats(0)
            self._update_pane_status()
        except Exception:
            pass


    @QtCore.pyqtSlot(str)
    def _on_directory_loaded(self, loaded_path:str):
        key = loaded_path.lower()
        if key in self._dirload_timer:
            ms = self._dirload_timer[key].elapsed()
            dlog(f"directoryLoaded: '{loaded_path}' in {ms} ms")
            self._dirload_timer.pop(key, None)


        if self._pending_normal_root is None:
            return

        if (os.path.normcase(loaded_path) == os.path.normcase(self.current_path())
            and self._using_fast and self._pending_normal_root
            and os.path.normcase(self._pending_normal_root) == os.path.normcase(loaded_path)
            and not self._search_mode):
            try:
                src_idx = self.source_model.index(loaded_path)
                self.view.setModel(self.proxy)
                self.view.setRootIndex(self.proxy.mapFromSource(self.stat_proxy.mapFromSource(src_idx)))
                self._using_fast = False
                self._pending_normal_root = None


                self._hook_selection_model()

                self._configure_header_browse()
                if not self.view.isSortingEnabled():
                    self.view.setSortingEnabled(True)

                self._apply_saved_sort()
                self._request_visible_stats(0)
            except Exception:
                pass

    def current_path(self)->str: return self.path_bar._current_path or QDir.homePath()
    def go_back(self):
        if not self._back_stack: return
        dst=self._back_stack.pop(); self._fwd_stack.append(self.current_path()); self.set_path(dst, push_history=False)
    def go_forward(self):
        if not self._fwd_stack: return
        dst=self._fwd_stack.pop(); self._back_stack.append(self.current_path()); self.set_path(dst, push_history=False)
    def go_up(self):
        parent=Path(self.current_path()).parent; self.set_path(str(parent), push_history=True)
    def refresh(self):
        self.hard_refresh()
    def hard_refresh(self):
        if self._search_mode:
            self._apply_filter(); return
        try: self._cancel_fast_stat_worker()
        except Exception: pass
        try: self._cancel_enum_worker(wait_ms=100)
        except Exception: pass
        try: self.stat_proxy.clear_cache()
        except Exception: pass
        self.set_path(self.current_path(), push_history=False)
        self.host.flash_status("Hard refresh")


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


    def copy_selection(self):
        paths=self._selected_paths()
        if not paths: return
        self.host.set_clipboard({"op":"copy","paths":paths}); self.host.flash_status(f"Copied {len(paths)} item(s)")
    def cut_selection(self):
        paths=self._selected_paths()
        if not paths: return
        self.host.set_clipboard({"op":"cut","paths":paths}); self.host.flash_status(f"Cut {len(paths)} item(s)")

    def _external_clipboard_payload(self):
        try:
            cb = QApplication.clipboard()
            md = cb.mimeData()
        except Exception:
            return None
        if not md or not md.hasUrls():
            return None

        urls = md.urls()

        paths = list(dict.fromkeys(u.toLocalFile() for u in urls if u.isLocalFile()))
        if not paths:
            return None

        def _decode_drop_effect_from_qt():
            if sys.platform != "win32":
                return None
            fmt = 'application/x-qt-windows-mime;value="Preferred DropEffect"'
            if not md.hasFormat(fmt):
                return None
            try:
                data = md.data(fmt)
                if data and len(data) >= 4:
                    return int.from_bytes(bytes(data)[:4], byteorder="little", signed=False)
            except Exception:
                return None

        def _decode_drop_effect_from_win32():
            if not HAS_PYWIN32:
                return None
            try:
                fmt = win32clipboard.RegisterClipboardFormat("Preferred DropEffect")
                win32clipboard.OpenClipboard()
                try:
                    data = win32clipboard.GetClipboardData(fmt)
                    if data and len(bytes(data)) >= 4:
                        return int.from_bytes(bytes(data)[:4], byteorder="little", signed=False)
                finally:
                    win32clipboard.CloseClipboard()
            except Exception:
                return None

        effect = _decode_drop_effect_from_qt()
        if effect is None:
            effect = _decode_drop_effect_from_win32()

        op = "copy"
        if effect is not None:
            if effect & 2:
                op = "move"
            elif effect & 1:
                op = "copy"

        return {"op": op, "paths": paths}

    def paste_into_current(self):
        clip=self.host.get_clipboard() or self._external_clipboard_payload()
        if not clip:
            self.host.flash_status("Clipboard has no files to paste")
            return
        dst_dir=self.current_path(); op=clip.get("op"); srcs=clip.get("paths") or []
        if not srcs:
            self.host.flash_status("Clipboard has no files to paste")
            return
        self._start_bg_op("copy" if op=="copy" else "move", srcs, dst_dir)
        if op=="cut": self.host.clear_clipboard()

    def _start_bg_op(self, op, srcs, dst_dir):
        cur_worker = getattr(self, "_file_worker", None)
        if cur_worker and cur_worker.isRunning():
            self.host.flash_status("Another file operation is already running")
            return

        valid_srcs = []
        skipped_same = []
        blocked_nested = []
        auto_map = {}

        for src in srcs:
            if not src or not os.path.exists(src):
                continue

            base = os.path.basename(src.rstrip("\\/")) or os.path.basename(src)
            dst = os.path.join(dst_dir, base)

            if _paths_same(src, dst):
                if op == "copy":

                    auto_map[src] = "copy"
                    valid_srcs.append(src)
                else:

                    skipped_same.append(src)
                continue

            if os.path.isdir(src) and not os.path.islink(src) and _is_subpath(dst, src):
                blocked_nested.append(src)
                continue

            valid_srcs.append(src)

        if blocked_nested:
            sample = "\n".join(blocked_nested[:5])
            more = "\n..." if len(blocked_nested) > 5 else ""
            QMessageBox.warning(
                self,
                f"{op.title()} blocked",
                "Cannot copy/move a folder into its own subfolder:\n\n"
                f"{sample}{more}",
            )

        if not valid_srcs:
            if skipped_same:
                self.host.flash_status("Nothing to move (same source and destination)")
            return


        conflicts=[]
        for src in valid_srcs:
            if src in auto_map:
                continue
            base=os.path.basename(src.rstrip("\\/")) or os.path.basename(src)
            dst=os.path.join(dst_dir, base)
            if os.path.exists(dst): conflicts.append((src,dst))

        conflict_map=dict(auto_map)
        if conflicts:
            dlg=ConflictResolutionDialog(self, conflicts, dst_dir)
            if dlg.exec_()!=QDialog.Accepted: return

            conflict_map.update(dlg.result_map())


        old_dlg = getattr(self, "_op_progress_dialog", None)
        if old_dlg:
            try:
                old_dlg.close()
                old_dlg.deleteLater()
            except Exception:
                pass
            self._op_progress_dialog = None

        worker=FileOpWorker(op, valid_srcs, dst_dir, conflict_map=conflict_map, parent=self)
        dlgp=QProgressDialog(f"{op.title()} in progress...", "Cancel", 0, 100, self)
        dlgp.setWindowTitle(f"{op.title()} files"); dlgp.setWindowModality(Qt.WindowModal)
        dlgp.setAutoClose(True); dlgp.setAutoReset(True)
        dlgp.setMinimumDuration(0)

        worker.progress.connect(dlgp.setValue)
        worker.status.connect(lambda s: self.host.statusBar().showMessage(s,2000))

        def _on_error(msg):
            if msg == "Operation cancelled.":
                self.host.flash_status(f"{op.title()} cancelled")
                return
            QMessageBox.critical(self,f"{op.title()} failed",msg)

        def _finish_ok():
            try: dlgp.setValue(100); dlgp.close()
            except Exception: pass
            if not self._using_fast and not self._search_mode: self.stat_proxy.clear_cache()
            self._request_visible_stats(0); self._update_pane_status()
            self.host.flash_status(f"{op.title()} complete")

        def _cleanup_worker():
            if getattr(self, "_file_worker", None) is worker:
                self._file_worker = None
            try:
                if getattr(self, "_op_progress_dialog", None) is dlgp:
                    self._op_progress_dialog = None
                dlgp.close()
                dlgp.deleteLater()
            except Exception:
                pass
            worker.deleteLater()

        worker.error.connect(_on_error)
        worker.finished.connect(_cleanup_worker)
        worker.finished_ok.connect(_finish_ok); dlgp.canceled.connect(worker.cancel)
        self._file_worker = worker
        self._op_progress_dialog = dlgp
        worker.start()
        dlgp.open()


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


    def _enter_browse_mode(self):

        try:
            if hasattr(self, "_cancel_search_worker"):
                self._cancel_search_worker()
        except Exception:
            pass
        self._search_pending_items = {}
        self._search_stats_done = set()
        self._search_model = None
        self._search_proxy = None

        if not self._search_mode:
            self._request_visible_stats(0)
            return

        self._search_mode = False

        if self._using_fast:
            self.view.setModel(self._fast_proxy)
            self.view.setRootIndex(QtCore.QModelIndex())
        else:
            self.view.setModel(self.proxy)


        self._hook_selection_model()

        path = self.current_path()
        if not self._using_fast:
            try:
                src_idx = self.source_model.index(path)
                self.view.setRootIndex(self.proxy.mapFromSource(self.stat_proxy.mapFromSource(src_idx)))
            except Exception:
                pass

        self._configure_header_browse()
        if not self.view.isSortingEnabled():
            self.view.setSortingEnabled(True)

        self._apply_saved_sort()

        self._request_visible_stats(0)

    def _enter_search_mode(self, model:QStandardItemModel):
        self._cancel_fast_stat_worker()
        self._using_fast=False
        self._search_mode=True
        self._search_model=model
        self._search_proxy=FsSortProxy(self)
        self._search_proxy.setDynamicSortFilter(False)
        self._search_proxy.setSourceModel(self._search_model)
        self.view.setModel(self._search_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())


        self._hook_selection_model()

        header = self.view.header()
        self._configure_header_search()
        if not self.view.isSortingEnabled(): self.view.setSortingEnabled(True)
        header.setSortIndicator(0, Qt.AscendingOrder); self.view.sortByColumn(0, Qt.AscendingOrder)


    def _apply_filter(self):
        pattern = self.filter_edit.text().strip()
        if not pattern:

            self._enter_browse_mode()
            return


        self._cancel_search_worker()

        base = self.current_path()


        model = SearchResultModel(self)
        model.setHorizontalHeaderLabels(["Name", "Size", "Date Modified", "Folder"])
        self._enter_search_mode(model)


        self._search_pending_items = {}
        self._search_stats_done = set()
        self._search_stat_worker = None
        self._search_stat_queue = []
        self._search_stat_pending = set()


        w = SearchWorker(base, pattern, self, max_results=SEARCH_RESULT_LIMIT)
        self._search_worker = w
        w.batchReady.connect(self._on_search_batch, Qt.QueuedConnection)
        w.error.connect(lambda msg: self.host.statusBar().showMessage(f"Search error: {msg}", 4000))
        w.truncated.connect(lambda n: self.host.statusBar().showMessage(
            f"Search capped at {n} results. Refine filter to narrow results.", 6000
        ))
        w.finished.connect(self._on_search_finished, Qt.QueuedConnection)

        QApplication.setOverrideCursor(Qt.WaitCursor)
        w.start()


    def _fill_search_visible_icons(self):
        if not self._search_mode or not self._search_proxy or not self._search_model:
            return

        root_ix = self.view.rootIndex()
        vp = self.view.viewport()
        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        rc = self._search_proxy.rowCount(root_ix)
        end = bot_ix.row() if bot_ix.isValid() else min(start + 200, rc - 1)
        start = max(0, start - 40)
        end = min(rc - 1, end + 100)
        if end < start:
            end = start

        paths_need_stat = []

        for r in range(start, end + 1):
            prx_ix = self._search_proxy.index(r, 0, root_ix)
            if not prx_ix.isValid():
                continue
            src_ix = self._search_proxy.mapToSource(prx_ix)
            if not src_ix.isValid():
                continue
            item_name = self._search_model.item(src_ix.row(), 0)
            item_size = self._search_model.item(src_ix.row(), 1)
            item_date = self._search_model.item(src_ix.row(), 2)
            if not item_name:
                continue

            p = item_name.data(Qt.UserRole)
            isdir = bool(item_name.data(IS_DIR_ROLE))


            if p and not bool(item_name.data(SEARCH_ICON_READY_ROLE)):
                idx = self.source_model.index(p)
                if idx.isValid():
                    icon = self.source_model.fileIcon(idx)
                    if icon and not icon.isNull():
                        item_name.setIcon(icon)
                item_name.setData(True, SEARCH_ICON_READY_ROLE)


            if p and not isdir and p not in getattr(self, "_search_stats_done", set()):
                self._search_pending_items[p] = (item_size, item_date)
                paths_need_stat.append(p)
                self._search_stats_done.add(p)

        if paths_need_stat:
            self._enqueue_search_stat_paths(paths_need_stat, batch_limit=220)

    def _build_fallback_new_actions(self, menu: QMenu):
        return {
            menu.addAction(label): (kind, default_name, ext, status)
            for label, kind, default_name, ext, status in self._FALLBACK_NEW_ACTION_SPECS
        }

    def _create_fallback_new_item(self, dst_dir: str, kind: str, default_name: str, ext: str | None):
        if kind == "folder":
            newp = unique_dest_path(dst_dir, default_name)
            os.makedirs(newp, exist_ok=False)
            return
        _create_new_file_with_template(dst_dir, default_name, ext or "")

    def _context_menu_screen_point(self, pos) -> tuple[int, int]:
        try:
            cx, cy = win32api.GetCursorPos()
            return int(cx), int(cy)
        except Exception:
            g = self.view.viewport().mapToGlobal(pos)
            return g.x(), g.y()

    def _try_native_context_menu(self, pos, owner_hwnd: int, paths: list[str]) -> bool:
        if not HAS_PYWIN32:
            return False
        screen_pt = self._context_menu_screen_point(pos)
        if paths:
            return show_explorer_context_menu(owner_hwnd, paths, screen_pt)
        return show_explorer_background_menu(owner_hwnd, self.current_path(), screen_pt)

    def _on_context_menu(self, pos):
        owner_hwnd = int(self.window().winId()) if HAS_PYWIN32 else 0
        paths = self._selected_paths()


        if self._try_native_context_menu(pos, owner_hwnd, paths):
            return


        if paths:
            global_pt = QCursor.pos()
            menu = QMenu(self)

            act_open = menu.addAction("Open")
            act_rename = menu.addAction("Rename")
            if len(paths) != 1:
                act_rename.setEnabled(False)
            act_delete = menu.addAction("Delete")
            menu.addSeparator()
            act_copy = menu.addAction("Copy")
            act_cut = menu.addAction("Cut")
            act_paste = menu.addAction("Paste")

            payload = self.host.get_clipboard() or self._external_clipboard_payload()
            act_paste.setEnabled(bool(payload and payload.get("paths")))

            action = menu.exec_(global_pt)
            if action == act_open:
                self._open_current()
            elif action == act_rename:
                self.rename_selection()
            elif action == act_delete:
                self.delete_selection()
            elif action == act_copy:
                self.copy_selection()
            elif action == act_cut:
                self.cut_selection()
            elif action == act_paste:
                self.paste_into_current()
            return


        dst_dir = self.current_path()
        global_pt = QCursor.pos()
        menu = QMenu(self)
        action_map = self._build_fallback_new_actions(menu)
        action = menu.exec_(global_pt)
        selected = action_map.get(action)
        if not selected:
            return
        kind, default_name, ext, status = selected

        try:
            self._create_fallback_new_item(dst_dir, kind, default_name, ext)
            self.hard_refresh()
            self.host.flash_status(status)
        except Exception as e:
            QMessageBox.critical(self, "Create failed", str(e))


    def _on_selection_changed(self,*_):
        self._request_selection_status_update()
    def _update_statusbar_selection(self):
        self._render_selection_status(update_statusbar=True, update_label=False, update_free=False)

    def _drive_label(self, path:str)->str:
        if path.startswith("\\\\"):
            comps=[c for c in path.split("\\") if c]
            return f"\\\\{comps[0]}\\{comps[1]}" if len(comps)>=2 else "\\\\"
        drv,_=os.path.splitdrive(path); return drv if drv else os.sep

    def _is_network_path(self, path:str)->bool:
        if not path:
            return False
        try:
            path = os.path.abspath(path)
        except Exception:
            pass
        path = path.replace("/", "\\")
        if path.startswith("\\\\"):
            return True
        drv, _ = os.path.splitdrive(path)
        if not drv:
            return False
        try:
            DRIVE_REMOTE = 4
            return ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(drv + "\\")) == DRIVE_REMOTE
        except Exception:
            return False

    def _update_pane_status(self):
        self._render_selection_status(update_statusbar=False, update_label=True, update_free=True)

    def closeEvent(self, e):
        self.shutdown(wait_ms=1000)
        super().closeEvent(e)



class MultiExplorer(QMainWindow):
    namedBookmarksChanged=pyqtSignal(list)
    def __init__(self, pane_count:int=6, start_paths=None, initial_theme:str="dark"):
        super().__init__()
        self.theme=initial_theme if initial_theme in ("dark","light") else "dark"
        self._layout_states=[4,6,8]; self._layout_idx=self._layout_states.index(pane_count) if pane_count in self._layout_states else 1
        self.setWindowTitle(f"Multi-Pane File Explorer - {pane_count} panes"); self.resize(1500,900)


        top=QWidget(self); top_lay=QHBoxLayout(top); top_lay.setContentsMargins(6,2,6,2); top_lay.setSpacing(ROW_SPACING)
        self.btn_layout=QToolButton(top); self.btn_layout.setToolTip("Toggle layout (4 / 6 / 8)"); self.btn_layout.setFixedHeight(UI_H)
        self.btn_theme=QToolButton(top); self.btn_theme.setToolTip("Toggle Light/Dark"); self.btn_theme.setFixedHeight(UI_H)
        self.btn_bm_edit=QToolButton(top); self.btn_bm_edit.setToolTip("Edit Bookmarks"); self.btn_bm_edit.setFixedHeight(UI_H)


        self.btn_session=QToolButton(top)
        self.btn_session.setToolTip("Session (save/load all pane paths)")
        self.btn_session.setFixedHeight(UI_H)

        self.btn_about=QToolButton(top); self.btn_about.setToolTip("About"); self.btn_about.setFixedHeight(UI_H)
        top_lay.addWidget(self.btn_layout,0); top_lay.addWidget(self.btn_theme,0); top_lay.addWidget(self.btn_bm_edit,0)
        top_lay.addWidget(self.btn_session,0)
        top_lay.addWidget(self.btn_about,0); top_lay.addStretch(1)

        self.central=QWidget(self); self.setCentralWidget(self.central)
        vmain=QVBoxLayout(self.central); vmain.setContentsMargins(0,0,0,0); vmain.setSpacing(ROW_SPACING)
        vmain.addWidget(top,0); self.grid=QGridLayout(); vmain.addLayout(self.grid,1)

        self.named_bookmarks=migrate_legacy_favorites_into_named(load_named_bookmarks()); save_named_bookmarks(self.named_bookmarks)
        self._clipboard=None; self._bm_dlg=None

        self._update_layout_icon(); self._update_theme_icon()
        self.btn_layout.clicked.connect(self._cycle_layout); self.btn_theme.clicked.connect(self._toggle_theme)
        self.btn_bm_edit.clicked.connect(self._open_bookmark_editor)
        self.btn_session.clicked.connect(self._open_session_manager)
        self.btn_about.clicked.connect(self._show_about)

        self.panes=[]; self.build_panes(pane_count, start_paths or []); self._update_theme_dependent_icons()
        self._install_focus_tracker()
        if getattr(self, "panes", None):
            self.mark_active_pane(self.panes[0])

        self.statusBar().showMessage("Ready", 1500)

        self._wd_timer = None
        if DEBUG:
            self._wd_timer=QTimer(self); self._wd_timer.setInterval(50); self._wd_last=time.perf_counter()
            def _wd_tick():
                now=time.perf_counter(); gap=(now-self._wd_last)*1000
                if gap>200: dlog(f"[STALL] UI event loop blocked ~{gap:.0f} ms")
                self._wd_last=now
            self._wd_timer.timeout.connect(_wd_tick); self._wd_timer.start()

        settings=QSettings(ORG_NAME, APP_NAME); geo=settings.value("window/geometry")
        if isinstance(geo, QtCore.QByteArray): self._safe_restore_geometry(geo)

    def mark_active_pane(self, pane):
        try:
            self._active_pane = pane
            for p in getattr(self, "panes", []):
                try:
                    is_active = (p is pane)

                    p.path_bar.set_active(is_active)

                    if hasattr(p, "set_active_visual"):
                        p.set_active_visual(is_active)
                except Exception:
                    pass
            dlog(f"[active] pane={getattr(pane, 'pane_id', '?')}")
        except Exception:
            pass

    def _install_focus_tracker(self):
        app = QApplication.instance()
        if not app:
            return

        try:
            self._focus_tracker_connected
        except AttributeError:
            self._focus_tracker_connected = False
        if self._focus_tracker_connected:
            return
        app.focusChanged.connect(self._on_focus_changed)
        self._focus_tracker_connected = True

    def _on_focus_changed(self, old, now):
        try:
            if not now:
                return
            if not isinstance(now, QWidget):
                return
            for p in getattr(self, "panes", []):

                if p.isAncestorOf(now):
                    self.mark_active_pane(p)
                    return
        except Exception:
            pass


    def _update_layout_icon(self):
        states=getattr(self,"_layout_states",[4,6,8]); idx=getattr(self,"_layout_idx",0)
        if not states: states=[4,6,8]
        if idx>=len(states) or idx<0: idx=0; self._layout_idx=0
        state=states[idx]; self.btn_layout.setIcon(icon_grid_layout(state, self.theme))
    def _update_theme_icon(self):
        if hasattr(self, "btn_theme") and self.btn_theme:
            self.btn_theme.setIcon(icon_theme_toggle(self.theme))
        if hasattr(self, "btn_bm_edit") and self.btn_bm_edit:
            self.btn_bm_edit.setIcon(icon_bookmark_edit(self.theme))
        if hasattr(self, "btn_session") and self.btn_session:
            self.btn_session.setIcon(icon_session(self.theme))
        if hasattr(self, "btn_about") and self.btn_about:
            self.btn_about.setIcon(icon_info(self.theme))

    def _update_theme_dependent_icons(self):
        self._update_layout_icon(); self._update_theme_icon()
        for p in getattr(self,"panes",[]):
            try:
                p.btn_star.setIcon(icon_star(p.btn_star.isChecked(), self.theme))
                p.btn_cmd.setIcon(icon_cmd(self.theme))

                if getattr(getattr(p, "path_bar", None), "_btn_copy", None):
                    p.path_bar._btn_copy.setIcon(icon_copy_squares(self.theme))
            except Exception:
                pass

    def _cycle_layout(self):
        self._layout_idx=(self._layout_idx+1)%len(self._layout_states); n=self._layout_states[self._layout_idx]; self.build_panes(n, self._current_paths())
    def _toggle_theme(self):
        self.theme="light" if self.theme=="dark" else "dark"
        app=QApplication.instance()
        apply_dark_style(app) if self.theme=="dark" else apply_light_style(app)
        self._update_theme_dependent_icons()
        s=QSettings(ORG_NAME, APP_NAME); s.setValue("ui/theme", self.theme); s.sync()

    def build_panes(self, n:int, start_paths):
        was_max = self.isMaximized()


        try:
            prev_count = len(getattr(self, "panes", []))
        except Exception:
            prev_count = 0
        if prev_count > 0:
            try:
                prev_paths = self._current_paths()
                s = QSettings(ORG_NAME, APP_NAME)
                s.setValue(f"layout/last_paths_{prev_count}", prev_paths)
                s.sync()
            except Exception:
                pass




        final_paths = list(start_paths or [])[:n]
        if len(final_paths) < n:
            s = QSettings(ORG_NAME, APP_NAME)
            saved = s.value(f"layout/last_paths_{n}", [])
            if not isinstance(saved, list):
                saved = []
            base_len = len(final_paths)
            for i in range(base_len, n):
                cand = saved[i] if i < len(saved) else None
                if not cand or not os.path.exists(str(cand)):
                    cand = QDir.homePath()
                final_paths.append(str(cand))


        old_panes = list(getattr(self, "panes", []))
        for p in old_panes:
            try:
                p.shutdown(wait_ms=600)
            except Exception:
                pass

        vmain = self.centralWidget().layout() if self.centralWidget() else None
        if hasattr(self, "grid") and isinstance(self.grid, QGridLayout):
            while self.grid.count():
                it = self.grid.takeAt(0)
                w = it.widget()
                if w:
                    w.setParent(None)
                    w.deleteLater()
            if vmain:
                try:
                    vmain.removeItem(self.grid)
                except Exception:
                    pass
            try:
                self.grid.setParent(None)
            except Exception:
                pass


        cols = {4: 2, 6: 3, 8: 4}.get(n, 3)
        gap = GRID_GAPS.get(cols, 3)
        margin_lr = GRID_MARG_LR.get(cols, 6)

        self.grid = QGridLayout()
        self.grid.setSpacing(gap)
        self.grid.setContentsMargins(margin_lr, 2, margin_lr, 4)
        if vmain:
            vmain.addLayout(self.grid, 1)


        for c in range(cols):
            self.grid.setColumnStretch(c, 1)
            self.grid.setColumnMinimumWidth(c, 0)
        rows = (n + cols - 1) // cols
        for r in range(rows):
            self.grid.setRowStretch(r, 1)
            self.grid.setRowMinimumHeight(r, 0)


        self.panes = []
        self.setUpdatesEnabled(False)
        for i in range(n):
            spath = final_paths[i] if i < len(final_paths) else None
            pane = ExplorerPane(None, start_path=spath, pane_id=i + 1, host_main=self)
            self.panes.append(pane)
            rr = i // cols
            cc = i % cols
            self.grid.addWidget(pane, rr, cc)
        self.setUpdatesEnabled(True)


        self.setWindowTitle(f"Multi-Pane File Explorer - {n} panes")
        self._update_theme_dependent_icons()


        if was_max:
            QTimer.singleShot(0, self._unmax_then_remax)
        else:
            QTimer.singleShot(0, self._kick_layout)


    def _unmax_then_remax(self):
        try:

            if not self.isMaximized():
                self._kick_layout()
                return


            self.showNormal()
            QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.LayoutRequest)
            QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents)


            self._kick_layout()


            self.showMaximized()
            QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.LayoutRequest)
            QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents)


            self._kick_layout()
        except Exception:
            pass

    def _kick_layout(self):
        try:
            cw = self.centralWidget()
            lay = cw.layout() if cw else None

            keep_geom = not self.isMaximized()
            if keep_geom:
                g = self.geometry()

                old_min = self.minimumSize()
                old_max = self.maximumSize()
                self.setMinimumSize(g.size())
                self.setMaximumSize(g.size())


            if lay:
                lay.invalidate()
                lay.activate()

            if hasattr(self, "grid"):
                self.grid.invalidate()
                self.grid.update()


            for p in getattr(self, "panes", []):
                try:
                    p.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                    if hasattr(p, "view") and p.view:
                        p.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                        p.view.updateGeometries()
                        p.view.doItemsLayout()
                        p.view.viewport().update()

                    if hasattr(p, "path_bar") and hasattr(p.path_bar, "_pin_to_right"):
                        p.path_bar._pin_to_right()
                except Exception:
                    pass


            QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.LayoutRequest)
            QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents)

            if lay:
                lay.activate()
            if hasattr(self, "grid"):
                self.grid.invalidate()
            self.repaint()

        except Exception:
            pass
        finally:

            try:
                if keep_geom:

                    try:
                        self.setMinimumSize(old_min)
                        self.setMaximumSize(old_max)
                    except Exception:

                        self.setMinimumSize(QtCore.QSize(0, 0))
                        self.setMaximumSize(QtCore.QSize(16777215, 16777215))

                    if self.geometry() != g:
                        self.setGeometry(g)
            except Exception:
                pass


    def _safe_restore_geometry(self, ba: QtCore.QByteArray):
        try:
            ok = self.restoreGeometry(ba)
            if not ok:
                return


            win = self.windowHandle()
            screen = (win.screen().availableGeometry() if win and win.screen()
                      else QApplication.primaryScreen().availableGeometry())
            sg = screen

            g = self.geometry()
            new_w = min(max(600, g.width()), sg.width())
            new_h = min(max(350, g.height()), sg.height())
            new_x = min(max(sg.left(), g.x()), sg.right() - new_w)
            new_y = min(max(sg.top(),  g.y()), sg.bottom() - new_h)


            if (new_w != g.width()) or (new_h != g.height()) or (new_x != g.x()) or (new_y != g.y()):
                self.setGeometry(new_x, new_y, new_w, new_h)
        except Exception:
            pass


    def _current_paths(self): return [p.current_path() for p in self.panes]
    def set_clipboard(self,payload:dict): self._clipboard=payload
    def get_clipboard(self): return self._clipboard
    def clear_clipboard(self): self._clipboard=None
    def flash_status(self,text:str):
        try: self.statusBar().showMessage(text,2000)
        except Exception: pass


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
            "<div style='color:#000; font-size:12pt;'><b>Multi-Pane File Explorer v2.0.0</b></div>"
            "<div style='color:#111; margin-top:6px;'>A compact multi-pane file explorer for Windows (PyQt5).</div>"
            "<div style='color:#111; margin-top:6px;'>For feedback, contact <b>kkongt2.kang</b>.</div>"
        )
        lay.addWidget(lbl)
        btns=QDialogButtonBox(QDialogButtonBox.Ok, dlg); lay.addWidget(btns); btns.accepted.connect(dlg.accept)
        pal=dlg.palette(); pal.setColor(QPalette.Window, QColor(255,255,255)); pal.setColor(QPalette.WindowText, QColor(0,0,0))
        dlg.setPalette(pal); dlg.setStyleSheet("QLabel { color: #000; } QDialog { background: #FFF; }")
        dlg.resize(380,180); dlg.exec_()

    def closeEvent(self,e):
        paths = []
        try:
            paths = self._current_paths()
        except Exception:
            paths = []

        for p in list(getattr(self, "panes", [])):
            try:
                p.shutdown(wait_ms=1000)
            except Exception:
                pass

        settings=QSettings(ORG_NAME, APP_NAME)
        settings.setValue("window/geometry", self.saveGeometry())
        settings.setValue("layout/pane_count", len(paths) if paths else len(self.panes))
        for i,p in enumerate(paths if paths else [x.current_path() for x in self.panes]):
            settings.setValue(f"layout/pane_{i}_path", p)
        settings.sync(); super().closeEvent(e)


    def _get_sessions(self) -> list:
        s = QSettings(ORG_NAME, APP_NAME)
        val = s.value("sessions/items", [])
        out = []
        if isinstance(val, list):
            for it in val:
                try:
                    name = str(it.get("name", "")).strip()
                    paths = [str(p) for p in it.get("paths", [])]
                    panes = int(it.get("panes", len(paths) or len(self.panes) or 6))
                    ts = float(it.get("ts", time.time()))
                    if name and paths:
                        out.append({"name": name, "paths": paths, "panes": panes, "ts": ts})
                except Exception:
                    pass
        return out

    def _set_sessions(self, items: list):
        s = QSettings(ORG_NAME, APP_NAME)
        s.setValue("sessions/items", items); s.sync()

    def _save_session(self, name: str):
        name = (name or "").strip()
        if not name:
            QMessageBox.information(self, "Save Session", "Please enter a session name.")
            return
        paths = self._current_paths()
        panes = len(self.panes)
        items = self._get_sessions()

        lowered = name.lower()
        replaced = False
        for i, it in enumerate(items):
            if it.get("name","").lower() == lowered:
                items[i] = {"name": name, "paths": paths, "panes": panes, "ts": time.time()}
                replaced = True
                break
        if not replaced:
            items.append({"name": name, "paths": paths, "panes": panes, "ts": time.time()})
        self._set_sessions(items)
        try: self.statusBar().showMessage(f"Session '{name}' saved.", 2000)
        except Exception: pass

    def _delete_session(self, name: str):
        items = self._get_sessions()
        new_items = [it for it in items if it.get("name","") != name]
        self._set_sessions(new_items)

    def _load_session(self, name: str):
        items = self._get_sessions()
        target = None
        for it in items:
            if it.get("name","") == name:
                target = it; break
        if not target:
            QMessageBox.warning(self, "Load Session", "Session not found."); return
        paths = list(target.get("paths", []))
        panes = int(target.get("panes", len(paths)))
        if panes <= 0 or not paths:
            QMessageBox.warning(self, "Load Session", "Session data is empty or invalid."); return


        if panes != len(self.panes):
            self.build_panes(panes, paths)
        else:
            for i, p in enumerate(paths):
                if i < len(self.panes) and os.path.exists(p):
                    try:
                        self.panes[i].set_path(p, push_history=False)
                    except Exception:
                        pass
        try: self.statusBar().showMessage(f"Session '{name}' loaded.", 2000)
        except Exception: pass

    def _open_session_manager(self):
        dlg = SessionManagerDialog(self, self._get_sessions())
        if dlg.exec_() == QDialog.Accepted:
            pass

class SessionManagerDialog(QDialog):
    def __init__(self, parent: MultiExplorer, sessions: list):
        super().__init__(parent)
        self.setWindowTitle("Session Manager")
        self.resize(560, 400)

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Name", "Panes", "Saved"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.btn_save = QPushButton("Save Current", self)
        self.btn_load = QPushButton("Load Selected", self)
        self.btn_delete = QPushButton("Delete Selected", self)
        self.btn_close = QPushButton("Close", self)

        btns = QHBoxLayout()
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_load)
        btns.addWidget(self.btn_delete)
        btns.addStretch(1)
        btns.addWidget(self.btn_close)

        lay = QVBoxLayout(self)
        lay.addWidget(self.table, 1)
        lay.addLayout(btns)

        self._sessions = []
        self.set_sessions(sessions)

        self.btn_close.clicked.connect(self.accept)
        self.btn_load.clicked.connect(self._on_load)
        self.btn_delete.clicked.connect(self._on_delete)
        self.btn_save.clicked.connect(self._on_save)

    def set_sessions(self, items: list):
        self._sessions = list(items or [])
        self.table.setRowCount(len(self._sessions))
        for r, it in enumerate(self._sessions):
            name = it.get("name","")
            panes = int(it.get("panes", 0))
            ts = float(it.get("ts", time.time()))
            dt = QDateTime.fromSecsSinceEpoch(int(ts)).toString("yyyy-MM-dd HH:mm:ss")

            self.table.setItem(r, 0, QTableWidgetItem(name))
            self.table.setItem(r, 1, QTableWidgetItem(str(panes)))
            self.table.setItem(r, 2, QTableWidgetItem(dt))
        self.table.resizeColumnsToContents()

    def _selected_name(self) -> str | None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        r = rows[0].row()
        it = self.table.item(r, 0)
        return it.text().strip() if it else None

    def _on_load(self):
        name = self._selected_name()
        if not name:
            QMessageBox.information(self, "Load Session", "Please select a session.")
            return
        try:
            self.parent()._load_session(name)
        except Exception as e:
            QMessageBox.critical(self, "Load Session", str(e))

    def _on_delete(self):
        name = self._selected_name()
        if not name:
            QMessageBox.information(self, "Delete Session", "Please select a session.")
            return
        btn = QMessageBox.question(self, "Delete Session", f"Delete this session?\n{name}",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if btn != QMessageBox.Yes:
            return
        try:
            self.parent()._delete_session(name)
            self.set_sessions(self.parent()._get_sessions())
        except Exception as e:
            QMessageBox.critical(self, "Delete Session", str(e))

    def _on_save(self):
        name, ok = QInputDialog.getText(self, "Save Session", "Session name:",
                                        text=time.strftime("Session %Y-%m-%d %H-%M-%S"))
        if not ok or not name.strip():
            return
        name = name.strip()

        exists = any(s.get("name","").lower() == name.lower() for s in self.parent()._get_sessions())
        if exists:
            btn = QMessageBox.question(self, "Save Session",
                                       f"A session with this name already exists.\nOverwrite it?\n{name}",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if btn != QMessageBox.Yes:
                return
        try:
            self.parent()._save_session(name)
            self.set_sessions(self.parent()._get_sessions())
        except Exception as e:
            QMessageBox.critical(self, "Save Session", str(e))



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
        btn = QToolButton(path_wrap); btn.setText("..."); btn.setFixedHeight(UI_H)
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
    ap.add_argument("--debug", action="store_true", help="Enable debug logs (or set MULTIPANE_DEBUG=1)")
    return ap.parse_args()

def main():
    global DEBUG
    args=parse_args()
    DEBUG = bool(args.debug or _env_flag("MULTIPANE_DEBUG"))
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

