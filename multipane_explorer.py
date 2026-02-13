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
    import win32con, win32gui, win32api, win32clipboard
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

def icon_bookmark_edit(theme: str):
    """
    '북마크 편집' 아이콘:
      - 큼직한 노란 별(⭐)
      - 우하단에 연필(편집) 오버레이
    """
    def paint(p: QPainter, w, h):
        p.setRenderHint(QPainter.Antialiasing, True)

        # ── 별(북마크)
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

        # ── 연필(편집) 오버레이
        p.save()
        body = QColor(100, 180, 255) if theme == "dark" else QColor(40, 120, 220)
        tip  = QColor(240, 200, 80)

        # 우하단에 기울여 배치
        p.translate(w * 0.64, h * 0.68)
        p.rotate(-25)

        # 연필 몸통(정수 좌표 OK)
        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(body))
        p.drawRect(-5, -2, 12, 4)

        # 연필 촉(삼각형)
        p.setBrush(QBrush(tip))
        tri = QPolygonF([
            QtCore.QPointF(7, -2),
            QtCore.QPointF(7,  2),
            QtCore.QPointF(10, 0)
        ])
        p.drawPolygon(tri)

        # 지우개 부분 (부동소수 → QRectF 사용)
        eraser = QColor(230, 230, 240) if theme == "dark" else QColor(250, 250, 255)
        p.setBrush(QBrush(eraser))
        p.drawRect(QtCore.QRectF(-6.5, -2.2, 2.6, 4.4))

        p.restore()

    return _make_icon(22, 22, paint)


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

        /* 기본 crumb 버튼 */
        QPushButton#crumb {
            background: rgba(255,255,255,0.05);
            border: 1px solid #2B2E34;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #E6E9EE;
        }
        QPushButton#crumb:hover { background: rgba(255,255,255,0.09); }
        QLabel#crumbSep { color: #7F8796; }

        /* breadcrumb 바(스크롤 영역) 자체 외곽선/배경 제거 */
        QScrollArea#crumbScroll { border: 0px solid transparent; }
        QScrollArea#crumbScroll[active="true"] { border: 0px solid transparent; }
        QScrollArea#crumbScroll > QWidget#crumbViewport { background: transparent; }

        /* 활성 Pane일 때: crumb 하나하나만 은은한 파랑 배경 */
        QWidget#paneRoot[active="true"] QPushButton#crumb {
            background: rgba(94,155,255,0.16);
            border-color: rgba(94,155,255,0.40);
        }
        QWidget#paneRoot[active="true"] QPushButton#crumb:hover {
            background: rgba(94,155,255,0.22);
        }

        /* Pane 전체 하이라이트(유지) */
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

        /* 기본 crumb 버튼 */
        QPushButton#crumb {
            background: rgba(0,0,0,0.04);
            border: 1px solid #E5E8EE;
            padding: 0 6px; border-radius: 6px; text-align: left; color: #1C1C1E;
        }
        QPushButton#crumb:hover { background: rgba(0,0,0,0.07); }
        QLabel#crumbSep { color: #7A7F89; }

        /* breadcrumb 바(스크롤 영역) 자체 외곽선/배경 제거 */
        QScrollArea#crumbScroll { border: 0px solid transparent; }
        QScrollArea#crumbScroll[active="true"] { border: 0px solid transparent; }
        QScrollArea#crumbScroll > QWidget#crumbViewport { background: transparent; }

        /* 활성 Pane일 때: crumb 하나하나만 은은한 파랑 배경 */
        QWidget#paneRoot[active="true"] QPushButton#crumb {
            background: rgba(64,128,255,0.12);
            border-color: rgba(64,128,255,0.40);
        }
        QWidget#paneRoot[active="true"] QPushButton#crumb:hover {
            background: rgba(64,128,255,0.18);
        }

        /* Pane 전체 하이라이트(유지) */
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


# -------------------- Vector Icons --------------------
def icon_copy_squares(theme: str):
    def paint(p: QPainter, w, h):
        # 색상: 흰 채움 + 테마별 스트로크
        stroke = QColor(210, 214, 225) if theme == "dark" else QColor(85, 95, 115)
        fill   = QColor(255, 255, 255)

        p.setRenderHint(QPainter.Antialiasing, True)
        pen = QPen(stroke, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        p.setPen(pen)
        p.setBrush(QBrush(fill))

        # 슬래시(/) 방향 배치
        front_rect = QtCore.QRect(6, 3, 11, 11)  # 위쪽(앞) 사각형
        back_rect  = QtCore.QRect(3, 6, 11, 11)  # 아래쪽(뒤) 사각형
        radius = 3

        # 1) 위 사각형 (흰 채움 + 윤곽)
        p.drawRoundedRect(front_rect, radius, radius)

        # 2) 아래 사각형을 마지막으로 그려 앞 사각형 윤곽을 덮음
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
    """
    '세션' 아이콘:
      - 겹쳐진 3장의 탭(여러 창/상태를 저장하는 느낌)
      - 우측하단에 작은 별로 '저장된 세션' 강조
    """
    def paint(p: QPainter, w, h):
        p.setRenderHint(QPainter.Antialiasing, True)

        # 색 설정
        line = QColor(190, 195, 210) if theme == "dark" else QColor(90, 100, 120)
        fill = QColor(60, 66, 80) if theme == "dark" else QColor(245, 247, 250)
        tab  = QColor(100, 150, 255) if theme == "dark" else QColor(80, 120, 230)
        star_fill  = QColor(255, 210, 60)
        star_edge  = QColor(160, 120, 0)

        # 겹친 탭 3장
        p.setPen(QPen(line, 1.3))
        p.setBrush(QBrush(fill))
        p.drawRoundedRect(QtCore.QRectF(3.0, 6.5, 12.5, 9.0), 2.5, 2.5)
        p.drawRoundedRect(QtCore.QRectF(5.0, 5.0, 12.5, 9.0), 2.5, 2.5)
        p.setBrush(QBrush(tab))
        p.drawRoundedRect(QtCore.QRectF(7.0, 3.5, 12.5, 9.0), 2.5, 2.5)

        # 우하단 별(작지만 선명하게)
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

    # ── 속성(프로퍼티) 처리: 환경별 3단 폴백 ─────────────────────────
    if verb and verb.lower() in ("properties","prop","property"):
        try:
            target = paths[0] if (paths and len(paths)>0) else work_dir
            target = _normalize_fs_path(target)

            ok = False
            # 1) pywin32가 제공하는 SHObjectProperties가 있으면 사용
            try:
                fn = getattr(shell, "SHObjectProperties", None)
                if callable(fn):
                    # SHOP_FILEPATH = 0x00000002
                    fn(int(owner_hwnd), 0x00000002, target, None)
                    ok = True
            except Exception as e:
                if DEBUG: print("[ctx] SHObjectProperties (pywin32) failed:", e)

            # 2) ShellExecuteEx + 'properties' verb
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

            # 3) ctypes로 SHObjectPropertiesW 직접 호출 (확실한 최후 폴백)
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
            if c==0:
                return r["name"]
            if c==1:
                # ▶ 폴더의 크기는 표시하지 않음 (정렬용 값은 EditRole/SIZE_BYTES_ROLE에서 0으로 유지)
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
                # ▶ 폴더는 0, 파일은 실제 바이트 (정렬에 사용)
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
            # ▶ 폴더는 0, 파일은 실제 바이트
            if r.get("is_dir", False): return 0
            return 0 if r["size"] is None else int(r["size"])

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
        self.setObjectName("pathbar")

        self._host=QWidget(self); self._hlay=QHBoxLayout(self._host)
        self._hlay.setContentsMargins(4,0,4,0); self._hlay.setSpacing(max(0, ROW_SPACING-2))

        self._scroll=QScrollArea(self); self._scroll.setObjectName("crumbScroll")
        self._scroll.setWidget(self._host)
        self._scroll.setWidgetResizable(True); self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff); self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setViewportMargins(0,0,0,0); self._scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._scroll.setFixedHeight(UI_H); self.setFixedHeight(UI_H)

        # ── viewport에도 이름/배경 허용 속성을 부여(스타일 대상) ──
        try:
            vp = self._scroll.viewport()
            vp.setObjectName("crumbViewport")
            vp.setAttribute(Qt.WA_StyledBackground, True)
        except Exception:
            pass

        # 활성 표시 초기값
        self._scroll.setProperty("active", False)

        self._edit=QLineEdit(self); self._edit.hide(); self._edit.setClearButtonEnabled(True); self._edit.setFixedHeight(UI_H)
        self._edit.returnPressed.connect(self._on_edit_return)

        # --- Copy Path 버튼 (우측 고정) ---
        self._btn_copy = QToolButton(self)
        self._btn_copy.setToolTip("Copy current path")
        self._btn_copy.setFixedHeight(UI_H)
        # 아이콘: 테마 감지
        theme = getattr(getattr(parent, "host", None), "theme", "dark")
        try:
            self._btn_copy.setIcon(icon_copy_squares(theme))
        except Exception:
            # 폴백: 텍스트
            self._btn_copy.setText("Copy")
        self._btn_copy.clicked.connect(self._copy_current_path)

        # 레이아웃: [scroll(breadcrumb) | edit(overlap) | copy-button]
        wrap=QHBoxLayout(self); wrap.setContentsMargins(0,0,0,0); wrap.setSpacing(0)
        wrap.addWidget(self._scroll, 1)
        wrap.addWidget(self._edit, 1)
        wrap.addWidget(self._btn_copy, 0)

        self._host.installEventFilter(self); self._edit.installEventFilter(self)
        self.set_path(self._current_path)

    def _copy_current_path(self):
        # 편집 중이면 입력값 우선, 아니면 현재 경로
        t = self._edit.text().strip() if self._edit.isVisible() else self._current_path
        if not t:
            t = self._current_path
        QApplication.clipboard().setText(t)
        QToolTip.showText(QCursor.pos(), f"Copied: {t}", self)


    def set_active(self, active: bool):
        """
        Path bar를 활성/비활성 시각 상태로 전환한다.
        (QSS: QScrollArea#crumbScroll[active="true"] 와 그 viewport)
        """
        try:
            self._scroll.setProperty("active", bool(active))
            # 재적용 (scroll + viewport 모두)
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

        # 빌드 직후에도 항상 오른쪽(가장 하위 폴더)이 보이도록 고정
        QTimer.singleShot(0, self._pin_to_right)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        # 리사이즈 때도 우측 고정 유지
        QTimer.singleShot(0, self._pin_to_right)

    def _pin_to_right(self):
        try:
            if hasattr(self, "_hbar") and self._hbar:
                self._hbar.setValue(self._hbar.maximum())
        except Exception:
            pass

# -------------------- SearchResultModel --------------------
class SearchResultModel(QStandardItemModel):
    """
    검색 결과 전용 모델.
    - Size 컬럼은 사람이 읽기 쉬운 단위로 표시
    - 외부 드래그를 위해 text/uri-list (file://...) 로 QMimeData를 생성
    """
    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and index.column() == 1:
            b = super().data(index, SIZE_BYTES_ROLE)
            if b is None:
                b = super().data(index, Qt.EditRole)
            if isinstance(b, (int, float)) and b:
                return human_size(int(b))
            return ""
        return super().data(index, role)

    # ---- 드래그 지원: 외부 앱으로 file:// URL 전달 ----
    def mimeTypes(self):
        # 외부 드롭 타겟들이 인식하는 표준 파일 목록 MIME
        return ["text/uri-list"]

    def mimeData(self, indexes):
        md = QtCore.QMimeData()
        # 같은 행이 여러 컬럼으로 중복 들어오는 것을 방지
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
            path = it.data(Qt.UserRole)  # ExplorerPane._apply_filter 에서 저장한 실제 경로
            if path:
                urls.append(QUrl.fromLocalFile(path))

        md.setUrls(urls)
        return md

    def flags(self, index):
        # 선택/활성 + 드래그 가능 플래그 부여
        f = super().flags(index)
        if index.isValid():
            f |= Qt.ItemIsDragEnabled
        return f

    def supportedDragActions(self):
        # 보통 외부 드래그는 '복사' 의미로 전달
        return Qt.CopyAction

    def startDrag(self, supportedActions):
        # 외부 앱으로 드래그할 때 항상 파일 URL과 텍스트를 함께 제공
        from PyQt5.QtGui import QDrag  # 로컬 임포트(추가 전역 임포트 불필요)
        md = QtCore.QMimeData()

        # 현재 선택 경로(브라우즈/검색 모드 모두 대응)
        paths = [p for p in self.pane._selected_paths() if p and os.path.exists(p)]
        if not paths:
            return

        # 호환성: text/uri-list + text/plain
        md.setUrls([QUrl.fromLocalFile(p) for p in paths])
        md.setText("\r\n".join(paths))

        drag = QDrag(self)
        drag.setMimeData(md)

        # 기본은 Copy로(일부 앱이 Move 거부하는 문제 방지)
        drag.exec_(Qt.CopyAction | Qt.MoveAction, Qt.CopyAction)


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
        self._drag_start_pos = None
        self._drag_start_index = QtCore.QModelIndex()
        self._drag_start_modifiers = Qt.NoModifier
        self._drag_start_was_selected = False
        self._drag_ready = False
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
        if e.button() == Qt.LeftButton:
            self._drag_start_pos = e.pos()
            self._drag_start_index = self.indexAt(e.pos())
            self._drag_start_modifiers = e.modifiers()
            sm = self.selectionModel()
            self._drag_start_was_selected = bool(self._drag_start_index.isValid() and sm and sm.isSelected(self._drag_start_index))
            self._drag_ready = False

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

    def mouseMoveEvent(self, e):
        if (e.buttons() & Qt.LeftButton) and self._drag_start_pos is not None:
            # 일부 Windows 환경에서 드래그-아웃 대신 내부 다중선택이 시작되는 문제를 회피
            if (e.pos() - self._drag_start_pos).manhattanLength() >= QApplication.startDragDistance():
                if not self._drag_ready:
                    ix = self._drag_start_index
                    sm = self.selectionModel()
                    if ix.isValid() and sm and sm.isSelected(ix):
                        # Shift는 범위 선택 제스처이므로 강제 드래그를 막는다.
                        # Ctrl은 클릭 시점에 이미 선택된 항목에서 시작한 경우(복사 드래그 의도)만 허용한다.
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


# -------------------- Explorer Pane --------------------
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

        # 저장된 정렬 설정 복원
        self._load_sort_settings()

        self.set_path(start_path or QDir.homePath(), push_history=False)
        self._update_star_button(); self._rebuild_quick_bookmark_buttons()

        self._connect_signals()
        self._register_shortcuts()

        # 초기 정렬 적용 (기본값: Name 오름차순)
        self._apply_saved_sort()
        self._update_pane_status()

    def _init_state(self):
        self._search_mode=False; self._search_model=None; self._search_proxy=None
        self._back_stack=[]; self._fwd_stack=[]; self._undo_stack=[]
        self._last_hover_index=QtCore.QModelIndex(); self._tooltip_last_ms=0.0; self._tooltip_interval_ms=180; self._tooltip_last_text=""
        self._fast_model=FastDirModel(self); self._fast_proxy=FsSortProxy(self); self._fast_proxy.setSourceModel(self._fast_model)
        self._using_fast=False; self._fast_stat_worker=None; self._enum_worker=None; self._pending_normal_root=None
        self._dirload_timer={}
        self._sort_column = 0  # 기본값: Name
        self._sort_order = Qt.AscendingOrder

    def _build_toolbar(self):
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
        return row_toolbar

    def _build_path_row(self):
        self.path_bar=PathBar(self); self.path_bar.setToolTip("Breadcrumb — Double-click or F4/Ctrl+L to enter path")
        row_path=QHBoxLayout(); row_path.setContentsMargins(0,0,0,0); row_path.setSpacing(0); row_path.addWidget(self.path_bar,1)
        return row_path

    def _build_filter_row(self):
        self.filter_label=QLabel("Filter:", self)
        self.filter_edit=QLineEdit(self); self.filter_edit.setPlaceholderText("Filter (*.pdf, *file*.xls*, …)"); self.filter_edit.setClearButtonEnabled(True); self.filter_edit.setFixedHeight(UI_H)
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
        self._generic_icons=GenericIconProvider(self.style())
        if ALWAYS_GENERIC_ICONS: self.source_model.setIconProvider(self._generic_icons)

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

    def _configure_header_browse(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, 90)
        header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)

    def _configure_header_fast(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.resizeSection(1, 90)
        header.resizeSection(3, 150)
        self.view.setColumnHidden(2, True)

    def _configure_header_search(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.resizeSection(1, 90)
        header.resizeSection(2, 150)

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
        self.btn_new_file.clicked.connect(self.create_text_file)  # 새 문서 생성
        self.btn_refresh.clicked.connect(self.hard_refresh)  # 하드 리프레시
        self.view.activated.connect(self._on_double_click)
        self.view.viewport().installEventFilter(self)
        self.view.installEventFilter(self)
        self.path_bar.installEventFilter(self)
        self.filter_edit.installEventFilter(self)  # ← 추가: 필터창에서 ESC 감지
        self._sel_model = None
        self._hook_selection_model()
        self.filter_edit.returnPressed.connect(self._apply_filter)
        self.btn_search.clicked.connect(self._apply_filter)
        self.filter_edit.textChanged.connect(self._on_filter_text_changed)  # ← 추가: x로 지우면 브라우즈로
        try: self.view.verticalScrollBar().valueChanged.connect(lambda _v: self._schedule_visible_stats())
        except Exception: pass

    def _register_shortcuts(self):
        def add_sc(seq, slot):
            sc=QShortcut(QKeySequence(seq), self.view)
            sc.setContext(Qt.WidgetWithChildrenShortcut); sc.activated.connect(slot); return sc
        add_sc("Backspace", self.go_back); add_sc("Alt+Left", self.go_back); add_sc("Alt+Right", self.go_forward)
        add_sc("Alt+Up", self.go_up)  # ← Alt+Up으로 상위 폴더 이동
        add_sc("Ctrl+L", self.path_bar.start_edit); add_sc("F4", self.path_bar.start_edit)
        add_sc("F3", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+F", lambda:(self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        add_sc("Ctrl+C", self.copy_selection); add_sc("Ctrl+X", self.cut_selection); add_sc("Ctrl+V", self.paste_into_current);add_sc("Ctrl+Z", self.undo_last)
        add_sc("Delete", self.delete_selection); add_sc("Shift+Delete", lambda: self.delete_selection(permanent=True)); add_sc("F2", self.rename_selection)
        add_sc(Qt.Key_Return, self._open_current); add_sc(Qt.Key_Enter, self._open_current); add_sc("Ctrl+O", self._open_current)

        # 경로 복사 단축키
        add_sc("Ctrl+Shift+C", lambda: self._copy_path_shortcut(False))
        add_sc("Alt+Shift+C",  lambda: self._copy_path_shortcut(True))

    def _load_sort_settings(self):
        """저장된 정렬 설정 불러오기"""
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            self._sort_column = s.value(f"pane_{self.pane_id}/sort_column", 0, type=int)
            order_val = s.value(f"pane_{self.pane_id}/sort_order", Qt.AscendingOrder, type=int)
            self._sort_order = Qt.DescendingOrder if order_val == Qt.DescendingOrder else Qt.AscendingOrder
        except Exception:
            self._sort_column = 0
            self._sort_order = Qt.AscendingOrder
    
    def _save_sort_settings(self):
        """현재 정렬 설정 저장"""
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            s.setValue(f"pane_{self.pane_id}/sort_column", self._sort_column)
            s.setValue(f"pane_{self.pane_id}/sort_order", int(self._sort_order))
            s.sync()
        except Exception:
            pass
    
    def _apply_saved_sort(self):
        """저장된 정렬 설정을 현재 뷰에 적용"""
        try:
            v = self.view
            if not v.isSortingEnabled():
                v.setSortingEnabled(True)
            v.header().setSortIndicator(self._sort_column, self._sort_order)
            v.sortByColumn(self._sort_column, self._sort_order)
        except Exception:
            pass


    def set_active_visual(self, active: bool):
        """
        Pane 전체에 'active' 속성을 주어 QSS 하이라이트가 적용되도록 한다.
        - 루트 위젯 objectName을 'paneRoot'로 보장
        - 스타일 재적용(repolish)로 즉시 반영
        - breadcrumb의 각 crumb(QPushButton#crumb)까지 확실히 repolish
        """
        try:
            # paneRoot 보장 + active 속성 토글
            if self.objectName() != "paneRoot":
                self.setObjectName("paneRoot")
            self.setProperty("active", bool(active))

            # 루트와 주요 자식 위젯들 repolish
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

            # breadcrumb의 모든 crumb 버튼도 명시적으로 repolish
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
        """
        뷰의 모델을 교체하면 selectionModel도 새로 생긴다.
        항상 최신 selectionModel에 selectionChanged를 다시 연결한다.
        """
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

        # 현재 상태를 즉시 반영
        QTimer.singleShot(0, self._update_pane_status)

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

    def _on_filter_text_changed(self, text: str):
        # 공백 제거 후 비어 있으면 검색 결과 표시를 중단하고 브라우즈로 복귀
        if not (text or "").strip():
            self._enter_browse_mode()

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
        생성 후 목록을 새로고침하고, 방금 만든 파일을 선택한 뒤
        뷰로 포커스를 되돌려 단축키(Del, F2 등)가 곧바로 동작하도록 합니다.
        """
        base_dir = self.current_path()
        try:
            name = f"New Document {time.strftime('%Y%m%d-%H%M%S')}.txt"
            new_path = _create_new_file_with_template(base_dir, name, ".txt")
        except Exception as e:
            QMessageBox.critical(self, "Create failed", str(e))
            return

        # 목록 갱신
        self.hard_refresh()

        # 방금 만든 파일을 선택하고 포커스를 뷰로 돌리는 시도
        def _try_select():
            # 단축키가 바로 동작하도록 뷰로 포커스 강제
            try:
                if self.view and not self.view.hasFocus():
                    self.view.setFocus(Qt.ShortcutFocusReason)
            except Exception:
                pass

            # 검색 모드에선 선택 스킵 (탐색 모드에서만 선택)
            if self._search_mode:
                return

            try:
                if self._using_fast:
                    # Fast 모델에서 경로로 행 찾기
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
                    # Normal 모델 경로 → 인덱스 매핑
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

        # fast→normal 전환 타이밍을 고려해 여러 번 재시도
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
        w=self._fast_stat_worker
        if w and w.isRunning(): w.cancel(); w.wait(100)
        self._fast_stat_worker=None

    def _schedule_visible_stats(self):
        """
        화면에 보이는 영역에 대해서만 stat(크기/수정시각) 계산을 예약한다.
        - 검색 모드: 검색 전용 경량 경로만 처리하고 즉시 반환 (프록시 혼동 방지)
        - fast 모델: fast 프록시가 실제로 뷰에 연결돼 있을 때만 처리
        - 일반 모델: self.proxy 가 뷰에 연결돼 있을 때만 처리
        """
        # 검색 모드에서는 검색 전용 경로만 갱신하고, 다른 프록시로의 mapToSource 호출을 피한다.
        if self._search_mode:
            self._fill_search_visible_icons()
            return

        # 현재 뷰에 연결된 모델을 확인 (프록시 불일치 시 아무 것도 하지 않음)
        current_model = self.view.model()
        vp = self.view.viewport()

        # ---- Fast(임시) 모델 경로 ----
        if self._using_fast:
            if current_model is not self._fast_proxy:
                return
            rc = self._fast_proxy.rowCount()
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
                prx_ix = self._fast_proxy.index(r, 0)
                src_ix = self._fast_proxy.mapToSource(prx_ix)  # -> FastDirModel 인덱스
                row = src_ix.row()
                if row is None or row < 0:
                    continue
                if not self._fast_model.has_stat(row):
                    to_rows.append(row)

                # OS 아이콘 즉시 채우기
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
            self._fast_stat_worker = w
            w.start()
            return

        # ---- 일반(정상) 모델 경로 ----
        if current_model is not self.proxy:
            return

        model = self.proxy
        stat_proxy = self.stat_proxy
        src_model = self.source_model

        paths = []
        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        end = bot_ix.row() if bot_ix.isValid() else min(start + 120, model.rowCount() - 1)
        start = max(0, start - 40)
        end = min(model.rowCount() - 1, end + 80)
        if end < start:
            end = start

        for r in range(start, end + 1):
            prx_ix = model.index(r, 0)                 # -> FsSortProxy(self.proxy) 인덱스
            st_ix  = model.mapToSource(prx_ix)         # -> StatOverlayProxy 인덱스
            src_ix = stat_proxy.mapToSource(st_ix)     # -> QFileSystemModel 인덱스
            try:
                p = src_model.filePath(src_ix)
            except Exception:
                p = None
            if p:
                paths.append(p)

        if paths:
            stat_proxy.request_paths(paths)

    def _on_header_clicked(self, col:int):
        """헤더 클릭 시 정렬 변경 및 저장"""
        v=self.view
        if not v.isSortingEnabled(): 
            v.setSortingEnabled(True)
        
        # 같은 컬럼 클릭 시 방향 토글, 다른 컬럼은 오름차순
        if col == self._sort_column:
            new_order = Qt.DescendingOrder if self._sort_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            new_order = Qt.AscendingOrder
        
        self._sort_column = col
        self._sort_order = new_order
        
        v.header().setSortIndicator(col, new_order)
        v.sortByColumn(col, new_order)
        
        # 설정 저장
        self._save_sort_settings()

    def _mark_self_active(self):
        """이 Pane을 활성 Pane으로 표시."""
        try:
            if hasattr(self.host, "mark_active_pane"):
                self.host.mark_active_pane(self)
        except Exception:
            pass

    def eventFilter(self, obj, ev):
        # 뷰포트(파일 리스트) 영역: 기존 동작 유지
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

        # 필터 입력창: Esc로 필터 해제 + 브라우즈 모드 복귀
        if obj is self.filter_edit:
            if ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Escape:
                try:
                    self.filter_edit.clear()
                finally:
                    # 검색 모드였다면 브라우즈 모드로 복귀
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

        # 1) cmd.exe 경로 결정 (ComSpec 우선)
        comspec = os.environ.get("ComSpec") or r"C:\Windows\System32\cmd.exe"

        # 2) 우선: pywin32가 있으면 ShellExecute로 안전 실행
        if HAS_PYWIN32:
            try:
                # /K 로 창 유지 + 해당 폴더로 이동
                # title 지정은 선택적(디버깅 편의를 위해)
                params = f'/K title Multi-Pane File Explorer & cd /d "{path}"'
                win32api.ShellExecute(
                    int(self.window().winId()) if self.window() else 0,
                    "open",
                    comspec,
                    params,
                    path,  # 작업 디렉터리
                    win32con.SW_SHOWNORMAL
                )
                return
            except Exception:
                pass  # 아래 단계로 폴백

        # 3) 표준 subprocess 경로 (새 콘솔/프로세스 그룹으로 확실히 분리)
        try:
            flags = 0
            flags |= getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
            flags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)

            # 창 표시 정보(일부 환경에서 창이 즉시 사라지는 걸 방지)
            si = None
            try:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = 1  # SW_SHOWNORMAL
            except Exception:
                si = None

            # /K 로 유지 + /d 로 드라이브 전환 포함
            subprocess.Popen(
                [comspec, "/K", f'cd /d "{path}"'],
                cwd=path,
                creationflags=flags,
                startupinfo=si
            )
            return
        except Exception:
            pass  # 마지막 폴백

        # 4) 최종 폴백: cmd 내장 명령 'start' 사용 (shell=True 필요)
        try:
            # 빈 제목("") 필수, /D로 작업폴더 지정 후 cmd 실행
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

    def _use_fast_model(self, path: str):
        """
        FastDirModel을 사용해 먼저 빠르게 목록을 보여주고,
        디렉터리 열거가 끝난 뒤 한 번만 정렬/리페인트/아이콘 채우기를 수행하여
        대용량(수천 개+) 폴더 진입 속도를 높인다.
        ※ 아이콘 채우기는 열거 중에는 일시 비활성화하고, 끝난 뒤 복구한다.
        """
        # 진행 중인 워커/캐시 정리
        self._cancel_fast_stat_worker()
        if self._enum_worker and self._enum_worker.isRunning():
            self._enum_worker.cancel()
            self._enum_worker.wait(100)
        try:
            self.stat_proxy.clear_cache()
        except Exception:
            pass

        # Fast 모델로 전환
        self._using_fast = True
        self._fast_model.reset_dir(path)
        self.view.setModel(self._fast_proxy)
        self._configure_header_fast()

        # ── 아이콘 채우기 임시 비활성화(열거 중만) ─────────────────────
        #   Fast 모드에서 가시 영역 통계 계산 시 OS 아이콘을 요청하면
        #   큰 지연이 발생할 수 있어, 열거가 끝날 때까지 무시합니다.
        orig_apply_icon = getattr(self._fast_model, "apply_icon", None)

        def _noop_apply_icon(_row, _icon):
            # 열거 중에는 실제 아이콘을 적용하지 않음
            return None

        try:
            self._fast_model.apply_icon = _noop_apply_icon
        except Exception:
            orig_apply_icon = None  # 복구 불가해도 동작에는 지장 없음

        # 열거 중에는 재정렬이 반복되지 않도록 정렬/업데이트를 잠시 끔
        was_sorting = self.view.isSortingEnabled()
        if was_sorting:
            self.view.setSortingEnabled(False)
        self.view.setUpdatesEnabled(False)

        self._fast_batch_counter = 0
        self._enum_worker = DirEnumWorker(path)

        # 배치 도착: 모델에 추가 + 과도한 갱신은 스로틀링
        def _on_batch(rows):
            self._fast_model.append_rows(rows)
            self._fast_batch_counter += 1
            if (self._fast_batch_counter % 6) == 0:
                QTimer.singleShot(0, self._schedule_visible_stats)

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

        # 열거 종료: 정렬/업데이트 복구 + 아이콘 채우기 복구 후 즉시 가시영역 아이콘/통계를 채움
        def _on_finished():
            try:
                # 화면 업데이트 재개
                self.view.setUpdatesEnabled(True)

                self._apply_saved_sort()
                self.view.setSortingEnabled(True)

                # 아이콘 apply 함수 원복 → 이제 가시영역 아이콘을 실제로 채움
                _restore_apply_icon()

                # 가시 영역 통계/아이콘을 즉시 2회 스케줄(첫 패스 후 아이콘 캐시 안정화용)
                QTimer.singleShot(0, self._schedule_visible_stats)
                QTimer.singleShot(80, self._schedule_visible_stats)
            finally:
                # 사용자가 정렬을 꺼두고 썼다면 원 상태 유지
                if not was_sorting:
                    self.view.setSortingEnabled(False)

        self._enum_worker.finished.connect(_on_finished)

        # 초기 화면 빈칸 방지를 위해 첫 패스 예약
        QTimer.singleShot(0, self._schedule_visible_stats)

        # 백그라운드 열거 시작
        self._enum_worker.start()


    def _start_normal_model_loading(self, path:str):
        """
        큰 폴더 진입 시 UI 스톨을 줄이기 위해:
          - 항목 수가 매우 많으면(Q&D 카운트) QFileSystemModel 로딩을 생략하고 fast 모델만 유지
          - 중간 규모 이상이면 아이콘을 제너릭으로 강제하여 셸 아이콘 조회 비용 최소화
        """
        self._pending_normal_root = path
        t = QElapsedTimer(); t.start()
        self._dirload_timer[path.lower()] = t

        # ── 폴더 크기 빠른 추정: threshold를 넘으면 즉시 중단
        count = 0
        is_huge = False
        HUGE_THRESHOLD = 3000   # 이 이상이면 normal 모델 전환 생략
        GENERIC_THRESHOLD = 1200  # 이 이상이면 제너릭 아이콘 강제
        try:
            with os.scandir(path) as it:
                for _ in it:
                    count += 1
                    if count >= HUGE_THRESHOLD:
                        is_huge = True
                        break
        except Exception:
            # 읽기 실패해도 normal 로딩을 시도할 수 있게 둔다
            pass

        if is_huge:
            # 매우 큰 폴더: fast 모델만 유지 (normal 전환 생략)
            dlog(f"[perf] Skip QFileSystemModel for huge folder (>= {HUGE_THRESHOLD} items): {path}")
            self._pending_normal_root = None
            QTimer.singleShot(0, self._schedule_visible_stats)
            return

        # 중간 규모 이상: 제너릭 아이콘 강제(셸 아이콘 조회 비용 절감)
        try:
            if ALWAYS_GENERIC_ICONS or count >= GENERIC_THRESHOLD:
                self.source_model.setIconProvider(self._generic_icons)
        except Exception:
            pass

        # 일반 모델 비동기 로딩 시작
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
        """
        Return (accessible, prompted).
        - accessible: path became accessible after this call
        - prompted: a credential/login flow was launched
        """
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
                nr.dwType = 1  # RESOURCETYPE_DISK
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
                        "네트워크 인증 창을 열었습니다.\n인증 후 같은 경로를 다시 열어주세요.",
                    )
                    return
            if not os.path.exists(path):
                QMessageBox.warning(self, "Path not found", path)
                return

            # 검색 결과 뷰에서 다른 폴더로 이동할 때는 먼저 검색 모드를 정리한다.
            # (_search_mode가 남아 있으면 normal 모델 전환이 막혀 외부 드래그가 실패할 수 있음)
            if self._search_mode:
                self._enter_browse_mode()

            cur = getattr(self.path_bar, "_current_path", None)
            if push_history and cur and os.path.normcase(cur) != os.path.normcase(path):
                self._back_stack.append(cur)
                self._fwd_stack.clear()

            # 경로 표시/즐겨찾기 상태 반영
            self.path_bar.set_path(path)
            self._update_star_button()

            # ── 파일시스템 워처 바인딩(자동 새로고침) ──
            try:
                self._bind_fs_watcher(path)
            except Exception:
                pass

            # 빠른 모델로 즉시 표시 → 백그라운드로 QFileSystemModel 로딩
            self._use_fast_model(path)
            QTimer.singleShot(0, lambda: self._start_normal_model_loading(path))

            # 상태바/선택정보 업데이트 예약
            QTimer.singleShot(50, self._update_pane_status)
            self._update_statusbar_selection()

    def _bind_fs_watcher(self, folder_path: str):
        """
        현재 Pane이 보고 있는 폴더에 QFileSystemWatcher를 바인딩한다.
        - 디렉터리 변경 이벤트를 디바운스(합치기)해서 과도한 새로고침을 방지
        - 일반 모드(QFileSystemModel)에서는 모델이 스스로 갱신되므로 최소 작업만,
          빠른 모드(FastDirModel/검색 모드)에서는 가벼운 재로딩을 수행
        """
        # 워처/디바운스 타이머 지연 생성
        if not hasattr(self, "_fswatch"):
            self._fswatch = QtCore.QFileSystemWatcher(self)
            self._fswatch.directoryChanged.connect(self._on_fs_changed)
            self._fswatch.fileChanged.connect(self._on_fs_changed)

        if not hasattr(self, "_fswatch_debounce"):
            self._fswatch_debounce = QTimer(self)
            self._fswatch_debounce.setSingleShot(True)
            self._fswatch_debounce.setInterval(600)  # ms: 변경 폭주 대비
            self._fswatch_debounce.timeout.connect(self._apply_fs_change)

        # 기존 감시 경로 제거 후 새 경로 등록
        try:
            dirs = list(self._fswatch.directories())
            if dirs:
                self._fswatch.removePaths(dirs)
        except Exception:
            pass

        try:
            # 존재하는 디렉터리만 감시
            if os.path.isdir(folder_path):
                self._fswatch.addPath(folder_path)
        except Exception:
            # 감시 실패는 기능상 치명적이지 않음(무시)
            pass

    def _on_fs_changed(self, _path: str):
        """
        워처 이벤트 수신 시 디바운스 타이머만 갱신한다.
        (실제 갱신은 _apply_fs_change 에서 일괄 처리)
        """
        try:
            if self._fswatch_debounce.isActive():
                self._fswatch_debounce.stop()
            self._fswatch_debounce.start()
        except Exception:
            pass

    def _apply_fs_change(self):
        """
        디바운스 후 실제 반영 로직.
        - 검색 모드: 필터가 있으면 재검색, 없으면 브라우즈 모드로 복귀
        - 빠른 모드: 현재 경로를 빠르게 재열거
        - 일반 모드(QFileSystemModel): 모델이 자동 반영 → 가시 영역 통계만 갱신
        """
        try:
            # 검색 모드일 때
            if getattr(self, "_search_mode", False):
                pattern = self.filter_edit.text().strip()
                if pattern:
                    # 현재 필터 그대로 재검색
                    self._apply_filter()
                else:
                    # 필터가 비어 있으면 브라우즈 모드로 복귀
                    self._enter_browse_mode()
                return

            # 빠른 모드일 때는 빠르게 재열거
            if getattr(self, "_using_fast", False):
                self._use_fast_model(self.current_path())
                return

            # 일반 모드(QFileSystemModel)라면 자동으로 반영되므로 가시영역만 갱신
            QTimer.singleShot(0, self._schedule_visible_stats)
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

        # 큰 폴더로 판단해 normal 전환을 건너뛰었으면 아무 것도 하지 않음
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

                # ★ 선택모델 재연결
                self._hook_selection_model()

                self._configure_header_browse()
                if not self.view.isSortingEnabled():
                    self.view.setSortingEnabled(True)
                # 저장된 정렬 적용
                self._apply_saved_sort()
                QTimer.singleShot(0, self._schedule_visible_stats)
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

    def _external_clipboard_payload(self):
        try:
            cb = QApplication.clipboard()
            md = cb.mimeData()
        except Exception:
            return None
        if not md or not md.hasUrls():
            return None

        urls = md.urls()
        # Keep order while removing duplicates to avoid redundant work
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
        # 검색 관련 상태/워커 정리
        try:
            if hasattr(self, "_cancel_search_worker"):
                self._cancel_search_worker()
        except Exception:
            pass
        self._search_item_by_path = {}
        self._search_stats_done = set()
        self._search_model = None
        self._search_proxy = None

        if not self._search_mode:
            QTimer.singleShot(0, self._schedule_visible_stats)
            return

        self._search_mode = False

        if self._using_fast:
            self.view.setModel(self._fast_proxy)
        else:
            self.view.setModel(self.proxy)

        # ★ 선택모델 재연결
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
        # 저장된 정렬 적용
        self._apply_saved_sort()

        QTimer.singleShot(0, self._schedule_visible_stats)

    def _enter_search_mode(self, model:QStandardItemModel):
        self._cancel_fast_stat_worker()
        self._using_fast=False
        self._search_mode=True
        self._search_model=model
        self._search_proxy=FsSortProxy(self)
        self._search_proxy.setSourceModel(self._search_model)
        self.view.setModel(self._search_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())

        # ★ 선택모델 재연결
        self._hook_selection_model()

        header = self.view.header()
        self._configure_header_search()
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

        # 먼저 윈도우 탐색기 네이티브 메뉴 시도
        if self._try_native_context_menu(pos, owner_hwnd, paths):
            return

        # ── Fallback: 'New' 항목만 상위 레벨에 바로 표시 ──
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


    def _on_selection_changed(self,*_): self._update_statusbar_selection(); self._update_pane_status()
    def _update_statusbar_selection(self):
        # 상태바는 가볍게: 디렉터리 재귀 계산 없이, 전부 파일일 때만 합계
        sel = self._selected_paths()
        cnt = len(sel)
        msg = f"Pane {self.pane_id} — selected {cnt} item(s)"
        if cnt and all(os.path.isfile(p) for p in sel):
            total = 0
            for p in sel:
                try:
                    total += os.path.getsize(p)
                except Exception:
                    pass
            msg += f" / {human_size(total)}"
        try:
            self.host.statusBar().showMessage(msg, 2000)
        except Exception:
            pass

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
        sel = self._selected_paths()
        cnt = len(sel)

        # 좌측 상태 텍스트: 선택 개수, 전부 파일일 때는 총 용량
        text = ""
        if cnt:
            only_files = all(os.path.isfile(p) for p in sel)
            if only_files:
                total = 0
                for p in sel:
                    try:
                        total += os.path.getsize(p)
                    except Exception:
                        pass
                text = f"{cnt} selected — {human_size(total)}"
            else:
                text = f"{cnt} selected"
        self.lbl_sel.setText(text)

        # 우측 드라이브 여유 공간
        path = self.current_path()
        if self._is_network_path(path):
            self.lbl_free.setText("")
        else:
            try:
                total, used, free = shutil.disk_usage(path)
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

        # ▶ Session 버튼 추가 (Edit Bookmarks 오른쪽)
        self.btn_session=QToolButton(top)
        self.btn_session.setToolTip("Session (save/load all pane paths)")
        self.btn_session.setFixedHeight(UI_H)

        self.btn_about=QToolButton(top); self.btn_about.setToolTip("About"); self.btn_about.setFixedHeight(UI_H)
        top_lay.addWidget(self.btn_layout,0); top_lay.addWidget(self.btn_theme,0); top_lay.addWidget(self.btn_bm_edit,0)
        top_lay.addWidget(self.btn_session,0)  # ← 추가된 버튼
        top_lay.addWidget(self.btn_about,0); top_lay.addStretch(1)

        self.central=QWidget(self); self.setCentralWidget(self.central)
        vmain=QVBoxLayout(self.central); vmain.setContentsMargins(0,0,0,0); vmain.setSpacing(ROW_SPACING)
        vmain.addWidget(top,0); self.grid=QGridLayout(); vmain.addLayout(self.grid,1)

        self.named_bookmarks=migrate_legacy_favorites_into_named(load_named_bookmarks()); save_named_bookmarks(self.named_bookmarks)
        self._clipboard=None; self._bm_dlg=None

        self._update_layout_icon(); self._update_theme_icon()
        self.btn_layout.clicked.connect(self._cycle_layout); self.btn_theme.clicked.connect(self._toggle_theme)
        self.btn_bm_edit.clicked.connect(self._open_bookmark_editor)
        self.btn_session.clicked.connect(self._open_session_manager)  # ← 세션 매니저 열기
        self.btn_about.clicked.connect(self._show_about)

        self.panes=[]; self.build_panes(pane_count, start_paths or []); self._update_theme_dependent_icons()
        self._install_focus_tracker()
        if getattr(self, "panes", None):
            self.mark_active_pane(self.panes[0])

        self.statusBar().showMessage("Ready", 1500)

        self._wd_timer=QTimer(self); self._wd_timer.setInterval(50); self._wd_last=time.perf_counter()
        def _wd_tick():
            now=time.perf_counter(); gap=(now-self._wd_last)*1000
            if gap>200: dlog(f"[STALL] UI event loop blocked ~{gap:.0f} ms")
            self._wd_last=now
        self._wd_timer.timeout.connect(_wd_tick); self._wd_timer.start()

        settings=QSettings(ORG_NAME, APP_NAME); geo=settings.value("window/geometry")
        if isinstance(geo, QtCore.QByteArray): self.restoreGeometry(geo)

    def mark_active_pane(self, pane):
        """
        활성 Pane을 설정하고, 각 Pane의 PathBar와 Pane 전체 시각 상태를 반영한다.
        """
        try:
            self._active_pane = pane
            for p in getattr(self, "panes", []):
                try:
                    is_active = (p is pane)
                    # Path bar 하이라이트
                    p.path_bar.set_active(is_active)
                    # Pane 전체 하이라이트
                    if hasattr(p, "set_active_visual"):
                        p.set_active_visual(is_active)
                except Exception:
                    pass
            dlog(f"[active] pane={getattr(pane, 'pane_id', '?')}")
        except Exception:
            pass

    def _install_focus_tracker(self):
        """앱 전역 포커스 변화에 따라 활성 Pane을 갱신한다."""
        app = QApplication.instance()
        if not app:
            return
        # 중복 연결 방지
        try:
            self._focus_tracker_connected
        except AttributeError:
            self._focus_tracker_connected = False
        if self._focus_tracker_connected:
            return
        app.focusChanged.connect(self._on_focus_changed)
        self._focus_tracker_connected = True

    def _on_focus_changed(self, old, now):
        """포커스가 바뀔 때, 포커스를 품고 있는 Pane을 활성화한다."""
        try:
            if not now:
                return
            if not isinstance(now, QWidget):
                return
            for p in getattr(self, "panes", []):
                # now가 Pane의 자손 위젯이면 해당 Pane을 활성화
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
            self.btn_bm_edit.setIcon(icon_bookmark_edit(self.theme))  # 북마크 편집(별+연필)
        if hasattr(self, "btn_session") and self.btn_session:
            self.btn_session.setIcon(icon_session(self.theme))        # ★ 세션 아이콘 지정
        if hasattr(self, "btn_about") and self.btn_about:
            self.btn_about.setIcon(icon_info(self.theme))

    def _update_theme_dependent_icons(self):
        self._update_layout_icon(); self._update_theme_icon()
        for p in getattr(self,"panes",[]):
            try:
                p.btn_star.setIcon(icon_star(p.btn_star.isChecked(), self.theme))
                p.btn_cmd.setIcon(icon_cmd(self.theme))
                # PathBar의 복사 아이콘도 갱신
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
        """
        - 현재(직전) 창 개수에 대한 경로들을 last_paths_{count} 로 저장
        - 4→6, 6→8 등 창 개수를 늘릴 때, 새로 생기는 창(5~n)은
          마지막으로 사용했던 same-count(6 또는 8) 레이아웃의 경로를 사용
        - 레이아웃은 새 GridLayout으로 재구성하고, 최대화 상태를 유지/복원
        """
        was_max = self.isMaximized()

        # 1) 직전 레이아웃의 경로 저장 (예: 4개 사용 중이었다면 last_paths_4 로 저장)
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

        # 2) 대상 레이아웃에서 사용할 경로 구성
        #    - start_paths(호출자가 넘긴 현재 표시 경로들)로 우선 채움
        #    - 부족한 나머지는 last_paths_{n}에서 보충 (주로 5~n 채우기)
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

        # 3) 기존 그리드/페인 완전 제거
        vmain = self.centralWidget().layout() if self.centralWidget() else None
        if hasattr(self, "grid") and isinstance(self.grid, QGridLayout):
            while self.grid.count():
                it = self.grid.takeAt(0)
                w = it.widget()
                if w:
                    w.setParent(None)
            if vmain:
                try:
                    vmain.removeItem(self.grid)
                except Exception:
                    pass
            try:
                self.grid.setParent(None)
            except Exception:
                pass

        # 4) 새 그리드 구성
        cols = {4: 2, 6: 3, 8: 4}.get(n, 3)
        gap = GRID_GAPS.get(cols, 3)
        margin_lr = GRID_MARG_LR.get(cols, 6)

        self.grid = QGridLayout()
        self.grid.setSpacing(gap)
        self.grid.setContentsMargins(margin_lr, 2, margin_lr, 4)
        if vmain:
            vmain.addLayout(self.grid, 1)

        # 스트레치 설정
        for c in range(cols):
            self.grid.setColumnStretch(c, 1)
            self.grid.setColumnMinimumWidth(c, 0)
        rows = (n + cols - 1) // cols
        for r in range(rows):
            self.grid.setRowStretch(r, 1)
            self.grid.setRowMinimumHeight(r, 0)

        # 5) 새 페인 생성/배치 (final_paths를 그대로 사용)
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

        # 6) 타이틀/아이콘/레이아웃 마무리
        self.setWindowTitle(f"Multi-Pane File Explorer — {n} panes")
        self._update_theme_dependent_icons()

        # 최대화 상태였다면: 해제→재최대화로 레이아웃을 확실히 채움
        if was_max:
            QTimer.singleShot(0, self._unmax_then_remax)
        else:
            QTimer.singleShot(0, self._kick_layout)


    def _unmax_then_remax(self):
        try:
            # 최대화 상태가 아니면 일반 킥만
            if not self.isMaximized():
                self._kick_layout()
                return

            # 1) 최대화 해제 (창 크기 변경 경고 피하기 위해 showNormal 사용)
            self.showNormal()
            QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.LayoutRequest)
            QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents)

            # 2) 레이아웃 재계산
            self._kick_layout()

            # 3) 다시 최대화
            self.showMaximized()
            QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.LayoutRequest)
            QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents)

            # 4) 마지막으로 한 번 더 레이아웃 킥
            self._kick_layout()
        except Exception:
            pass

    def _kick_layout(self):
        """
        레이아웃을 재계산하되, '최대화가 아닌 상태'에서는 현재 창 크기/위치를
        절대 변경하지 않는다. (전환 시 창이 커지는 문제 방지)
        """
        try:
            cw = self.centralWidget()
            lay = cw.layout() if cw else None

            keep_geom = not self.isMaximized()   # 최대화가 아니라면 현재 지오메트리 고정
            if keep_geom:
                g = self.geometry()
                # 현재 크기로 일시 고정(레이아웃 중 창이 커지는 것 방지)
                old_min = self.minimumSize()
                old_max = self.maximumSize()
                self.setMinimumSize(g.size())
                self.setMaximumSize(g.size())

            # ── 레이아웃 강제 패스 (창 크기 변경 없이) ─────────────────
            if lay:
                lay.invalidate()
                lay.activate()

            if hasattr(self, "grid"):
                self.grid.invalidate()
                self.grid.update()

            # 각 Pane 내부 위젯 갱신 (확장 정책만 보정: 창 크기에는 영향 없음)
            for p in getattr(self, "panes", []):
                try:
                    p.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                    if hasattr(p, "view") and p.view:
                        p.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                        p.view.updateGeometries()
                        p.view.doItemsLayout()
                        p.view.viewport().update()
                    # 경로바는 항상 우측(가장 하위 폴더) 고정
                    if hasattr(p, "path_bar") and hasattr(p.path_bar, "_pin_to_right"):
                        p.path_bar._pin_to_right()
                except Exception:
                    pass

            # 레이아웃 요청 처리(사용자 입력 제외)
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
            # ── 잠금 해제 & 지오메트리 복원 ─────────────────────────────
            try:
                if keep_geom:
                    # 원래 min/max 복원
                    try:
                        self.setMinimumSize(old_min)
                        self.setMaximumSize(old_max)
                    except Exception:
                        # 실패 시 일반 복원값으로
                        self.setMinimumSize(QtCore.QSize(0, 0))
                        self.setMaximumSize(QtCore.QSize(16777215, 16777215))
                    # 레이아웃 중 변경됐으면 원 지오메트리로 되돌림
                    if self.geometry() != g:
                        self.setGeometry(g)
            except Exception:
                pass


    def _safe_restore_geometry(self, ba: QtCore.QByteArray):
        """
        저장된 지오메트리를 복원하되, 현재 모니터의 사용 가능 영역을 넘지 않도록 클램프한다.
        """
        try:
            ok = self.restoreGeometry(ba)
            if not ok:
                return

            # 현재 화면의 사용 가능 영역
            win = self.windowHandle()
            screen = (win.screen().availableGeometry() if win and win.screen()
                      else QApplication.primaryScreen().availableGeometry())
            sg = screen

            g = self.geometry()
            new_w = min(max(600, g.width()), sg.width())
            new_h = min(max(350, g.height()), sg.height())
            new_x = min(max(sg.left(), g.x()), sg.right() - new_w)
            new_y = min(max(sg.top(),  g.y()), sg.bottom() - new_h)

            # 창 크기/위치가 화면을 벗어나면 보정
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
            "<div style='color:#000; font-size:12pt;'><b>Multi-Pane File Explorer v1.3.0</b></div>"
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

    # ---- Sessions (save/load all pane paths) ---------------------------------
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
            QMessageBox.information(self, "Save Session", "세션 이름을 입력해 주세요.")
            return
        paths = self._current_paths()
        panes = len(self.panes)
        items = self._get_sessions()
        # 이름 중복은 대소문자 무시하고 교체
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
            QMessageBox.warning(self, "Load Session", "세션을 찾을 수 없습니다."); return
        paths = list(target.get("paths", []))
        panes = int(target.get("panes", len(paths)))
        if panes <= 0 or not paths:
            QMessageBox.warning(self, "Load Session", "세션 데이터가 비어 있습니다."); return

        # 현재 레이아웃과 다르면 재구성
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
            pass  # 필요 시 후처리

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
            QMessageBox.information(self, "Load Session", "세션을 선택해 주세요.")
            return
        try:
            self.parent()._load_session(name)
        except Exception as e:
            QMessageBox.critical(self, "Load Session", str(e))

    def _on_delete(self):
        name = self._selected_name()
        if not name:
            QMessageBox.information(self, "Delete Session", "세션을 선택해 주세요.")
            return
        btn = QMessageBox.question(self, "Delete Session", f"삭제하시겠습니까?\n{name}",
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
        # 이미 존재하면 덮어쓸지 확인
        exists = any(s.get("name","").lower() == name.lower() for s in self.parent()._get_sessions())
        if exists:
            btn = QMessageBox.question(self, "Save Session",
                                       f"동일한 이름의 세션이 있습니다.\n덮어쓰시겠습니까?\n{name}",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if btn != QMessageBox.Yes:
                return
        try:
            self.parent()._save_session(name)
            self.set_sessions(self.parent()._get_sessions())
        except Exception as e:
            QMessageBox.critical(self, "Save Session", str(e))


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
