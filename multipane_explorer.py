

import os, sys, fnmatch, argparse, shutil, ctypes, math, subprocess, time, re, uuid, errno, stat
from contextlib import contextmanager
from pathlib import Path

from PyQt5 import QtCore
from PyQt5.QtCore import (
    Qt, QDir, QUrl, QDateTime, QSortFilterProxyModel,
    pyqtSignal, QSettings, QEvent, QTimer, QSize, QAbstractTableModel,
    QIdentityProxyModel, QElapsedTimer, QStringListModel, QPropertyAnimation,
    QSequentialAnimationGroup, QPauseAnimation
)
from PyQt5.QtGui import (
    QDesktopServices, QPalette, QColor, QKeySequence, QIcon,
    QStandardItemModel, QStandardItem, QPainter, QPixmap, QPen, QBrush,
    QCursor, QPolygonF, QGuiApplication, QFont, QImage
)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTreeView, QFileSystemModel,
    QLineEdit, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout,
    QAction, QInputDialog, QMessageBox, QAbstractItemView,
    QMenu, QStyle, QHeaderView, QScrollArea, QFrame, QLabel, QShortcut,
    QToolButton, QDialog, QDialogButtonBox, QTableWidget, QTableWidgetItem,
    QCheckBox, QFileDialog, QProgressBar, QToolTip, QSizePolicy, QFileIconProvider,
    QComboBox, QSpacerItem, QCompleter, QSpinBox, QStyledItemDelegate,
    QGraphicsOpacityEffect
)



def app_resource_path(filename: str) -> str:
    """Return a bundled resource path in both source and PyInstaller builds."""
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, filename)

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
APP_VERSION = "2.6.2"


BASE_FONT_PT = 9.5
FONT_PT = BASE_FONT_PT
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
SEARCH_PROGRESS_INTERVAL = 250
DEFAULT_SEARCH_EXCLUDE_DIRS = {
    ".git", ".hg", ".svn", "node_modules", ".venv", "venv",
    "__pycache__", ".mypy_cache", ".pytest_cache", "dist", "build",
}
FILEOP_ERROR_DETAIL_LIMIT = 50
LARGE_FOLDER_THRESHOLD = 3000
FILEOP_FAST_PROGRESS_SCAN_LIMIT = 4000
GENERIC_ICON_THRESHOLD = 1200
SHELL_ICON_FAILURE_TTL_S = 300
PATH_HISTORY_LIMIT = 30
BOOKMARK_LIMIT = 30
QUICK_BOOKMARK_MIN_W = 42
QUICK_BOOKMARK_MAX_W = 78
QUICK_BOOKMARK_MORE_W = 30
VALID_THEMES = ("dark", "light")
SIZE_COL_WIDTH = 60
DATE_COL_WIDTH = 122
SEARCH_FOLDER_COL_WIDTH = 240
LIST_DATETIME_FMT = "yyyy-MM-dd HH:mm"
HOVER_TOOLTIP_DURATION_MULTIPLIER = 9
APP_ICON_FILENAME = "images-5.ico"
ABOUT_IMAGE_FILENAME = "images-6.ico"

GLOBAL_SHELL_ICON_CACHE = {}
GLOBAL_SHELL_ICON_FAILURES = {}

# Keep this list in sync with the README keyboard-shortcuts section.
KEYBOARD_SHORTCUTS = [
    ("F1", "Open shortcuts help"),
    ("Backspace / Alt+Left", "Back"),
    ("Alt+Right", "Forward"),
    ("Alt+Up", "Parent folder"),
    ("Enter / Ctrl+O", "Open"),
    ("Ctrl+L / F4", "Edit path"),
    ("Ctrl+F / F3", "Focus filter"),
    ("Esc (in filter input)", "Clear filter + return to browse mode"),
    ("F5", "Hard refresh"),
    ("Ctrl+C / Ctrl+X / Ctrl+V", "Copy / Cut / Paste"),
    ("Ctrl+Z", "Undo (new folder / rename / safe copy-move actions)"),
    ("Delete / Shift+Delete", "Recycle / Permanent delete"),
    ("F2", "Rename"),
    ("Ctrl+Shift+R", "Bulk rename"),
    ("Ctrl+Shift+C", "Copy full path"),
    ("Ctrl+Shift+D", "Open the selected item's containing folder"),
    ("Alt+Shift+C", "Copy folder path (parent folder if a file is selected)"),
]


HAS_PYWIN32 = True
try:
    import pythoncom
    import win32con, win32gui, win32api, win32clipboard
    from win32com.shell import shell, shellcon
except Exception:
    HAS_PYWIN32 = False


def _dedupe_local_paths(paths):
    out = []
    seen = set()
    for raw in paths or ():
        if not raw:
            continue
        try:
            path = os.path.normpath(os.fspath(raw))
        except Exception:
            continue
        key = os.path.normcase(path) if sys.platform == "win32" else path
        if key in seen:
            continue
        seen.add(key)
        out.append(path)
    return out


def _decode_preferred_drop_effect(raw):
    try:
        data = bytes(raw) if raw is not None else b""
    except Exception:
        return None
    if len(data) < 4:
        return None
    return int.from_bytes(data[:4], byteorder="little", signed=False)


def _drop_effect_to_operation(effect):
    if effect is None:
        return None
    if effect & 2:
        return "move"
    if effect & 1:
        return "copy"
    return None


def _read_windows_file_clipboard_payload():
    if sys.platform != "win32" or not HAS_PYWIN32:
        return None
    try:
        drop_effect_fmt = win32clipboard.RegisterClipboardFormat("Preferred DropEffect")
        win32clipboard.OpenClipboard()
        try:
            paths = []
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                data = win32clipboard.GetClipboardData(win32con.CF_HDROP)
                if isinstance(data, (tuple, list)):
                    paths = [os.fspath(p) for p in data if p]
                elif data:
                    paths = [os.fspath(data)]
            effect = None
            if win32clipboard.IsClipboardFormatAvailable(drop_effect_fmt):
                effect = _decode_preferred_drop_effect(
                    win32clipboard.GetClipboardData(drop_effect_fmt)
                )
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None

    paths = _dedupe_local_paths(paths)
    if not paths:
        return None
    return {"op": _drop_effect_to_operation(effect) or "copy", "paths": paths}


def _normalize_file_clipboard_payload(payload):
    if not isinstance(payload, dict):
        return None
    paths = _dedupe_local_paths(payload.get("paths") or [])
    if not paths:
        return None
    op = str(payload.get("op") or "copy").strip().lower()
    if op not in {"copy", "cut", "move"}:
        op = "copy"
    return {"op": op, "paths": paths}


def _clipboard_operation_to_drop_effect(op):
    return 2 if str(op).strip().lower() in {"cut", "move"} else 1


def _clipboard_payload_matches(left, right) -> bool:
    left = _normalize_file_clipboard_payload(left)
    right = _normalize_file_clipboard_payload(right)
    if not left or not right:
        return False
    left_op = "move" if left["op"] in {"cut", "move"} else "copy"
    right_op = "move" if right["op"] in {"cut", "move"} else "copy"
    if left_op != right_op:
        return False
    return [_path_key(p) for p in left["paths"]] == [_path_key(p) for p in right["paths"]]


def execute_bulk_rename_transaction(operations) -> list[tuple[str, str]]:
    """Rename all items atomically as a group, rolling back every completed step on failure."""
    temp_pairs = []
    committed = []
    try:
        for src, dst in operations:
            parent = os.path.dirname(src) or os.curdir
            temp = os.path.join(parent, f".__mprn_tmp_{uuid.uuid4().hex}")
            while os.path.lexists(temp):
                temp = os.path.join(parent, f".__mprn_tmp_{uuid.uuid4().hex}")
            os.rename(src, temp)
            temp_pairs.append((src, temp, dst))

        for src, temp, dst in temp_pairs:
            os.rename(temp, dst)
            committed.append((dst, src))
        return committed
    except Exception as original_error:
        rollback_errors = []

        # Final names must be restored first because original names are still free.
        for dst, src in reversed(committed):
            if not os.path.lexists(dst):
                continue
            try:
                os.rename(dst, src)
            except Exception as exc:
                rollback_errors.append(f"{dst} -> {src}: {exc}")

        committed_sources = {_path_key(src) for _dst, src in committed}
        for src, temp, _dst in reversed(temp_pairs):
            if _path_key(src) in committed_sources or not os.path.lexists(temp):
                continue
            try:
                os.rename(temp, src)
            except Exception as exc:
                rollback_errors.append(f"{temp} -> {src}: {exc}")

        if rollback_errors:
            details = "\n".join(rollback_errors[:10])
            more = "\n..." if len(rollback_errors) > 10 else ""
            raise RuntimeError(
                f"{original_error}\n\nRollback was incomplete:\n{details}{more}"
            ) from original_error
        raise


def _write_windows_file_clipboard_payload(payload):
    payload = _normalize_file_clipboard_payload(payload)
    if sys.platform != "win32" or not HAS_PYWIN32 or not payload:
        return False

    class _Point(ctypes.Structure):
        _fields_ = [
            ("x", ctypes.c_long),
            ("y", ctypes.c_long),
        ]

    class _DropFiles(ctypes.Structure):
        _fields_ = [
            ("pFiles", ctypes.c_uint32),
            ("pt", _Point),
            ("fNC", ctypes.c_int),
            ("fWide", ctypes.c_int),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    global_alloc = kernel32.GlobalAlloc
    global_alloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    global_alloc.restype = ctypes.c_void_p
    global_lock = kernel32.GlobalLock
    global_lock.argtypes = [ctypes.c_void_p]
    global_lock.restype = ctypes.c_void_p
    global_unlock = kernel32.GlobalUnlock
    global_unlock.argtypes = [ctypes.c_void_p]
    global_unlock.restype = ctypes.c_int
    global_free = kernel32.GlobalFree
    global_free.argtypes = [ctypes.c_void_p]
    global_free.restype = ctypes.c_void_p
    open_clipboard = user32.OpenClipboard
    open_clipboard.argtypes = [ctypes.c_void_p]
    open_clipboard.restype = ctypes.c_int
    empty_clipboard = user32.EmptyClipboard
    empty_clipboard.argtypes = []
    empty_clipboard.restype = ctypes.c_int
    set_clipboard_data = user32.SetClipboardData
    set_clipboard_data.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    set_clipboard_data.restype = ctypes.c_void_p
    close_clipboard = user32.CloseClipboard
    close_clipboard.argtypes = []
    close_clipboard.restype = ctypes.c_int
    register_clipboard_format = user32.RegisterClipboardFormatW
    register_clipboard_format.argtypes = [ctypes.c_wchar_p]
    register_clipboard_format.restype = ctypes.c_uint
    GMEM_MOVEABLE = 0x0002
    GMEM_ZEROINIT = 0x0040

    def _alloc_global_bytes(raw):
        handle = global_alloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(raw))
        if not handle:
            raise OSError("GlobalAlloc failed")
        ptr = global_lock(handle)
        if not ptr:
            global_free(handle)
            raise OSError("GlobalLock failed")
        try:
            ctypes.memmove(ptr, raw, len(raw))
        finally:
            global_unlock(handle)
        return handle

    file_list = ("\0".join(payload["paths"]) + "\0\0").encode("utf-16le")
    dropfiles = _DropFiles()
    dropfiles.pFiles = ctypes.sizeof(_DropFiles)
    dropfiles.fNC = 0
    dropfiles.fWide = 1
    hdrop = None
    heffect = None

    try:
        hdrop = _alloc_global_bytes(bytes(dropfiles) + file_list)
        heffect = _alloc_global_bytes(
            _clipboard_operation_to_drop_effect(payload["op"]).to_bytes(4, "little")
        )
        effect_fmt = register_clipboard_format("Preferred DropEffect")

        opened = False
        for _ in range(8):
            if open_clipboard(None):
                opened = True
                break
            time.sleep(0.03)
        if not opened:
            return False
        try:
            if not empty_clipboard():
                return False
            if not set_clipboard_data(win32con.CF_HDROP, hdrop):
                return False
            hdrop = None
            if not set_clipboard_data(effect_fmt, heffect):
                return False
            heffect = None
            return True
        finally:
            close_clipboard()
    except Exception:
        return False
    finally:
        if hdrop:
            global_free(hdrop)
        if heffect:
            global_free(heffect)


def _clear_windows_clipboard():
    if sys.platform != "win32":
        return False
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    open_clipboard = user32.OpenClipboard
    open_clipboard.argtypes = [ctypes.c_void_p]
    open_clipboard.restype = ctypes.c_int
    empty_clipboard = user32.EmptyClipboard
    empty_clipboard.argtypes = []
    empty_clipboard.restype = ctypes.c_int
    close_clipboard = user32.CloseClipboard
    close_clipboard.argtypes = []
    close_clipboard.restype = ctypes.c_int
    opened = False
    for _ in range(8):
        if open_clipboard(None):
            opened = True
            break
        time.sleep(0.03)
    if not opened:
        return False
    try:
        return bool(empty_clipboard())
    finally:
        close_clipboard()


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

def _path_exists_for_delete(path: str) -> bool:
    try:
        return os.path.lexists(path)
    except Exception:
        return os.path.exists(path)

def _is_junction(path: str) -> bool:
    isjunction = getattr(os.path, "isjunction", None)
    if isjunction:
        try:
            if isjunction(path):
                return True
        except Exception:
            pass
    if os.name == "nt":
        try:
            st = os.lstat(path)
            tag = getattr(st, "st_reparse_tag", 0)
            junction_tag = getattr(stat, "IO_REPARSE_TAG_MOUNT_POINT", 0xA0000003)
            return bool(tag == junction_tag)
        except Exception:
            pass
    return False

def _is_dir_link(path: str) -> bool:
    try:
        return bool(os.path.islink(path) or _is_junction(path))
    except Exception:
        return False

def _link_target_is_directory(path: str) -> bool:
    if os.path.isdir(path):
        return True
    if os.name == "nt":
        try:
            attrs = getattr(os.lstat(path), "st_file_attributes", 0)
            directory_attr = getattr(stat, "FILE_ATTRIBUTE_DIRECTORY", 0x10)
            return bool(attrs & directory_attr)
        except Exception:
            pass
    return False

def _same_filesystem(src: str, dst_parent: str) -> bool:
    """Return whether an atomic rename can stay on one filesystem."""
    try:
        src_dev = os.stat(src, follow_symlinks=False).st_dev
        dst_dev = os.stat(dst_parent, follow_symlinks=False).st_dev
        return src_dev == dst_dev
    except Exception:
        if os.name == "nt":
            try:
                src_drive = os.path.splitdrive(os.path.abspath(src))[0]
                dst_drive = os.path.splitdrive(os.path.abspath(dst_parent))[0]
                return bool(src_drive and dst_drive and src_drive.casefold() == dst_drive.casefold())
            except Exception:
                pass
        return False

def _is_cross_device_error(exc: BaseException) -> bool:
    return (
        getattr(exc, "errno", None) == errno.EXDEV
        or getattr(exc, "winerror", None) == 17  # ERROR_NOT_SAME_DEVICE
    )

def unique_dest_path(dst_dir: str, name: str) -> str:
    base, ext = os.path.splitext(name); candidate = name; i = 1
    while os.path.lexists(os.path.join(dst_dir, candidate)):
        suffix = " - Copy" if i == 1 else f" - Copy ({i})"
        candidate = f"{base}{suffix}{ext}"; i += 1
    return os.path.join(dst_dir, candidate)

def remove_any(path: str):
    if not _path_exists_for_delete(path):
        return
    if os.path.isdir(path) and not _is_dir_link(path):
        shutil.rmtree(path)
    elif os.path.isdir(path) and _is_dir_link(path):
        os.rmdir(path)
    else:
        os.remove(path)

class DeleteCancelled(Exception):
    pass

def _entry_is_junction(entry) -> bool:
    fn = getattr(entry, "is_junction", None)
    if fn:
        try:
            return bool(fn())
        except Exception:
            pass
    return _is_junction(getattr(entry, "path", ""))

def _make_delete_writable(path: str):
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass

def _is_not_empty_error(exc: BaseException) -> bool:
    return (
        getattr(exc, "errno", None) in (errno.ENOTEMPTY, errno.EEXIST)
        or getattr(exc, "winerror", None) in (145, 183)
    )

def _delete_error_message(path: str, exc: BaseException) -> str:
    return f"{path}: {exc}"

def _is_qt_main_thread() -> bool:
    try:
        app = QApplication.instance()
        return app is None or QtCore.QThread.currentThread() == app.thread()
    except Exception:
        return True

def _scan_delete_item_counts(path: str, should_cancel=None) -> tuple[int, dict[str, int]]:
    """Return the total item count and subtree counts without following directory links."""
    counts: dict[str, int] = {}

    def check_cancel():
        if should_cancel and should_cancel():
            raise DeleteCancelled()

    def visit(p: str) -> int:
        check_cancel()
        key = _path_key(p)
        if not _path_exists_for_delete(p):
            counts[key] = 1
            return 1

        if not (os.path.isdir(p) and not _is_dir_link(p)):
            counts[key] = 1
            return 1

        total = 1  # the directory itself
        try:
            with os.scandir(p) as it:
                entries = list(it)
        except OSError:
            counts[key] = total
            return total

        for entry in entries:
            total += visit(entry.path)
        counts[key] = total
        return total

    path = _normalize_fs_path(path)
    return visit(path), counts


def delete_any_permanent_best_effort(
    path: str,
    should_cancel=None,
    on_items_done=None,
    item_count_of=None,
) -> tuple[int, list[str]]:
    """Delete everything possible under path, returning (deleted_count, errors)."""
    deleted = 0
    errors: list[str] = []

    def check_cancel():
        if should_cancel and should_cancel():
            raise DeleteCancelled()

    def planned_count(p: str) -> int:
        if item_count_of:
            try:
                return max(1, int(item_count_of(p)))
            except Exception:
                pass
        return 1

    def mark_done(units: int = 1):
        if on_items_done:
            try:
                on_items_done(max(0, int(units)))
            except Exception:
                pass

    def record_error(p: str, exc: BaseException):
        errors.append(_delete_error_message(p, exc))

    def remove_file_like(p: str) -> bool:
        nonlocal deleted
        check_cancel()
        if not _path_exists_for_delete(p):
            mark_done(1)
            return True
        last_exc = None
        for attempt in range(2):
            try:
                if _is_dir_link(p) and os.path.isdir(p):
                    os.rmdir(p)
                else:
                    os.remove(p)
                deleted += 1
                mark_done(1)
                return True
            except PermissionError as exc:
                last_exc = exc
                if attempt == 0:
                    _make_delete_writable(p)
                    continue
                break
            except OSError as exc:
                last_exc = exc
                if attempt == 0 and getattr(exc, "errno", None) in (errno.EACCES, errno.EPERM):
                    _make_delete_writable(p)
                    continue
                break
        if last_exc:
            record_error(p, last_exc)
        mark_done(1)
        return False

    def remove_dir(p: str) -> bool:
        nonlocal deleted
        check_cancel()
        child_failed = False
        try:
            with os.scandir(p) as it:
                entries = list(it)
        except OSError as exc:
            record_error(p, exc)
            mark_done(planned_count(p))
            return False

        for entry in entries:
            check_cancel()
            child = entry.path
            try:
                is_dir = entry.is_dir(follow_symlinks=False)
                is_link = entry.is_symlink()
                is_junction = _entry_is_junction(entry)
            except OSError as exc:
                record_error(child, exc)
                mark_done(planned_count(child))
                child_failed = True
                continue

            if is_dir and not is_link and not is_junction:
                if not remove_dir(child):
                    child_failed = True
            elif not remove_file_like(child):
                child_failed = True

        if not _path_exists_for_delete(p):
            mark_done(1)
            return not child_failed

        last_exc = None
        for attempt in range(2):
            try:
                os.rmdir(p)
                deleted += 1
                mark_done(1)
                return not child_failed
            except PermissionError as exc:
                last_exc = exc
                if attempt == 0:
                    _make_delete_writable(p)
                    continue
                break
            except OSError as exc:
                last_exc = exc
                if attempt == 0 and getattr(exc, "errno", None) in (errno.EACCES, errno.EPERM):
                    _make_delete_writable(p)
                    continue
                break

        if last_exc and not (child_failed and _is_not_empty_error(last_exc)):
            record_error(p, last_exc)
        mark_done(1)
        return False

    path = _normalize_fs_path(path)
    check_cancel()
    if not _path_exists_for_delete(path):
        mark_done(planned_count(path))
        return 0, []
    if os.path.isdir(path) and not _is_dir_link(path):
        remove_dir(path)
    else:
        remove_file_like(path)
    return deleted, errors

def _qt_move_path_to_trash(path: str) -> bool:
    if not path or not hasattr(QtCore.QFile, "moveToTrash") or not _is_qt_main_thread():
        return False
    try:
        result = QtCore.QFile.moveToTrash(_normalize_fs_path(path))
        if isinstance(result, tuple):
            return bool(result[0])
        return bool(result)
    except Exception as e:
        if DEBUG: print("[delete] QFile.moveToTrash failed:", e)
    return False

def _win_shell_move_path_to_trash(path: str, hwnd: int = 0) -> bool:
    if sys.platform != "win32" or not path:
        return False
    try:
        class _SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", ctypes.c_void_p),
                ("wFunc", ctypes.c_uint),
                ("pFrom", ctypes.c_void_p),
                ("pTo", ctypes.c_void_p),
                ("fFlags", ctypes.c_ushort),
                ("fAnyOperationsAborted", ctypes.c_int),
                ("hNameMappings", ctypes.c_void_p),
                ("lpszProgressTitle", ctypes.c_void_p),
            ]

        shell_path = _normalize_fs_path(os.path.abspath(path))
        from_buf = ctypes.create_unicode_buffer(shell_path + "\0\0")
        op = _SHFILEOPSTRUCTW()
        op.hwnd = ctypes.c_void_p(int(hwnd or 0))
        op.wFunc = 3  # FO_DELETE
        op.pFrom = ctypes.cast(from_buf, ctypes.c_void_p)
        op.pTo = None
        op.fFlags = 0x0040 | 0x0010 | 0x0400 | 0x0004  # ALLOWUNDO, NOCONFIRMATION, NOERRORUI, SILENT
        op.fAnyOperationsAborted = 0
        op.hNameMappings = None
        op.lpszProgressTitle = None

        shfile_operation = ctypes.WinDLL("shell32").SHFileOperationW
        shfile_operation.argtypes = [ctypes.POINTER(_SHFILEOPSTRUCTW)]
        shfile_operation.restype = ctypes.c_int
        res = shfile_operation(ctypes.byref(op))
        return (res == 0) and (not op.fAnyOperationsAborted) and (not os.path.exists(path))
    except Exception as e:
        if DEBUG: print("[delete] SHFileOperationW(ctypes) failed:", e)
    return False

def recycle_path_to_trash(path: str, hwnd: int = 0) -> bool:
    if not path or not _path_exists_for_delete(path):
        return True
    if _qt_move_path_to_trash(path) and not _path_exists_for_delete(path):
        return True
    if HAS_SEND2TRASH:
        try:
            send2trash(path)
            if not _path_exists_for_delete(path):
                return True
        except Exception as e:
            if DEBUG: print("[delete] send2trash failed:", e)
    if HAS_PYWIN32:
        try:
            pFrom = (_normalize_fs_path(path) + "\0\0")
            flags = (shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION |
                     shellcon.FOF_NOERRORUI | shellcon.FOF_SILENT)
            res, aborted = shell.SHFileOperation((int(hwnd), shellcon.FO_DELETE, pFrom, None, flags, False, None, None))
            if (res == 0) and (not aborted) and not _path_exists_for_delete(path):
                return True
        except Exception as e:
            if DEBUG: print("[delete] SHFileOperation failed:", e)
    if _win_shell_move_path_to_trash(path, hwnd):
        return True
    if DEBUG:
        print("[delete] refusing permanent fallback for Recycle Bin delete:", path)
    return False

def recycle_any_best_effort(
    path: str,
    hwnd: int = 0,
    should_cancel=None,
    on_items_done=None,
    item_count_of=None,
) -> tuple[int, list[str]]:
    """Move as much as possible to Recycle Bin while reporting subtree progress."""
    deleted = 0
    errors: list[str] = []

    def check_cancel():
        if should_cancel and should_cancel():
            raise DeleteCancelled()

    def planned_count(p: str) -> int:
        if item_count_of:
            try:
                return max(1, int(item_count_of(p)))
            except Exception:
                pass
        return 1

    def mark_done(units: int):
        if on_items_done:
            try:
                on_items_done(max(0, int(units)))
            except Exception:
                pass

    def record_error(p: str):
        errors.append(f"{p}: Could not move item to Recycle Bin.")

    def dir_has_entries(p: str) -> bool:
        try:
            with os.scandir(p) as it:
                for _entry in it:
                    return True
        except OSError:
            return True
        return False

    def recycle_whole(p: str, units: int | None = None) -> bool:
        nonlocal deleted
        check_cancel()
        work_units = planned_count(p) if units is None else max(1, int(units))
        if not _path_exists_for_delete(p):
            mark_done(work_units)
            return True
        if recycle_path_to_trash(p, hwnd):
            deleted += work_units
            mark_done(work_units)
            return True
        return False

    def recycle_dir_contents(p: str) -> bool:
        check_cancel()
        child_failed = False
        try:
            with os.scandir(p) as it:
                entries = list(it)
        except OSError as exc:
            errors.append(_delete_error_message(p, exc))
            mark_done(planned_count(p))
            return False

        for entry in entries:
            check_cancel()
            child = entry.path
            if recycle_whole(child):
                continue
            try:
                is_dir = entry.is_dir(follow_symlinks=False)
                is_link = entry.is_symlink()
                is_junction = _entry_is_junction(entry)
            except OSError as exc:
                errors.append(_delete_error_message(child, exc))
                mark_done(planned_count(child))
                child_failed = True
                continue

            if is_dir and not is_link and not is_junction:
                if not recycle_dir_contents(child):
                    child_failed = True
            else:
                record_error(child)
                mark_done(1)
                child_failed = True

        if not _path_exists_for_delete(p):
            mark_done(1)
            return not child_failed
        if recycle_whole(p, units=1):
            return not child_failed
        if not (child_failed and dir_has_entries(p)):
            record_error(p)
        mark_done(1)
        return False

    path = _normalize_fs_path(path)
    check_cancel()
    if not _path_exists_for_delete(path):
        mark_done(planned_count(path))
        return 0, []
    if recycle_whole(path):
        return deleted, []
    if os.path.isdir(path) and not _is_dir_link(path):
        recycle_dir_contents(path)
    else:
        record_error(path)
        mark_done(1)
    return deleted, errors

def recycle_to_trash(paths: list, hwnd: int = 0) -> bool:
    if not paths: return True
    ok = True
    for p in paths:
        if not recycle_path_to_trash(p, hwnd):
            ok = False
    return ok

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



class FileOperationManager(QtCore.QObject):
    busyChanged = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._workers = set()
        self._queue = []

    def _prune_finished(self):
        for worker in list(self._workers):
            if worker in self._queue:
                continue
            try:
                running = worker.isRunning()
            except Exception:
                running = False
            if running:
                continue
            self._workers.discard(worker)
            try:
                worker.deleteLater()
            except Exception:
                pass

    def is_busy(self) -> bool:
        self._prune_finished()
        return any(worker.isRunning() for worker in list(self._workers))

    def has_pending(self) -> bool:
        self._prune_finished()
        return self.is_busy() or bool(self._queue)

    def owns(self, worker) -> bool:
        return worker in self._workers

    def cancel_worker(self, worker, wait_ms: int = 300) -> bool:
        if worker not in self._workers:
            return False
        if worker in self._queue:
            self._queue.remove(worker)
            self._workers.discard(worker)
            worker.deleteLater()
            self.busyChanged.emit(self.has_pending())
            return True
        try:
            if worker.isRunning() and hasattr(worker, "cancel"):
                worker.cancel()
                return bool(worker.wait(max(0, int(wait_ms))))
            return True
        except Exception:
            return False

    def register(self, worker) -> bool:
        if worker is None or self.is_busy():
            return False
        worker.setParent(self)
        self._workers.add(worker)
        worker.finished.connect(self._on_worker_finished)
        self.busyChanged.emit(True)
        return True

    def submit(self, worker) -> str:
        if worker is None:
            return "rejected"
        busy = self.is_busy()
        worker.setParent(self)
        self._workers.add(worker)
        worker.finished.connect(self._on_worker_finished)
        if busy:
            self._queue.append(worker)
            self.busyChanged.emit(True)
            return "queued"
        worker.start()
        self.busyChanged.emit(True)
        return "started"

    @QtCore.pyqtSlot()
    def _on_worker_finished(self):
        worker = self.sender()
        if worker in self._workers:
            self._workers.discard(worker)
            try:
                worker.deleteLater()
            except Exception:
                pass
        self._start_next_queued()
        self.busyChanged.emit(self.has_pending())

    def _start_next_queued(self):
        if any(worker.isRunning() for worker in list(self._workers)):
            return
        while self._queue:
            worker = self._queue.pop(0)
            if worker not in self._workers:
                continue
            try:
                worker.start()
                return
            except Exception:
                self._workers.discard(worker)

    def queued_count(self) -> int:
        return len(self._queue)

    def cancel_all(self, wait_ms: int = 8000) -> bool:
        workers = list(self._workers)
        self._queue.clear()
        for worker in workers:
            try:
                if worker.isRunning() and hasattr(worker, "cancel"):
                    worker.cancel()
            except Exception:
                pass

        deadline = time.monotonic() + max(0, int(wait_ms)) / 1000.0
        all_stopped = True
        for worker in workers:
            try:
                if not worker.isRunning():
                    continue
                remaining_ms = max(0, int((deadline - time.monotonic()) * 1000))
                if remaining_ms <= 0 or not worker.wait(remaining_ms):
                    all_stopped = False
            except Exception:
                all_stopped = False

        if all_stopped:
            self._prune_finished()
            self.busyChanged.emit(False)
        return all_stopped


def _cancel_and_wait_child_threads(owner: QtCore.QObject, wait_ms: int) -> bool:
    """Cancel and audit child QThreads before their QObject owner is destroyed."""
    workers = list(owner.findChildren(QtCore.QThread))
    for worker in workers:
        try:
            if worker.isRunning() and hasattr(worker, "cancel"):
                worker.cancel()
        except Exception:
            pass

    deadline = time.monotonic() + max(0, int(wait_ms)) / 1000.0
    all_stopped = True
    for worker in workers:
        try:
            if not worker.isRunning():
                continue
            remaining_ms = max(0, int((deadline - time.monotonic()) * 1000))
            if remaining_ms <= 0 or not worker.wait(remaining_ms):
                all_stopped = False
        except Exception:
            all_stopped = False
    return all_stopped


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
        self._total_bytes = 0
        self._done_bytes = 0
        self._total_items = 0
        self._done_items = 0
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._src_progress_cache: dict[str, tuple[int, int]] = {}
        self._progress_estimated = False
        self.errors = []
        self.error_count = 0
        self.undo_remove_paths = []
        self.undo_move_pairs = []
        self.successful_source_keys = set()
        self.clipboard_payload = None
        self._ui_op = op

    def cancel(self):
        self._cancel = True

    def remaining_source_paths(self) -> list[str]:
        remaining = []
        for src in self.srcs:
            if _path_key(src) in self.successful_source_keys:
                continue
            if _path_exists_for_delete(src):
                remaining.append(src)
        return _dedupe_local_paths(remaining)

    def _mark_source_success(self, src: str):
        self.successful_source_keys.add(_path_key(src))

    def _scan_source_progress(self, path: str) -> tuple[int, int]:
        total_bytes = 0
        total_items = 0
        if os.path.isdir(path) and not _is_dir_link(path):
            for root, dirs, files in os.walk(path):
                if self._cancel:
                    break
                total_items += 1
                if total_items >= FILEOP_FAST_PROGRESS_SCAN_LIMIT:
                    self._progress_estimated = True
                    break
                traversable_dirs = []
                for dirname in dirs:
                    child_dir = os.path.join(root, dirname)
                    if _is_dir_link(child_dir):
                        total_items += 1
                    else:
                        traversable_dirs.append(dirname)
                dirs[:] = traversable_dirs
                for filename in files:
                    if self._cancel:
                        break
                    fp = os.path.join(root, filename)
                    total_items += 1
                    if total_items >= FILEOP_FAST_PROGRESS_SCAN_LIMIT:
                        self._progress_estimated = True
                        break
                    if not os.path.islink(fp):
                        try:
                            total_bytes += max(0, int(os.path.getsize(fp)))
                        except Exception:
                            pass
        else:
            total_items = 1
            try:
                total_bytes = max(0, int(os.path.getsize(path)))
            except Exception:
                total_bytes = 0
        return total_bytes, max(1, total_items)

    def _calc_total(self):
        self._total_bytes = 0
        self._done_bytes = 0
        self._total_items = 0
        self._done_items = 0
        self._src_progress_cache = {}
        for idx, src in enumerate(self.srcs, start=1):
            if self._cancel:
                break
            name = os.path.basename(src.rstrip("\\/")) or src
            self.status.emit(f"Scanning {idx}/{len(self.srcs)}: {name}")
            stats = self._scan_source_progress(src)
            self._src_progress_cache[_path_key(src)] = stats
            self._total_bytes += stats[0]
            self._total_items += stats[1]
            if self._progress_estimated:
                self.status.emit("Large operation detected; starting with estimated progress ...")
                break
        self._total_items = max(1, self._total_items)
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._emit_progress()

    def _emit_progress(self, force: bool = False):
        if force:
            pct = 100
        else:
            item_ratio = min(1.0, self._done_items / max(1, self._total_items))
            if self._total_bytes > 0:
                byte_ratio = min(1.0, self._done_bytes / self._total_bytes)
                item_weight = 0.35 * (self._total_items / (self._total_items + 20.0))
                ratio = (byte_ratio * (1.0 - item_weight)) + (item_ratio * item_weight)
            else:
                ratio = item_ratio
            cap = 95 if self._progress_estimated else 99
            pct = min(cap, max(0, int(ratio * 100)))

        now = time.perf_counter()
        should_emit = (
            force
            or self._last_progress_pct < 0
            or (
                pct > self._last_progress_pct
                and (
                    (pct - self._last_progress_pct) >= 1
                    or (now - self._last_progress_emit_ts) >= 0.05
                )
            )
        )
        if not should_emit:
            return
        self._last_progress_pct = pct
        self._last_progress_emit_ts = now
        self.progress.emit(pct)

    def _tick_progress(self, delta_bytes: int = 0, delta_items: int = 0):
        self._done_bytes += max(0, int(delta_bytes or 0))
        self._done_items += max(0, int(delta_items or 0))
        if self._progress_estimated:
            self._total_bytes = max(self._total_bytes, self._done_bytes + max(1, delta_bytes or 0))
            self._total_items = max(self._total_items, self._done_items + max(1, delta_items or 0))
        self._emit_progress()

    def _source_progress(self, src: str) -> tuple[int, int]:
        try:
            stats = self._src_progress_cache.get(_path_key(src))
        except Exception:
            stats = None
        if stats is not None:
            return stats
        if self._progress_estimated:
            return 0, 1
        return self._scan_source_progress(src)

    def _skip_source_progress(self, src):
        total_bytes, total_items = self._source_progress(src)
        self._tick_progress(total_bytes, total_items)

    def _emit_source_done(self):
        self._emit_progress()

    def _new_sibling_work_path(self, dst: str, kind: str) -> str:
        parent = os.path.dirname(dst) or os.curdir
        os.makedirs(parent, exist_ok=True)
        candidate = os.path.join(parent, f".__mprn_{kind}_{uuid.uuid4().hex}")
        while os.path.lexists(candidate):
            candidate = os.path.join(parent, f".__mprn_{kind}_{uuid.uuid4().hex}")
        return candidate

    def _backup_destination(self, dst: str) -> str | None:
        if not os.path.lexists(dst):
            return None
        backup = self._new_sibling_work_path(dst, "backup")
        os.replace(dst, backup)
        return backup

    def _cleanup_path(self, path: str):
        if not path or not os.path.lexists(path):
            return
        remove_any(path)

    def _restore_backup(self, dst: str, backup: str | None):
        if not backup:
            return
        restore_errors = []
        try:
            self._cleanup_path(dst)
        except Exception as exc:
            restore_errors.append(f"Could not remove partial destination {dst}: {exc}")
        try:
            if os.path.lexists(backup):
                os.replace(backup, dst)
        except Exception as exc:
            restore_errors.append(f"Could not restore original destination {dst}: {exc}")
        if restore_errors:
            raise RuntimeError("; ".join(restore_errors))

    def _discard_backup(self, backup: str | None, dst: str):
        if not backup or not os.path.lexists(backup):
            return
        try:
            self._cleanup_path(backup)
        except Exception as exc:
            self._record_copy_error(backup, dst, f"Operation succeeded, but backup cleanup failed: {exc}")

    def _rollback_destination(self, dst: str, backup: str | None):
        try:
            if backup:
                self._restore_backup(dst, backup)
            else:
                self._cleanup_path(dst)
        except Exception as exc:
            self._record_copy_error(dst, dst, f"Rollback failed: {exc}")

    def _copy_file(self, src, dst):
        copied = 0
        temp_path = None
        try:
            os.makedirs(os.path.dirname(dst) or os.curdir, exist_ok=True)
            temp_path = self._new_sibling_work_path(dst, "partial")
            with open(src, "rb") as fsrc, open(temp_path, "xb") as fdst:
                while True:
                    if self._cancel:
                        return False
                    buf = fsrc.read(1024 * 1024)
                    if not buf:
                        break
                    fdst.write(buf)
                    copied += len(buf)
                    self._tick_progress(delta_bytes=len(buf))
                fdst.flush()
                os.fsync(fdst.fileno())
            try:
                shutil.copystat(src, temp_path, follow_symlinks=True)
            except Exception:
                pass
            if self._cancel:
                return False
            os.replace(temp_path, dst)
            temp_path = None
            self._tick_progress(delta_items=1)
            return True
        except Exception:
            self._skip_file_progress(src, copied)
            raise
        finally:
            if temp_path and os.path.lexists(temp_path):
                try:
                    self._cleanup_path(temp_path)
                except Exception:
                    pass

    def _copy_link(self, src: str, dst: str) -> bool:
        temp_path = None
        try:
            if _is_junction(src):
                raise OSError("Copying Windows directory junctions is not supported; the junction was left unchanged.")
            target = os.readlink(src)
            temp_path = self._new_sibling_work_path(dst, "link")
            os.symlink(target, temp_path, target_is_directory=_link_target_is_directory(src))
            try:
                shutil.copystat(src, temp_path, follow_symlinks=False)
            except Exception:
                pass
            if self._cancel:
                return False
            os.replace(temp_path, dst)
            temp_path = None
            self._tick_progress(delta_items=1)
            return True
        finally:
            if temp_path and os.path.lexists(temp_path):
                try:
                    remove_any(temp_path)
                except Exception:
                    pass

    def _skip_file_progress(self, src, copied_bytes: int = 0):
        try:
            size = max(0, int(os.path.getsize(src)))
        except Exception:
            size = 0
        remaining = max(0, size - max(0, int(copied_bytes or 0)))
        self._tick_progress(delta_bytes=remaining, delta_items=1)

    def _record_copy_error(self, src, dst, exc):
        if self._cancel:
            return
        self.error_count += 1
        if len(self.errors) < FILEOP_ERROR_DETAIL_LIMIT:
            self.errors.append(f"{src} -> {dst}: {exc}")
        name = os.path.basename(str(src).rstrip("\\/")) or str(src)
        self.status.emit(f"Failed: {name}")

    def _can_undo_new_destination(self, existed_before: bool, action: str | None) -> bool:
        return (not existed_before) or action == "copy"

    def _remember_created_for_undo(self, path: str):
        if not path:
            return
        key = _path_key(path)
        if any(_path_key(p) == key for p in self.undo_remove_paths):
            return
        self.undo_remove_paths.append(path)

    def _remember_move_for_undo(self, final_path: str, original_path: str):
        if not final_path or not original_path:
            return
        self.undo_move_pairs.append((final_path, original_path))

    def _copy_dir_recursive(self, src_dir, dst_dir):
        if self._cancel:
            return False
        self._tick_progress(delta_items=1)
        ok = True
        try:
            entries = list(os.scandir(src_dir))
        except Exception as exc:
            self._record_copy_error(src_dir, dst_dir, exc)
            return False

        for entry in entries:
            if self._cancel:
                return False
            src_path = entry.path
            dst_path = os.path.join(dst_dir, entry.name)
            try:
                if _is_junction(src_path):
                    raise OSError("Copying Windows directory junctions is not supported; the junction was left unchanged.")
                if entry.is_symlink():
                    if not self._copy_link(src_path, dst_path):
                        return False
                elif entry.is_dir(follow_symlinks=False):
                    os.makedirs(dst_path, exist_ok=False)
                    if not self._copy_dir_recursive(src_path, dst_path):
                        ok = False
                elif not self._copy_file(src_path, dst_path):
                    return False
            except Exception as exc:
                self._record_copy_error(src_path, dst_path, exc)
                self._skip_source_progress(src_path)
                ok = False
        return ok

    def _copy_to_new_path(self, src: str, dst: str) -> bool:
        if _is_junction(src):
            raise OSError("Copying Windows directory junctions is not supported; the junction was left unchanged.")
        if os.path.islink(src):
            return self._copy_link(src, dst)
        if os.path.isdir(src):
            os.makedirs(dst, exist_ok=False)
            return self._copy_dir_recursive(src, dst)
        return self._copy_file(src, dst)

    def _copy_source_transactional(self, src: str, dst: str, action: str | None, existed: bool) -> bool:
        backup = None
        if existed and action not in {"skip", "copy", "overwrite"}:
            self._record_copy_error(src, dst, "No conflict resolution was selected; destination was left unchanged.")
            self._skip_source_progress(src)
            return False
        if existed and action == "skip":
            self._skip_source_progress(src)
            return True
        if existed and action == "copy":
            dst = unique_dest_path(self.dst_dir, os.path.basename(dst))
            existed = False
        if existed and action == "overwrite":
            try:
                backup = self._backup_destination(dst)
            except Exception as exc:
                self._record_copy_error(src, dst, f"Could not protect existing destination: {exc}")
                self._skip_source_progress(src)
                return False

        created_for_undo = self._can_undo_new_destination(existed, action)
        copied_ok = False
        try:
            copied_ok = self._copy_to_new_path(src, dst)
            if copied_ok and not self._cancel:
                self._discard_backup(backup, dst)
                if created_for_undo:
                    self._remember_created_for_undo(dst)
                return True
        except Exception as exc:
            self._record_copy_error(src, dst, exc)
        self._rollback_destination(dst, backup)
        return False

    def _move_source_transactional(self, src: str, dst: str, action: str | None, existed: bool) -> bool:
        if existed and action not in {"skip", "copy", "overwrite"}:
            self._record_copy_error(src, dst, "No conflict resolution was selected; destination was left unchanged.")
            self._skip_source_progress(src)
            return False
        if existed and action == "skip":
            self._skip_source_progress(src)
            return False
        if existed and action == "copy":
            dst = unique_dest_path(self.dst_dir, os.path.basename(dst))
            existed = False

        can_undo_move = self._can_undo_new_destination(existed, action)
        src_progress = self._source_progress(src)
        same_filesystem = _same_filesystem(src, os.path.dirname(dst) or self.dst_dir)

        if same_filesystem:
            backup = None
            if existed and action == "overwrite":
                try:
                    backup = self._backup_destination(dst)
                except Exception as exc:
                    self._record_copy_error(src, dst, f"Could not protect existing destination: {exc}")
                    self._skip_source_progress(src)
                    return False
            try:
                os.replace(src, dst)
                self._tick_progress(src_progress[0], src_progress[1])
                self._discard_backup(backup, dst)
                self._mark_source_success(src)
                if can_undo_move:
                    self._remember_move_for_undo(dst, src)
                return True
            except Exception as exc:
                self._rollback_destination(dst, backup)
                if not _is_cross_device_error(exc):
                    self._record_copy_error(src, dst, exc)
                    self._skip_source_progress(src)
                    return False
                self.status.emit("Filesystem boundary detected; switching to safe copy-and-cleanup move ...")

        if _is_junction(src):
            self._record_copy_error(
                src,
                dst,
                "Moving a Windows directory junction across filesystems is not supported; the junction was left unchanged.",
            )
            self._skip_source_progress(src)
            return False

        # Cross-filesystem moves are deliberately copy-first.  shutil.move() cannot
        # be cancelled and may partially delete the source before reporting an error.
        staging = self._new_sibling_work_path(dst, "move")
        try:
            copied_ok = self._copy_to_new_path(src, staging)
        except Exception as exc:
            copied_ok = False
            self._record_copy_error(src, staging, exc)

        if not copied_ok or self._cancel:
            try:
                self._cleanup_path(staging)
            except Exception as exc:
                self._record_copy_error(staging, dst, f"Could not remove incomplete staging copy: {exc}")
            return False

        backup = None
        try:
            if existed and action == "overwrite":
                backup = self._backup_destination(dst)
            os.replace(staging, dst)
            staging = None
        except Exception as exc:
            self._record_copy_error(src, dst, f"Could not promote the completed staging copy: {exc}")
            try:
                if staging and os.path.lexists(staging):
                    self._cleanup_path(staging)
            except Exception as cleanup_exc:
                self._record_copy_error(staging, dst, f"Could not remove staging copy: {cleanup_exc}")
            self._rollback_destination(dst, backup)
            return False

        # The destination is now a complete copy.  Never roll it back if source
        # cleanup is cancelled or fails: it may be the only complete copy left.
        self._discard_backup(backup, dst)
        if self._cancel:
            return False

        cleanup_errors = []
        try:
            if os.path.isdir(src) and not _is_dir_link(src):
                _deleted, cleanup_errors = delete_any_permanent_best_effort(
                    src,
                    should_cancel=lambda: self._cancel,
                )
            else:
                remove_any(src)
        except DeleteCancelled:
            return False
        except Exception as exc:
            cleanup_errors = [str(exc)]

        if not _path_exists_for_delete(src):
            self._mark_source_success(src)
            if can_undo_move:
                self._remember_move_for_undo(dst, src)
            return True

        detail = cleanup_errors[0] if cleanup_errors else "source still exists after cleanup"
        self._record_copy_error(
            src,
            dst,
            f"Destination copy is complete, but source cleanup was incomplete: {detail}",
        )
        return False

    def run(self):
        try:
            self.status.emit(f"Scanning items for {self.op} ...")
            self._calc_total()
            if self._cancel:
                self.error.emit("Operation cancelled.")
                return
            self.status.emit(f"Preparing {self.op} ...")

            for src in self.srcs:
                if self._cancel:
                    break
                if not os.path.lexists(src):
                    self._mark_source_success(src)
                    self._skip_source_progress(src)
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

                if os.path.isdir(src) and not _is_dir_link(src) and _is_subpath(dst, src):
                    self.status.emit(f"Skipped nested destination: {base}")
                    self._skip_source_progress(src)
                    self._emit_source_done()
                    continue

                existed = os.path.lexists(dst)
                action = self.conflict_map.get(src) if existed else None
                if self.op == "copy":
                    self._copy_source_transactional(src, dst, action, existed)
                else:
                    self._move_source_transactional(src, dst, action, existed)
                self._emit_source_done()

            if self._cancel:
                self.error.emit("Operation cancelled.")
                return
            self._done_bytes = max(self._done_bytes, self._total_bytes)
            self._done_items = max(self._done_items, self._total_items)
            self._emit_progress(force=True)
            self.finished_ok.emit()
        except Exception as exc:
            self.error.emit(str(exc))

class DeleteWorker(QtCore.QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished_ok = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, paths: list, permanent: bool = False, hwnd: int = 0, parent=None):
        super().__init__(parent)
        self.paths = [p for p in paths if p]
        self.permanent = permanent
        self.hwnd = int(hwnd or 0)
        self._cancel = False
        self._total = 1
        self._done = 0
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._subtree_item_counts: dict[str, int] = {}
        self.deleted_count = 0
        self.errors = []

    def cancel(self):
        self._cancel = True

    def _emit_progress(self, force: bool = False):
        if force:
            pct = 100
        else:
            total = max(1, int(self._total or 1))
            pct = min(99, max(0, int(self._done * 100 / total)))
        now = time.perf_counter()
        should_emit = (
            force
            or self._last_progress_pct < 0
            or (
                pct > self._last_progress_pct
                and (
                    (pct - self._last_progress_pct) >= 1
                    or (now - self._last_progress_emit_ts) >= 0.05
                )
            )
        )
        if not should_emit:
            return
        self._last_progress_pct = pct
        self._last_progress_emit_ts = now
        self.progress.emit(pct)

    def _scan_total_items(self):
        self._total = 0
        self._done = 0
        self._subtree_item_counts = {}
        for idx, path in enumerate(self.paths, start=1):
            if self._cancel:
                raise DeleteCancelled()
            name = os.path.basename(path.rstrip("\\/")) or os.path.basename(path) or path
            self.status.emit(f"Scanning {idx}/{len(self.paths)}: {name}")
            count, subtree_counts = _scan_delete_item_counts(
                path,
                should_cancel=lambda: self._cancel,
            )
            self._total += count
            self._subtree_item_counts.update(subtree_counts)
        self._total = max(1, self._total)
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._emit_progress()

    def _item_count_of(self, path: str) -> int:
        return max(1, int(self._subtree_item_counts.get(_path_key(path), 1)))

    def _on_items_done(self, units: int):
        self._done += max(0, int(units or 0))
        self._emit_progress()

    def run(self):
        coinit = False
        try:
            self.status.emit("Scanning items for delete ...")
            self._scan_total_items()
            if self._cancel:
                self.error.emit("Operation cancelled.")
                return

            if sys.platform == "win32" and HAS_PYWIN32:
                try:
                    pythoncom.CoInitialize()
                    coinit = True
                except Exception:
                    coinit = False

            verb = "Deleting" if self.permanent else "Sending to Recycle Bin"
            total_paths = len(self.paths)
            if total_paths == 0:
                self._done = self._total
                self._emit_progress(force=True)
                self.finished_ok.emit()
                return

            for idx, path in enumerate(self.paths, start=1):
                if self._cancel:
                    self.error.emit("Operation cancelled.")
                    return

                name = os.path.basename(path.rstrip("\\/")) or os.path.basename(path) or path
                self.status.emit(f"{verb} {idx}/{total_paths}: {name}")

                try:
                    if self.permanent:
                        deleted, errors = delete_any_permanent_best_effort(
                            path,
                            should_cancel=lambda: self._cancel,
                            on_items_done=self._on_items_done,
                            item_count_of=self._item_count_of,
                        )
                    else:
                        deleted, errors = recycle_any_best_effort(
                            path,
                            hwnd=self.hwnd,
                            should_cancel=lambda: self._cancel,
                            on_items_done=self._on_items_done,
                            item_count_of=self._item_count_of,
                        )
                    self.deleted_count += deleted
                    self.errors.extend(errors)
                except DeleteCancelled:
                    self.error.emit("Operation cancelled.")
                    return
                except Exception as e:
                    self.errors.append(f"{path}: {e}")
                    self._on_items_done(self._item_count_of(path))

            self._done = max(self._done, self._total)
            self._emit_progress(force=True)
            self.finished_ok.emit()
        except DeleteCancelled:
            self.error.emit("Operation cancelled.")
        except Exception as e:
            self.error.emit(str(e))
        finally:
            if coinit:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass



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
    QToolButton#quickBookmarkBtn {{ text-align: left; padding-left: 4px; padding-right: 4px; }}
    QToolButton#quickBookmarkMoreBtn {{ padding-left: 4px; padding-right: 4px; }}
    QLabel#modeBadge {{ padding: 0 6px; border-radius: 6px; }}
    QLabel#crumbSep {{ padding: 0 0px; margin: 0; }}
    """

def _star_polygon(cx, cy, r, inner_ratio=0.45):
    return QPolygonF([
        QtCore.QPointF(cx + math.cos(-math.pi/2 + i * math.pi/5) * (r if i % 2 == 0 else r * inner_ratio),
                       cy + math.sin(-math.pi/2 + i * math.pi/5) * (r if i % 2 == 0 else r * inner_ratio))
        for i in range(10)
    ])

def _setup_readonly_table(table: QTableWidget, labels, resize_modes, row_count=None):
    table.setColumnCount(len(labels)); table.setHorizontalHeaderLabels(list(labels))
    if row_count is not None: table.setRowCount(row_count)
    header = table.horizontalHeader()
    for col, mode in enumerate(resize_modes):
        header.setSectionResizeMode(col, mode)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    return table

def _set_table_row_items(table: QTableWidget, row: int, *values):
    for col, value in enumerate(values):
        table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))

def _add_dialog_button_box(layout, parent, buttons, accept_slot, reject_slot=None):
    btns = QDialogButtonBox(buttons, parent)
    if accept_slot: btns.accepted.connect(accept_slot)
    if reject_slot: btns.rejected.connect(reject_slot)
    layout.addWidget(btns)
    return btns

def _empty_bookmark_item():
    return {"enabled": False, "name": "", "path": ""}

def _apply_palette_colors(widget, colors):
    pal = widget.palette()
    for role, rgb in colors.items(): pal.setColor(role, QColor(*rgb))
    widget.setPalette(pal)

_THEME_PALETTE_SHARED = {
    QPalette.ToolTipBase: (255, 255, 255),
    QPalette.Highlight: (64, 128, 255),
    QPalette.HighlightedText: (255, 255, 255),
}
_THEME_CSS_SHARED = {
    "busy_border": "#5E9BFF",
    "tree_selected_fg": "#FFFFFF",
    "active_root_border": "#5E9BFF",
    "message_box_css": "",
}
_THEME_CSS_TEMPLATE = """
        QMainWindow { background: %(window_bg)s; }
        QLineEdit, QPushButton, QToolButton { background: %(control_bg)s; border: 1px solid %(control_border)s; border-radius: 8px; color: %(control_fg)s; min-height: 0px; }
        QToolButton[busy="true"] { background: %(busy_bg)s; border: 1px solid %(busy_border)s; color: %(busy_fg)s; font-weight: 600; }
        QLineEdit:focus, QPushButton:focus, QToolButton:focus, QComboBox:focus { border: 2px solid %(focus_border)s; }
        QTreeView:focus { border: 2px solid %(focus_border)s; }
        QTreeView { background: %(tree_bg)s; alternate-background-color: %(tree_alt_bg)s; border: 1px solid %(tree_border)s; border-radius: 10px; }
        QTreeView::item { color: %(tree_fg)s; }
        QTreeView::item:selected { background: %(tree_selected_bg)s; color: %(tree_selected_fg)s; }
        QTreeView::item:hover { %(tree_hover_rule)s; }
        QHeaderView::section { background: %(header_bg)s; color: %(header_fg)s; border: 0; border-right: 1px solid %(header_border)s; }
        QMenu { background-color: %(menu_bg)s; color: %(menu_fg)s; border: 1px solid %(menu_border)s; border-radius: 8px; }
        QMenu::item { padding: 6px 12px; }
        QMenu::item:selected { background: %(menu_selected_bg)s; }
        QPushButton#crumb { background: %(crumb_bg)s; border: 1px solid %(crumb_border)s; padding: 0 6px; border-radius: 6px; text-align: left; color: %(crumb_fg)s; }
        QPushButton#crumb:hover { background: %(crumb_hover_bg)s; }
        QLabel#crumbSep { color: %(crumb_sep)s; }
        QLabel#modeBadge { color: %(mode_badge_fg)s; background: %(mode_badge_bg)s; border: 1px solid %(mode_badge_border)s; }
        QScrollArea#crumbScroll, QScrollArea#crumbScroll[active="true"] { border: 0px solid transparent; }
        QScrollArea#crumbScroll > QWidget#crumbViewport { background: transparent; }
        QWidget#paneRoot { border: 1px solid transparent; border-radius: 10px; }
        QWidget#paneRoot[active="true"] { border: 1px solid %(active_root_border)s; background: %(active_root_bg)s; }
        QWidget#paneRoot[active="true"] QPushButton#crumb { background: %(active_crumb_bg)s; border-color: %(active_crumb_border)s; }
        QWidget#paneRoot[active="true"] QPushButton#crumb:hover { background: %(active_crumb_hover_bg)s; }
        QWidget#paneRoot[active="true"] QTreeView { border-color: %(active_tree_border)s; }
        QWidget#paneRoot[active="true"] QLineEdit { border: 1px solid %(active_lineedit_border)s; }
        QWidget#paneRoot[active="true"] QLineEdit:focus { border: 1px solid %(active_root_border)s; }
        QWidget#paneRoot[active="true"] QToolButton, QWidget#paneRoot[active="true"] QPushButton { border-color: %(active_btn_border)s; }
        QWidget#paneRoot[drop_target="true"] { border: 2px solid %(drop_border)s; background: %(drop_bg)s; }
        QWidget#paneRoot[drop_target="true"] QTreeView { border-color: %(drop_tree_border)s; }
        QWidget#paneRoot[drop_target="true"] QLineEdit { border-color: %(drop_lineedit_border)s; }
        %(message_box_css)s
    """
_THEME_STYLE_SPECS = {
    "dark": {
        "palette": _THEME_PALETTE_SHARED | {
            QPalette.Window: (28, 30, 34), QPalette.Base: (22, 24, 28), QPalette.AlternateBase: (30, 32, 36),
            QPalette.Text: (230, 233, 238), QPalette.ButtonText: (230, 233, 238), QPalette.WindowText: (230, 233, 238),
            QPalette.ToolTipText: (30, 30, 30), QPalette.Button: (38, 40, 46),
        },
        "css": _THEME_CSS_SHARED | {
            "window_bg": "#1C1E22", "control_bg": "#26282E", "control_border": "#33363D", "control_fg": "#E6E9EE",
            "busy_bg": "#2B3E62", "busy_fg": "#EAF1FF", "focus_border": "#8FC1FF",
            "tree_bg": "#16181C", "tree_alt_bg": "#1E2026", "tree_border": "#2B2E34", "tree_fg": "#E6E9EE",
            "tree_selected_bg": "#4068FF", "tree_hover_rule": "background: rgba(160,190,255,0.22)",
            "header_bg": "#20232A", "header_fg": "#D6DAE2", "header_border": "#2E3138",
            "menu_bg": "#20232A", "menu_fg": "#E6E9EE", "menu_border": "#2B2E34", "menu_selected_bg": "#2D3550",
            "crumb_bg": "rgba(255,255,255,0.05)", "crumb_border": "#2B2E34", "crumb_fg": "#E6E9EE",
            "crumb_hover_bg": "rgba(255,255,255,0.09)", "crumb_sep": "#7F8796",
            "mode_badge_fg": "#DDE8FF", "mode_badge_bg": "rgba(94,155,255,0.16)", "mode_badge_border": "rgba(94,155,255,0.38)",
            "active_root_bg": "rgba(94, 155, 255, 0.06)", "active_crumb_bg": "rgba(94,155,255,0.16)",
            "active_crumb_border": "rgba(94,155,255,0.40)", "active_crumb_hover_bg": "rgba(94,155,255,0.22)",
            "active_tree_border": "rgba(94,155,255,0.45)", "active_lineedit_border": "rgba(94,155,255,0.35)",
            "active_btn_border": "rgba(94,155,255,0.25)", "drop_border": "#34C88A", "drop_bg": "rgba(52, 200, 138, 0.14)",
            "drop_tree_border": "rgba(52,200,138,0.70)", "drop_lineedit_border": "rgba(52,200,138,0.58)",
            "message_box_css": """
        QMessageBox { background: #FFFFFF; color: #000000; }
        QMessageBox QLabel { color: #000000; }
        QMessageBox QPushButton { color: #000000; background: #F2F4F8; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px 10px; }
        QMessageBox QPushButton:hover { background: #EAEFFF; }
    """,
        },
    },
    "light": {
        "palette": _THEME_PALETTE_SHARED | {
            QPalette.Window: (248, 249, 251), QPalette.Base: (255, 255, 255), QPalette.AlternateBase: (246, 248, 250),
            QPalette.Text: (28, 28, 30), QPalette.ButtonText: (28, 28, 30), QPalette.WindowText: (28, 28, 30),
            QPalette.ToolTipText: (28, 28, 30), QPalette.Button: (242, 244, 248),
        },
        "css": _THEME_CSS_SHARED | {
            "window_bg": "#F8F9FB", "control_bg": "#FFFFFF", "control_border": "#DDE1E6", "control_fg": "#1C1C1E",
            "busy_bg": "#EAF2FF", "busy_fg": "#1A3B8A", "focus_border": "#2A63FF",
            "tree_bg": "#FFFFFF", "tree_alt_bg": "#F6F8FA", "tree_border": "#DDE1E6", "tree_fg": "#1A1A1A",
            "tree_selected_bg": "#2A63FF", "tree_hover_rule": "background: rgba(64,104,255,0.14); color: #1A1A1A",
            "header_bg": "#F1F3F7", "header_fg": "#333", "header_border": "#E5E8EE",
            "menu_bg": "#FFFFFF", "menu_fg": "#1C1C1E", "menu_border": "#CED3DB", "menu_selected_bg": "#EAEFFF",
            "crumb_bg": "rgba(0,0,0,0.04)", "crumb_border": "#E5E8EE", "crumb_fg": "#1C1C1E",
            "crumb_hover_bg": "rgba(0,0,0,0.07)", "crumb_sep": "#7A7F89",
            "mode_badge_fg": "#1A3B8A", "mode_badge_bg": "#EAF2FF", "mode_badge_border": "#BBD0FF",
            "active_root_bg": "rgba(64, 128, 255, 0.06)", "active_crumb_bg": "rgba(64,128,255,0.12)",
            "active_crumb_border": "rgba(64,128,255,0.40)", "active_crumb_hover_bg": "rgba(64,128,255,0.18)",
            "active_tree_border": "rgba(64,128,255,0.40)", "active_lineedit_border": "rgba(64,128,255,0.35)",
            "active_btn_border": "rgba(64,128,255,0.25)", "drop_border": "#22A96A", "drop_bg": "rgba(34, 169, 106, 0.10)",
            "drop_tree_border": "rgba(34,169,106,0.62)", "drop_lineedit_border": "rgba(34,169,106,0.50)",
        },
    },
}

def _apply_theme(app: QApplication, theme: str):
    spec = _THEME_STYLE_SPECS["light" if theme == "light" else "dark"]
    _apply_palette_colors(app, spec["palette"])
    app.setStyleSheet(_common_css() + (_THEME_CSS_TEMPLATE % spec["css"]))

def apply_dark_style(app: QApplication): _apply_theme(app, "dark")
def apply_light_style(app: QApplication): _apply_theme(app, "light")
def apply_theme_by_name(app: QApplication, theme: str): _apply_theme(app, theme)



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
        p.setPen(QPen(star_edge, 1.0))
        p.setBrush(QBrush(star_fill))
        p.drawPolygon(_star_polygon(w - 6.0, h - 6.0, 3.2, inner_ratio=0.44))
    return _make_icon(22, 22, paint)
def icon_star(checked: bool, theme: str):
    def paint(p: QPainter, w, h):
        poly = _star_polygon(w/2, h/2, min(w,h)/2.6)
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

def icon_shortcuts(theme: str):
    def paint(p: QPainter, w, h):
        border = QColor(180, 200, 255) if theme == "dark" else QColor(70, 100, 210)
        key_fill = QColor(80, 92, 116) if theme == "dark" else QColor(240, 244, 255)
        line = QColor(220, 228, 250) if theme == "dark" else QColor(90, 115, 210)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(QPen(border, 1.6))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(2, 4, w - 4, h - 7, 4, 4)
        p.setBrush(QBrush(key_fill))
        for y in (8, 12):
            for x in (6, 10, 14):
                p.drawRoundedRect(x, y, 3, 2, 0.8, 0.8)
        p.setPen(QPen(line, 1.8))
        p.drawLine(6, 16, w - 6, 16)
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

def icon_explorer(theme: str):
    def paint(p: QPainter, w, h):
        line = QColor(190, 195, 210) if theme=="dark" else QColor(90, 100, 120)
        fill = QColor(238, 190, 72) if theme=="dark" else QColor(255, 205, 85)
        tab = QColor(255, 220, 120) if theme=="dark" else QColor(255, 232, 150)
        arrow = QColor(95, 180, 255) if theme=="dark" else QColor(35, 115, 220)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(QPen(line, 1.3)); p.setBrush(QBrush(tab))
        p.drawRoundedRect(QtCore.QRectF(3.5, 5.0, 7.0, 4.5), 1.5, 1.5)
        p.setBrush(QBrush(fill))
        p.drawRoundedRect(QtCore.QRectF(3.0, 7.0, 16.0, 10.5), 2.3, 2.3)
        p.setPen(QPen(arrow, 2.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(9, 12, 15, 12)
        p.drawLine(12, 9, 15, 12)
        p.drawLine(12, 15, 15, 12)
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

def load_recent_path_history() -> list[str]:
    s = QSettings(ORG_NAME, APP_NAME)
    val = s.value("pathbar/recent_paths", [])
    out = []
    if isinstance(val, list):
        for p in val:
            try:
                sp = str(p).strip()
                if sp:
                    out.append(_normalize_fs_path(sp))
            except Exception:
                pass
    seen = set()
    uniq = []
    for p in out:
        k = os.path.normcase(_normalize_fs_path(p))
        if k in seen:
            continue
        seen.add(k)
        uniq.append(p)
    return uniq[:PATH_HISTORY_LIMIT]

def save_recent_path_history(items: list[str]):
    out = []
    seen = set()
    for p in list(items or []):
        try:
            sp = str(p).strip()
        except Exception:
            continue
        if not sp:
            continue
        np = _normalize_fs_path(sp)
        k = os.path.normcase(np)
        if k in seen:
            continue
        seen.add(k)
        out.append(np)
        if len(out) >= PATH_HISTORY_LIMIT:
            break
    s = QSettings(ORG_NAME, APP_NAME)
    s.setValue("pathbar/recent_paths", out)
    s.sync()


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
        return out[:BOOKMARK_LIMIT]
    return []

def save_named_bookmarks(items: list):
    s = QSettings(ORG_NAME, APP_NAME)
    s.setValue("bookmarks/named_items", items[:BOOKMARK_LIMIT]); s.sync()

def _derive_name_from_path(p: str) -> str:
    try:
        p = nice_path(p)
        if p.endswith(os.sep) and len(p) <= 3: return p
        base = os.path.basename(p.rstrip("\\/")); return base or p
    except Exception: return p

def file_extension_label(name_or_path: str, is_dir: bool = False) -> str:
    if is_dir:
        return ""
    try:
        base = os.path.basename(str(name_or_path or ""))
        ext = os.path.splitext(base)[1]
        return ext[1:].lower() if ext.startswith(".") and len(ext) > 1 else ""
    except Exception:
        return ""

def migrate_legacy_favorites_into_named(items: list) -> list:
    try:
        s = QSettings(ORG_NAME, APP_NAME)
        favs = s.value("favorites/paths", [])
        if not favs: return items[:BOOKMARK_LIMIT]
        existing = {os.path.normcase(x.get("path","")) for x in items}
        for p in favs:
            np = nice_path(str(p))
            if os.path.normcase(np) in existing: continue
            if len(items) >= BOOKMARK_LIMIT: break
            items.append({"enabled": True, "name": _derive_name_from_path(np), "path": np})
        s.remove("favorites/paths"); return items[:BOOKMARK_LIMIT]
    except Exception:
        return items[:BOOKMARK_LIMIT]


IS_DIR_ROLE = Qt.UserRole + 99
SIZE_BYTES_ROLE = Qt.UserRole + 100
SEARCH_ICON_READY_ROLE = Qt.UserRole + 101
NAME_FOLD_ROLE = Qt.UserRole + 102
ICON_KEY_ROLE = Qt.UserRole + 103

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
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.TextAlignmentRole:
            if section == 1:
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return super().headerData(section, orientation, role)
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


def _icon_cache_key(path: str, is_dir: bool) -> str:
    if is_dir:
        return "folder"
    try:
        ext = os.path.splitext(os.path.basename(str(path or "")))[1].lower()
    except Exception:
        ext = ""
    # These file types can have path-specific artwork. Everything else is
    # cached by extension so thousands of files do not trigger Shell lookups.
    if ext in {".exe", ".ico", ".lnk", ".url"}:
        return "path:" + _path_key(path)
    return "ext:" + (ext or "<none>")


def _load_windows_shell_icon_bgra(path: str, is_dir: bool, size: int = 24):
    """Return (BGRA bytes, width, height) without creating QPixmap off-thread."""
    if sys.platform != "win32":
        return None, 0, 0
    try:
        from ctypes import wintypes

        class SHFILEINFOW(ctypes.Structure):
            _fields_ = [
                ("hIcon", wintypes.HICON),
                ("iIcon", ctypes.c_int),
                ("dwAttributes", wintypes.DWORD),
                ("szDisplayName", wintypes.WCHAR * 260),
                ("szTypeName", wintypes.WCHAR * 80),
            ]

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", wintypes.DWORD),
                ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long),
                ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD),
                ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD),
                ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long),
                ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD),
            ]

        class RGBQUAD(ctypes.Structure):
            _fields_ = [
                ("rgbBlue", ctypes.c_ubyte),
                ("rgbGreen", ctypes.c_ubyte),
                ("rgbRed", ctypes.c_ubyte),
                ("rgbReserved", ctypes.c_ubyte),
            ]

        class BITMAPINFO(ctypes.Structure):
            _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", RGBQUAD * 1)]

        shell32 = ctypes.WinDLL("shell32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

        sh_get = shell32.SHGetFileInfoW
        sh_get.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(SHFILEINFOW), ctypes.c_uint, ctypes.c_uint]
        sh_get.restype = ctypes.c_size_t
        gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
        gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
        gdi32.CreateDIBSection.argtypes = [ctypes.c_void_p, ctypes.POINTER(BITMAPINFO), ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p, ctypes.c_uint]
        gdi32.CreateDIBSection.restype = ctypes.c_void_p
        gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        gdi32.SelectObject.restype = ctypes.c_void_p
        gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
        gdi32.DeleteObject.restype = wintypes.BOOL
        gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
        gdi32.DeleteDC.restype = wintypes.BOOL
        user32.DrawIconEx.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int, wintypes.HICON, ctypes.c_int, ctypes.c_int, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint]
        user32.DrawIconEx.restype = wintypes.BOOL
        user32.DestroyIcon.argtypes = [wintypes.HICON]
        user32.DestroyIcon.restype = wintypes.BOOL

        SHGFI_ICON = 0x000000100
        SHGFI_SMALLICON = 0x000000001
        SHGFI_USEFILEATTRIBUTES = 0x000000010
        FILE_ATTRIBUTE_DIRECTORY = 0x00000010
        FILE_ATTRIBUTE_NORMAL = 0x00000080

        p = str(path or "")
        ext = os.path.splitext(p)[1].lower()
        path_specific = (not is_dir) and ext in {".exe", ".ico", ".lnk", ".url"}
        attrs = FILE_ATTRIBUTE_DIRECTORY if is_dir else FILE_ATTRIBUTE_NORMAL
        flags = SHGFI_ICON | SHGFI_SMALLICON
        if not path_specific:
            flags |= SHGFI_USEFILEATTRIBUTES

        info = SHFILEINFOW()
        if not sh_get(p, attrs, ctypes.byref(info), ctypes.sizeof(info), flags) or not info.hIcon:
            return None, 0, 0

        hdc = gdi32.CreateCompatibleDC(None)
        if not hdc:
            user32.DestroyIcon(info.hIcon)
            return None, 0, 0

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = int(size)
        bmi.bmiHeader.biHeight = -int(size)  # top-down DIB
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0  # BI_RGB
        bits = ctypes.c_void_p()
        hbm = gdi32.CreateDIBSection(hdc, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
        if not hbm or not bits:
            gdi32.DeleteDC(hdc)
            user32.DestroyIcon(info.hIcon)
            return None, 0, 0

        old = gdi32.SelectObject(hdc, hbm)
        try:
            ctypes.memset(bits, 0, int(size) * int(size) * 4)
            DI_NORMAL = 0x0003
            if not user32.DrawIconEx(hdc, 0, 0, info.hIcon, int(size), int(size), 0, None, DI_NORMAL):
                return None, 0, 0
            raw = bytearray(ctypes.string_at(bits, int(size) * int(size) * 4))
            # Some legacy icon handlers return RGB with a zero alpha channel.
            if raw and not any(raw[3::4]):
                for i in range(0, len(raw), 4):
                    if raw[i] or raw[i + 1] or raw[i + 2]:
                        raw[i + 3] = 255
            return bytes(raw), int(size), int(size)
        finally:
            if old:
                gdi32.SelectObject(hdc, old)
            gdi32.DeleteObject(hbm)
            gdi32.DeleteDC(hdc)
            user32.DestroyIcon(info.hIcon)
    except Exception:
        return None, 0, 0


class ShellIconWorker(QtCore.QThread):
    iconReady = pyqtSignal(str, bytes, int, int)
    finishedCycle = pyqtSignal(object)

    def __init__(self, jobs, parent=None):
        super().__init__(parent)
        self._jobs = list(jobs or [])
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            for key, path, is_dir in self._jobs:
                if self._cancel:
                    break
                raw, w, h = _load_windows_shell_icon_bgra(path, bool(is_dir))
                if raw:
                    self.iconReady.emit(str(key), raw, int(w), int(h))
        finally:
            self.finishedCycle.emit(self._jobs)


class FastDirModel(QAbstractTableModel):
    HEADERS = ["Name", "Size", "Ext", "Date Modified"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._root = ""
        self._rows = []
        self._icon_cache = {}
        self._icon_rows = {}
        self._icon_file = None
        self._icon_dir = None

    def rootPath(self):
        return self._root

    def reset_dir(self, path: str):
        self.beginResetModel()
        self._root = path
        self._rows = []
        self._icon_cache.clear()
        self._icon_rows.clear()
        self.endResetModel()

    @QtCore.pyqtSlot(list)
    def append_rows(self, rows: list):
        if not rows:
            return
        prepared = []
        for rec in rows:
            rec = dict(rec)
            key = rec.get("icon_key") or _icon_cache_key(rec.get("path", ""), bool(rec.get("is_dir")))
            rec["icon_key"] = key
            prepared.append(rec)
        start = len(self._rows)
        self.beginInsertRows(QtCore.QModelIndex(), start, start + len(prepared) - 1)
        self._rows.extend(prepared)
        for offset, rec in enumerate(prepared):
            self._icon_rows.setdefault(rec["icon_key"], []).append(start + offset)
        self.endInsertRows()

    def row_path(self, row: int) -> str:
        return self._rows[row]["path"] if 0 <= row < len(self._rows) else ""

    def row_is_dir(self, row: int) -> bool:
        return bool(self._rows[row].get("is_dir")) if 0 <= row < len(self._rows) else False

    def icon_key(self, row: int) -> str:
        return str(self._rows[row].get("icon_key", "")) if 0 <= row < len(self._rows) else ""

    def has_stat(self, row: int) -> bool:
        if 0 <= row < len(self._rows):
            rec = self._rows[row]
            return rec.get("mtime") is not None and (rec.get("is_dir", False) or rec.get("size") is not None)
        return False

    def has_icon(self, row: int) -> bool:
        key = self.icon_key(row)
        return bool(key and key in self._icon_cache)

    @QtCore.pyqtSlot(int, object, object)
    def apply_stat(self, row: int, size_val, mtime_val):
        if not (0 <= row < len(self._rows)):
            return
        changed = []
        if self._rows[row].get("size") is None and size_val is not None:
            self._rows[row]["size"] = int(size_val)
            changed.append(1)
        if self._rows[row].get("mtime") is None and mtime_val is not None:
            self._rows[row]["mtime"] = float(mtime_val)
            changed.append(3)
        for col in changed:
            ix = self.index(row, col)
            self.dataChanged.emit(ix, ix, [Qt.DisplayRole, Qt.EditRole, SIZE_BYTES_ROLE])

    @QtCore.pyqtSlot(str, object)
    def apply_icon_key(self, key: str, icon):
        if not key or not isinstance(icon, QIcon) or icon.isNull():
            return
        self._icon_cache[key] = icon
        rows = self._icon_rows.get(key, [])
        if not rows:
            return
        # Rows sharing one extension normally form several ranges after sorting;
        # emitting one broad range is much cheaper than one signal per row.
        first, last = min(rows), max(rows)
        self.dataChanged.emit(self.index(first, 0), self.index(last, 0), [Qt.DecorationRole])

    def rowCount(self, parent=QtCore.QModelIndex()):
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return 4

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        if orientation == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return int(Qt.AlignRight | Qt.AlignVCenter) if section in (1, 2, 3) else int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def flags(self, index):
        base = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.isValid():
            base |= Qt.ItemIsDragEnabled
        return base

    def mimeTypes(self):
        return ["text/uri-list"]

    def mimeData(self, indexes):
        md = QtCore.QMimeData()
        rows = sorted({ix.row() for ix in indexes if ix.isValid()})
        paths = [self.row_path(row) for row in rows if self.row_path(row)]
        if paths:
            md.setUrls([QUrl.fromLocalFile(p) for p in paths])
            md.setText("\r\n".join(paths))
        return md

    def supportedDragActions(self):
        return Qt.CopyAction | Qt.MoveAction

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        rec = self._rows[index.row()]
        col = index.column()

        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignRight | Qt.AlignVCenter) if col in (1, 2, 3) else int(Qt.AlignLeft | Qt.AlignVCenter)

        if role == Qt.DecorationRole and col == 0:
            icon = self._icon_cache.get(rec.get("icon_key"))
            if icon is not None:
                return icon
            try:
                if self._icon_file is None or self._icon_dir is None:
                    st = QApplication.instance().style()
                    self._icon_file = st.standardIcon(QStyle.SP_FileIcon) if st else QIcon()
                    self._icon_dir = st.standardIcon(QStyle.SP_DirIcon) if st else QIcon()
            except Exception:
                return None
            return self._icon_dir if rec.get("is_dir") else self._icon_file

        if role == Qt.DisplayRole:
            if col == 0:
                return rec.get("name", "")
            if col == 1:
                if rec.get("is_dir") or rec.get("size") is None:
                    return ""
                return human_size(int(rec["size"]))
            if col == 2:
                return rec.get("ext", "")
            if col == 3:
                if rec.get("mtime") is None:
                    return ""
                return QDateTime.fromSecsSinceEpoch(int(rec["mtime"])).toString(LIST_DATETIME_FMT)

        if role == Qt.EditRole:
            if col == 0:
                return rec.get("name", "")
            if col == 1:
                return 0 if rec.get("is_dir") or rec.get("size") is None else int(rec["size"])
            if col == 2:
                return rec.get("ext", "")
            if col == 3:
                return QDateTime.fromSecsSinceEpoch(int(rec["mtime"])) if rec.get("mtime") is not None else QDateTime()

        if role == Qt.ToolTipRole:
            return rec.get("path", "")
        if role == Qt.UserRole:
            return rec.get("path", "")
        if role == IS_DIR_ROLE:
            return bool(rec.get("is_dir"))
        if role == SIZE_BYTES_ROLE:
            return 0 if rec.get("is_dir") or rec.get("size") is None else int(rec["size"])
        if role == NAME_FOLD_ROLE and col == 0:
            return rec.get("name_l") or str(rec.get("name", "")).lower()
        if role == ICON_KEY_ROLE:
            return rec.get("icon_key", "")
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
                    size_val=0 if stat.S_ISDIR(st.st_mode) else int(st.st_size)
                    mtime_val=float(st.st_mtime)
                except Exception:
                    size_val=0; mtime_val=None
                self.statReady.emit(row,size_val,mtime_val)
        finally:
            self.finishedCycle.emit()

class DirEnumWorker(QtCore.QThread):
    batchReady=QtCore.pyqtSignal(list); finished=QtCore.pyqtSignal(); error=QtCore.pyqtSignal(str)
    def __init__(self, root:str, parent=None, preload_size: bool = False, preload_mtime: bool = False):
        super().__init__(parent)
        self.root=root
        self._cancel=False
        self._preload_size = bool(preload_size)
        self._preload_mtime = bool(preload_mtime)
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
                    ext = file_extension_label(name, is_dir)
                    size_val = None
                    mtime_val = None
                    if self._preload_size or self._preload_mtime:
                        try:
                            st = entry.stat(follow_symlinks=False)
                            if self._preload_size:
                                size_val = 0 if is_dir else int(st.st_size)
                            if self._preload_mtime:
                                mtime_val = float(st.st_mtime)
                        except Exception:
                            if self._preload_size:
                                size_val = 0 if is_dir else None
                            if self._preload_mtime:
                                mtime_val = None
                    batch.append({
                        "name": name,
                        "name_l": name.lower(),
                        "path": p,
                        "is_dir": is_dir,
                        "ext": ext,
                        "size": size_val,
                        "mtime": mtime_val,
                        "icon_key": _icon_cache_key(p, is_dir),
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
                    size_val=0 if stat.S_ISDIR(st.st_mode) else int(st.st_size)
                    mtime_val=float(st.st_mtime)
                except Exception:
                    size_val=0; mtime_val=None
                self.statReady.emit(p,size_val,mtime_val)
        finally:
            self.finishedCycle.emit()

class SearchWorker(QtCore.QThread):
    batchReady = pyqtSignal(str, list)
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    truncated = pyqtSignal(int)

    def __init__(self, base_path: str, pattern_str: str, parent=None, max_results: int = SEARCH_RESULT_LIMIT):
        super().__init__(parent)
        self.base = base_path
        self._cancel = False
        self._max_results = max(1, int(max_results))
        self._matches = 0
        self._visited_dirs = 0
        self._visited_entries = 0
        self._last_progress_emit = 0
        self._truncated = False

        raw = (pattern_str or "").replace(",", " ").replace(";", " ").split()
        self._patterns = [p.lower() for p in raw] if raw else ["*"]

        self._tests = []
        for p in self._patterns:
            simple_ext = (p.startswith("*.") and ("*" not in p[2:]) and ("?" not in p) and ("[" not in p) and ("]" not in p))
            has_wildcard = any(ch in p for ch in "*?[")
            if simple_ext:
                ext = p[1:]
                self._tests.append(lambda name, ext=ext: name.endswith(ext))
            elif not has_wildcard:
                # Plain text works as a case-insensitive substring filter.
                # Example: "abc" matches "abc.txt", "my_abc_file.cpp", etc.
                self._tests.append(lambda name, needle=p: needle in name)
            else:
                self._tests.append(lambda name, pat=p: fnmatch.fnmatchcase(name, pat))

    def cancel(self): self._cancel = True

    def _emit_progress(self, folder: str = ""):
        now = time.monotonic()
        interval_s = max(0.05, SEARCH_PROGRESS_INTERVAL / 1000.0)
        if self._last_progress_emit and (now - self._last_progress_emit) < interval_s:
            return
        self._last_progress_emit = now
        self.progress.emit(self._visited_dirs, self._visited_entries, folder)

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
                self._visited_dirs += 1
                self._emit_progress(d)
                try:
                    with os.scandir(d) as it:
                        for entry in it:
                            if self._cancel:
                                break
                            self._visited_entries += 1
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
                                if entry.name.lower() in DEFAULT_SEARCH_EXCLUDE_DIRS:
                                    continue
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
            self.progress.emit(self._visited_dirs, self._visited_entries, "")
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
        self._refresh_after_pending = set()

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
        self._refresh_after_pending.clear()
        self._cancel_worker()

    def _cancel_worker(self):
        w = self._worker
        if w and w.isRunning():
            w.cancel()
            w.wait(100)
        self._worker = None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and section == 2:
            return "Ext"
        return super().headerData(section, orientation, role)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        col = index.column()
        if col not in (1, 2, 3):
            return super().data(index, role)

        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignRight | Qt.AlignVCenter)

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

        # File metadata is populated only by NormalStatWorker; never block data() with fileInfo().
        info = None

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

            if info is not None:
                try:
                    size_val = max(0, int(info.size()))
                    if role == SIZE_BYTES_ROLE:
                        return size_val
                    if role == Qt.EditRole:
                        return size_val
                    if role == Qt.DisplayRole:
                        return human_size(size_val)
                except Exception:
                    pass

            if role in (SIZE_BYTES_ROLE, Qt.EditRole):
                return 0
            if role == Qt.DisplayRole:
                return ""
            return super().data(index, role)

        if col == 2:
            ext = file_extension_label(p, is_dir)
            if role in (Qt.DisplayRole, Qt.EditRole):
                return ext
            return super().data(index, role)

        if col == 3:
            if rec and rec[1] is not None:
                dt = QDateTime.fromSecsSinceEpoch(int(rec[1]))
                if role == Qt.DisplayRole:
                    return dt.toString(LIST_DATETIME_FMT)
                if role == Qt.EditRole:
                    return dt

            if info is not None:
                try:
                    dt = info.lastModified()
                    if dt and dt.isValid():
                        if role == Qt.DisplayRole:
                            return dt.toString(LIST_DATETIME_FMT)
                        if role == Qt.EditRole:
                            return dt
                except Exception:
                    pass

            if role == Qt.DisplayRole:
                return ""
            if role == Qt.EditRole:
                return QDateTime()
            return super().data(index, role)

    def request_paths(self, paths: list[str], batch_limit: int = 256, force: bool = False):
        try:
            self._batch_limit = max(1, int(batch_limit))
        except Exception:
            self._batch_limit = 256

        added = False
        for p in paths:
            if not p:
                continue
            if force:
                self._cache.pop(p, None)
            elif p in self._cache:
                continue

            if p in self._pending:
                if force:
                    self._refresh_after_pending.add(p)
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
        retry = []
        for p in batch:
            self._pending.discard(p)
            if p in self._refresh_after_pending:
                self._refresh_after_pending.discard(p)
                retry.append(p)
        self._worker = None
        for p in retry:
            if p in self._pending:
                continue
            self._cache.pop(p, None)
            self._pending.add(p)
            self._queue.append(p)
        self._start_next_batch()

class PathBar(QWidget):
    pathSubmitted=pyqtSignal(str)
    _shared_recent_paths: list[str] | None = None

    def __init__(self, parent=None):
        super().__init__(parent); self._current_path=QDir.homePath()
        self.setObjectName("pathbar")

        if PathBar._shared_recent_paths is None:
            PathBar._shared_recent_paths = load_recent_path_history()
        self._recent_paths = list(PathBar._shared_recent_paths or [])

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

        self._suggest_timer = QTimer(self)
        self._suggest_timer.setSingleShot(True)
        self._suggest_timer.setInterval(110)
        self._suggest_timer.timeout.connect(self._refresh_edit_completer)
        self._edit_model = QStringListModel(self)
        self._edit_completer = QCompleter(self._edit_model, self)
        self._edit_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._edit_completer.setCompletionMode(QCompleter.PopupCompletion)
        self._edit_completer.setModelSorting(QCompleter.CaseInsensitivelySortedModel)
        self._edit.setCompleter(self._edit_completer)
        self._edit.textEdited.connect(lambda _t: self._queue_suggestions_update(False))

        self._btn_hist = QToolButton(self)
        self._btn_hist.setToolTip("Recent paths")
        self._btn_hist.setFixedHeight(UI_H)
        self._btn_hist.setArrowType(Qt.DownArrow)
        self._btn_hist.clicked.connect(self._show_recent_paths_menu)

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
        wrap.addWidget(self._btn_hist, 0)
        wrap.addWidget(self._btn_copy, 0)

        self._host.installEventFilter(self); self._edit.installEventFilter(self)
        self.set_path(self._current_path)

    def _copy_current_path(self):
        t = self._edit.text().strip() if self._edit.isVisible() else self._current_path
        if not t:
            t = self._current_path
        QApplication.clipboard().setText(t)
        QToolTip.showText(QCursor.pos(), f"Copied: {t}", self)

    def _set_recent_paths(self, items: list[str]):
        seen = set()
        out = []
        for p in list(items or []):
            sp = str(p).strip()
            if not sp:
                continue
            np = _normalize_fs_path(sp)
            key = os.path.normcase(np)
            if key in seen:
                continue
            seen.add(key)
            out.append(np)
            if len(out) >= PATH_HISTORY_LIMIT:
                break
        self._recent_paths = out
        PathBar._shared_recent_paths = list(out)
        save_recent_path_history(out)

    def _reload_recent_paths(self):
        if PathBar._shared_recent_paths is None:
            PathBar._shared_recent_paths = load_recent_path_history()
        self._recent_paths = list(PathBar._shared_recent_paths or [])

    def remember_path(self, path: str):
        try:
            np = _normalize_fs_path(nice_path(path))
        except Exception:
            np = _normalize_fs_path(str(path))
        if not np:
            return
        key = os.path.normcase(np)
        merged = [np]
        for p in self._recent_paths:
            if os.path.normcase(_normalize_fs_path(p)) != key:
                merged.append(_normalize_fs_path(p))
        self._set_recent_paths(merged)

    def _list_root_paths(self) -> list[str]:
        out = []
        try:
            for fi in QDir.drives():
                try:
                    p = fi.absoluteFilePath()
                except Exception:
                    p = ""
                if p:
                    out.append(_normalize_fs_path(p))
        except Exception:
            pass
        if not out:
            out.append(_normalize_fs_path(QDir.rootPath()))
        seen = set()
        uniq = []
        for p in out:
            k = os.path.normcase(p)
            if k in seen:
                continue
            seen.add(k)
            uniq.append(p)
        return uniq

    def _collect_recent_suggestions(self, typed: str, max_items: int = 30) -> list[str]:
        self._reload_recent_paths()
        t = (typed or "").strip().lower()
        if not t:
            return self._recent_paths[:max_items]
        starts = [p for p in self._recent_paths if p.lower().startswith(t)]
        contains = [p for p in self._recent_paths if t in p.lower() and not p.lower().startswith(t)]
        return (starts + contains)[:max_items]

    def _collect_filesystem_suggestions(self, typed: str, max_items: int = 45) -> list[str]:
        t = (typed or "").strip().strip('"')
        if not t:
            return self._list_root_paths()[:max_items]

        t = t.replace("/", os.sep)
        if os.name == "nt" and len(t) == 2 and t[1] == ":":
            t = t + os.sep

        if t.endswith(("\\", "/")):
            parent = _normalize_fs_path(t)
            prefix = ""
        else:
            parent = _normalize_fs_path(os.path.dirname(t))
            prefix = os.path.basename(t)

        if not parent:
            roots = self._list_root_paths()
            tl = t.lower()
            if not tl:
                return roots[:max_items]
            return [r for r in roots if r.lower().startswith(tl)][:max_items]

        if not os.path.isdir(parent):
            return []

        out = []
        pref_l = prefix.lower()
        try:
            with os.scandir(parent) as it:
                for entry in it:
                    try:
                        if not entry.is_dir(follow_symlinks=False):
                            continue
                    except Exception:
                        continue
                    name = entry.name
                    if pref_l and not name.lower().startswith(pref_l):
                        continue
                    out.append(_normalize_fs_path(os.path.join(parent, name)))
                    if len(out) >= max_items:
                        break
        except Exception:
            return []
        if not pref_l and os.path.isdir(parent):
            out.insert(0, _normalize_fs_path(parent))
        return out[:max_items]

    def _collect_edit_suggestions(self, typed: str) -> list[str]:
        out = []
        seen = set()
        def add_path(p):
            if not p:
                return
            np = _normalize_fs_path(str(p))
            key = os.path.normcase(np)
            if key in seen:
                return
            seen.add(key)
            out.append(np)

        raw = (typed or "").strip().strip('"')
        if raw and os.path.isdir(_normalize_fs_path(raw)):
            add_path(_normalize_fs_path(raw))

        for p in self._collect_recent_suggestions(raw, 35):
            add_path(p)
        for p in self._collect_filesystem_suggestions(raw, 50):
            add_path(p)
        return out[:80]

    def _queue_suggestions_update(self, immediate: bool = False):
        if immediate:
            self._refresh_edit_completer()
            return
        if self._suggest_timer.isActive():
            self._suggest_timer.stop()
        self._suggest_timer.start()

    def _refresh_edit_completer(self):
        typed = self._edit.text().strip()
        items = self._collect_edit_suggestions(typed)
        self._edit_model.setStringList(items)
        self._edit_completer.setCompletionPrefix(typed)
        if self._edit.isVisible() and self._edit.hasFocus() and items:
            self._edit_completer.complete()

    def _show_recent_paths_menu(self):
        self._reload_recent_paths()
        menu = QMenu(self)
        action_map = {}
        for p in self._recent_paths[:PATH_HISTORY_LIMIT]:
            act = menu.addAction(p)
            action_map[act] = p
        if action_map:
            menu.addSeparator()
            act_clear = menu.addAction("Clear History")
        else:
            act_none = menu.addAction("(No recent paths)")
            act_none.setEnabled(False)
            act_clear = None
        picked = menu.exec_(self._btn_hist.mapToGlobal(QtCore.QPoint(0, self._btn_hist.height())))
        if not picked:
            return
        if act_clear is not None and picked == act_clear:
            self._set_recent_paths([])
            self._queue_suggestions_update(True)
            return
        selected_path = action_map.get(picked)
        if not selected_path:
            return
        if self._edit.isVisible():
            self._edit.setText(selected_path)
            self._edit.setFocus()
            self._edit.selectAll()
            self._queue_suggestions_update(True)
        else:
            self.pathSubmitted.emit(selected_path)

    def set_active(self, active: bool):
        try:
            active = bool(active)
            if bool(self._scroll.property("active")) == active:
                return
            self._scroll.setProperty("active", active)
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
        if obj is self._edit and ev.type()==QEvent.FocusOut:
            try:
                comp = self._edit.completer()
                pop = comp.popup() if comp else None
                if pop and pop.isVisible():
                    return False
            except Exception:
                pass
            self.cancel_edit()
        return super().eventFilter(obj, ev)

    def start_edit(self):
        self._reload_recent_paths()
        self._edit.setText(self._current_path); self._scroll.hide(); self._edit.show()
        self._edit.setFocus(); self._edit.selectAll()
        self._queue_suggestions_update(True)

    def cancel_edit(self): self._edit.hide(); self._scroll.show()

    def _on_edit_return(self):
        t=self._edit.text().strip(); self.cancel_edit()
        if t: self.pathSubmitted.emit(t)

    def set_path(self, path:str):
        self._current_path = nice_path(path)
        self.remember_path(self._current_path)
        self._rebuild()

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


class SearchResultModel(QAbstractTableModel):
    HEADERS = ["Name", "Size", "Ext", "Date Modified", "Folder"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows = []
        self._row_by_path = {}
        self._icon_cache = {}
        self._icon_rows = {}
        self._icon_file = None
        self._icon_dir = None

    @QtCore.pyqtSlot(list)
    def append_rows(self, rows: list):
        if not rows:
            return
        prepared = []
        for incoming in rows:
            rec = dict(incoming)
            path = str(rec.get("path", ""))
            is_dir = bool(rec.get("is_dir"))
            rec.update({
                "name_l": str(rec.get("name", "")).lower(),
                "ext": file_extension_label(rec.get("name", ""), is_dir),
                "size": rec.get("size"),
                "mtime": rec.get("mtime"),
                "icon_key": rec.get("icon_key") or _icon_cache_key(path, is_dir),
            })
            prepared.append(rec)
        first = len(self._rows)
        self.beginInsertRows(QtCore.QModelIndex(), first, first + len(prepared) - 1)
        self._rows.extend(prepared)
        for offset, rec in enumerate(prepared):
            row = first + offset
            self._row_by_path[rec.get("path", "")] = row
            self._icon_rows.setdefault(rec["icon_key"], []).append(row)
        self.endInsertRows()

    def rowCount(self, parent=QtCore.QModelIndex()):
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return 5

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        if orientation == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return int(Qt.AlignRight | Qt.AlignVCenter) if section in (1, 2, 3, 4) else int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def row_path(self, row: int) -> str:
        return self._rows[row].get("path", "") if 0 <= row < len(self._rows) else ""

    def row_is_dir(self, row: int) -> bool:
        return bool(self._rows[row].get("is_dir")) if 0 <= row < len(self._rows) else False

    def icon_key(self, row: int) -> str:
        return str(self._rows[row].get("icon_key", "")) if 0 <= row < len(self._rows) else ""

    def has_icon(self, row: int) -> bool:
        key = self.icon_key(row)
        return bool(key and key in self._icon_cache)

    def has_stat(self, row: int) -> bool:
        if not (0 <= row < len(self._rows)):
            return False
        rec = self._rows[row]
        return rec.get("mtime") is not None and (rec.get("is_dir") or rec.get("size") is not None)

    @QtCore.pyqtSlot(str, object, object)
    def apply_stat(self, path: str, size_val, mtime_val):
        row = self._row_by_path.get(path)
        if row is None or not (0 <= row < len(self._rows)):
            return
        rec = self._rows[row]
        changed = []
        if rec.get("size") is None and size_val is not None:
            rec["size"] = int(size_val)
            changed.append(1)
        if rec.get("mtime") is None and mtime_val is not None:
            rec["mtime"] = float(mtime_val)
            changed.append(3)
        for col in changed:
            ix = self.index(row, col)
            self.dataChanged.emit(ix, ix, [Qt.DisplayRole, Qt.EditRole, SIZE_BYTES_ROLE])

    @QtCore.pyqtSlot(str, object)
    def apply_icon_key(self, key: str, icon):
        if not key or not isinstance(icon, QIcon) or icon.isNull():
            return
        self._icon_cache[key] = icon
        rows = self._icon_rows.get(key, [])
        if rows:
            self.dataChanged.emit(self.index(min(rows), 0), self.index(max(rows), 0), [Qt.DecorationRole])

    def flags(self, index):
        base = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.isValid():
            base |= Qt.ItemIsDragEnabled
        return base

    def mimeTypes(self):
        return ["text/uri-list"]

    def mimeData(self, indexes):
        md = QtCore.QMimeData()
        rows = sorted({ix.row() for ix in indexes if ix.isValid()})
        paths = [self.row_path(r) for r in rows if self.row_path(r)]
        if paths:
            md.setUrls([QUrl.fromLocalFile(p) for p in paths])
            md.setText("\r\n".join(paths))
        return md

    def supportedDragActions(self):
        return Qt.CopyAction | Qt.MoveAction

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        rec = self._rows[index.row()]
        col = index.column()

        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignRight | Qt.AlignVCenter) if col in (1, 2, 3, 4) else int(Qt.AlignLeft | Qt.AlignVCenter)

        if role == Qt.DecorationRole and col == 0:
            icon = self._icon_cache.get(rec.get("icon_key"))
            if icon is not None:
                return icon
            try:
                if self._icon_file is None or self._icon_dir is None:
                    st = QApplication.instance().style()
                    self._icon_file = st.standardIcon(QStyle.SP_FileIcon) if st else QIcon()
                    self._icon_dir = st.standardIcon(QStyle.SP_DirIcon) if st else QIcon()
            except Exception:
                return None
            return self._icon_dir if rec.get("is_dir") else self._icon_file

        if role == Qt.DisplayRole:
            if col == 0:
                return rec.get("name", "")
            if col == 1:
                if rec.get("is_dir") or rec.get("size") is None:
                    return ""
                return human_size(int(rec["size"]))
            if col == 2:
                return rec.get("ext", "")
            if col == 3:
                if rec.get("mtime") is None:
                    return ""
                return QDateTime.fromSecsSinceEpoch(int(rec["mtime"])).toString(LIST_DATETIME_FMT)
            if col == 4:
                return rec.get("folder", "")

        if role == Qt.EditRole:
            if col == 0:
                return rec.get("name", "")
            if col == 1:
                return 0 if rec.get("is_dir") or rec.get("size") is None else int(rec["size"])
            if col == 2:
                return rec.get("ext", "")
            if col == 3:
                return QDateTime.fromSecsSinceEpoch(int(rec["mtime"])) if rec.get("mtime") is not None else QDateTime()
            if col == 4:
                return rec.get("folder", "")

        if role == Qt.ToolTipRole:
            return rec.get("path", "")
        if role == Qt.UserRole:
            return rec.get("path", "")
        if role == IS_DIR_ROLE:
            return bool(rec.get("is_dir"))
        if role == SIZE_BYTES_ROLE:
            return 0 if rec.get("is_dir") or rec.get("size") is None else int(rec["size"])
        if role == NAME_FOLD_ROLE and col == 0:
            return rec.get("name_l", "")
        if role == ICON_KEY_ROLE:
            return rec.get("icon_key", "")
        return None


class SearchFolderDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignRight | Qt.AlignVCenter
        option.textElideMode = Qt.ElideLeft


class BulkRenameDialog(QDialog):
    _INVALID_WIN_CHARS = set('<>:"/\\|?*')

    def __init__(self, parent, paths: list[str]):
        super().__init__(parent)
        self.setWindowTitle("Bulk Rename")
        self.resize(980, 560)
        self._paths = [p for p in list(paths or []) if p and os.path.exists(p)]
        self._plan = []

        lay = QVBoxLayout(self)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        self.ed_prefix = QLineEdit(self); self.ed_prefix.setPlaceholderText("Prefix")
        self.ed_suffix = QLineEdit(self); self.ed_suffix.setPlaceholderText("Suffix")
        self.ed_find = QLineEdit(self); self.ed_find.setPlaceholderText("Find text")
        self.ed_replace = QLineEdit(self); self.ed_replace.setPlaceholderText("Replace with")

        self.chk_case = QCheckBox("Case sensitive replace", self)
        self.chk_number = QCheckBox("Append number", self)
        self.spin_start = QSpinBox(self); self.spin_start.setRange(1, 999999); self.spin_start.setValue(1)
        self.spin_step = QSpinBox(self); self.spin_step.setRange(1, 9999); self.spin_step.setValue(1)
        self.spin_pad = QSpinBox(self); self.spin_pad.setRange(1, 8); self.spin_pad.setValue(3)
        self.ed_sep = QLineEdit(self); self.ed_sep.setText("_"); self.ed_sep.setMaxLength(8)

        grid.addWidget(QLabel("Prefix"), 0, 0)
        grid.addWidget(self.ed_prefix, 0, 1)
        grid.addWidget(QLabel("Suffix"), 0, 2)
        grid.addWidget(self.ed_suffix, 0, 3)
        grid.addWidget(QLabel("Find"), 1, 0)
        grid.addWidget(self.ed_find, 1, 1)
        grid.addWidget(QLabel("Replace"), 1, 2)
        grid.addWidget(self.ed_replace, 1, 3)
        grid.addWidget(self.chk_case, 2, 0, 1, 2)
        grid.addWidget(self.chk_number, 2, 2, 1, 2)
        grid.addWidget(QLabel("Start"), 3, 0)
        grid.addWidget(self.spin_start, 3, 1)
        grid.addWidget(QLabel("Step"), 3, 2)
        grid.addWidget(self.spin_step, 3, 3)
        grid.addWidget(QLabel("Padding"), 4, 0)
        grid.addWidget(self.spin_pad, 4, 1)
        grid.addWidget(QLabel("Separator"), 4, 2)
        grid.addWidget(self.ed_sep, 4, 3)
        lay.addLayout(grid)

        self.lbl_summary = QLabel("", self)
        lay.addWidget(self.lbl_summary)
        self.tbl = _setup_readonly_table(
            QTableWidget(self),
            ["Current Name", "New Name", "Folder", "Status"],
            [QHeaderView.ResizeToContents, QHeaderView.ResizeToContents, QHeaderView.Stretch, QHeaderView.ResizeToContents],
        )
        lay.addWidget(self.tbl, 1)
        btns = _add_dialog_button_box(lay, self, QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self.accept, self.reject)
        self.btn_ok = btns.button(QDialogButtonBox.Ok)
        for w in (self.ed_prefix, self.ed_suffix, self.ed_find, self.ed_replace, self.ed_sep):
            w.textChanged.connect(self._rebuild_preview)
        for w in (self.chk_case, self.chk_number):
            w.toggled.connect(self._rebuild_preview)
        for w in (self.spin_start, self.spin_step, self.spin_pad):
            w.valueChanged.connect(self._rebuild_preview)

        self._rebuild_preview()

    def _transform_stem(self, stem: str, idx: int) -> str:
        text = stem
        find = self.ed_find.text()
        repl = self.ed_replace.text()
        if find:
            if self.chk_case.isChecked():
                text = text.replace(find, repl)
            else:
                try:
                    text = re.sub(re.escape(find), repl, text, flags=re.IGNORECASE)
                except Exception:
                    text = text.replace(find, repl)
        text = f"{self.ed_prefix.text()}{text}{self.ed_suffix.text()}"
        if self.chk_number.isChecked():
            seq = self.spin_start.value() + (idx * self.spin_step.value())
            num = str(seq).zfill(self.spin_pad.value())
            text = f"{text}{self.ed_sep.text()}{num}"
        return text

    def _is_invalid_name(self, name: str) -> str | None:
        if not name:
            return "empty name"
        if os.name == "nt":
            if any(ch in self._INVALID_WIN_CHARS for ch in name):
                return "invalid character"
            if name.endswith(" ") or name.endswith("."):
                return "trailing space/dot"
        if os.sep in name:
            return "contains path separator"
        if os.altsep and os.altsep in name:
            return "contains path separator"
        return None

    def _build_plan(self) -> list[dict]:
        rows = []
        selected_keys = set()
        for p in self._paths:
            try:
                key = os.path.normcase(os.path.normpath(p))
                selected_keys.add(key)
            except Exception:
                pass
            rows.append({
                "src": p,
                "folder": os.path.dirname(p),
                "name": os.path.basename(p),
                "is_dir": os.path.isdir(p),
                "new_name": "",
                "dst": "",
                "status": "",
                "error": False,
            })

        by_folder = {}
        for r in rows:
            by_folder.setdefault(r["folder"], []).append(r)

        for folder, items in by_folder.items():
            items.sort(key=lambda x: x["name"].lower())
            for idx, r in enumerate(items):
                old_name = r["name"]
                if r["is_dir"]:
                    stem = old_name
                    ext = ""
                else:
                    stem, ext = os.path.splitext(old_name)
                new_stem = self._transform_stem(stem, idx)
                new_name = f"{new_stem}{ext}"
                r["new_name"] = new_name
                r["dst"] = os.path.join(folder, new_name)

                msg = self._is_invalid_name(new_name)
                if msg:
                    r["status"] = msg
                    r["error"] = True
                    continue

                if os.path.normcase(os.path.normpath(r["src"])) == os.path.normcase(os.path.normpath(r["dst"])):
                    r["status"] = "unchanged"
                else:
                    r["status"] = "ready"

            by_dst = {}
            for r in items:
                key = os.path.normcase(os.path.normpath(r["dst"]))
                by_dst.setdefault(key, []).append(r)
            for dup in by_dst.values():
                if len(dup) <= 1:
                    continue
                for r in dup:
                    r["status"] = "duplicate target name"
                    r["error"] = True

            for r in items:
                if r["error"]:
                    continue
                dst = r["dst"]
                dst_key = os.path.normcase(os.path.normpath(dst))
                src_key = os.path.normcase(os.path.normpath(r["src"]))
                if src_key == dst_key:
                    continue
                if os.path.exists(dst) and dst_key not in selected_keys:
                    r["status"] = "target already exists"
                    r["error"] = True

        return rows

    def _rebuild_preview(self):
        self._plan = self._build_plan()
        self.tbl.setRowCount(len(self._plan))
        changed = 0
        errors = 0
        for i, r in enumerate(self._plan):
            folder = r.get("folder", "")
            old_name = r.get("name", "")
            new_name = r.get("new_name", "")
            status = r.get("status", "")
            if status == "ready":
                changed += 1
            if r.get("error"):
                errors += 1
            _set_table_row_items(self.tbl, i, old_name, new_name, folder, status)
        self.tbl.resizeColumnsToContents()
        self.btn_ok.setEnabled(changed > 0 and errors == 0)
        self.lbl_summary.setText(
            f"Selected {len(self._plan)} item(s) / will rename {changed} / errors {errors}"
        )

    def result_operations(self) -> list[tuple[str, str]]:
        return [
            (r.get("src", ""), r.get("dst", ""))
            for r in self._plan
            if not r.get("error") and r.get("status") == "ready" and r.get("src") and r.get("dst")
        ]


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


        self.tbl = _setup_readonly_table(
            QTableWidget(self),
            ["Name", "Destination", "Action"],
            [QHeaderView.ResizeToContents, QHeaderView.Stretch, QHeaderView.ResizeToContents],
            row_count=len(conflicts),
        )
        self._combos = []
        for r, (src, dst) in enumerate(conflicts):
            name = os.path.basename(src)
            _set_table_row_items(self.tbl, r, name, dst)
            combo = QComboBox(self.tbl)
            combo.addItems(["Overwrite", "Skip", "Copy"])
            self.tbl.setCellWidget(r, 2, combo)
            self._combos.append(combo)
        lay.addWidget(self.tbl, 1)
        _add_dialog_button_box(lay, self, QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self.accept, self.reject)
        btn_over.clicked.connect(lambda: self._apply_all("Overwrite"))
        btn_skip.clicked.connect(lambda: self._apply_all("Skip"))
        btn_copy.clicked.connect(lambda: self._apply_all("Copy"))
        theme = getattr(getattr(parent, "host", None), "theme", "dark")
        if theme == "dark":
            _apply_palette_colors(self, {
                QPalette.Window: (255, 255, 255), QPalette.Base: (255, 255, 255), QPalette.AlternateBase: (245, 245, 245),
                QPalette.Text: (0, 0, 0), QPalette.ButtonText: (0, 0, 0), QPalette.WindowText: (0, 0, 0),
            })
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
        return {
            src: (choice if choice in ("overwrite", "skip", "copy") else "overwrite")
            for (src, _dst), combo in zip(self._conflicts, self._combos)
            for choice in [combo.currentText().strip().lower()]
        }


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

    def ensure_drag_ready(self):
        self._clear_drag_state()
        try:
            self.setDragEnabled(True)
            self.setAcceptDrops(True)
            self.setDropIndicatorShown(True)
            self.setDefaultDropAction(Qt.MoveAction)
            if self.dragDropMode() != QAbstractItemView.DragDrop:
                self.setDragDropMode(QAbstractItemView.DragDrop)
            if self.selectionBehavior() != QAbstractItemView.SelectRows:
                self.setSelectionBehavior(QAbstractItemView.SelectRows)
        except Exception:
            pass

    def _row_selected_for_drag(self, sm, index):
        if not (sm and index.isValid()):
            return False
        try:
            if sm.isRowSelected(index.row(), index.parent()):
                return True
        except Exception:
            pass
        try:
            return bool(sm.isSelected(index))
        except Exception:
            return False

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            try:
                self.pane.set_drop_target_visual(True)
                self.pane._mark_self_active()
            except Exception:
                pass
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)
    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            try:
                self.pane.set_drop_target_visual(True)
                self.pane._mark_self_active()
            except Exception:
                pass
            if e.keyboardModifiers() & Qt.ControlModifier:
                e.setDropAction(Qt.CopyAction)
            else:
                e.setDropAction(Qt.MoveAction)
            e.accept()
        else:
            super().dragMoveEvent(e)
    def dragLeaveEvent(self, e):
        try:
            self.pane.set_drop_target_visual(False)
        except Exception:
            pass
        super().dragLeaveEvent(e)
    def dropEvent(self, e):
        try:
            if e.mimeData().hasUrls():
                urls=e.mimeData().urls(); srcs=[u.toLocalFile() for u in urls if u.isLocalFile()]
                if srcs:
                    op="copy" if (e.dropAction()==Qt.CopyAction or (e.keyboardModifiers() & Qt.ControlModifier)) else "move"
                    self.pane._start_bg_op(op, srcs, self.pane.current_path()); e.acceptProposedAction(); return
            super().dropEvent(e)
        finally:
            try:
                self.pane.set_drop_target_visual(False)
            except Exception:
                pass


    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F5:
            try:
                self.pane.hard_refresh()
                self.ensure_drag_ready()
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
            self._drag_start_was_selected = self._row_selected_for_drag(sm, self._drag_start_index)
            self._drag_ready = False


        try:
            if not self.hasFocus():
                self.setFocus(Qt.MouseFocusReason)
        except Exception:
            pass


        if e.button() == Qt.LeftButton and e.modifiers() == Qt.NoModifier:
            clicked = self.indexAt(e.pos())
            sm = self.selectionModel()
            prev_rows = sm.selectedRows(0) if sm else []
            same_single = (clicked.isValid() and len(prev_rows)==1 and
                           prev_rows[0].row()==clicked.row() and prev_rows[0].parent()==clicked.parent())
            super().mousePressEvent(e)
            if same_single and sm and len(sm.selectedRows(0)) == 0:
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
                        and (not self._row_selected_for_drag(sm, ix))):
                        try:
                            sm.select(ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                        except Exception:
                            pass
                    if ix.isValid() and sm and self._row_selected_for_drag(sm, ix):


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
        self.setObjectName("paneRoot")
        self.setProperty("active", False)
        self.setProperty("drop_target", False)
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
        self._search_running = False
        self._search_results_stale = False
        self._back_stack=[]; self._fwd_stack=[]; self._undo_stack=[]
        self._last_hover_index=QtCore.QModelIndex(); self._tooltip_last_ms=0.0; self._tooltip_interval_ms=180; self._tooltip_last_text=""
        self._tooltip_display_ms = 36000
        try:
            st = QApplication.instance().style()
            base_ms = int(st.styleHint(QStyle.SH_ToolTip_FallAsleepDelay)) if st else 0
            if base_ms > 0:
                self._tooltip_display_ms = max(1000, int(base_ms * HOVER_TOOLTIP_DURATION_MULTIPLIER))
        except Exception:
            pass
        self._fast_model=FastDirModel(self); self._fast_proxy=FsSortProxy(self); self._fast_proxy.setSourceModel(self._fast_model)
        self._using_fast=False; self._fast_stat_worker=None; self._enum_worker=None
        self._fast_enum_count = 0
        self._fast_enum_root = ""
        self._fast_enum_done = False
        self._large_folder_mode = False
        self._file_worker=None
        self._icon_cache = {}
        self._icon_failed = GLOBAL_SHELL_ICON_FAILURES
        self._icon_pending = set()
        self._icon_queue = []
        self._icon_worker = None
        self._op_progress_dialog=None
        self._dirload_timer={}
        self._sort_column = 0
        self._sort_order = Qt.AscendingOrder
        self._search_sort_column = 0
        self._search_sort_order = Qt.AscendingOrder
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
        self._fs_change_generation = 0

    def _build_toolbar(self):
        self.btn_star=QToolButton(self); self.btn_star.setCheckable(True)
        self.btn_star.setIcon(icon_star(False, getattr(self.host,"theme","dark"))); self.btn_star.setToolTip("Add bookmark for this folder"); self.btn_star.setFixedHeight(UI_H)
        self._bm_btn_container=QWidget(self); self._bm_btn_layout=QHBoxLayout(self._bm_btn_container)
        self._bm_btn_layout.setContentsMargins(0,0,0,0); self._bm_btn_layout.setSpacing(ROW_SPACING)
        self._quick_bm_buttons = []
        self._quick_bm_more_btn = QToolButton(self._bm_btn_container)
        self._quick_bm_more_btn.setObjectName("quickBookmarkMoreBtn")
        self._quick_bm_more_btn.setText("...")
        self._quick_bm_more_btn.setToolTip("More bookmarks")
        self._quick_bm_more_btn.setFixedHeight(UI_H)
        self._quick_bm_more_btn.setFixedWidth(QUICK_BOOKMARK_MORE_W)
        self._quick_bm_more_btn.setPopupMode(QToolButton.InstantPopup)
        self._quick_bm_more_btn.hide()

        self.btn_cmd=QToolButton(self); self.btn_cmd.setIcon(icon_cmd(self.host.theme)); self.btn_cmd.setToolTip("Open Command Prompt here"); self.btn_cmd.setFixedHeight(UI_H)
        self.btn_explorer=QToolButton(self); self.btn_explorer.setIcon(icon_explorer(self.host.theme)); self.btn_explorer.setToolTip("Open this folder in Windows Explorer"); self.btn_explorer.setFixedHeight(UI_H)
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
        row_toolbar.addWidget(self.btn_explorer)
        row_toolbar.addWidget(self.btn_up)
        row_toolbar.addWidget(self.btn_new)
        row_toolbar.addWidget(self.btn_new_file)
        row_toolbar.addWidget(self.btn_refresh)
        self._row_toolbar=row_toolbar


        _tight_css = "QToolButton{padding-left:4px;padding-right:4px;}"
        for b in (self.btn_cmd, self.btn_explorer, self.btn_up, self.btn_new, self.btn_new_file, self.btn_refresh):
            b.setStyleSheet(_tight_css)
            b.setAutoRaise(True)
        return row_toolbar

    def _build_path_row(self):
        self.path_bar=PathBar(self); self.path_bar.setToolTip("Breadcrumb - Double-click or F4/Ctrl+L to enter path")
        row_path=QHBoxLayout(); row_path.setContentsMargins(0,0,0,0); row_path.setSpacing(0); row_path.addWidget(self.path_bar,1)
        return row_path

    def _build_filter_row(self):
        self.filter_label=QLabel("Filter:", self)
        self.filter_edit=QLineEdit(self); self.filter_edit.setPlaceholderText("Filter (abc, *.pdf, *file*.xls*)"); self.filter_edit.setClearButtonEnabled(True); self.filter_edit.setFixedHeight(UI_H)
        self.filter_label.setFixedHeight(UI_H); self.filter_label.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.btn_search=QToolButton(self); self.btn_search.setText("Search"); self.btn_search.setToolTip("Run recursive search"); self.btn_search.setFixedHeight(UI_H)
        self.btn_search.setProperty("busy", False)
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
        self.view.setSelectionBehavior(QAbstractItemView.SelectRows)
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
        header.resizeSection(1, SIZE_COL_WIDTH)
        header.resizeSection(2, 44)
        header.resizeSection(3, DATE_COL_WIDTH)
        self.view.setColumnHidden(2, False)
        self._schedule_browse_name_autofit()

    def _configure_header_fast(self):
        header = self.view.header()
        header.setStretchLastSection(False)
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        header.resizeSection(1, SIZE_COL_WIDTH)
        header.resizeSection(2, 44)
        header.resizeSection(3, DATE_COL_WIDTH)
        self.view.setColumnHidden(2, False)
        self._schedule_browse_name_autofit()

    def _configure_header_search(self):
        header = self.view.header()
        name_width = max(self._browse_name_min_width, header.sectionSize(0))
        size_width = self._load_search_header_width(1, SIZE_COL_WIDTH)
        ext_width = self._load_search_header_width(2, 44)
        date_width = self._load_search_header_width(3, DATE_COL_WIDTH)
        folder_width = self._load_search_header_width(4, SEARCH_FOLDER_COL_WIDTH)
        self._header_resize_guard = True
        try:
            header.blockSignals(True)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            header.setSectionResizeMode(1, QHeaderView.Interactive)
            header.setSectionResizeMode(2, QHeaderView.Interactive)
            header.setSectionResizeMode(3, QHeaderView.Interactive)
            header.setSectionResizeMode(4, QHeaderView.Interactive)
            header.resizeSection(0, name_width)
            header.resizeSection(1, size_width)
            header.resizeSection(2, ext_width)
            header.resizeSection(3, date_width)
            header.resizeSection(4, folder_width)
            self.view.setColumnHidden(2, True)
        finally:
            header.blockSignals(False)
            self._header_resize_guard = False

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

        vp_w = v.viewport().width()
        if vp_w <= 0:
            return
        fixed_w = 0
        for col in (1, 2, 3):
            if not v.isColumnHidden(col):
                fixed_w += header.sectionSize(col)
        target = vp_w - fixed_w
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
        if self._header_resize_guard:
            return

        if self._search_mode:
            if logical_index in (1, 2, 3, 4):
                self._save_search_header_width(logical_index, _new_size)
            return

        if logical_index in (1, 2, 3):
            self._schedule_browse_name_autofit()

    def _build_status_row(self):
        self.lbl_sel=QLabel("", self); self.lbl_free=QLabel("", self)
        self.lbl_mode=QLabel("", self)
        self.lbl_mode.setObjectName("modeBadge")
        self.lbl_mode.setFixedHeight(UI_H)
        self.lbl_mode.hide()
        self.op_progress_bar = QProgressBar(self)
        self.op_progress_bar.setObjectName("paneOpProgress")
        self.op_progress_bar.setFixedHeight(UI_H)
        self.op_progress_bar.setMinimumWidth(120)
        self.op_progress_bar.setMaximumWidth(240)
        self.op_progress_bar.setTextVisible(True)
        self.op_progress_bar.hide()
        self.btn_op_cancel = QToolButton(self)
        self.btn_op_cancel.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))
        self.btn_op_cancel.setToolTip("Cancel file operation")
        self.btn_op_cancel.setFixedHeight(UI_H)
        self.btn_op_cancel.setAutoRaise(True)
        self.btn_op_cancel.clicked.connect(self._request_file_op_cancel)
        self.btn_op_cancel.hide()
        row_status=QHBoxLayout(); row_status.setContentsMargins(0,0,0,0); row_status.setSpacing(ROW_SPACING)
        row_status.addWidget(self.lbl_sel,0); row_status.addWidget(self.lbl_mode,0); row_status.addStretch(1)
        row_status.addWidget(self.op_progress_bar,0); row_status.addWidget(self.btn_op_cancel,0)
        row_status.addWidget(self.lbl_free,0)
        self._row_status=row_status
        return row_status

    def _request_file_op_cancel(self):
        worker = getattr(self, "_file_worker", None)
        if worker and worker.isRunning():
            try:
                worker.cancel()
            except Exception:
                pass
            self._set_pane_progress_status("Cancelling...")
            try:
                self.btn_op_cancel.setEnabled(False)
            except Exception:
                pass

    def _show_pane_progress(self, label: str, busy: bool = False):
        bar = getattr(self, "op_progress_bar", None)
        btn = getattr(self, "btn_op_cancel", None)
        if not bar or not btn:
            return
        label = str(label or "Working")
        if busy:
            bar.setRange(0, 0)
            bar.setFormat(label)
        else:
            bar.setRange(0, 100)
            bar.setValue(0)
            bar.setFormat(f"{label} %p%")
        bar.setToolTip(label)
        btn.setEnabled(True)
        bar.show()
        btn.show()

    def _set_pane_progress_status(self, status: str):
        bar = getattr(self, "op_progress_bar", None)
        if not bar:
            return
        bar.setToolTip(str(status or ""))

    def _set_pane_progress_value(self, value: int):
        bar = getattr(self, "op_progress_bar", None)
        if not bar:
            return
        try:
            if bar.maximum() != 0:
                bar.setValue(int(value))
        except Exception:
            pass

    def _hide_pane_progress(self):
        bar = getattr(self, "op_progress_bar", None)
        btn = getattr(self, "btn_op_cancel", None)
        if bar:
            try:
                if bar.maximum() != 0:
                    bar.setValue(100)
                bar.hide()
            except Exception:
                pass
        if btn:
            try:
                btn.hide()
                btn.setEnabled(True)
            except Exception:
                pass

    def _set_large_folder_mode(self, active: bool, count: int | None = None, complete: bool = False):
        self._large_folder_mode = bool(active)
        lbl = getattr(self, "lbl_mode", None)
        if not lbl:
            return
        if not active:
            lbl.clear()
            lbl.hide()
            return
        text = "Large folder mode"
        if count is not None:
            try:
                n = max(0, int(count))
                text += f": {n:,}{'' if complete else '+'} items"
            except Exception:
                pass
        lbl.setText(text)
        lbl.setToolTip("Using fast listing for a large folder; native details may load differently.")
        lbl.show()

    def _apply_layout(self, row_toolbar, row_path, row_filter, row_status):
        root_layout=QVBoxLayout(self); root_layout.setContentsMargins(*PANE_MARGIN); root_layout.setSpacing(max(1, ROW_SPACING//2))
        root_layout.addLayout(row_toolbar); root_layout.addLayout(row_path); root_layout.addLayout(row_filter)
        root_layout.addWidget(self.view,1); root_layout.addLayout(row_status)

    def _connect_signals(self):
        self.host.namedBookmarksChanged.connect(self._on_bookmarks_changed)
        self.path_bar.pathSubmitted.connect(lambda p: self.set_path(p, push_history=True))
        self.btn_star.clicked.connect(self._on_star_toggle)
        self.btn_cmd.clicked.connect(self._open_cmd_here)
        self.btn_explorer.clicked.connect(self._open_current_path_in_explorer)
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
        self.btn_search.clicked.connect(self._on_search_button_clicked)
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
        add_sc("Ctrl+Shift+R", self.bulk_rename_selection)
        add_sc(Qt.Key_Return, self._open_current); add_sc(Qt.Key_Enter, self._open_current); add_sc("Ctrl+O", self._open_current)


        add_sc("Ctrl+Shift+C", lambda: self._copy_path_shortcut(False))
        add_sc("Ctrl+Shift+D", self._open_selected_item_container)
        add_sc("Alt+Shift+C",  lambda: self._copy_path_shortcut(True))

    def _set_search_button_state(self, running: bool):
        self._search_running = bool(running)
        btn = getattr(self, "btn_search", None)
        if not btn:
            return
        btn.setText("Cancel" if self._search_running else "Search")
        btn.setToolTip("Cancel running search" if self._search_running else "Run recursive search")
        btn.setProperty("busy", self._search_running)
        try:
            st = btn.style()
            st.unpolish(btn)
            st.polish(btn)
            btn.update()
        except Exception:
            pass

    def _on_search_button_clicked(self):
        w = getattr(self, "_search_worker", None)
        if w and w.isRunning():
            self._cancel_search_worker()
            self._set_search_button_state(False)
            self.host.flash_status("Search cancelled")
            return
        self._apply_filter()

    def _load_sort_settings(self):
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            self._sort_column = s.value(f"pane_{self.pane_id}/sort_column", 0, type=int)
            order_val = s.value(f"pane_{self.pane_id}/sort_order", Qt.AscendingOrder, type=int)
            self._sort_order = Qt.DescendingOrder if order_val == Qt.DescendingOrder else Qt.AscendingOrder
        except Exception:
            self._sort_column = 0
            self._sort_order = Qt.AscendingOrder
        self._search_sort_column = 0
        self._search_sort_order = Qt.AscendingOrder

    def _save_sort_settings(self):
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            s.setValue(f"pane_{self.pane_id}/sort_column", self._sort_column)
            s.setValue(f"pane_{self.pane_id}/sort_order", int(self._sort_order))
            s.sync()
        except Exception:
            pass

    def _get_sort_state(self, search_mode: bool | None = None) -> tuple[int, Qt.SortOrder]:
        is_search = self._search_mode if search_mode is None else bool(search_mode)
        if is_search:
            col = getattr(self, "_search_sort_column", 0)
            order = getattr(self, "_search_sort_order", Qt.AscendingOrder)
            max_col = 4
        else:
            col = getattr(self, "_sort_column", 0)
            order = getattr(self, "_sort_order", Qt.AscendingOrder)
            max_col = 3
        try:
            col = int(col)
        except Exception:
            col = 0
        if col < 0 or col > max_col:
            col = 0
        order = Qt.DescendingOrder if order == Qt.DescendingOrder else Qt.AscendingOrder
        return col, order

    def _set_sort_state(self, column: int, order, search_mode: bool | None = None):
        is_search = self._search_mode if search_mode is None else bool(search_mode)
        col, normalized_order = self._get_sort_state(search_mode=is_search)
        try:
            col = int(column)
        except Exception:
            pass
        max_col = 4 if is_search else 3
        if col < 0 or col > max_col:
            col = 0
        normalized_order = Qt.DescendingOrder if order == Qt.DescendingOrder else Qt.AscendingOrder
        if is_search:
            self._search_sort_column = col
            self._search_sort_order = normalized_order
        else:
            self._sort_column = col
            self._sort_order = normalized_order
            self._save_sort_settings()
        return col, normalized_order

    def _sync_sort_state_from_view(self):
        v = getattr(self, "view", None)
        if v is None:
            return
        header = v.header()
        if header is None:
            return
        try:
            col = header.sortIndicatorSection()
            order = header.sortIndicatorOrder()
        except Exception:
            return
        self._set_sort_state(col, order, search_mode=self._search_mode)

    def _load_search_header_width(self, logical_index: int, default: int) -> int:
        fallback = max(24, int(default))
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            width = s.value(f"pane_{self.pane_id}/search_width_{logical_index}", fallback, type=int)
            return max(24, int(width))
        except Exception:
            return fallback

    def _save_search_header_width(self, logical_index: int, width: int):
        try:
            s = QSettings(ORG_NAME, APP_NAME)
            s.setValue(f"pane_{self.pane_id}/search_width_{logical_index}", max(24, int(width)))
            s.sync()
        except Exception:
            pass

    def _apply_saved_sort(self, search_mode: bool | None = None):
        try:
            v = self.view
            col, order = self._get_sort_state(search_mode=search_mode)
            model = v.model()
            if model is not None:
                try:
                    max_col = max(0, int(model.columnCount()) - 1)
                except Exception:
                    max_col = 0
                if col > max_col:
                    col = 0
            if not v.isSortingEnabled():
                v.setSortingEnabled(True)
            v.header().setSortIndicator(col, order)
            v.sortByColumn(col, order)
        except Exception:
            pass


    def set_active_visual(self, active: bool):
        try:
            active = bool(active)
            if self.objectName() != "paneRoot":
                self.setObjectName("paneRoot")
            if bool(self.property("active")) == active:
                return
            self.setProperty("active", active)


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
    def set_drop_target_visual(self, active: bool):
        try:
            if self.objectName() != "paneRoot":
                self.setObjectName("paneRoot")
            self.setProperty("drop_target", bool(active))

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

    def _select_visible_path(self, target_path: str, focus: bool = False) -> bool:
        target_path = nice_path(target_path)
        target_key = os.path.normcase(target_path)

        try:
            if focus and self.view and not self.view.hasFocus():
                self.view.setFocus(Qt.ShortcutFocusReason)
        except Exception:
            pass

        if self._search_mode:
            return False

        try:
            if self._using_fast:
                rows = self._fast_model.rowCount()
                for r in range(rows):
                    rp = self._fast_model.row_path(r)
                    if rp and os.path.normcase(rp) == target_key:
                        prx_ix = self._fast_proxy.index(r, 0)
                        sm = self.view.selectionModel()
                        sm.clearSelection()
                        sm.select(prx_ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                        self.view.scrollTo(prx_ix, QAbstractItemView.PositionAtCenter)
                        self.view.setCurrentIndex(prx_ix)
                        return True
            else:
                src_ix = self.source_model.index(target_path)
                if src_ix.isValid():
                    st_ix = self.stat_proxy.mapFromSource(src_ix)
                    prx_ix = self.proxy.mapFromSource(st_ix)
                    sm = self.view.selectionModel()
                    sm.clearSelection()
                    sm.select(prx_ix, QtCore.QItemSelectionModel.Select | QtCore.QItemSelectionModel.Rows)
                    self.view.scrollTo(prx_ix, QAbstractItemView.PositionAtCenter)
                    self.view.setCurrentIndex(prx_ix)
                    return True
        except Exception:
            pass
        return False

    def _schedule_select_visible_path(self, target_path: str, focus: bool = False):
        target_path = nice_path(target_path)
        for delay in (0, 80, 200, 450, 900, 1500):
            QTimer.singleShot(delay, lambda p=target_path, f=focus: self._select_visible_path(p, focus=f))

    def _open_selected_item_container(self):
        if not self._search_mode:
            try:
                self.host.statusBar().showMessage("This shortcut is available in search mode.", 2000)
            except Exception:
                pass
            return

        sel = self._selected_paths()
        if len(sel) != 1:
            try:
                self.host.statusBar().showMessage("Select exactly one item.", 2000)
            except Exception:
                pass
            return

        target_path = nice_path(sel[0])
        try:
            container_path = str(Path(target_path).parent)
        except Exception:
            container_path = os.path.dirname(target_path.rstrip("\\/")) or target_path

        if not container_path:
            container_path = target_path

        if not os.path.isdir(container_path):
            try:
                self.host.statusBar().showMessage("Containing folder is not available.", 2000)
            except Exception:
                pass
            return

        self.set_path(container_path, push_history=True)
        self._schedule_select_visible_path(target_path, focus=True)
        try:
            self.host.flash_status("Opened containing folder")
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
        stopped = self._stop_worker_thread(getattr(self, "_search_worker", None), 120, "search")
        self._search_worker = None
        stopped = self._stop_worker_thread(
            getattr(self, "_search_stat_worker", None), 80, "search-stat"
        ) and stopped
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
        self._set_search_button_state(False)
        return stopped

    @QtCore.pyqtSlot(str, list)
    def _on_search_batch(self, base_path: str, rows: list):
        if not self._search_mode or not isinstance(self._search_model, SearchResultModel):
            return
        self._search_model.append_rows(rows)
        self._request_visible_stats(0)

    @QtCore.pyqtSlot(int, int, str)
    def _on_search_progress(self, dirs: int, entries: int, folder: str):
        if not getattr(self, "_search_running", False):
            return
        tail = ""
        if folder:
            try:
                tail = f" — {nice_path(folder)}"
            except Exception:
                tail = f" — {folder}"
        self.host.statusBar().showMessage(
            f"Searching... {int(dirs)} folders, {int(entries)} items{tail}",
            1200,
        )

    @QtCore.pyqtSlot()
    def _on_search_finished(self):
        worker = self.sender()
        if worker is not getattr(self, "_search_worker", None):
            return
        try:
            if QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()
        except Exception:
            pass
        self._search_worker = None
        self._search_running = False
        self._set_search_button_state(False)
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
        try:
            rows = self._search_model.rowCount() if self._search_model is not None else 0
        except Exception:
            rows = 0
        stale = " Results may be stale." if getattr(self, "_search_results_stale", False) else ""
        self.host.statusBar().showMessage(f"Search complete: {rows} result(s).{stale}", 4000)

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
        w.finishedCycle.connect(lambda b=batch, worker=w: self._on_search_stat_cycle_finished(worker, b), Qt.QueuedConnection)
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

    def _on_search_stat_cycle_finished(self, worker, batch):
        if worker is not getattr(self, "_search_stat_worker", None):
            return
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
        if self.sender() is not getattr(self, "_search_stat_worker", None):
            return
        model = getattr(self, "_search_model", None)
        if isinstance(model, SearchResultModel):
            model.apply_stat(path, size_val, mtime_val)



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

    def _apply_icon_to_models(self, key: str, icon: QIcon):
        try:
            self._fast_model.apply_icon_key(key, icon)
        except Exception:
            pass
        model = getattr(self, "_search_model", None)
        if isinstance(model, SearchResultModel):
            try:
                model.apply_icon_key(key, icon)
            except Exception:
                pass

    @QtCore.pyqtSlot(str, bytes, int, int)
    def _apply_async_icon_raw(self, key: str, raw: bytes, width: int, height: int):
        if not raw or width <= 0 or height <= 0:
            return
        try:
            image = QImage(raw, int(width), int(height), int(width) * 4, QImage.Format_ARGB32).copy()
            icon = QIcon(QPixmap.fromImage(image))
            if icon.isNull():
                return
            self._icon_cache[key] = icon
            GLOBAL_SHELL_ICON_CACHE[key] = icon
            self._icon_failed.pop(key, None)
            self._apply_icon_to_models(key, icon)
        except Exception:
            pass

    def _queue_async_icons(self, jobs):
        added = False
        now = time.monotonic()
        for key, path, is_dir in jobs:
            key = str(key or "")
            if not key:
                continue
            cached = self._icon_cache.get(key) or GLOBAL_SHELL_ICON_CACHE.get(key)
            if cached is not None:
                self._icon_cache[key] = cached
                self._icon_failed.pop(key, None)
                self._apply_icon_to_models(key, cached)
                continue
            failed_at = self._icon_failed.get(key)
            if failed_at is not None and (now - float(failed_at)) < SHELL_ICON_FAILURE_TTL_S:
                continue
            if failed_at is not None:
                self._icon_failed.pop(key, None)
            if key in self._icon_pending:
                continue
            self._icon_pending.add(key)
            self._icon_queue.append((key, str(path or ""), bool(is_dir)))
            added = True
        if added:
            self._start_next_icon_worker()

    def _start_next_icon_worker(self, batch_limit: int = 64):
        cur = getattr(self, "_icon_worker", None)
        if cur and cur.isRunning():
            return
        if not self._icon_queue:
            self._icon_worker = None
            return
        batch = self._icon_queue[:max(1, int(batch_limit))]
        del self._icon_queue[:len(batch)]
        worker = ShellIconWorker(batch, self)
        worker.iconReady.connect(self._apply_async_icon_raw, Qt.QueuedConnection)
        worker.finishedCycle.connect(self._on_icon_cycle_finished, Qt.QueuedConnection)
        self._icon_worker = worker
        worker.start()

    @QtCore.pyqtSlot(object)
    def _on_icon_cycle_finished(self, jobs):
        for key, _path, _is_dir in list(jobs or []):
            self._icon_pending.discard(key)
            if key not in self._icon_cache:
                self._icon_failed[key] = time.monotonic()
        self._icon_worker = None
        self._start_next_icon_worker()

    def _cancel_icon_worker(self):
        self._icon_queue = []
        self._icon_pending.clear()
        stopped = self._stop_worker_thread(getattr(self, "_icon_worker", None), 150, "shell-icon")
        self._icon_worker = None
        return stopped

    def _cancel_fast_stat_worker(self):
        stopped = self._stop_worker_thread(self._fast_stat_worker, 120, "fast-stat")
        self._fast_stat_worker=None
        return stopped

    def _cancel_enum_worker(self, wait_ms: int = 150):
        stopped = self._stop_worker_thread(self._enum_worker, wait_ms, "dir-enum")
        self._enum_worker = None
        return stopped

    def _cancel_file_worker(self, wait_ms: int = 300):
        worker = getattr(self, "_file_worker", None)
        manager = getattr(getattr(self, "host", None), "file_ops", None)
        if worker and manager and manager.owns(worker):
            stopped = manager.cancel_worker(worker, wait_ms)
        else:
            stopped = self._stop_worker_thread(worker, wait_ms, "file-op")
        self._file_worker = None
        try:
            self._hide_pane_progress()
        except Exception:
            pass
        dlg = getattr(self, "_op_progress_dialog", None)
        if dlg:
            try:
                dlg.close()
                dlg.deleteLater()
            except Exception:
                pass
            self._op_progress_dialog = None
        return stopped

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
            self._cancel_icon_worker()
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

        # References may have been cleared after a timed-out cancellation.  The
        # workers remain QObject children, so audit every child thread before the
        # pane is allowed to be destroyed.
        return _cancel_and_wait_child_threads(self, wait_ms)

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

    def _visible_browse_stat_paths(self, margin_before: int = 40, margin_after: int = 80) -> list[str]:
        if self._search_mode or self._using_fast:
            return []
        if self.view.model() is not self.proxy:
            return []

        model = self.proxy
        stat_proxy = self.stat_proxy
        src_model = self.source_model
        root_ix = self.view.rootIndex()
        vp = self.view.viewport()

        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        rc = model.rowCount(root_ix)
        end = bot_ix.row() if bot_ix.isValid() else min(start + 120, rc - 1)
        start = max(0, start - max(0, int(margin_before)))
        end = min(rc - 1, end + max(0, int(margin_after)))
        if end < start:
            end = start

        paths = []
        for r in range(start, end + 1):
            prx_ix = model.index(r, 0, root_ix)
            if not prx_ix.isValid():
                continue
            st_ix = model.mapToSource(prx_ix)
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
        return paths

    def _refresh_visible_browse_stats(self, force: bool = False, generation: int | None = None):
        if generation is not None and generation != getattr(self, "_fs_change_generation", 0):
            return
        paths = self._visible_browse_stat_paths()
        if paths:
            self.stat_proxy.request_paths(paths, force=force)

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
        try:
            selected = list(self.view.selectionModel().selectedRows(0))
        except Exception:
            selected = []
        sig = tuple(self._index_to_full_path(ix) or "" for ix in selected)
        now = time.perf_counter()
        if sig == self._selection_cache_sig and (now - self._selection_cache_ts) <= 0.2:
            return self._selection_cache_data

        count = len(selected)
        only_files = count > 0
        total = 0
        for ix in selected:
            try:
                if bool(ix.data(IS_DIR_ROLE)):
                    only_files = False
                    total = 0
                    break
                size_value = ix.sibling(ix.row(), 1).data(SIZE_BYTES_ROLE)
                total += max(0, int(size_value or 0))
            except Exception:
                only_files = False
                total = 0
                break

        data = (count, only_files, total)
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
            icon_jobs = []
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
                    key = self._fast_model.icon_key(row)
                    if p and key:
                        icon_jobs.append((key, p, self._fast_model.row_is_dir(row)))

            if icon_jobs:
                self._queue_async_icons(icon_jobs)
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

        self._refresh_visible_browse_stats(force=False)

    def _on_header_clicked(self, col:int):
        v=self.view
        cur_col, cur_order = self._get_sort_state()

        if col == cur_col:
            new_order = Qt.DescendingOrder if cur_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            new_order = Qt.AscendingOrder

        col, new_order = self._set_sort_state(col, new_order, search_mode=self._search_mode)

        v.header().setSortIndicator(col, new_order)
        if self._search_mode and self._search_running:
            return
        if not v.isSortingEnabled():
            v.setSortingEnabled(True)
        v.sortByColumn(col, new_order)

    def _mark_self_active(self):
        try:
            if hasattr(self.host, "mark_active_pane"):
                self.host.mark_active_pane(self)
        except Exception:
            pass

    def eventFilter(self, obj, ev):

        if obj is self.view.viewport():
            if ev.type() == QEvent.ToolTip:
                # Suppress default model tooltip so our custom duration is not overridden.
                return True
            if ev.type()==QEvent.MouseButtonPress:
                if ev.button()==Qt.XButton1: self.go_back(); return True
                if ev.button()==Qt.XButton2: self.go_forward(); return True
            if ev.type()==QEvent.MouseMove:
                ix=self.view.indexAt(ev.pos())
                if ix!=self._last_hover_index: self._last_hover_index=ix
                if not self._search_mode:
                    QToolTip.hideText()
                    self._tooltip_last_text = ""
                elif ix.isValid():
                    now_ms=time.perf_counter()*1000.0
                    if (now_ms-self._tooltip_last_ms)>=self._tooltip_interval_ms:
                        name=ix.sibling(ix.row(),0).data(Qt.DisplayRole); full=self._index_to_full_path(ix)
                        tip=full if full else name
                        if tip!=self._tooltip_last_text:
                            QToolTip.showText(QCursor.pos(), tip, self.view.viewport(), QtCore.QRect(), self._tooltip_display_ms)
                            self._tooltip_last_text=tip; self._tooltip_last_ms=now_ms
                else:
                    QToolTip.hideText()
                    self._tooltip_last_text = ""
            if ev.type() == QEvent.Leave:
                QToolTip.hideText()
                self._tooltip_last_text = ""
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

    def _open_current_path_in_explorer(self):
        path = self.current_path() or os.getcwd()
        try:
            path = os.path.abspath(path)
        except Exception:
            pass

        if not os.path.isdir(path):
            QMessageBox.warning(self, "Windows Explorer", "Current folder is not available.")
            return

        if sys.platform == "win32":
            if HAS_PYWIN32:
                try:
                    win32api.ShellExecute(
                        int(self.window().winId()) if self.window() else 0,
                        "open",
                        "explorer.exe",
                        f'"{path}"',
                        None,
                        win32con.SW_SHOWNORMAL
                    )
                    return
                except Exception:
                    pass

            try:
                subprocess.Popen(["explorer.exe", path])
                return
            except Exception as e:
                QMessageBox.critical(self, "Windows Explorer", f"Failed to open Windows Explorer:\n{e}")
                return

        if not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
            QMessageBox.critical(self, "Open Folder", f"Failed to open folder:\n{path}")


    def _on_bookmarks_changed(self,*_): self._update_star_button(); self._rebuild_quick_bookmark_buttons()
    def _on_star_toggle(self): self.host.toggle_bookmark(self.current_path())
    def _update_star_button(self):
        idx,it=self.host.is_path_bookmarked(self.current_path()); checked=bool(it and it.get("enabled"))
        self.btn_star.setChecked(checked); self.btn_star.setIcon(icon_star(checked, getattr(self.host,"theme","dark")))
        self.btn_star.setToolTip("Remove bookmark for this folder" if checked else "Add bookmark for this folder")
    def _quick_bookmark_button_width(self, text: str) -> int:
        try:
            text_w = self._bm_btn_container.fontMetrics().horizontalAdvance(str(text or ""))
        except Exception:
            text_w = len(str(text or "")) * 7
        return max(QUICK_BOOKMARK_MIN_W, min(QUICK_BOOKMARK_MAX_W, text_w + 18))

    def _update_quick_bookmark_more_menu(self, overflow_buttons):
        more_btn = getattr(self, "_quick_bm_more_btn", None)
        if more_btn is None:
            return
        old_menu = more_btn.menu()
        if old_menu is not None:
            more_btn.setMenu(None)
            old_menu.deleteLater()
        if not overflow_buttons:
            more_btn.setToolTip("More bookmarks")
            return

        menu = QMenu(more_btn)
        for btn in overflow_buttons:
            name = str(btn.property("fullText") or btn.text() or "")
            path = str(btn.property("bookmarkPath") or "")
            act = QAction(name, menu)
            act.setToolTip(path)
            act.triggered.connect(lambda _=False, p=path: self.set_path(p, push_history=True))
            menu.addAction(act)
        more_btn.setMenu(menu)
        more_btn.setToolTip(f"More bookmarks ({len(overflow_buttons)})")

    def _rebuild_quick_bookmark_buttons(self):
        while self._bm_btn_layout.count():
            it=self._bm_btn_layout.takeAt(0); w=it.widget()
            if w and w is not getattr(self, "_quick_bm_more_btn", None): w.deleteLater()
        self._quick_bm_buttons = []
        for it in self.host.get_enabled_bookmarks():
            name=it.get("name") or _derive_name_from_path(it.get("path","")); p=it.get("path","")
            btn=QToolButton(self._bm_btn_container)
            btn.setObjectName("quickBookmarkBtn")
            btn.setText(str(name))
            btn.setProperty("fullText", str(name))
            btn.setProperty("bookmarkPath", str(p))
            btn.setToolTip(p)
            btn.setFixedHeight(UI_H)
            btn.setFixedWidth(self._quick_bookmark_button_width(str(name)))
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            btn.clicked.connect(lambda _=False, path=p: self.set_path(path, push_history=True))
            self._bm_btn_layout.addWidget(btn)
            self._quick_bm_buttons.append(btn)
        self._bm_btn_layout.addWidget(self._quick_bm_more_btn)
        self._bm_btn_layout.addStretch(1)
        QTimer.singleShot(0, self._refresh_quick_bookmark_button_texts)

    def _refresh_quick_bookmark_button_texts(self):
        try:
            buttons = list(getattr(self, "_quick_bm_buttons", []))
            more_btn = getattr(self, "_quick_bm_more_btn", None)
            if more_btn is None:
                return
            if not buttons:
                more_btn.hide()
                self._update_quick_bookmark_more_menu([])
                return

            available = max(0, int(self._bm_btn_container.contentsRect().width()))
            spacing = max(0, int(self._bm_btn_layout.spacing()))
            widths = [self._quick_bookmark_button_width(str(btn.property("fullText") or "")) for btn in buttons]

            visible_count = 0
            used = 0
            total = len(buttons)
            for i, w in enumerate(widths):
                spacing_before = spacing if visible_count else 0
                remaining_after = total - (i + 1)
                more_reserve = (spacing + QUICK_BOOKMARK_MORE_W) if remaining_after else 0
                if used + spacing_before + w + more_reserve <= available:
                    used += spacing_before + w
                    visible_count += 1
                else:
                    break

            overflow_buttons = buttons[visible_count:]
            for i, btn in enumerate(buttons):
                visible = i < visible_count
                btn.setVisible(visible)
                if not visible:
                    continue
                full = str(btn.property("fullText") or btn.text() or "")
                btn_w = widths[i]
                btn.setFixedWidth(btn_w)
                elided = btn.fontMetrics().elidedText(full, Qt.ElideRight, max(24, btn_w - 12))
                if btn.text() != elided:
                    btn.setText(elided)
            self._update_quick_bookmark_more_menu(overflow_buttons)
            more_btn.setVisible(bool(overflow_buttons))
        except Exception:
            pass

    def _selected_paths(self):
        paths=[]; sel=self.view.selectionModel().selectedRows(0)
        for ix in sel:
            p=self._index_to_full_path(ix)
            if p: paths.append(p)
        seen=set(); out=[]
        for p in paths:
            key = _path_key(p)
            if key in seen:
                continue
            seen.add(key)
            out.append(p)
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
        self._cancel_enum_worker(wait_ms=150)
        try:
            self.stat_proxy.clear_cache()
        except Exception:
            pass

        self._using_fast = True
        self._fast_model.reset_dir(path)
        self.view.setModel(self._fast_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())
        self._hook_selection_model()
        self._configure_header_fast()
        self._set_large_folder_mode(False)
        self._fast_enum_count = 0
        self._fast_enum_root = path
        self._fast_enum_done = False

        sort_col, sort_order = self._get_sort_state(search_mode=False)
        preload_size = (sort_col == 1)
        preload_mtime = (sort_col == 3)
        live_sort_during_enum = False
        was_sorting = self.view.isSortingEnabled()
        old_dynamic_sort = self._fast_proxy.dynamicSortFilter()
        self._fast_proxy.setDynamicSortFilter(False)
        if was_sorting:
            self.view.setSortingEnabled(False)

        self._fast_batch_counter = 0
        worker = DirEnumWorker(path, self, preload_size=preload_size, preload_mtime=preload_mtime)
        self._enum_worker = worker

        def _on_batch(rows):
            if worker is not self._enum_worker or os.path.normcase(path) != os.path.normcase(self.current_path()):
                return
            self._fast_model.append_rows(rows)
            self._fast_enum_count += len(rows or [])
            if self._fast_enum_count >= LARGE_FOLDER_THRESHOLD:
                self._set_large_folder_mode(True, count=self._fast_enum_count, complete=False)
            self._fast_batch_counter += 1
            if (self._fast_batch_counter % 4) == 0:
                self._request_visible_stats(0)

        worker.batchReady.connect(_on_batch, Qt.QueuedConnection)
        worker.error.connect(lambda msg: self.host.statusBar().showMessage(f"List error: {msg}", 4000))

        def _on_finished():
            if worker is not self._enum_worker:
                return
            self._fast_enum_done = True
            self._enum_worker = None
            self._set_large_folder_mode(
                self._fast_enum_count >= LARGE_FOLDER_THRESHOLD,
                count=self._fast_enum_count,
                complete=True,
            )
            self._fast_proxy.setDynamicSortFilter(old_dynamic_sort)
            self.view.setSortingEnabled(True)
            self._apply_saved_sort(search_mode=False)
            if not was_sorting:
                self.view.setSortingEnabled(False)
            self._request_visible_stats(0)
            self._request_visible_stats(80)

        worker.finished.connect(_on_finished, Qt.QueuedConnection)
        self._request_visible_stats(0)
        worker.start()

    def _start_normal_model_loading(self, path: str, known_count: int | None = None):
        # FastDirModel is now the only browse model. Retain this compatibility
        # stub for older call sites without triggering QFileSystemModel scanning.
        return

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


            # FastDirModel is the sole browse model. This performs one directory
            # enumeration instead of scanning once here and again in QFileSystemModel.
            self._use_fast_model(path)


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
                    self._search_results_stale = True
                    self.host.statusBar().showMessage(
                        "Folder changed; search results may be stale. Press Search to refresh.",
                        6000,
                    )
                else:

                    self._enter_browse_mode()
                return


            if getattr(self, "_using_fast", False):
                self.host.statusBar().showMessage("Folder changed; refreshing listing ...", 1500)
                self._use_fast_model(self.current_path())
                return


            self._fs_change_generation += 1
            generation = self._fs_change_generation
            self._refresh_visible_browse_stats(force=True, generation=generation)
            for delay in (350, 1200, 2500):
                QTimer.singleShot(
                    delay,
                    lambda g=generation: self._refresh_visible_browse_stats(force=True, generation=g),
                )
            self._update_pane_status()
        except Exception:
            pass


    @QtCore.pyqtSlot(str)
    def _on_directory_loaded(self, loaded_path: str):
        key = loaded_path.lower()
        timer = self._dirload_timer.pop(key, None)
        if timer is not None:
            dlog(f"directoryLoaded: '{loaded_path}' in {timer.elapsed()} ms")
        # Kept only for QFileSystemModel compatibility; browse rows are supplied
        # by FastDirModel and never switch models after enumeration.

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
        self._sync_sort_state_from_view()
        if self._search_mode:
            self._apply_filter()
            try:
                self.view.ensure_drag_ready()
            except Exception:
                pass
            return
        try: self._cancel_fast_stat_worker()
        except Exception: pass
        try: self._cancel_enum_worker(wait_ms=100)
        except Exception: pass
        try: self.stat_proxy.clear_cache()
        except Exception: pass
        self.set_path(self.current_path(), push_history=False)
        try:
            self.view.ensure_drag_ready()
        except Exception:
            pass
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
        if self.host.set_clipboard({"op":"copy","paths":paths}):
            self.host.flash_status(f"Copied {len(paths)} item(s)")
        else:
            self.host.flash_status("Failed to copy items to clipboard")
    def cut_selection(self):
        paths=self._selected_paths()
        if not paths: return
        if self.host.set_clipboard({"op":"cut","paths":paths}):
            self.host.flash_status(f"Cut {len(paths)} item(s)")
        else:
            self.host.flash_status("Failed to cut items to clipboard")

    def _external_clipboard_payload(self):
        try:
            cb = QApplication.clipboard()
            md = cb.mimeData()
        except Exception:
            md = None

        if sys.platform == "win32":
            native_payload = _read_windows_file_clipboard_payload()
            if native_payload:
                return native_payload

        if not md:
            return None

        try:
            paths = _dedupe_local_paths(
                u.toLocalFile() for u in md.urls() if u.isLocalFile()
            )
        except Exception:
            paths = []
        if not paths:
            return None

        effect = None
        if sys.platform == "win32":
            fmt = 'application/x-qt-windows-mime;value="Preferred DropEffect"'
            if md.hasFormat(fmt):
                try:
                    effect = _decode_preferred_drop_effect(md.data(fmt))
                except Exception:
                    effect = None

        return {"op": _drop_effect_to_operation(effect) or "copy", "paths": paths}

    def paste_into_current(self):
        clip = self.host.get_clipboard() or self._external_clipboard_payload()
        clip = _normalize_file_clipboard_payload(clip)
        if not clip:
            self.host.flash_status("Clipboard has no files to paste")
            return
        dst_dir = self.current_path()
        op = clip.get("op")
        srcs = clip.get("paths") or []
        if not srcs:
            self.host.flash_status("Clipboard has no files to paste")
            return
        move_payload = clip if op in ("cut", "move") else None
        self._start_bg_op(
            "copy" if op == "copy" else "move",
            srcs,
            dst_dir,
            clipboard_payload=move_payload,
        )

    def _push_file_op_undo(self, worker, op: str):
        remove_paths = list(getattr(worker, "undo_remove_paths", []) or [])
        move_pairs = list(getattr(worker, "undo_move_pairs", []) or [])
        actions = []
        if move_pairs:
            actions.append({"type": "move_back", "pairs": move_pairs})
        if remove_paths:
            actions.append({"type": "remove_created", "paths": remove_paths})
        if not actions:
            return
        act = actions[0] if len(actions) == 1 else {"type": "compound", "actions": actions}
        act["label"] = f"Undo {op}"
        self._undo_stack.append(act)

    def _sync_move_clipboard(self, worker):
        expected = getattr(worker, "clipboard_payload", None)
        if not expected:
            return
        try:
            current = self.host.get_clipboard()
        except Exception:
            current = None
        # Never erase a clipboard that the user changed while the operation was running.
        if not _clipboard_payload_matches(current, expected):
            return
        remaining = worker.remaining_source_paths()
        if remaining:
            self.host.set_clipboard({"op": "cut", "paths": remaining})
        else:
            self.host.clear_clipboard()

    @QtCore.pyqtSlot(str)
    def _on_file_worker_error(self, msg):
        worker = self.sender()
        if isinstance(worker, FileOpWorker):
            self._sync_move_clipboard(worker)
        op = getattr(worker, "_ui_op", "operation")
        if msg == "Operation cancelled.":
            self.host.flash_status(f"{str(op).title()} cancelled")
            return
        QMessageBox.critical(self, f"{str(op).title()} failed", msg)

    @QtCore.pyqtSlot()
    def _on_file_worker_finished_ok(self):
        worker = self.sender()
        if not isinstance(worker, FileOpWorker):
            return
        op = getattr(worker, "_ui_op", worker.op)
        self._hide_pane_progress()
        if not self._using_fast and not self._search_mode:
            self.stat_proxy.clear_cache()
        self._request_visible_stats(0)
        self._update_pane_status()
        self._push_file_op_undo(worker, op)
        self._sync_move_clipboard(worker)

        failed = int(getattr(worker, "error_count", 0) or len(getattr(worker, "errors", [])))
        if failed:
            details_list = list(getattr(worker, "errors", []) or [])
            details = "\n".join(details_list)
            if failed > len(details_list):
                details = (details + "\n" if details else "") + f"... {failed - len(details_list)} more error(s)."
            QMessageBox.warning(
                self,
                f"{str(op).title()} completed with errors",
                f"{str(op).title()} finished, but {failed} file(s) could not be processed.\n\n{details[:2000]}",
            )
            self.host.flash_status(f"{str(op).title()} finished with {failed} error(s)")
            return
        self.host.flash_status(f"{str(op).title()} complete")

    @QtCore.pyqtSlot()
    def _on_file_worker_thread_finished(self):
        worker = self.sender()
        if getattr(self, "_file_worker", None) is worker:
            self._file_worker = None
        self._hide_pane_progress()

    def _start_bg_op(self, op, srcs, dst_dir, clipboard_payload=None):
        manager = getattr(self.host, "file_ops", None)

        valid_srcs = []
        skipped_same = []
        blocked_nested = []
        auto_map = {}
        seen_src_keys = set()

        for src in srcs:
            src_key = _path_key(src)
            if src_key in seen_src_keys:
                continue
            seen_src_keys.add(src_key)
            if not src or not os.path.lexists(src):
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
            if os.path.isdir(src) and not _is_dir_link(src) and _is_subpath(dst, src):
                blocked_nested.append(src)
                continue
            valid_srcs.append(src)

        if blocked_nested:
            sample = "\n".join(blocked_nested[:5])
            more = "\n..." if len(blocked_nested) > 5 else ""
            QMessageBox.warning(
                self,
                f"{op.title()} blocked",
                "Cannot copy/move a folder into its own subfolder:\n\n" f"{sample}{more}",
            )

        if not valid_srcs:
            if skipped_same:
                self.host.flash_status("Nothing to move (same source and destination)")
            return False

        conflicts = []
        for src in valid_srcs:
            if src in auto_map:
                continue
            base = os.path.basename(src.rstrip("\\/")) or os.path.basename(src)
            dst = os.path.join(dst_dir, base)
            if os.path.lexists(dst):
                conflicts.append((src, dst))

        conflict_map = dict(auto_map)
        if conflicts:
            dlg = ConflictResolutionDialog(self, conflicts, dst_dir)
            if dlg.exec_() != QDialog.Accepted:
                return False
            conflict_map.update(dlg.result_map())

        worker = FileOpWorker(op, valid_srcs, dst_dir, conflict_map=conflict_map, parent=None)
        worker.clipboard_payload = _normalize_file_clipboard_payload(clipboard_payload)
        worker._ui_op = op

        self._show_pane_progress(op.title(), busy=False)
        worker.progress.connect(self._set_pane_progress_value)
        worker.status.connect(self._set_pane_progress_status)
        worker.status.connect(self.host.show_operation_status)
        worker.started.connect(lambda label=op.title(): self._show_pane_progress(label, busy=False))
        worker.error.connect(self._on_file_worker_error)
        worker.finished_ok.connect(self._on_file_worker_finished_ok)
        worker.finished.connect(self._on_file_worker_thread_finished)
        self._file_worker = worker
        self._op_progress_dialog = None
        if manager:
            state = manager.submit(worker)
            if state == "queued":
                self._show_pane_progress(f"{op.title()} queued", busy=True)
                self.host.flash_status(f"{op.title()} queued")
            elif state != "started":
                self.host.flash_status(f"Could not start {op}")
                return False
        else:
            worker.start()
        return True

    @QtCore.pyqtSlot(str)
    def _on_delete_worker_error(self, msg):
        if msg == "Operation cancelled.":
            QTimer.singleShot(0, self.refresh)
            self.host.flash_status("Delete cancelled")
            return
        QTimer.singleShot(0, self.refresh)
        QMessageBox.critical(self, "Delete failed", msg)

    @QtCore.pyqtSlot()
    def _on_delete_worker_finished_ok(self):
        worker = self.sender()
        if not isinstance(worker, DeleteWorker):
            return
        permanent = bool(getattr(worker, "_ui_permanent", False))
        self._hide_pane_progress()
        if not self._using_fast and not self._search_mode:
            self.stat_proxy.clear_cache()
        self.refresh()
        self._request_visible_stats(0)
        self._update_pane_status()

        if worker.errors:
            details = "\n".join(worker.errors)[:2000]
            failed = len(worker.errors)
            success_msg = (
                f"Deleted {worker.deleted_count} item(s)"
                if permanent else
                f"Sent {worker.deleted_count} item(s) to Recycle Bin"
            )
            if worker.deleted_count > 0:
                QMessageBox.warning(
                    self,
                    "Delete completed with errors",
                    f"{success_msg}, but {failed} failed.\n\n{details}",
                )
            else:
                QMessageBox.critical(self, "Delete failed", details or "Could not delete the selected items.")
            self.host.flash_status(f"Delete finished with {failed} error(s)")
            return

        if permanent:
            self.host.flash_status(f"Deleted {worker.deleted_count} item(s)")
        else:
            self.host.flash_status(f"Sent {worker.deleted_count} item(s) to Recycle Bin")

    @QtCore.pyqtSlot()
    def _on_delete_worker_thread_finished(self):
        worker = self.sender()
        if getattr(self, "_file_worker", None) is worker:
            self._file_worker = None
        self._hide_pane_progress()

    def _start_delete_op(self, paths, permanent: bool = False):
        manager = getattr(self.host, "file_ops", None)

        valid_paths = [p for p in paths if p]
        if not valid_paths:
            return False

        hwnd = int(self.window().winId()) if (not permanent and sys.platform == "win32") else 0
        worker = DeleteWorker(valid_paths, permanent=permanent, hwnd=hwnd, parent=None)
        worker._ui_permanent = permanent

        self._show_pane_progress("Delete" if permanent else "Recycle", busy=False)
        worker.progress.connect(self._set_pane_progress_value)
        worker.status.connect(self._set_pane_progress_status)
        worker.status.connect(self.host.show_operation_status)
        worker.started.connect(lambda label=("Delete" if permanent else "Recycle"): self._show_pane_progress(label, busy=False))
        worker.error.connect(self._on_delete_worker_error)
        worker.finished_ok.connect(self._on_delete_worker_finished_ok)
        worker.finished.connect(self._on_delete_worker_thread_finished)
        self._file_worker = worker
        self._op_progress_dialog = None
        if manager:
            state = manager.submit(worker)
            if state == "queued":
                self._show_pane_progress("Delete queued" if permanent else "Recycle queued", busy=True)
                self.host.flash_status("Delete queued")
            elif state != "started":
                self.host.flash_status("Could not start delete")
                return False
        else:
            worker.start()
        return True


    def delete_selection(self, permanent:bool=False):
        paths=self._selected_paths()
        if not paths: return
        title="Delete permanently" if permanent else "Delete"
        action="permanently delete" if permanent else "move to Recycle Bin"
        msg=f"{len(paths)} item(s) will be {action}.\n\nAre you sure?"
        btn=QMessageBox.question(self,title,msg,QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
        if btn!=QMessageBox.Yes: return
        self._start_delete_op(paths, permanent=permanent)

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

    def bulk_rename_selection(self):
        paths = self._selected_paths()
        if not paths:
            self.host.flash_status("Select items to bulk rename")
            return

        dlg = BulkRenameDialog(self, paths)
        if dlg.exec_() != QDialog.Accepted:
            return

        ops = dlg.result_operations()
        if not ops:
            self.host.flash_status("No items to rename")
            return

        try:
            committed = execute_bulk_rename_transaction(ops)
        except Exception as exc:
            QMessageBox.critical(self, "Bulk Rename failed", str(exc))
            self.refresh()
            return

        if committed:
            self._undo_stack.append({"type": "move_back", "pairs": committed})
        self.refresh()
        self.host.flash_status(f"Renamed {len(committed)} item(s)")

    def _undo_remove_created(self, paths: list[str]) -> bool:
        failed = []
        hwnd = int(self.window().winId()) if sys.platform == "win32" else 0
        for p in reversed(list(paths or [])):
            if not p or not os.path.exists(p):
                continue
            if not recycle_path_to_trash(p, hwnd):
                failed.append(p)
        if failed:
            sample = "\n".join(failed[:8])
            more = "\n..." if len(failed) > 8 else ""
            raise RuntimeError(
                "Could not move copied item(s) to Recycle Bin, so they were left in place:\n"
                f"{sample}{more}"
            )
        return True

    def _apply_undo_action(self, act: dict) -> bool:
        t = act.get("type")
        if t=="mkdir":
            path=act["path"]
            try: os.rmdir(path)
            except OSError:
                QMessageBox.information(self,"Undo New Folder","Folder is not empty; cannot undo safely.")
                return False
        elif t=="delete":
            for p in act.get("paths",[]): remove_any(p)
        elif t=="remove_created":
            return self._undo_remove_created(act.get("paths", []))
        elif t=="move_back":
            for dst,src in reversed(list(act.get("pairs",[]))):
                target_dir=os.path.dirname(src)
                if target_dir:
                    os.makedirs(target_dir, exist_ok=True)
                if os.path.exists(src):
                    base=os.path.basename(src); src=unique_dest_path(target_dir, base)
                worker = FileOpWorker("move", [dst], target_dir or os.curdir)
                if not worker._move_source_transactional(dst, src, None, False):
                    details = "\n".join(worker.errors) or f"Could not safely move {dst} back to {src}."
                    raise RuntimeError(details)
        else:
            return False
        return True

    def undo_last(self):
        if not self._undo_stack: self.host.flash_status("Nothing to undo"); return
        act=self._undo_stack.pop()
        try:
            if act.get("type") == "compound":
                for sub in reversed(list(act.get("actions", []))):
                    if not self._apply_undo_action(sub):
                        return
            elif not self._apply_undo_action(act):
                return
            self.refresh(); self.host.flash_status("Undone")
        except Exception as e:
            QMessageBox.critical(self,"Undo failed",str(e))


    def _dispose_search_models(self, model, proxy):
        if proxy is not None:
            try:
                proxy.setSourceModel(None)
            except Exception:
                pass
            try:
                proxy.deleteLater()
            except Exception:
                pass
        if model is not None:
            try:
                model.deleteLater()
            except Exception:
                pass

    def _enter_browse_mode(self):
        self._sync_sort_state_from_view()
        old_search_model = getattr(self, "_search_model", None)
        old_search_proxy = getattr(self, "_search_proxy", None)

        try:
            if hasattr(self, "_cancel_search_worker"):
                self._cancel_search_worker()
        except Exception:
            pass
        self._search_pending_items = {}
        self._search_stats_done = set()
        self._search_model = None
        self._search_proxy = None
        self._set_search_button_state(False)
        QToolTip.hideText()
        self._tooltip_last_text = ""

        if not self._search_mode:
            self._dispose_search_models(old_search_model, old_search_proxy)
            self._request_visible_stats(0)
            return

        self._search_mode = False

        self._using_fast = True
        self.view.setModel(self._fast_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())


        self._hook_selection_model()

        path = self.current_path()

        self._configure_header_browse()
        if not self.view.isSortingEnabled():
            self.view.setSortingEnabled(True)

        self._apply_saved_sort(search_mode=False)

        self._dispose_search_models(old_search_model, old_search_proxy)
        self._request_visible_stats(0)

    def _enter_search_mode(self, model:SearchResultModel):
        self._sync_sort_state_from_view()
        self._cancel_fast_stat_worker()
        old_search_model = getattr(self, "_search_model", None)
        old_search_proxy = getattr(self, "_search_proxy", None)
        self._search_mode=True
        self._search_model=model
        self._search_proxy=FsSortProxy(self)
        self._search_proxy.setDynamicSortFilter(False)
        self._search_proxy.setSourceModel(self._search_model)
        self.view.setModel(self._search_proxy)
        self.view.setRootIndex(QtCore.QModelIndex())
        if not hasattr(self, "_search_folder_delegate"):
            self._search_folder_delegate = SearchFolderDelegate(self.view)
        self.view.setItemDelegateForColumn(4, self._search_folder_delegate)


        self._hook_selection_model()

        self._configure_header_search()
        col, order = self._get_sort_state(search_mode=True)
        self.view.header().setSortIndicator(col, order)
        # Sorting every incoming batch is expensive; sort once when the search ends.
        self.view.setSortingEnabled(False)
        if old_search_model is not model:
            self._dispose_search_models(old_search_model, old_search_proxy)


    def _apply_filter(self):
        pattern = self.filter_edit.text().strip()
        if not pattern:
            self._enter_browse_mode()
            return


        self._cancel_search_worker()

        base = self.current_path()


        model = SearchResultModel(self)
        self._enter_search_mode(model)


        self._search_pending_items = {}
        self._search_stats_done = set()
        self._search_stat_worker = None
        self._search_stat_queue = []
        self._search_stat_pending = set()
        self._search_results_stale = False


        w = SearchWorker(base, pattern, self, max_results=SEARCH_RESULT_LIMIT)
        self._search_worker = w
        w.batchReady.connect(self._on_search_batch, Qt.QueuedConnection)
        w.progress.connect(self._on_search_progress, Qt.QueuedConnection)
        w.error.connect(lambda msg: self.host.statusBar().showMessage(f"Search error: {msg}", 4000))
        w.truncated.connect(lambda n: self.host.statusBar().showMessage(
            f"Search capped at {n} results. Refine filter to narrow results.", 6000
        ))
        w.finished.connect(self._on_search_finished, Qt.QueuedConnection)

        QApplication.setOverrideCursor(Qt.WaitCursor)
        self._search_running = True
        self._set_search_button_state(True)
        self.host.flash_status("Searching...")
        w.start()


    def _fill_search_visible_icons(self):
        model = getattr(self, "_search_model", None)
        proxy = getattr(self, "_search_proxy", None)
        if not self._search_mode or not isinstance(model, SearchResultModel) or proxy is None:
            return

        root_ix = self.view.rootIndex()
        vp = self.view.viewport()
        top_ix = self.view.indexAt(QtCore.QPoint(1, 1))
        bot_ix = self.view.indexAt(QtCore.QPoint(1, max(1, vp.height() - 2)))
        start = top_ix.row() if top_ix.isValid() else 0
        rc = proxy.rowCount(root_ix)
        end = bot_ix.row() if bot_ix.isValid() else min(start + 160, rc - 1)
        start = max(0, start - 30)
        end = min(rc - 1, end + 70)
        if end < start:
            return

        stat_paths = []
        icon_jobs = []
        for proxy_row in range(start, end + 1):
            pidx = proxy.index(proxy_row, 0, root_ix)
            sidx = proxy.mapToSource(pidx)
            row = sidx.row()
            if row < 0:
                continue
            path = model.row_path(row)
            is_dir = model.row_is_dir(row)
            key = model.icon_key(row)
            if key and not model.has_icon(row):
                icon_jobs.append((key, path, is_dir))
            if path and not model.has_stat(row) and path not in self._search_stat_pending:
                stat_paths.append(path)
                if len(stat_paths) >= 220:
                    break

        if icon_jobs:
            self._queue_async_icons(icon_jobs)
        if stat_paths:
            self._enqueue_search_stat_paths(stat_paths, batch_limit=220)

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
        if not self.shutdown(wait_ms=2000):
            e.ignore()
            return
        super().closeEvent(e)



class MultiExplorer(QMainWindow):
    namedBookmarksChanged=pyqtSignal(list)
    def __init__(self, pane_count:int=6, start_paths=None, initial_theme:str="dark"):
        super().__init__()
        self.setWindowIcon(QIcon(app_resource_path(APP_ICON_FILENAME)))
        self.theme=initial_theme if initial_theme in VALID_THEMES else "dark"
        self._layout_states=[4,6,8]; self._layout_idx=self._layout_states.index(pane_count) if pane_count in self._layout_states else 1
        self.setWindowTitle(f"Multi-Pane File Explorer - {pane_count} panes"); self.resize(1500,900)
        top=QWidget(self); top_lay=QHBoxLayout(top); top_lay.setContentsMargins(6,2,6,2); top_lay.setSpacing(ROW_SPACING)
        for name, tip, slot in (
            ("btn_layout", "Toggle layout (4 / 6 / 8)", self._cycle_layout),
            ("btn_theme", "Toggle Light/Dark", self._toggle_theme),
            ("btn_bm_edit", "Edit Bookmarks", self._open_bookmark_editor),
            ("btn_session", "Session (save/load all pane paths)", self._open_session_manager),
            ("btn_shortcuts", "Keyboard Shortcuts", self._show_shortcuts),
            ("btn_about", "About", self._show_about),
        ):
            btn=QToolButton(top); btn.setToolTip(tip); btn.setFixedHeight(UI_H)
            setattr(self, name, btn); top_lay.addWidget(btn,0); btn.clicked.connect(slot)
        top_lay.addStretch(1)
        self.central=QWidget(self); self.setCentralWidget(self.central)
        vmain=QVBoxLayout(self.central); vmain.setContentsMargins(0,0,0,0); vmain.setSpacing(ROW_SPACING)
        vmain.addWidget(top,0); self.grid=QGridLayout(); vmain.addLayout(self.grid,1)
        self.named_bookmarks=migrate_legacy_favorites_into_named(load_named_bookmarks()); save_named_bookmarks(self.named_bookmarks)
        self._clipboard=None; self._bm_dlg=None
        self.file_ops = FileOperationManager(self)
        self._update_layout_icon(); self._update_theme_icon()
        self._help_shortcut = QShortcut(QKeySequence("F1"), self)
        self._help_shortcut.setContext(Qt.ApplicationShortcut)
        self._help_shortcut.activated.connect(self._show_shortcuts)
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
        # Geometry is tracked from actual screen/DPI events. Merely activating the
        # application never forces a resize, which avoids mixed-DPI focus flicker.
        self._stable_normal_geometry = QtCore.QRect(self.geometry())
        self._stable_window_state = self.windowState()
        self._geometry_restore_guard = False
        self._screen_transition_active = False
        self._screen_transition_generation = 0
        self._pre_transition_geometry = QtCore.QRect()
        self._connected_window_handle = None
        self._connected_screen = None
        self._normal_geometry_by_screen = {}
        QTimer.singleShot(0, self._setup_screen_tracking)

    def mark_active_pane(self, pane):
        try:
            previous = getattr(self, "_active_pane", None)
            if previous is pane:
                return
            self._active_pane = pane
            for p in (previous, pane):
                if p is None:
                    continue
                try:
                    is_active = (p is pane)
                    p.path_bar.set_active(is_active)
                    if hasattr(p, "set_active_visual"):
                        p.set_active_visual(is_active)
                except Exception:
                    pass
            try:
                if pane and hasattr(pane, "view") and hasattr(pane.view, "ensure_drag_ready"):
                    pane.view.ensure_drag_ready()
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
        for name, icon_fn in (
            ("btn_theme", icon_theme_toggle),
            ("btn_bm_edit", icon_bookmark_edit),
            ("btn_session", icon_session),
            ("btn_shortcuts", icon_shortcuts),
            ("btn_about", icon_info),
        ):
            btn = getattr(self, name, None)
            if btn:
                btn.setIcon(icon_fn(self.theme))

    def _update_theme_dependent_icons(self):
        self._update_layout_icon(); self._update_theme_icon()
        for p in getattr(self,"panes",[]):
            try:
                p.btn_star.setIcon(icon_star(p.btn_star.isChecked(), self.theme))
                p.btn_cmd.setIcon(icon_cmd(self.theme))
                p.btn_explorer.setIcon(icon_explorer(self.theme))

                if getattr(getattr(p, "path_bar", None), "_btn_copy", None):
                    p.path_bar._btn_copy.setIcon(icon_copy_squares(self.theme))
            except Exception:
                pass

    @QtCore.pyqtSlot(str)
    def show_operation_status(self, text: str):
        try:
            self.statusBar().showMessage(str(text), 2000)
        except Exception:
            pass

    def _cycle_layout(self):
        if getattr(self, "file_ops", None) and self.file_ops.is_busy():
            self.flash_status("Finish or cancel the file operation before changing the layout")
            QMessageBox.information(
                self,
                "File operation in progress",
                "The pane layout cannot be changed while a copy, move, or delete operation is running.",
            )
            return
        next_idx = (self._layout_idx + 1) % len(self._layout_states)
        n = self._layout_states[next_idx]
        if self.build_panes(n, self._current_paths()):
            self._layout_idx = next_idx
            self._update_layout_icon()

    def _apply_theme(self, theme_name: str, persist: bool = True):
        if theme_name not in VALID_THEMES:
            theme_name = "dark"
        self.theme = theme_name
        app = QApplication.instance()
        if app:
            apply_theme_by_name(app, self.theme)
        self._update_theme_dependent_icons()
        if persist:
            s=QSettings(ORG_NAME, APP_NAME); s.setValue("ui/theme", self.theme); s.sync()

    def _toggle_theme(self):
        self._apply_theme("light" if self.theme == "dark" else "dark", persist=True)

    def build_panes(self, n:int, start_paths):
        if getattr(self, "panes", None) and getattr(self, "file_ops", None) and self.file_ops.has_pending():
            self.flash_status("Finish or cancel the file operation before rebuilding panes")
            return False
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
        shutdown_failed = False
        for p in old_panes:
            try:
                if not p.shutdown(wait_ms=2000):
                    shutdown_failed = True
            except Exception:
                shutdown_failed = True
        if shutdown_failed:
            QMessageBox.warning(
                self,
                "Could not change layout safely",
                "A background folder, search, or icon task has not stopped yet. "
                "The existing panes were kept to prevent a thread-lifecycle crash.",
            )
            return False

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
        return True


    def _unmax_then_remax(self):
        # Rebuilding the pane grid does not require changing the native window state.
        # showNormal()/showMaximized() can repeatedly trigger DPI and resize messages on
        # mixed-DPI multi-monitor systems, so keep the current state unchanged.
        self._kick_layout()

    def _kick_layout(self):
        try:
            cw = self.centralWidget()
            lay = cw.layout() if cw else None

            if lay:
                lay.invalidate()
                lay.activate()

            if hasattr(self, "grid"):
                self.grid.invalidate()
                self.grid.activate()

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

            if cw:
                cw.updateGeometry()
                cw.update()
        except Exception:
            pass

    def _is_normal_window_state(self) -> bool:
        try:
            state = self.windowState()
            return not bool(state & (Qt.WindowMaximized | Qt.WindowMinimized | Qt.WindowFullScreen))
        except Exception:
            return not (self.isMaximized() or self.isMinimized() or self.isFullScreen())

    def _screen_key(self, screen=None) -> str:
        try:
            screen = screen or (self.windowHandle().screen() if self.windowHandle() else None)
            if screen is None:
                return ""
            return f"{screen.name()}|{screen.logicalDotsPerInchX():.1f}|{screen.logicalDotsPerInchY():.1f}"
        except Exception:
            return ""

    def _screen_for_rect(self, rect: QtCore.QRect):
        app = QApplication.instance()
        if app is None:
            return None
        try:
            screen = app.screenAt(rect.center())
            if screen is not None:
                return screen
        except Exception:
            pass
        best = None
        best_area = -1
        try:
            for screen in app.screens():
                inter = screen.availableGeometry().intersected(rect)
                area = max(0, inter.width()) * max(0, inter.height())
                if area > best_area:
                    best, best_area = screen, area
        except Exception:
            pass
        return best or app.primaryScreen()

    def _disconnect_screen_signals(self):
        screen = getattr(self, "_connected_screen", None)
        if screen is None:
            return
        for signal_name in ("availableGeometryChanged", "geometryChanged", "logicalDotsPerInchChanged"):
            try:
                getattr(screen, signal_name).disconnect(self._on_screen_metrics_changed)
            except Exception:
                pass
        self._connected_screen = None

    def _connect_screen_signals(self, screen):
        if screen is getattr(self, "_connected_screen", None):
            return
        self._disconnect_screen_signals()
        self._connected_screen = screen
        if screen is None:
            return
        for signal_name in ("availableGeometryChanged", "geometryChanged", "logicalDotsPerInchChanged"):
            try:
                getattr(screen, signal_name).connect(self._on_screen_metrics_changed)
            except Exception:
                pass

    def _setup_screen_tracking(self):
        win = self.windowHandle()
        if win is None:
            QTimer.singleShot(50, self._setup_screen_tracking)
            return
        if win is not getattr(self, "_connected_window_handle", None):
            old = getattr(self, "_connected_window_handle", None)
            if old is not None:
                try:
                    old.screenChanged.disconnect(self._on_window_screen_changed)
                except Exception:
                    pass
            self._connected_window_handle = win
            try:
                win.screenChanged.connect(self._on_window_screen_changed)
            except Exception:
                pass
        self._connect_screen_signals(win.screen())
        self._remember_stable_window_geometry(force=True)

    def _remember_stable_window_geometry(self, force: bool = False):
        if getattr(self, "_geometry_restore_guard", False) or getattr(self, "_screen_transition_active", False):
            return
        if not self._is_normal_window_state():
            self._stable_window_state = self.windowState()
            return
        if not force and not self.isActiveWindow():
            return
        geom = self.geometry()
        if geom.isValid() and geom.width() >= 300 and geom.height() >= 200:
            self._stable_normal_geometry = QtCore.QRect(geom)
            self._stable_window_state = self.windowState()
            key = self._screen_key()
            if key:
                self._normal_geometry_by_screen[key] = QtCore.QRect(geom)

    def _begin_screen_transition(self):
        if not hasattr(self, "_screen_transition_generation"):
            return
        if not self._screen_transition_active and self._is_normal_window_state():
            candidate = getattr(self, "_stable_normal_geometry", QtCore.QRect())
            if not candidate.isValid():
                candidate = self.geometry()
            self._pre_transition_geometry = QtCore.QRect(candidate)
        self._screen_transition_generation += 1
        generation = self._screen_transition_generation
        self._screen_transition_active = True
        # A single settling callback is tied to a real screen/DPI event, not focus.
        QTimer.singleShot(90, lambda g=generation: self._finish_screen_transition(g))

    @QtCore.pyqtSlot(object)
    def _on_window_screen_changed(self, screen):
        self._connect_screen_signals(screen)
        self._begin_screen_transition()

    def _on_screen_metrics_changed(self, *_args):
        self._begin_screen_transition()

    def _finish_screen_transition(self, generation: int):
        if generation != getattr(self, "_screen_transition_generation", -1):
            return
        try:
            if self._is_normal_window_state():
                screen = self.windowHandle().screen() if self.windowHandle() else self._screen_for_rect(self.geometry())
                if screen is not None:
                    avail = screen.availableGeometry()
                    geom = self.geometry()
                    before = getattr(self, "_pre_transition_geometry", QtCore.QRect())
                    # Preserve the logical size that existed before the actual DPI/screen
                    # transition. Do not interfere while the user is actively dragging.
                    dragging = bool(QApplication.mouseButtons() & Qt.LeftButton)
                    desired_w = geom.width() if dragging or not before.isValid() else before.width()
                    desired_h = geom.height() if dragging or not before.isValid() else before.height()
                    width = min(max(600, desired_w), avail.width())
                    height = min(max(350, desired_h), avail.height())
                    x = min(max(avail.left(), geom.x()), avail.right() - width + 1)
                    y = min(max(avail.top(), geom.y()), avail.bottom() - height + 1)
                    corrected = QtCore.QRect(x, y, width, height)
                    if corrected != geom:
                        self._geometry_restore_guard = True
                        self.setGeometry(corrected)
        finally:
            self._geometry_restore_guard = False
            self._screen_transition_active = False
            self._pre_transition_geometry = QtCore.QRect()
            self._remember_stable_window_geometry(force=True)

    def changeEvent(self, event):
        # Activation alone must never resize the window. Only remember the last
        # good geometry when leaving the application.
        if (hasattr(self, "_normal_geometry_by_screen")
            and event.type() == QEvent.ActivationChange
            and not self.isActiveWindow()):
            self._remember_stable_window_geometry(force=True)
        super().changeEvent(event)

    def nativeEvent(self, event_type, message):
        if sys.platform == "win32":
            try:
                msg = ctypes.cast(int(message), ctypes.POINTER(_MSG)).contents
                if msg.message == 0x02E0:  # WM_DPICHANGED
                    self._begin_screen_transition()
            except Exception:
                pass
        return super().nativeEvent(event_type, message)

    def moveEvent(self, event):
        super().moveEvent(event)
        if not getattr(self, "_geometry_restore_guard", False) and not getattr(self, "_screen_transition_active", False):
            if hasattr(self, "_stable_normal_geometry"):
                QTimer.singleShot(0, self._remember_stable_window_geometry)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not getattr(self, "_geometry_restore_guard", False) and not getattr(self, "_screen_transition_active", False):
            if hasattr(self, "_stable_normal_geometry"):
                QTimer.singleShot(0, self._remember_stable_window_geometry)

    def _safe_restore_geometry(self, ba: QtCore.QByteArray):
        try:
            ok = self.restoreGeometry(ba)
            if not ok:
                return


            g = self.geometry()
            target_screen = self._screen_for_rect(g)
            if target_screen is None:
                return
            sg = target_screen.availableGeometry()

            new_w = min(max(600, g.width()), sg.width())
            new_h = min(max(350, g.height()), sg.height())
            new_x = min(max(sg.left(), g.x()), sg.right() - new_w)
            new_y = min(max(sg.top(),  g.y()), sg.bottom() - new_h)


            if (new_w != g.width()) or (new_h != g.height()) or (new_x != g.x()) or (new_y != g.y()):
                self.setGeometry(new_x, new_y, new_w, new_h)
        except Exception:
            pass


    def _current_paths(self): return [p.current_path() for p in self.panes]
    def set_clipboard(self,payload:dict):
        payload = _normalize_file_clipboard_payload(payload)
        if sys.platform == "win32":
            if payload and _write_windows_file_clipboard_payload(payload):
                self._clipboard = _read_windows_file_clipboard_payload() or payload
            else:
                self._clipboard = None
        else:
            self._clipboard = payload
        return bool(self._clipboard)
    def get_clipboard(self):
        if sys.platform == "win32":
            self._clipboard = _read_windows_file_clipboard_payload()
        return self._clipboard
    def clear_clipboard(self):
        if sys.platform == "win32":
            _clear_windows_clipboard()
        self._clipboard=None
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
            if len(self.named_bookmarks)>=BOOKMARK_LIMIT and all(x.get("enabled") for x in self.named_bookmarks):
                QMessageBox.information(self,"Bookmarks",f"Bookmark limit reached ({BOOKMARK_LIMIT}). Please edit bookmarks to free a slot.")
                self._open_bookmark_editor(); return
            reused=False
            for i,it in enumerate(self.named_bookmarks):
                if not it.get("enabled") and not it.get("name") and not it.get("path"):
                    self.named_bookmarks[i]={"enabled":True,"name":_derive_name_from_path(np),"path":np}; reused=True; break
            if not reused: self.named_bookmarks.append({"enabled":True,"name":_derive_name_from_path(np),"path":np})
            if len(self.named_bookmarks)>BOOKMARK_LIMIT: self.named_bookmarks=self.named_bookmarks[:BOOKMARK_LIMIT]
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
            for it in new_items[:BOOKMARK_LIMIT]:
                cleaned.append({"name":it.get("name","").strip(),"path":it.get("path","").strip(),"enabled":bool(it.get("enabled",False))})
            self.named_bookmarks=cleaned[:BOOKMARK_LIMIT]; save_named_bookmarks(self.named_bookmarks); self.namedBookmarksChanged.emit(self.named_bookmarks)

    def _on_bmdlg_closed(self,*_):
        try:
            if getattr(self,"_bm_dlg",None): self.namedBookmarksChanged.disconnect(self._bm_dlg.set_items)
        except Exception: pass
        self._bm_dlg=None

    def _show_shortcuts(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Keyboard Shortcuts")
        dlg.resize(780, 460)
        lay = QVBoxLayout(dlg)
        lbl = QLabel("The same shortcut descriptions documented in README.md", dlg)
        lay.addWidget(lbl)
        table = _setup_readonly_table(
            QTableWidget(dlg), ["Key", "Action"], [QHeaderView.ResizeToContents, QHeaderView.Stretch], row_count=len(KEYBOARD_SHORTCUTS)
        )
        for r, (key_text, action_text) in enumerate(KEYBOARD_SHORTCUTS):
            _set_table_row_items(table, r, key_text, action_text)
        lay.addWidget(table, 1)
        _add_dialog_button_box(lay, dlg, QDialogButtonBox.Ok, dlg.accept)
        dlg.exec_()

    def _show_about(self):
        dlg=QDialog(self); dlg.setWindowTitle("About")
        dlg.setWindowIcon(QIcon(app_resource_path(ABOUT_IMAGE_FILENAME)))
        lay=QVBoxLayout(dlg)
        lbl=QLabel(dlg); lbl.setTextFormat(Qt.RichText)
        lbl.setText(
            f"<div style='color:#000; font-size:12pt;'><b>Multi-Pane File Explorer v{APP_VERSION}</b></div>"
            "<div style='color:#111; margin-top:6px;'>A compact multi-pane file explorer for Windows (PyQt5).</div>"
            "<div style='color:#111; margin-top:6px;'>For feedback, contact <b>kkongt2.kang</b>.</div>"
        )
        lay.addWidget(lbl)

        about_pixmap = QPixmap(app_resource_path(ABOUT_IMAGE_FILENAME))
        if not about_pixmap.isNull():
            img_lbl = QLabel(dlg)
            img_lbl.setAlignment(Qt.AlignCenter)
            img_lbl.setPixmap(about_pixmap.scaled(96, 96, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            opacity = QGraphicsOpacityEffect(img_lbl)
            opacity.setOpacity(0.0)
            img_lbl.setGraphicsEffect(opacity)
            lay.addWidget(img_lbl)

            fade_in = QPropertyAnimation(opacity, b"opacity", dlg)
            fade_in.setDuration(5000)
            fade_in.setStartValue(0.0)
            fade_in.setEndValue(1.0)

            hold = QPauseAnimation(3000, dlg)

            fade_out = QPropertyAnimation(opacity, b"opacity", dlg)
            fade_out.setDuration(5000)
            fade_out.setStartValue(1.0)
            fade_out.setEndValue(0.0)

            animation = QSequentialAnimationGroup(dlg)
            animation.addAnimation(fade_in)
            animation.addAnimation(hold)
            animation.addAnimation(fade_out)
            animation.setLoopCount(-1)
            dlg._about_animation = animation
            animation.start()

        _add_dialog_button_box(lay, dlg, QDialogButtonBox.Ok, dlg.accept)
        _apply_palette_colors(dlg, {QPalette.Window: (255, 255, 255), QPalette.WindowText: (0, 0, 0)})
        dlg.setStyleSheet("QLabel { color: #000; } QDialog { background: #FFF; }")
        dlg.resize(380,290); dlg.exec_()

    def closeEvent(self, e):
        manager = getattr(self, "file_ops", None)
        if manager and manager.has_pending():
            answer = QMessageBox.question(
                self,
                "File operation in progress",
                "A copy, move, or delete operation is still running or queued.\n\n"
                "Cancel the operation and exit?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                e.ignore()
                return
            if not manager.cancel_all(wait_ms=8000):
                QMessageBox.warning(
                    self,
                    "Could not exit safely",
                    "The file operation has not stopped yet. The window will remain open to prevent data corruption.",
                )
                e.ignore()
                return
            try:
                QApplication.processEvents(QtCore.QEventLoop.AllEvents, 100)
            except Exception:
                pass

        paths = []
        try:
            paths = self._current_paths()
        except Exception:
            paths = []

        pane_shutdown_failed = False
        for pane in list(getattr(self, "panes", [])):
            try:
                if not pane.shutdown(wait_ms=3000):
                    pane_shutdown_failed = True
            except Exception:
                pane_shutdown_failed = True
        if pane_shutdown_failed:
            QMessageBox.warning(
                self,
                "Could not exit safely",
                "A background folder, search, or icon task has not stopped yet. "
                "The window will remain open to prevent a thread-lifecycle crash.",
            )
            e.ignore()
            return

        settings = QSettings(ORG_NAME, APP_NAME)
        settings.setValue("window/geometry", self.saveGeometry())
        settings.setValue("layout/pane_count", len(paths) if paths else len(self.panes))
        for i, path in enumerate(paths if paths else [x.current_path() for x in self.panes]):
            settings.setValue(f"layout/pane_{i}_path", path)
        settings.sync()
        super().closeEvent(e)



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
            if not self.build_panes(panes, paths):
                return
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
        SessionManagerDialog(self, self._get_sessions()).exec_()

class SessionManagerDialog(QDialog):
    def __init__(self, parent: MultiExplorer, sessions: list):
        super().__init__(parent)
        self.setWindowTitle("Session Manager")
        self.resize(560, 400)

        self.table = _setup_readonly_table(
            QTableWidget(self),
            ["Name", "Panes", "Saved"],
            [QHeaderView.Stretch, QHeaderView.ResizeToContents, QHeaderView.ResizeToContents],
        )

        self.btn_save, self.btn_load, self.btn_delete, self.btn_close = [QPushButton(text, self) for text in (
            "Save Current", "Load Selected", "Delete Selected", "Close"
        )]
        btns = QHBoxLayout()
        for btn in (self.btn_save, self.btn_load, self.btn_delete): btns.addWidget(btn)
        btns.addStretch(1)
        btns.addWidget(self.btn_close)
        lay = QVBoxLayout(self)
        lay.addWidget(self.table, 1)
        lay.addLayout(btns)
        self._sessions = []
        self.set_sessions(sessions)
        for btn, slot in ((self.btn_close, self.accept), (self.btn_load, self._on_load), (self.btn_delete, self._on_delete), (self.btn_save, self._on_save)):
            btn.clicked.connect(slot)

    def set_sessions(self, items: list):
        self._sessions = list(items or [])
        self.table.setRowCount(len(self._sessions))
        for r, it in enumerate(self._sessions):
            name = it.get("name","")
            panes = int(it.get("panes", 0))
            ts = float(it.get("ts", time.time()))
            dt = QDateTime.fromSecsSinceEpoch(int(ts)).toString("yyyy-MM-dd HH:mm:ss")
            _set_table_row_items(self.table, r, name, panes, dt)
        self.table.resizeColumnsToContents()

    def _selected_name(self) -> str | None:
        rows = self.table.selectionModel().selectedRows()
        it = self.table.item(rows[0].row(), 0) if rows else None
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
        self.setWindowTitle(f"Edit Bookmarks (max {BOOKMARK_LIMIT})")
        self.resize(760, 520)
        self.table = _setup_readonly_table(
            QTableWidget(self),
            ["Enabled", "Name", "Path"],
            [QHeaderView.ResizeToContents, QHeaderView.ResizeToContents, QHeaderView.Stretch],
        )
        self.table.setRowCount(BOOKMARK_LIMIT)
        self._rows = []
        lay = QVBoxLayout(self); lay.addWidget(self.table, 1)
        _add_dialog_button_box(lay, self, QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self.accept, self.reject)
        items = list(items or [])
        for i in range(BOOKMARK_LIMIT):
            it = items[i] if i < len(items) else _empty_bookmark_item()
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
        return [
            {"enabled": enabled, "name": name, "path": path}
            for chk, name_edit, path_edit in self._rows
            for enabled, name, path in [(chk.isChecked(), name_edit.text().strip(), path_edit.text().strip())]
            if name or path or enabled
        ]

    def set_items(self, items: list):
        items = list(items or [])
        for r in range(BOOKMARK_LIMIT):
            it = items[r] if r < len(items) else _empty_bookmark_item()
            chk, name_edit, path_edit = self._rows[r]
            chk.setChecked(bool(it.get("enabled", False)))
            name_edit.setText(str(it.get("name", "")))
            path_edit.setText(str(it.get("path", "")))
        self.table.resizeColumnsToContents()


def _load_start_paths(desired_panes:int, cli_paths):
    s=QSettings(ORG_NAME, APP_NAME); cli_paths=list(cli_paths or []); paths=[]
    for i in range(desired_panes):
        if i<len(cli_paths) and os.path.exists(cli_paths[i]):
            paths.append(cli_paths[i]); continue
        p=s.value(f"layout/pane_{i}_path", QDir.homePath(), type=str)
        paths.append(p if p and os.path.exists(p) else QDir.homePath())
    return paths

def _resolve_pane_count(cli_panes, settings=None) -> int:
    if cli_panes in (4, 6, 8):
        return int(cli_panes)
    settings = settings or QSettings(ORG_NAME, APP_NAME)
    try:
        saved = int(settings.value("layout/pane_count", 6))
    except (TypeError, ValueError):
        saved = 6
    return saved if saved in (4, 6, 8) else 6

def parse_args(argv=None):
    ap=argparse.ArgumentParser(description="Multi-Pane File Explorer (PyQt5)")
    ap.add_argument("paths", nargs="*", help="Optional start paths per pane")
    ap.add_argument(
        "--panes", type=int, choices=[4,6,8], default=None,
        help="Number of panes: 4, 6 or 8 (default: restore the last layout)",
    )
    ap.add_argument("--debug", action="store_true", help="Enable debug logs (or set MULTIPANE_DEBUG=1)")
    return ap.parse_args(argv)

def main():
    global DEBUG
    args=parse_args()
    DEBUG = bool(args.debug or _env_flag("MULTIPANE_DEBUG"))
    _enable_win_per_monitor_v2()
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    # Qt requires this policy before QApplication is created. Applying it afterwards
    # can leave monitor-DPI transitions in an inconsistent native-window state.
    try:
        if hasattr(QGuiApplication, "setHighDpiScaleFactorRoundingPolicy"):
            QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
                Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
            )
    except Exception:
        pass
    app=QApplication(sys.argv)
    app.setWindowIcon(QIcon(app_resource_path(APP_ICON_FILENAME)))
    base_font=QFont("Segoe UI"); base_font.setPointSizeF(FONT_PT); app.setFont(base_font)
    app.setOrganizationName(ORG_NAME); app.setApplicationName(APP_NAME); app.setApplicationVersion(APP_VERSION)
    settings=QSettings(ORG_NAME, APP_NAME); theme=settings.value("ui/theme","dark")
    if theme not in VALID_THEMES: theme="dark"
    apply_theme_by_name(app, theme)
    pane_count=_resolve_pane_count(args.panes, settings)
    start_paths=_load_start_paths(pane_count, args.paths)
    w=MultiExplorer(pane_count=pane_count, start_paths=start_paths, initial_theme=theme); w.show()
    sys.exit(app.exec_())
if __name__=="__main__":
    main()
