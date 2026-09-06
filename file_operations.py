"""File transactions, recovery, deletion and Qt workers, independent of explorer widgets.

Filesystem mutations belong here. Workers publish results through signals; callers
own dialogs, pane refreshes and undo history. Private staging paths are the only
paths that failure cleanup may remove. Conflicting destination backups are retained
and reported when restoration cannot claim the original name.
"""

import copy
import ctypes
import errno
import os
import shutil
import stat
import sys
import tempfile
import time
import uuid
from pathlib import Path

from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSignal

DEBUG = os.environ.get("MULTIPANE_DEBUG", "").strip().lower() in {"1", "true", "yes", "on", "y"}
DURABLE_FILE_COPIES = os.environ.get("MULTIPANE_DURABLE_COPIES", "").strip().lower() in {"1", "true", "yes", "on", "y"}
FILEOP_ERROR_DETAIL_LIMIT = 50
FILEOP_FAST_PROGRESS_SCAN_LIMIT = 4000
FILE_COPY_BUFFER_SIZE = 4 * 1024 * 1024

try:
    import pythoncom
    from win32com.shell import shell, shellcon
    HAS_PYWIN32 = True
except Exception:
    HAS_PYWIN32 = False

try:
    from send2trash import send2trash
    HAS_SEND2TRASH = True
except Exception:
    HAS_SEND2TRASH = False


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


def execute_bulk_rename_transaction(operations, should_cancel=lambda: False) -> list[tuple[str, str]]:
    """Stage a rename group and restore original names on recoverable failures."""
    operations = list(operations)
    for column in (0, 1):
        keys = [_path_key(pair[column]) for pair in operations]
        if len(set(keys)) != len(keys):
            raise ValueError("Duplicate source or destination in rename plan")
    temp_pairs = []
    committed = []
    try:
        for src, dst in operations:
            if should_cancel():
                raise DeleteCancelled("Operation cancelled.")
            parent = os.path.dirname(src) or os.curdir
            temp = os.path.join(parent, f".__mprn_tmp_{uuid.uuid4().hex}")
            while os.path.lexists(temp):
                temp = os.path.join(parent, f".__mprn_tmp_{uuid.uuid4().hex}")
            _rename_no_replace(src, temp)
            temp_pairs.append((src, temp, dst))

        for src, temp, dst in temp_pairs:
            if should_cancel():
                raise DeleteCancelled("Operation cancelled.")
            _rename_no_replace(temp, dst)
            committed.append((dst, src))
        return committed
    except Exception as original_error:
        rollback_errors = []

        # Clear every committed name before restoring originals: final names can
        # occupy another item's original name in chains and cycles.
        temporary_names = {_path_key(src): temp for src, temp, _dst in temp_pairs}
        for dst, src in reversed(committed):
            if not os.path.lexists(dst):
                continue
            try:
                _rename_no_replace(dst, temporary_names[_path_key(src)])
            except Exception as exc:
                rollback_errors.append(f"{dst} -> {src}: {exc}")

        for src, temp, _dst in reversed(temp_pairs):
            if not os.path.lexists(temp):
                continue
            try:
                _rename_no_replace(temp, src)
            except Exception as exc:
                rollback_errors.append(f"{temp} -> {src}: {exc}")

        if rollback_errors:
            details = "\n".join(rollback_errors[:10])
            more = "\n..." if len(rollback_errors) > 10 else ""
            raise RuntimeError(
                f"{original_error}\n\nRollback was incomplete:\n{details}{more}"
            ) from original_error
        raise


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
        app = QtCore.QCoreApplication.instance()
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


def _rename_no_replace(src, dst):
    """Atomically claim a destination, refusing a concurrent occupant."""
    if os.name == "nt":
        os.rename(src, dst)
        return
    if sys.platform.startswith("linux"):
        libc = ctypes.CDLL(None, use_errno=True)
        rename = getattr(libc, "renameat2", None)
        if rename is not None:
            rename.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
            rename.restype = ctypes.c_int
            if rename(-100, os.fsencode(src), -100, os.fsencode(dst), 1) == 0:
                return
            code = ctypes.get_errno()
            raise OSError(code, os.strerror(code), dst)
    raise OSError("Atomic non-overwriting rename is unavailable on this platform")


def _source_signature(path):
    info = os.lstat(path)
    return (info.st_dev, info.st_ino, info.st_mode, info.st_size,
            info.st_mtime_ns, info.st_ctime_ns)


def _snapshot_move_source(root, should_cancel=lambda: False):
    """Record exactly the tree whose contents are about to be copied."""
    snapshot = {}
    pending = [root]
    while pending:
        if should_cancel():
            raise DeleteCancelled()
        path = pending.pop()
        if _is_junction(path):
            raise OSError(f"Cannot safely copy a directory junction: {path}")
        signature = _source_signature(path)
        snapshot[path] = signature
        if stat.S_ISDIR(signature[2]):
            with os.scandir(path) as entries:
                pending.extend(entry.path for entry in entries)
    return snapshot


def _cleanup_copied_source(root, snapshot, should_cancel=lambda: False):
    """Delete only unchanged copied entries; never recursively delete new data."""
    if _snapshot_move_source(root, should_cancel) != snapshot:
        return ["Source changed during copying; source was left in place."]
    errors = []
    for path, expected in reversed(list(snapshot.items())):
        if should_cancel():
            raise DeleteCancelled()
        try:
            # A replaced ancestor could redirect a child path into another tree.
            parent = os.path.dirname(path)
            while parent in snapshot:
                if _source_signature(parent)[:3] != snapshot[parent][:3]:
                    raise OSError(f"Source directory was replaced: {parent}")
                parent = os.path.dirname(parent)
            current = _source_signature(path)
            if stat.S_ISDIR(expected[2]):
                if current[:3] != expected[:3]:
                    raise OSError("Source directory was replaced")
                os.rmdir(path)  # Refuses any new entries, including late arrivals.
            else:
                if current != expected:
                    raise OSError("Source changed after copying")
                remove_any(path)
        except Exception as exc:
            errors.append(f"{path}: {exc}")
    return errors


class FileOpWorker(QtCore.QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished_ok = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(
        self,
        op: str,
        srcs: list,
        dst_dir: str,
        conflict_map: dict | None = None,
        parent=None,
        durable_copies: bool | None = None,
    ):
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
        self.durable_copies = DURABLE_FILE_COPIES if durable_copies is None else bool(durable_copies)
        self._copy_buffer = None

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
            # A same-filesystem move renames the selected root. Counting its
            # descendants can cost much more than the operation itself.
            if self.op == "move" and _same_filesystem(src, self.dst_dir):
                stats = (0, 1)
            else:
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

    def _backup_destination(self, dst: str, expected=None) -> str:
        if expected is not None and _source_signature(dst) != expected:
            raise OSError("Destination changed while the operation was running")
        backup = self._new_sibling_work_path(dst, "backup")
        _rename_no_replace(dst, backup)
        if expected is not None and _source_signature(backup)[:5] != expected[:5]:
            self._restore_backup(dst, backup)
            raise OSError("Destination was replaced while preparing overwrite")
        return backup

    def _cleanup_path(self, path: str):
        if not path or not os.path.lexists(path):
            return
        remove_any(path)

    def _restore_backup(self, dst: str, backup: str | None):
        if not backup or not os.path.lexists(backup):
            return
        try:
            _rename_no_replace(backup, dst)
        except Exception as exc:
            raise RuntimeError(
                f"Original destination is preserved at {backup}; could not restore {dst}: {exc}"
            ) from exc

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
        except Exception as exc:
            self._record_copy_error(dst, dst, f"Rollback failed: {exc}")

    def _copy_file(self, src, dst):
        copied = 0
        temp_path = None
        try:
            os.makedirs(os.path.dirname(dst) or os.curdir, exist_ok=True)
            temp_path = self._new_sibling_work_path(dst, "partial")
            with open(src, "rb") as fsrc, open(temp_path, "xb") as fdst:
                if self._copy_buffer is None:
                    self._copy_buffer = bytearray(FILE_COPY_BUFFER_SIZE)
                buf = self._copy_buffer
                view = memoryview(buf)
                while True:
                    if self._cancel:
                        return False
                    count = fsrc.readinto(buf)
                    if not count:
                        break
                    fdst.write(view[:count])
                    copied += count
                    self._tick_progress(delta_bytes=count)
                if self.durable_copies:
                    fdst.flush()
                    os.fsync(fdst.fileno())
            try:
                shutil.copystat(src, temp_path, follow_symlinks=True)
            except Exception:
                pass
            if self._cancel:
                return False
            _rename_no_replace(temp_path, dst)
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
            _rename_no_replace(temp_path, dst)
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
            with os.scandir(src_dir) as entries:
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
        except Exception as exc:
            self._record_copy_error(src_dir, dst_dir, exc)
            return False
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
        staging_dir = None
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
        try:
            expected = _source_signature(dst) if existed else None
            staging_dir = tempfile.mkdtemp(prefix=".__mprn_copy_", dir=os.path.dirname(dst) or os.curdir)
            staging = os.path.join(staging_dir, "payload")
            if not self._copy_to_new_path(src, staging) or self._cancel:
                return False
            if existed and action == "overwrite":
                backup = self._backup_destination(dst, expected)
            _rename_no_replace(staging, dst)
            self._discard_backup(backup, dst)
            if self._can_undo_new_destination(existed, action):
                self._remember_created_for_undo(dst)
            return True
        except Exception as exc:
            self._record_copy_error(src, dst, exc)
            self._rollback_destination(dst, backup)
            return False
        finally:
            if staging_dir:
                try:
                    self._cleanup_path(staging_dir)
                except Exception as exc:
                    self._record_copy_error(staging_dir, dst, f"Could not remove staging copy: {exc}")

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

        try:
            expected = _source_signature(dst) if existed else None
        except Exception as exc:
            self._record_copy_error(src, dst, exc)
            return False
        can_undo_move = self._can_undo_new_destination(existed, action)
        src_progress = self._source_progress(src)
        same_filesystem = _same_filesystem(src, os.path.dirname(dst) or self.dst_dir)

        if same_filesystem:
            backup = None
            if existed and action == "overwrite":
                try:
                    backup = self._backup_destination(dst, expected)
                except Exception as exc:
                    self._record_copy_error(src, dst, f"Could not protect existing destination: {exc}")
                    self._skip_source_progress(src)
                    return False
            try:
                _rename_no_replace(src, dst)
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
        staging_dir = tempfile.mkdtemp(prefix=".__mprn_move_", dir=os.path.dirname(dst) or os.curdir)
        staging = os.path.join(staging_dir, "payload")
        try:
            source_snapshot = _snapshot_move_source(src, lambda: self._cancel)
            copied_ok = self._copy_to_new_path(src, staging)
        except Exception as exc:
            copied_ok = False
            self._record_copy_error(src, staging, exc)

        if not copied_ok or self._cancel:
            try:
                self._cleanup_path(staging_dir)
            except Exception as exc:
                self._record_copy_error(staging, dst, f"Could not remove incomplete staging copy: {exc}")
            return False

        backup = None
        try:
            if existed and action == "overwrite":
                backup = self._backup_destination(dst, expected)
            _rename_no_replace(staging, dst)
            staging = None
        except Exception as exc:
            self._record_copy_error(src, dst, f"Could not promote the completed staging copy: {exc}")
            try:
                if staging and os.path.lexists(staging):
                    self._cleanup_path(staging_dir)
            except Exception as cleanup_exc:
                self._record_copy_error(staging, dst, f"Could not remove staging copy: {cleanup_exc}")
            self._rollback_destination(dst, backup)
            return False

        try:
            os.rmdir(staging_dir)
        except OSError as exc:
            self._record_copy_error(staging_dir, dst, f"Could not remove staging directory: {exc}")

        # The destination is now a complete copy.  Never roll it back if source
        # cleanup is cancelled or fails: it may be the only complete copy left.
        self._discard_backup(backup, dst)
        if self._cancel:
            return False

        cleanup_errors = []
        try:
            cleanup_errors = _cleanup_copied_source(
                src, source_snapshot, should_cancel=lambda: self._cancel,
            )
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

class UndoWorker(FileOpWorker):
    """Execute a private undo plan; the UI applies the remaining plan on finish."""
    def __init__(self, action, hwnd=0, parent=None):
        super().__init__("undo", [], "", parent=parent)
        self.remaining_action = copy.deepcopy(action)
        self.hwnd = hwnd
        self.failure_message = ""
        self.completed = False
        self._undo_done = 0
        self._undo_total = max(1, self._count_actions(action))

    @staticmethod
    def _count_actions(action):
        if action.get("type") == "compound":
            return sum(UndoWorker._count_actions(sub) for sub in action.get("actions", []))
        return len(action.get("pairs", action.get("paths", [None])))

    def _emit_progress(self, force=False):
        self.progress.emit(min(100, int(100 * self._undo_done / self._undo_total)))

    def _check_cancel(self):
        if self._cancel:
            raise DeleteCancelled("Operation cancelled.")

    def _apply_action(self, action):
        self._check_cancel()
        kind = action.get("type")
        if kind == "compound":
            while action.get("actions"):
                self._apply_action(action["actions"][-1])
                action["actions"].pop()
            return
        if kind == "mkdir":
            path = action["path"]
            if os.path.lexists(path):
                os.rmdir(path)
        elif kind in {"remove_created", "delete"}:
            # Older undo records must also use the Recycle Bin, never permanent deletion.
            while action.get("paths"):
                self._check_cancel()
                path = action["paths"][-1]
                self.status.emit(f"Undo: {path}")
                if os.path.lexists(path) and not recycle_path_to_trash(path, self.hwnd):
                    raise OSError(f"Could not move {path} to Recycle Bin; item was left in place.")
                action["paths"].pop()
                self._undo_done += 1
                self._emit_progress()
            return
        elif kind == "move_back":
            if action.get("rename_group"):
                pairs = action.get("pairs", [])
                execute_bulk_rename_transaction(pairs, lambda: self._cancel)
                self._undo_done += len(pairs)
                action["pairs"] = []
                self._emit_progress()
                return
            while action.get("pairs"):
                self._check_cancel()
                current, original = action["pairs"][-1]
                target_dir = os.path.dirname(original) or os.curdir
                os.makedirs(target_dir, exist_ok=True)
                if os.path.lexists(original):
                    original = unique_dest_path(target_dir, os.path.basename(original))
                self.status.emit(f"Undo: {current}")
                if not self._move_source_transactional(current, original, None, False):
                    self._check_cancel()
                    raise OSError("\n".join(self.errors) or f"Could not move {current} back to {original}")
                action["pairs"].pop()
                self._undo_done += 1
                self._emit_progress()
            return
        else:
            raise ValueError(f"Unsupported undo action: {kind}")
        self._undo_done += 1
        self._emit_progress()

    def run(self):
        try:
            self._apply_action(self.remaining_action)
            self.completed = True
            self.progress.emit(100)
            self.finished_ok.emit()
        except DeleteCancelled:
            self.failure_message = "Operation cancelled."
        except Exception as exc:
            self.failure_message = str(exc)


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

    def _prepare_progress(self):
        # Counting every descendant first doubles directory I/O and can delay a
        # Recycle Bin rename by minutes.  Progress is therefore based on selected
        # top-level paths while the delete helpers report the exact deleted count.
        self._total = max(1, len(self.paths))
        self._done = 0
        self._last_progress_pct = -1
        self._last_progress_emit_ts = 0.0
        self._emit_progress()

    def _on_path_done(self):
        self._done += 1
        self._emit_progress()

    def run(self):
        coinit = False
        try:
            self.status.emit("Preparing delete ...")
            self._prepare_progress()
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
                        )
                    else:
                        deleted, errors = recycle_any_best_effort(
                            path,
                            hwnd=self.hwnd,
                            should_cancel=lambda: self._cancel,
                        )
                    self.deleted_count += deleted
                    self.errors.extend(errors)
                    self._on_path_done()
                except DeleteCancelled:
                    self.error.emit("Operation cancelled.")
                    return
                except Exception as e:
                    self.errors.append(f"{path}: {e}")
                    self._on_path_done()

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
