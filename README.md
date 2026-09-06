# Multi-Pane File Explorer (PyQt5)

A multi-pane file explorer for Windows.
This README reflects the current behavior of `multipane_explorer.py`.

## Requirements
- Windows 10/11 (recommended)
- Python 3.10+
- PyQt5
- Optional: `pywin32`, `send2trash`

## Features
- 4/6/8 pane layout switching (top toolbar + `--panes`), with last layout/path restore
- Per-pane back/forward/up navigation history
- Path-bar edit mode with recent-path dropdown and background folder path autocomplete; stale suggestions are discarded when input changes
- Folders-first sorting, with proper size/date sorting for files
- Large-folder optimization: incremental `os.scandir` listing with a visible large-folder mode badge; directory snapshots and metadata queries are shared across panes and sort modes
- Auto-refresh via `QFileSystemWatcher` keeps the existing listing visible and applies changed rows after a successful scan, preserving selections for surviving files
- Background path validation, paste conflict checks, and free-space queries keep navigation responsive; cancelled navigation results are discarded
- Failed metadata queries back off and stop after three attempts until the listing is refreshed
- Filter/recursive search (wildcards like `*.txt`, `*report*.xlsx`, multi-pattern support)
- Copy/move/paste + drag-and-drop, with conflict actions: `Overwrite / Skip / Copy`
- Safe cross-filesystem moves use a completed staging copy before source cleanup; source changes detected during copying are preserved and reported, and cleanup never recursively removes newly added entries
- Same-filesystem moves skip descendant progress scans; folder copies process entries as they are enumerated and close enumeration handles on cancellation
- Copy/move promotion refuses concurrent destination conflicts; if restoring an overwritten destination is blocked, its backup is retained and its recovery path is reported
- Symbolic links are preserved when permitted, while cross-filesystem Windows junction copies are refused rather than traversed
- Bulk rename tool (prefix/suffix/find-replace/numbering) via `Ctrl+Shift+R`
- Per-pane file operation progress bar with cancellation
- Background undo with progress and cancellation; completed portions of cancelled operations remain undoable, and unfinished undo items can be retried
- Delete to Recycle Bin (`send2trash`/Shell API when available; no permanent fallback), `Shift+Delete` for permanent delete
- Two-row quick bookmark toolbar (top row first, up to 30 bookmarks with overflow menu), with a compact 1×1 star button, blank space below it, and more horizontal room for bookmarks; supports drag-to-reorder bookmark editing
- Session save/load/delete (pane count + pane paths)
- Dark/light theme toggle and active-pane highlighting
- Native Explorer context menu when `pywin32` is available, fallback menu otherwise
- Open Command Prompt in the current folder
- Customize command buttons (v2.9.0): top-left settings enable 0 (default), 2, 4, or 6 shared buttons between the CMD and Explorer columns; each uses a default icon or A-Z icon and runs a saved CMD command in the clicked pane's folder without showing a console. Settings persist, including disabled slots. Console output can be redirected to a file.

## Install
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install PyQt5 pywin32 send2trash
```

`pywin32` and `send2trash` are optional but recommended for native context-menu integration and reliable Recycle Bin behavior.

## Run
```powershell
python multipane_explorer.py [--panes 4|6|8] [--debug] [start_path1 start_path2 ...]
```

Examples:
```powershell
python multipane_explorer.py --panes 6
python multipane_explorer.py --panes 6 --debug
python multipane_explorer.py --panes 4 "C:\Windows" "D:\WS" "C:\Users\USER" "C:\Temp"
```

Enable debug logs with environment variable:
```powershell
$env:MULTIPANE_DEBUG=1; python multipane_explorer.py
```

Normal copies rely on the operating system's buffered writeback for throughput.
To force every copied file to stable storage before it is promoted into place:
```powershell
$env:MULTIPANE_DURABLE_COPIES=1; python multipane_explorer.py
```

## Test
```powershell
python -m unittest discover -s tests -v
```

The suite includes temporary-directory recovery checks and isolated, offscreen Qt
checks for responsive autocomplete and cancellable undo.

## Code layout
- `multipane_explorer.py`: application entry point, panes, models, search, path suggestions, settings, and dialogs
- `file_operations.py`: file transactions, source validation, rollback, bulk rename, Recycle Bin handling, operation queue, and copy/move/delete/undo workers; no explorer-widget dependency
- `tests/test_operation_recovery.py`: source changes, destination conflicts, cancellation, rename rollback, and partial undo
- `tests/test_async_operations.py`: Qt responsiveness and application startup checks in isolated subprocesses

## Search/Filter Behavior
- Type a filter and press `Enter` (or click `Search`) to run recursive search from the current folder
- While search is running, the same button becomes `Cancel`
- Pattern separators: space, `,`, `;` (OR matching)
- Press `Esc` in the filter input to clear filter and return to browse mode
- Search result cap: 50,000 items

## Keyboard Shortcuts
| Key | Action |
|---|---|
| `F1` | Open shortcuts help dialog |
| `Backspace` / `Alt+Left` | Back |
| `Alt+Right` | Forward |
| `Alt+Up` | Parent folder |
| `Enter` / `Ctrl+O` | Open |
| `Ctrl+L` / `F4` | Edit path |
| `Ctrl+F` / `F3` | Focus filter |
| `Esc` (in filter input) | Clear filter + return to browse mode |
| `F5` | Hard refresh |
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copy / Cut / Paste |
| `Ctrl+Z` | Undo (new folder / rename / safe copy-move actions) |
| `Delete` / `Shift+Delete` | Recycle / Permanent delete |
| `F2` | Rename |
| `Ctrl+Shift+R` | Bulk rename |
| `Ctrl+Shift+C` | Copy full path |
| `Ctrl+Shift+D` | Open the selected item's containing folder |
| `Alt+Shift+C` | Copy folder path (parent folder if a file is selected) |

## Settings
Uses `QSettings`:
- Organization: `MultiPane`
- Application: `Multi-Pane File Explorer`

Stored values include:
- Window geometry
- Theme
- Pane count and per-pane last path
- Bookmarks and sessions
- Per-pane sort column and sort order
- Per-pane search result column widths

## Build (optional)
Install the build requirements, then build from the checked-in PyInstaller spec:

```powershell
pip install -r requirements-build.txt
python -m PyInstaller --clean --noconfirm MultiPaneExplorer.spec
```

On Windows, you can run `makeExe.bat` to execute the same one-file build. The executable and taskbar use `images-5.ico`; the About dialog uses `images-6.ico`. The executable is written to `dist\MultiPaneExplorer.exe`.

To download a prebuilt executable from GitHub, open the **Build Windows EXE** workflow run for this branch or pull request, then download the `MultiPaneExplorer-windows-exe` artifact.

## License
MIT
