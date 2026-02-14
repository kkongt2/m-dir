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
- Folders-first sorting, with proper size/date sorting for files
- Large-folder optimization: fast incremental listing via `os.scandir`, then normal model handoff
- Auto-refresh on file system changes via `QFileSystemWatcher`
- Filter/recursive search (wildcards like `*.txt`, `*report*.xlsx`, multi-pattern support)
- Copy/move/paste + drag-and-drop, with conflict actions: `Overwrite / Skip / Copy`
- File operation progress dialog with cancellation
- Delete to Recycle Bin (`send2trash`/Shell API when available), `Shift+Delete` for permanent delete
- Bookmark editor and quick bookmark buttons (up to 10 bookmarks)
- Session save/load/delete (pane count + pane paths)
- Dark/light theme toggle and active-pane highlighting
- Native Explorer context menu when `pywin32` is available, fallback menu otherwise
- Open Command Prompt in the current folder

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

## Search/Filter Behavior
- Type a filter and press `Enter` (or click `Search`) to run recursive search from the current folder
- Pattern separators: space, `,`, `;` (OR matching)
- Press `Esc` in the filter input to clear filter and return to browse mode
- Search result cap: 50,000 items

## Keyboard Shortcuts
| Key | Action |
|---|---|
| `Backspace` / `Alt+Left` | Back |
| `Alt+Right` | Forward |
| `Alt+Up` | Parent folder |
| `Enter` / `Ctrl+O` | Open |
| `Ctrl+L` / `F4` | Edit path |
| `Ctrl+F` / `F3` | Focus filter |
| `Esc` (in filter input) | Clear filter + return to browse mode |
| `F5` | Hard refresh |
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copy / Cut / Paste |
| `Ctrl+Z` | Undo (currently focused on new folder / rename actions) |
| `Delete` / `Shift+Delete` | Recycle / Permanent delete |
| `F2` | Rename |
| `Ctrl+Shift+C` | Copy full path |
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

## Build (optional)
You can build with PyInstaller using `makeExe.bat` or `MultiPaneExplorer.spec`.

## License
MIT
