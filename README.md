# Multi-Pane File Explorer (PyQt5)

A fast, keyboard-friendly **multi-pane (4/6/8)** file explorer for Windows.  
**by kkongt2 · Built with GPT-5 Thinking**

> Windows 10/11 · Python 3.9–3.12

---

## Features
- **4 / 6 / 8 panes** with instant layout switch; restores last session & paths
- **Folders first, then files** sorting (size/date work as expected)
- **Watcher-based auto refresh** on file system changes
- **Large folder optimization**: `os.scandir` incremental listing + deferred size/icon stat
- **Breadcrumb chips**: only the active pane’s crumbs get a subtle blue highlight
- **Filter / search**: wildcards like `*.txt`, `*report*.xlsx`; **Esc** clears & returns to browse
- **One-click “New”**: folder / empty `.txt` (docx/xlsx/pptx templates when available)
- **Copy/Move conflict resolver**: per-item **Overwrite / Skip / Copy (keep both)**
- **Native Explorer context menu** (with pywin32), Properties supported
- **Bookmarks + named sessions** (save/restore all pane paths), Dark/Light themes

---

## Install
```bash
python -m venv .venv
# PowerShell
.\.venv\Scripts\Activate.ps1
pip install PyQt5 send2trash pywin32
```
> `pywin32` enables the native Explorer context menu/Properties.  
> `send2trash` is used for safe recycling.

---

## Run
```bash
python main.py [--panes 4|6|8] [start_path1 start_path2 ...]
```
Examples:
```bash
python main.py --panes 6
python main.py --panes 4 "C:\Windows" "D:\WS" "C:\Users\USER" "C:\Temp"
```

Settings are stored via `QSettings`.  
**Organization**: `MultiPane` · **Application**: `Multi-Pane File Explorer`

---

## Keyboard Shortcuts (essentials)
| Key | Action |
|---|---|
| Alt+Left / Alt+Right | Back / Forward |
| **Alt+Up** | Go to parent folder |
| Backspace | Back |
| Enter / Ctrl+O | Open / Navigate |
| F2 | Rename |
| Delete / Shift+Delete | Recycle / Delete permanently |
| Ctrl+C / X / V | Copy / Cut / Paste |
| **Ctrl+Shift+C** | Copy full path of selection |
| **Alt+Shift+C** | Copy folder path only (files → parent folder) |
| Ctrl+L / F4 | Focus path bar |
| **Ctrl+F / F3** | Focus filter (Esc clears & returns to browse) |
| F5 | Hard refresh |

---

## Tips
- Each pane filters independently.  
- The **active pane** gets a subtle blue accent and blue breadcrumb chips.  
- Very large directories are listed incrementally to keep the UI responsive.

---

## License
MIT
