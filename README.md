# Multi-Pane File Explorer (PyQt5)

*A fast, keyboard-friendly **multi-pane** file explorer for Windows.*  
**by kkongt2 · Built with GPT-5 Thinking**

> Windows 10/11 • Python 3.9–3.12

---

## Features

- **4 / 6 / 8 panes** with quick layout toggle (remembers last session & paths)
- **Folders first** sorting, then files (size/date work as expected)
- **Auto refresh (watcher-based)** when contents change
- **Super-fast listing** for big directories (scandir + incremental rows; switches away from `QFileSystemModel` for huge folders)
- **Breadcrumbs** with per-crumb highlight for the active pane
- **Filter / search** (wildcards like `*.txt`, `*report*.xlsx`), **Esc** clears & returns to browse
- **One-click New**: folder / empty `.txt` (and templates for docx/xlsx/pptx if available)
- **Copy/Move with conflict resolver** (Overwrite / Skip / Keep both per item)
- **Native Windows context menu** (incl. Properties) when available
- **Bookmarks** + **named sessions** (save/restore all pane paths)
- **Dark / Light themes**, compact toolbar

---

## Install

```bash
python -m venv .venv
# PowerShell
.\.venv\Scripts\Activate.ps1
pip install PyQt5 send2trash pywin32
```

> `pywin32`는 선택 사항이지만 탐색기 컨텍스트 메뉴/속성 창을 활성화합니다.  
> `send2trash`는 휴지통 이동에 사용됩니다.

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

Settings are stored via `QSettings`:

- **Organization**: `MultiPane`  
- **Application**: `Multi-Pane File Explorer`

---

## Keyboard Shortcuts (essentials)

| Key | Action |
|---|---|
| **Alt+Left / Alt+Right** | Back / Forward |
| **Backspace** | Up one folder |
| **Enter / Ctrl+O** | Open / Navigate |
| **F2** | Rename |
| **Delete / Shift+Delete** | Recycle / Delete permanently |
| **Ctrl+C / Ctrl+X / Ctrl+V** | Copy / Cut / Paste |
| **Ctrl+Shift+C** | Copy full path of selection |
| **Alt+Shift+C** | Copy folder path only (files → parent folder) |
| **Ctrl+L / F4** | Focus path bar |
| **Ctrl+F / F3** | Focus filter box |
| **Esc** (in filter) | Clear filter and return to browse |
| **F5** | Hard refresh |

---

## Tips

- Filtering supports wildcards. Each pane filters independently.  
- The **active pane** is subtly highlighted; its breadcrumb **chips** turn light blue.  
- Right-click shows the **native Explorer menu** when `pywin32` is installed; otherwise a lightweight **New …** menu appears.  
- On name conflicts during copy/move, choose **Overwrite / Skip / Copy (keep both)** per item.  
- Very large folders use an optimized fast model (`os.scandir`) and incremental stat/icon fill to keep the UI responsive.

---

## Platforms

- Windows 10/11 • Python 3.9–3.12  
- HiDPI / Per-monitor-V2 aware (Segoe UI default)

---

## License

MIT (or your preferred license). Add your copyright.

---

## Credits

Built with ❤️ on **PyQt5**. Optional integrations via **pywin32** and **send2trash**.  
**Author:** kkongt2 · **Based on GPT-5 Thinking**
