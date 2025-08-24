# Multi-Pane File Explorer (PyQt5)

A compact, high-performance **multi-pane file explorer for Windows** built with PyQt5.
Designed for power users who want to view and work across **4/6/8 panes** at once, with fast listing, Windows context menu integration, and thoughtful keyboard shortcuts.

> Tested on Windows 10/11, Python 3.9–3.12.

---

## ✨ Features

* **Multi-pane layouts**: switch instantly between **4 ↔ 6 ↔ 8** panes
* **Fast listing engine**: lazy stats, generic-icon fallback, and “huge folder” mode to prevent UI stalls
* **Windows shell context menu** (right-click) with `pywin32` (Properties, Open With, etc.)
* **Bookmarks** ⭐: one-click toggle per folder + quick-access buttons
* **Sessions** 💾: save/load *all* pane paths under a name (e.g., “Project-A – Dev”)
* **Breadcrumb path bar**: reduced spacing, auto-scroll pinned to the right (always shows deepest folder)
* **Filter & search**: `F3`/`Ctrl+F` opens filter; `Esc` clears and returns to browse
* **New items**:

  * New **Folder** button
  * New **Text File** button (creates a timestamped `.txt` and re-focuses the list)
* **Copy/move with conflict resolver**: per-file action (Overwrite / Skip / Copy both)
* **Trash / Delete**: Recycle Bin via `send2trash`/`pywin32`, with Shift+Delete for permanent
* **Status indicators**: free space (right), selection count & total size (left; files-only sum)
* **HiDPI aware**: crisp fonts and icons on per-monitor DPI setups
* **Persistent settings** via `QSettings` (pane count, last paths, theme, bookmarks, sessions)

---

## 📸 Screenshots

> Add images under `docs/` and reference them here, e.g.:
>
> ![6-pane dark](docs/6-pane-dark.png)
> ![Conflict dialog](docs/conflict-dialog.png)

---

## 🚀 Quick Start

```bash
# 1) Create and activate a venv (recommended)
py -3 -m venv .venv
.venv\Scripts\activate

# 2) Install dependencies
pip install -r requirements.txt
# or
pip install PyQt5 send2trash pywin32

# 3) Run
python multipane_explorer.py --panes 6  "C:\\"  "D:\work"  "C:\Windows"
```

**Command-line options**

* `paths ...` optional start path per pane (excess panes open Home)
* `--panes {4,6,8}` number of panes (default: 6)

---

## 🧠 Usage Tips

### Keyboard Shortcuts

| Action                          | Shortcut                                              |
| ------------------------------- | ----------------------------------------------------- |
| Back / Forward                  | `Alt+Left`, `Alt+Right` (mouse X1/X2 supported)       |
| Go up one folder                | (Button on toolbar)                                   |
| Hard refresh                    | `F5`                                                  |
| Path bar (edit)                 | `Ctrl+L`, `F4`                                        |
| Filter focus                    | `F3`, `Ctrl+F`                                        |
| Clear filter & return to browse | `Esc` (while filter focused)                          |
| Open / Enter                    | `Enter`, `Ctrl+O`                                     |
| Copy / Cut / Paste / Undo       | `Ctrl+C`, `Ctrl+X`, `Ctrl+V`, `Ctrl+Z`                |
| Delete / Permanent delete       | `Delete`, `Shift+Delete`                              |
| Rename                          | `F2`                                                  |
| Copy full path (file/folder)    | `Ctrl+Shift+C`                                        |
| Copy folder path only           | `Alt+Shift+C` *(for folders: same as `Ctrl+Shift+C`)* |

### Bookmarks & Sessions

* Click the ⭐ to toggle a bookmark for the current folder (max 10 quick slots).
* **Session** button → Save current layout & paths under a name; Load later with a click.

### Context Menu

* Right-click uses the **native Windows Explorer menu** when `pywin32` is available; otherwise a fallback menu is shown (with “New → Folder/Docx/Xlsx/Pptx/Text” etc.).

---

## ⚙️ Configuration & Tuning

Open the script to tweak these constants near the top:

* **Icons**
  `ALWAYS_GENERIC_ICONS = False`
  Set `True` to skip shell icon lookups globally (faster on slow systems).

* **Huge folder thresholds** *(to prevent UI stalls)*
  In `ExplorerPane._start_normal_model_loading`:

  ```python
  HUGE_THRESHOLD = 3000     # >= items → stay in fast model only
  GENERIC_THRESHOLD = 1200  # >= items → force generic icons
  ```

  Raise/lower to fit your machine.

* **Breadcrumb density**
  See `_common_css()`, `apply_dark_style()`, `apply_light_style()` (crumb button padding and separator padding), and `PathBar.__init__` (`self._hlay.setSpacing(0)`).

* **Themes**
  Light/Dark toggle in the toolbar; styles set by `apply_*_style()`.

* **Settings store**
  Uses `QSettings` with:

  ```python
  ORG_NAME = "MultiPane"
  APP_NAME = "Multi-Pane File Explorer"
  ```

  This saves: window geometry, pane count, per-pane last path, theme, bookmarks, sessions.

---

## 🗑️ Safe Delete Behavior

* If **`send2trash`** is installed (preferred): files go to the Recycle Bin.
* Else, if **`pywin32`** is available: uses `SHFileOperation` with undo.
* Else: falls back to **permanent deletion** (you’ll be asked to confirm).

---

## 🧩 Optional Dependencies

* `pywin32` — native context menu, Recycle Bin fallback, DPI awareness helpers
* `send2trash` — safe delete to Recycle Bin

> The app still runs without these, but with reduced Windows integration.

---

## 🛠️ Build a Standalone EXE (PyInstaller)

```bash
pip install pyinstaller
pyinstaller ^
  --noconfirm ^
  --windowed ^
  --name "MultiPaneExplorer" ^
  multipane_explorer.py
```

> If you use a custom icon: add `--icon=assets\app.ico`.

---

## 🧪 Performance Notes

* **Fast model first**: panes open with a light in-process lister (no shell icon lookups, lazy stat on visible rows).
* **Normal model** (QFileSystemModel) is attached **only when safe**.
  Extremely large folders (≥ `HUGE_THRESHOLD`) stay in fast mode to avoid UI stalls.
* **Generic icons** kick in for mid-sized folders (≥ `GENERIC_THRESHOLD`) to avoid expensive shell calls.

If you still see `[STALL] UI event loop blocked ~XXXX ms`, raise `HUGE_THRESHOLD` and/or keep `ALWAYS_GENERIC_ICONS=True`.

---

## 🧰 Troubleshooting

* **Size/Date not showing until refresh**
  Fixed via lazy stat on visible rows; ensure you’re on the latest code (look for `_schedule_visible_stats` and `StatOverlayProxy`).

* **Selection info not updating**
  The view’s `selectionModel()` is **re-hooked** whenever the model changes via `_hook_selection_model()`.

* **QSortFilterProxyModel: index from wrong model…**
  Addressed by checking the current view model before any `mapToSource()` calls.

* **Native context menu not appearing**
  Ensure `pywin32` is installed and Python is 64-bit on 64-bit Windows.

---

## 📦 Project Layout (suggested)

```
.
├─ multipane_explorer.py        # main script (all UI & logic)
├─ requirements.txt             # PyQt5, send2trash, pywin32
├─ docs/
│  ├─ 6-pane-dark.png
│  └─ conflict-dialog.png
└─ assets/
   └─ app.ico                   # optional app icon
```

**requirements.txt**

```txt
PyQt5>=5.15
send2trash>=1.8
pywin32>=306
```

---

## 🤝 Contributing

PRs welcome!
Ideas: ZIP/Extract integration, tabbed panes, quick rename templates, bulk hash checker.

1. Fork & create a feature branch
2. Keep PRs focused and include before/after notes or a short video/gif if UI-visible
3. Follow the current code style (PEP-8-ish, small helpers, no blocking in UI thread)

---

## 📄 License

No license has been specified yet.
Consider adding an open-source license (MIT is a common choice) by creating a `LICENSE` file at the repository root.

---

## 🙏 Acknowledgements

* Built on **PyQt5**
* Uses **send2trash** and **pywin32** for best-effort native Windows behavior

---

### One-liner

> A snappy, multi-pane Windows file explorer for power users — with real shell menus, sessions, bookmarks, and smart performance.
