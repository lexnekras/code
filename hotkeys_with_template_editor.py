# -*- coding: utf-8 -*-
"""
ALL-IN-ONE Hotkeys (Windows)

Объединение 3 скриптов в один, без конфликтов по RegisterHotKey ID и с сохранением логики:

1) Excel Clipboard Hotkey Paster
   - Ctrl + Alt + D  -> Excel (развернуть) → новая книга → вставка → правки диапазона

2) Excel NBSP Cleaner
   - Ctrl + Alt + L  -> очистить выделение в активном Excel (NBSP/узк. NBSP, приведение к числу, формат)

3) Встроенный редактор шаблонов + история буфера
   - Ctrl + 1..9     -> вставка шаблонов (подстановка {clip}/{clip1}/{clip2}...)
   - Ctrl + Alt + R  -> открыть редактор шаблонов

Общее:
   - Ctrl + Alt + Q  -> выход

Важно про двойной клик:
- Если файл сохранить как .py и запускать двойным кликом — откроется консоль (python.exe).
- Если сохранить как .pyw (или запускать через pythonw.exe) — будет "тихо" без консоли.

Требования:
  pip install pywin32 pyperclip
  (Нужен установленный Microsoft Excel; Windows)

"""

from __future__ import annotations

import time
import json
import os
import queue
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import threading
from pathlib import Path
from collections import deque
from contextlib import contextmanager
import ctypes
from ctypes import wintypes

if os.name == "nt":
    import pythoncom
    import pywintypes
    import win32com.client as win32
    import pyperclip
else:
    # Storage and the editor can be tested independently of Windows automation.
    pythoncom = pywintypes = win32 = pyperclip = None


# ===========================
#  UI / Messages
# ===========================
user32 = ctypes.WinDLL("user32", use_last_error=True) if os.name == "nt" else None
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True) if os.name == "nt" else None
oleacc = ctypes.OleDLL("oleacc") if os.name == "nt" else None

def info_box(text: str, title: str = "Hotkeys") -> None:
    try:
        user32.MessageBoxW(0, text, title, 0x40)
    except Exception:
        pass

def error_box(text: str, title: str = "Hotkeys - Ошибка") -> None:
    try:
        user32.MessageBoxW(0, text, title, 0x10)
    except Exception:
        pass


# ===========================
#  WinAPI constants
# ===========================
WM_HOTKEY = 0x0312
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_NOREPEAT = 0x4000

# Virtual keys
VK_0 = 0x30
VK_NUMPAD1 = 0x61
VK_Q = 0x51
VK_R = 0x52

VK_CONTROL = 0x11
VK_V = 0x56

# Excel window constants
XL_MAXIMIZED = -4137


# ============================================================
#  Feature #3: Local templates + clipboard history
# ============================================================
SETTINGS_DIR = Path(os.environ.get("APPDATA") or (Path.home() / ".config")) / "HotkeysTemplateEditor"
TEMPLATE_FILE = SETTINGS_DIR / "templates.json"
MAX_TEMPLATES = 9
PRESERVE_CLIPBOARD = True
PASTE_DELAY = 0.06
REGISTER_NUMPAD = False

# History
CLIP_HISTORY_SIZE = 5
CLIP_POLL_INTERVAL = 0.2
_clip_history = deque(maxlen=CLIP_HISTORY_SIZE)
_history_paused = 0
_history_lock = threading.RLock()
_shutdown_event = threading.Event()

@contextmanager
def _pause_history():
    global _history_paused
    with _history_lock:
        _history_paused += 1
    try:
        yield
    finally:
        with _history_lock:
            _history_paused -= 1

def wait_ctrl_released(timeout=0.9) -> None:
    end = time.time() + timeout
    while time.time() < end:
        ctrl_down = user32.GetAsyncKeyState(VK_CONTROL) & 0x8000
        if not ctrl_down:
            return
        time.sleep(0.01)

def _send_ctrl_v():
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_V, 0, 0, 0)
    user32.keybd_event(VK_V, 0, 2, 0)
    user32.keybd_event(VK_CONTROL, 0, 2, 0)

def paste_via_ctrl_v(text: str):
    if not text:
        return
    original = pyperclip.paste()
    with _pause_history():
        try:
            pyperclip.copy(text)
            time.sleep(PASTE_DELAY)
            _send_ctrl_v()
        finally:
            if PRESERVE_CLIPBOARD:
                time.sleep(PASTE_DELAY)
                pyperclip.copy(original)

def _init_clip_history():
    try:
        cur = pyperclip.paste() or ""
        with _history_lock:
            _clip_history.clear()
            if cur:
                _clip_history.appendleft(cur)
    except Exception:
        pass

def _clipboard_watcher():
    while not _shutdown_event.is_set():
        try:
            with _history_lock:
                if not _history_paused:
                    cur = pyperclip.paste()
                    if cur and (not _clip_history or cur != _clip_history[0]):
                        _clip_history.appendleft(cur)
        except Exception:
            pass
        _shutdown_event.wait(CLIP_POLL_INTERVAL)

def _get_clip(idx: int = 0) -> str:
    with _history_lock:
        if 0 <= idx < len(_clip_history):
            return _clip_history[idx]
    return ""

_clip_pattern = re.compile(r"\{clip(?:([0-9]+))?\}")

def substitute_clip_placeholders(text: str) -> str:
    # Capture once: all placeholders in one template refer to the same history.
    with _history_lock:
        try:
            current = pyperclip.paste() or ""
        except Exception:
            current = _get_clip()
        if current and (not _clip_history or current != _clip_history[0]):
            _clip_history.appendleft(current)
        history = list(_clip_history) if current else [""] + list(_clip_history)

    def repl(m: re.Match) -> str:
        g = m.group(1)
        idx = int(g) if g is not None else 0
        return history[idx] if idx < len(history) else ""
    return _clip_pattern.sub(repl, text)

templates: dict[str, str] = {}
_templates_lock = threading.Lock()
_paste_lock = threading.Lock()

def get_templates() -> dict[str, str]:
    with _templates_lock:
        return dict(templates)

def set_templates(values: dict[str, str]) -> None:
    global templates
    with _templates_lock:
        templates = dict(values)

def paste_text(text: str):
    wait_ctrl_released()
    paste_via_ctrl_v(text)

def do_paste(num_str: str):
    if not _paste_lock.acquire(blocking=False):
        return
    try:
        base = get_templates().get(num_str)
        if not base:
            return
        text = substitute_clip_placeholders(base)
        paste_text(text)
    except Exception as exc:
        error_box(f"Не удалось вставить шаблон:\n{exc}")
    finally:
        _paste_lock.release()

def on_hotkey_number(num_str: str):
    threading.Thread(target=do_paste, args=(num_str,), daemon=True).start()

# ============================================================
#  Feature #1: Excel Clipboard Hotkey Paster (Ctrl+Alt+D)
# ============================================================
# Settings from original script
PASTE_ANCHOR = "A1"
RETRY_PASTE_TIMES = 12
RETRY_DELAY_SEC = 0.12

DISABLE_WRAP_TEXT = True
AUTOFIT_COLUMNS = True
CONVERT_TEXT_NUMBERS = True

BRING_EXCEL_TO_FRONT = True
FORCE_EXCEL_FOREGROUND_BEFORE_PASTE = True

SHOW_ERROR_POPUPS_EXCEL_PASTE = True

_num_re = re.compile(r"""
^\s*
-?
(?:
  \d{1,3} (?:[ \u00A0]\d{3})+   # 1 234 567
  |
  \d+                          # 1234567
)
(?:,\d+)?                       # ,00
\s*$
""", re.VERBOSE)

def _excel_paste_msg(text: str, title: str = "Excel Paste Hotkey") -> None:
    if not SHOW_ERROR_POPUPS_EXCEL_PASTE:
        return
    error_box(text, title)

def bring_to_front(hwnd: int) -> None:
    try:
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass

def get_or_start_excel():
    """Excel.Application (используем уже запущенный, если есть)."""
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")

    excel.Visible = True
    excel.DisplayAlerts = True
    return excel

def _try_execute_mso(excel, mso: str) -> bool:
    """Пытается выполнить UI-команду Excel (как кнопка на ленте)."""
    try:
        excel.CommandBars.ExecuteMso(mso)
        return True
    except Exception:
        return False

def _coerce_ru_number(s: str):
    """Преобразует строку вида '4 375,00' / '406 875,00' / '93' в число."""
    if not isinstance(s, str):
        return s
    t = s.replace("\u00A0", " ").strip()  # NBSP -> space
    if not _num_re.match(t):
        return s
    t2 = t.replace(" ", "")
    if "," in t2:
        try:
            return float(t2.replace(",", "."))
        except Exception:
            return s
    try:
        return int(t2)
    except Exception:
        return s

def convert_text_numbers_in_range(rng) -> None:
    """Преобразует "числа как текст" внутри диапазона через массовую запись."""
    try:
        vals = rng.Value2
    except Exception:
        return
    if vals is None:
        return

    single_cell = False
    if not isinstance(vals, (tuple, list)):
        vals = [[vals]]
        single_cell = True
    else:
        if len(vals) > 0 and not isinstance(vals[0], (tuple, list)):
            vals = [list(vals)]
        else:
            vals = [list(row) for row in vals]

    changed = False
    for r in range(len(vals)):
        for c in range(len(vals[r])):
            v = vals[r][c]
            if isinstance(v, str):
                nv = _coerce_ru_number(v)
                if nv is not v and nv != v:
                    vals[r][c] = nv
                    changed = True

    if not changed:
        return

    try:
        rng.NumberFormat = "General"
    except Exception:
        pass

    try:
        if single_cell:
            rng.Value2 = vals[0][0]
        else:
            rng.Value2 = vals
    except Exception:
        pass

def get_pasted_range(excel, ws):
    """Берём выделение после вставки (обычно это вставленный диапазон)."""
    try:
        sel = excel.Selection
        _ = sel.Address
        return sel
    except Exception:
        pass
    try:
        return ws.UsedRange
    except Exception:
        return None

def paste_clipboard_to_new_workbook() -> None:
    """Новая книга → вставка → правки диапазона. Excel остаётся открытым."""
    pythoncom.CoInitialize()
    try:
        excel = get_or_start_excel()

        try:
            excel.WindowState = XL_MAXIMIZED
        except Exception:
            pass

        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Activate()
        ws.Range(PASTE_ANCHOR).Select()

        if BRING_EXCEL_TO_FRONT and FORCE_EXCEL_FOREGROUND_BEFORE_PASTE:
            try:
                bring_to_front(int(excel.Hwnd))
            except Exception:
                pass
            try:
                excel.WindowState = XL_MAXIMIZED
            except Exception:
                pass

        last_ok = False

        # 1) UI Paste
        for _ in range(RETRY_PASTE_TIMES):
            if _try_execute_mso(excel, "Paste"):
                last_ok = True
                break
            time.sleep(RETRY_DELAY_SEC)

        # 2) Paste source formatting
        if not last_ok:
            for _ in range(RETRY_PASTE_TIMES):
                if _try_execute_mso(excel, "PasteSourceFormatting"):
                    last_ok = True
                    break
                time.sleep(RETRY_DELAY_SEC)

        # 3) HTML PasteSpecial (если поддерживается)
        if not last_ok:
            try:
                ws.PasteSpecial(Format="HTML")
                last_ok = True
            except Exception:
                pass

        # 4) COM Paste
        if not last_ok:
            for _ in range(RETRY_PASTE_TIMES):
                try:
                    ws.Paste()
                    last_ok = True
                    break
                except Exception:
                    time.sleep(RETRY_DELAY_SEC)

        # 5) Range.PasteSpecial fallback
        if not last_ok:
            try:
                ws.Range(PASTE_ANCHOR).PasteSpecial()
                last_ok = True
            except Exception:
                pass

        if not last_ok:
            raise RuntimeError("Не удалось вставить из буфера. Проверьте 'Копировать с форматированием' в Acrobat Pro.")

        rng = get_pasted_range(excel, ws)
        if rng is not None:
            if DISABLE_WRAP_TEXT:
                try:
                    rng.WrapText = False
                except Exception:
                    pass
            if AUTOFIT_COLUMNS:
                try:
                    rng.Columns.AutoFit()
                except Exception:
                    pass
            if CONVERT_TEXT_NUMBERS:
                convert_text_numbers_in_range(rng)

        if BRING_EXCEL_TO_FRONT:
            try:
                bring_to_front(int(excel.Hwnd))
            except Exception:
                pass
            try:
                excel.WindowState = XL_MAXIMIZED
            except Exception:
                pass

    finally:
        pythoncom.CoUninitialize()


# ============================================================
#  Feature #2: Excel NBSP Cleaner (Ctrl+Alt+L)
# ============================================================
REMOVE_CHARS = ["\u00A0", "\u202F"]  # NBSP, narrow NBSP

CLEAN_CONSTANTS = True
CLEAN_FORMULAS = False

APPLY_THOUSAND_SEPARATOR = True
DECIMALS = 2
COERCE_TEXT_TO_NUMBER = True

MAX_CELLS = 200000

# Excel constants
xlPart = 2
xlByRows = 1
xlCellTypeConstants = 2
xlCellTypeFormulas = -4123
xlDecimalSeparator = 3
xlThousandsSeparator = 4

# oleacc constants
OBJID_NATIVEOM = 0xFFFFFFF0
HRESULT = ctypes.c_long
GA_ROOT = 2
CLSNAME_BUF = 256

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", ctypes.c_ubyte * 8),
    ]

    @staticmethod
    def from_string(g: str) -> "GUID":
        g = g.strip().strip("{}")
        parts = g.split("-")
        d1 = int(parts[0], 16)
        d2 = int(parts[1], 16)
        d3 = int(parts[2], 16)
        d4 = bytes.fromhex(parts[3] + parts[4])
        arr = (ctypes.c_ubyte * 8)(*d4)
        return GUID(d1, d2, d3, arr)

def _get_root_hwnd(hwnd: int) -> int:
    root = user32.GetAncestor(hwnd, GA_ROOT)
    return int(root) if root else int(hwnd)

def _get_class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(CLSNAME_BUF)
    user32.GetClassNameW(hwnd, buf, CLSNAME_BUF)
    return buf.value

def _find_excel7(hwnd_xlmain: int) -> int | None:
    """XLMAIN -> XLDESK -> EXCEL7"""
    FindWindowExW = user32.FindWindowExW
    FindWindowExW.argtypes = [wintypes.HWND, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR]
    FindWindowExW.restype = wintypes.HWND

    h_xldesk = FindWindowExW(wintypes.HWND(hwnd_xlmain), None, "XLDESK", None)
    if not h_xldesk:
        return None

    h_excel7 = FindWindowExW(h_xldesk, None, "EXCEL7", None)
    if not h_excel7:
        h_child = None
        while True:
            h_child = FindWindowExW(h_xldesk, h_child, None, None)
            if not h_child:
                break
            if _get_class_name(int(h_child)) == "EXCEL7":
                h_excel7 = h_child
                break

    return int(h_excel7) if h_excel7 else None

def get_excel_application_from_foreground() -> object | None:
    """Return Excel.Application for the active workbook window."""
    hwnd_fg = user32.GetForegroundWindow()
    if not hwnd_fg:
        return None
    hwnd_root = _get_root_hwnd(int(hwnd_fg))
    if _get_class_name(hwnd_root) != "XLMAIN":
        return None

    hwnd_excel7 = _find_excel7(hwnd_root)
    if hwnd_excel7 is None:
        return None

    iid_str = "{00020400-0000-0000-C000-000000000046}"  # IID_IDispatch
    guid = GUID.from_string(iid_str)
    p_disp = ctypes.c_void_p()

    AccessibleObjectFromWindow = oleacc.AccessibleObjectFromWindow
    AccessibleObjectFromWindow.restype = HRESULT
    AccessibleObjectFromWindow.argtypes = [
        wintypes.HWND,
        wintypes.DWORD,
        ctypes.POINTER(GUID),
        ctypes.POINTER(ctypes.c_void_p),
    ]

    hr = AccessibleObjectFromWindow(wintypes.HWND(hwnd_excel7), OBJID_NATIVEOM, ctypes.byref(guid), ctypes.byref(p_disp))
    if hr != 0 or not p_disp.value:
        return None

    iid = pywintypes.IID(iid_str)
    native_obj = pythoncom.ObjectFromAddress(int(p_disp.value), iid)
    disp = win32.Dispatch(native_obj)

    try:
        return disp.Application
    except Exception:
        return disp

def _safe_int(x):
    try:
        return int(x)
    except Exception:
        return None

def _safe_count(rng) -> int | None:
    try:
        return _safe_int(rng.CountLarge)
    except Exception:
        try:
            return _safe_int(rng.Count)
        except Exception:
            return None

def _choose_range(app):
    """Prefer RangeSelection then Selection then ActiveCell."""
    rs = sel = ac = None
    try:
        wnd = app.ActiveWindow
        rs = wnd.RangeSelection if wnd is not None else None
    except Exception:
        rs = None
    try:
        sel = app.Selection
    except Exception:
        sel = None
    try:
        ac = app.ActiveCell
    except Exception:
        ac = None

    for rng in (rs, sel, ac):
        if rng is not None and hasattr(rng, "Replace"):
            cnt = _safe_count(rng)
            if cnt is None:
                continue
            if cnt > MAX_CELLS:
                continue
            return rng, cnt
    return None, None

def _build_number_format_local(app) -> str:
    try:
        dec = str(app.International(xlDecimalSeparator))
    except Exception:
        dec = ","
    try:
        th = str(app.International(xlThousandsSeparator))
    except Exception:
        th = " "

    if DECIMALS <= 0:
        return f"#{th}##0"
    return f"#{th}##0{dec}" + ("0" * DECIMALS)

def _coerce_str_to_number(s: str, app):
    if not COERCE_TEXT_TO_NUMBER:
        return None
    try:
        dec = str(app.International(xlDecimalSeparator))
    except Exception:
        dec = ","
    try:
        th = str(app.International(xlThousandsSeparator))
    except Exception:
        th = " "

    t = s.strip()
    if not t:
        return None
    t = t.replace("\u00A0", "").replace("\u202F", "")
    t = t.replace(" ", "")
    if th and th != " ":
        t = t.replace(th, "")
    if dec and dec != ".":
        t = t.replace(dec, ".")
    if t.count(".") > 1:
        return None
    try:
        return float(t)
    except Exception:
        return None

def _process_single_cell(cell, app):
    try:
        v = cell.Value
    except Exception:
        return

    if isinstance(v, str):
        s = v
        for ch in REMOVE_CHARS:
            s = s.replace(ch, "")
        if s != v:
            try:
                cell.Value = s
                v = s
            except Exception:
                pass

    if isinstance(v, str):
        num = _coerce_str_to_number(v, app)
        if num is not None:
            try:
                cell.Value = num
                v = num
            except Exception:
                pass

    if APPLY_THOUSAND_SEPARATOR:
        fmt = _build_number_format_local(app)
        try:
            cell.NumberFormatLocal = fmt
        except Exception:
            try:
                cell.NumberFormat = fmt
            except Exception:
                pass

def _try_specialcells(rng, cell_type):
    try:
        sc = rng.SpecialCells(cell_type)
        if sc is not None and hasattr(sc, "Replace"):
            return sc
    except Exception:
        return None
    return None

def _replace_in_range(rng):
    for ch in REMOVE_CHARS:
        rng.Replace(ch, "", xlPart, xlByRows, False)

def _coerce_text_numbers_range(rng, app):
    if not COERCE_TEXT_TO_NUMBER:
        return
    try:
        cells = rng.Cells
        count = int(cells.Count)
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            c = cells.Item(i)
            v = c.Value
        except Exception:
            continue
        if not isinstance(v, str):
            continue
        num = _coerce_str_to_number(v, app)
        if num is not None:
            try:
                c.Value = num
            except Exception:
                pass

def _apply_number_format_range(rng, app):
    if not APPLY_THOUSAND_SEPARATOR:
        return
    fmt = _build_number_format_local(app)
    try:
        rng.NumberFormatLocal = fmt
    except Exception:
        try:
            rng.NumberFormat = fmt
        except Exception:
            pass

def clean_now():
    pythoncom.CoInitialize()
    try:
        app = get_excel_application_from_foreground()
        if app is None:
            raise RuntimeError("Не удалось получить Excel из активного окна. Сделай Excel активным и попробуй снова.")

        base, cnt = _choose_range(app)
        if base is None or cnt is None:
            raise RuntimeError("Не удалось определить выделение. Убедись, что выделены ячейки и Excel активен.")

        if cnt == 1:
            cell = app.ActiveCell
            _process_single_cell(cell, app)
            return

        did_any = False

        if CLEAN_CONSTANTS:
            const_rng = _try_specialcells(base, xlCellTypeConstants)
            if const_rng is not None:
                _replace_in_range(const_rng)
                _coerce_text_numbers_range(const_rng, app)
                _apply_number_format_range(const_rng, app)
                did_any = True

        if CLEAN_FORMULAS:
            form_rng = _try_specialcells(base, xlCellTypeFormulas)
            if form_rng is not None:
                _replace_in_range(form_rng)
                _apply_number_format_range(form_rng, app)
                did_any = True

        if not did_any:
            _replace_in_range(base)
            _coerce_text_numbers_range(base, app)
            _apply_number_format_range(base, app)

    finally:
        pythoncom.CoUninitialize()


# ============================================================
#  Hotkey registration (single unified loop)
# ============================================================
# We keep template hotkey IDs compatible with old script:
#   1..9           = templates
#   1000+1..9      = templates (numpad) if enabled
#   100            = open template editor
#   101            = exit
#
# Add non-conflicting IDs for Excel features:
ID_EXCEL_PASTE = 200   # Ctrl+Alt+D
ID_NBSP_CLEAN  = 201   # Ctrl+Alt+L

# ============================================================
#  Unified hotkey dispatcher
# ============================================================
_excel_paste_lock = threading.Lock()
_excel_clean_lock = threading.Lock()

def _run_excel_paste_async():
    if not _excel_paste_lock.acquire(blocking=False):
        return
    def worker():
        try:
            paste_clipboard_to_new_workbook()
        except Exception as e:
            _excel_paste_msg(str(e), "Excel Paste Hotkey - Ошибка")
        finally:
            _excel_paste_lock.release()
    threading.Thread(target=worker, daemon=True).start()

def _run_excel_clean_async():
    if not _excel_clean_lock.acquire(blocking=False):
        return
    def worker():
        try:
            clean_now()
        except Exception as e:
            error_box(str(e), "Excel NBSP Cleaner - Ошибка")
        finally:
            _excel_clean_lock.release()
    threading.Thread(target=worker, daemon=True).start()


# ============================================================
#  Template storage and optional one-time DOCX import
# ============================================================
class TemplateStore:
    def __init__(self, path: Path = TEMPLATE_FILE):
        self.path = Path(path)

    @staticmethod
    def validate(values) -> dict[str, str]:
        if not isinstance(values, dict):
            raise ValueError("Список шаблонов должен быть объектом JSON.")
        result = {}
        for key, value in values.items():
            if key not in {str(i) for i in range(1, MAX_TEMPLATES + 1)}:
                raise ValueError(f"Неизвестный номер шаблона: {key!r}")
            if not isinstance(value, str):
                raise ValueError(f"Шаблон {key} должен содержать текст.")
            result[key] = value
        return result

    def load(self) -> dict[str, str]:
        if not self.path.exists():
            return {}
        with self.path.open("r", encoding="utf-8-sig") as stream:
            document = json.load(stream)
        if not isinstance(document, dict) or document.get("version") != 1:
            raise ValueError("Неизвестный формат настроек шаблонов.")
        return self.validate(document.get("templates"))

    def save(self, values: dict[str, str]) -> None:
        values = self.validate(values)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=self.path.parent,
                prefix=self.path.name + ".", suffix=".tmp", delete=False,
            ) as stream:
                temporary = Path(stream.name)
                json.dump({"version": 1, "templates": values}, stream,
                          ensure_ascii=False, indent=2)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


def import_docx_templates(path: Path) -> dict[str, str]:
    """Read body paragraphs from DOCX using the standard library; Word is not needed."""
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
    result = {}
    for paragraph in document.findall("w:body/w:p", namespace):
        chunks = []
        for node in paragraph.iter():
            if node.tag == f"{{{namespace['w']}}}t":
                chunks.append(node.text or "")
            elif node.tag == f"{{{namespace['w']}}}tab":
                chunks.append("\t")
            elif node.tag in {f"{{{namespace['w']}}}br", f"{{{namespace['w']}}}cr"}:
                chunks.append("\n")
        text = "".join(chunks).strip()
        if text:
            result[str(len(result) + 1)] = text
            if len(result) == MAX_TEMPLATES:
                break
    return result


# ============================================================
#  Windows hotkeys: a dedicated message thread; Tk stays on main
# ============================================================
WM_QUIT = 0x0012
WM_REFRESH_HOTKEYS = 0x8001


def configure_windows_api():
    user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
    user32.RegisterHotKey.restype = wintypes.BOOL
    user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.UnregisterHotKey.restype = wintypes.BOOL
    user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
    user32.GetMessageW.restype = wintypes.BOOL
    user32.PeekMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT]
    user32.PeekMessageW.restype = wintypes.BOOL
    user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.PostThreadMessageW.restype = wintypes.BOOL
    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user32.GetAncestor.restype = wintypes.HWND
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    kernel32.GetCurrentThreadId.argtypes = []
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD


def own_window_is_foreground() -> bool:
    if user32 is None:
        return False
    process_id = wintypes.DWORD()
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return False
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    return process_id.value == os.getpid()


class HotkeyService:
    def __init__(self, events: queue.Queue):
        self.events = events
        self.thread_id = None
        self.stop_requested = threading.Event()
        self.thread = None
        self.registered = set()

    def start(self):
        self.thread = threading.Thread(target=self._run, name="WindowsHotkeys", daemon=True)
        self.thread.start()

    def refresh(self):
        if self.thread_id is not None:
            if not user32.PostThreadMessageW(self.thread_id, WM_REFRESH_HOTKEYS, 0, 0):
                self.events.put(("error", "Не удалось обновить горячие клавиши. Перезапустите программу."))

    def stop(self):
        self.stop_requested.set()
        if self.thread_id is not None:
            user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)

    def _bindings(self):
        bindings = {
            100: (MOD_CONTROL | MOD_ALT, VK_R, "Ctrl+Alt+R — редактор"),
            101: (MOD_CONTROL | MOD_ALT, VK_Q, "Ctrl+Alt+Q — выход"),
            ID_EXCEL_PASTE: (MOD_CONTROL | MOD_ALT, ord("D"), "Ctrl+Alt+D — Excel"),
            ID_NBSP_CLEAN: (MOD_CONTROL | MOD_ALT, ord("L"), "Ctrl+Alt+L — очистка Excel"),
        }
        for key, text in get_templates().items():
            if not text.strip():
                continue
            number = int(key)
            bindings[number] = (MOD_CONTROL, VK_0 + number, f"Ctrl+{key}")
            if REGISTER_NUMPAD:
                bindings[1000 + number] = (MOD_CONTROL, VK_NUMPAD1 + number - 1, f"Ctrl+NumPad{key}")
        return bindings

    def _sync(self):
        desired = self._bindings()
        for hotkey_id in self.registered - desired.keys():
            user32.UnregisterHotKey(None, hotkey_id)
        self.registered.intersection_update(desired)
        failed = {}
        for hotkey_id, (modifiers, key, label) in desired.items():
            if hotkey_id not in self.registered:
                if user32.RegisterHotKey(None, hotkey_id, modifiers | MOD_NOREPEAT, key):
                    self.registered.add(hotkey_id)
                else:
                    failed[hotkey_id] = label
        self.events.put(("hotkeys", {"registered": set(self.registered), "failed": failed}))

    def _run(self):
        try:
            message = wintypes.MSG()
            # PostThreadMessage requires an existing message queue.
            user32.PeekMessageW(ctypes.byref(message), None, 0, 0, 0)
            self.thread_id = kernel32.GetCurrentThreadId()
            if self.stop_requested.is_set():
                return
            self._sync()
            while not self.stop_requested.is_set():
                result = user32.GetMessageW(ctypes.byref(message), None, 0, 0)
                if result == -1:
                    raise ctypes.WinError(ctypes.get_last_error())
                if result == 0:
                    break
                if message.message == WM_REFRESH_HOTKEYS:
                    self._sync()
                elif message.message == WM_HOTKEY:
                    hotkey_id = int(message.wParam)
                    if 1 <= hotkey_id <= MAX_TEMPLATES:
                        self.events.put(("paste", str(hotkey_id)))
                    elif 1000 < hotkey_id <= 1000 + MAX_TEMPLATES:
                        self.events.put(("paste", str(hotkey_id - 1000)))
                    elif hotkey_id == 100:
                        self.events.put(("show", None))
                    elif hotkey_id == 101:
                        self.events.put(("exit", None))
                    elif hotkey_id == ID_EXCEL_PASTE:
                        _run_excel_paste_async()
                    elif hotkey_id == ID_NBSP_CLEAN:
                        _run_excel_clean_async()
        except Exception as exc:
            self.events.put(("error", f"Служба горячих клавиш остановлена:\n{exc}"))
        finally:
            for hotkey_id in self.registered:
                user32.UnregisterHotKey(None, hotkey_id)
            self.registered.clear()
            self.thread_id = None
            self.events.put(("stopped", None))

# ============================================================
#  Built-in template editor (all Tk operations on the main thread)
# ============================================================
class TemplateEditor(tk.Tk):
    def __init__(self, store: TemplateStore, service: HotkeyService, events: queue.Queue):
        super().__init__()
        self.store = store
        self.service = service
        self.events = events
        self._saved = {str(i): get_templates().get(str(i), "") for i in range(1, MAX_TEMPLATES + 1)}
        self._draft = dict(self._saved)
        self._current = "1"
        self._loading_text = False
        self._closing = False
        self._registered = set()
        self._failed = {}
        self._poll_id = None
        self.geometry("1100x740")
        self.minsize(930, 680)
        self.configure(background="#F4F6FA")
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#F4F6FA")
        style.configure("TLabel", background="#F4F6FA", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"))
        style.configure("Heading.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=34)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        self._build_ui()
        self._load_text()
        self._refresh_rows()
        self.protocol("WM_DELETE_WINDOW", self.hide_editor)
        self.bind("<Control-s>", self._save_shortcut)
        self.bind("<Escape>", lambda event: self.hide_editor())
        self._poll_id = self.after(80, self._poll_events)

    def _build_ui(self):
        menu = tk.Menu(self)
        actions = tk.Menu(menu, tearoff=False)
        actions.add_command(label="Сохранить шаблоны", accelerator="Ctrl+S", command=self.save_templates)
        actions.add_command(label="Импортировать старые шаблоны из Word…", command=self.import_from_word)
        actions.add_separator()
        actions.add_command(label="Свернуть в фон", command=self.hide_editor)
        actions.add_command(label="Завершить программу", command=self.exit_program)
        menu.add_cascade(label="Программа", menu=actions)
        self.configure(menu=menu)

        container = ttk.Frame(self, padding=20)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="Мои шаблоны", style="Title.TLabel").pack(anchor="w")
        ttk.Label(container, text="Выберите сочетание слева, введите текст и нажмите «Сохранить».").pack(anchor="w", pady=(5, 14))
        self.runtime_status = tk.StringVar(value="Подключение горячих клавиш…")
        self.runtime_label = ttk.Label(container, textvariable=self.runtime_status, wraplength=1030)
        self.runtime_label.pack(anchor="w", pady=(0, 12))

        panels = ttk.Panedwindow(container, orient="horizontal")
        panels.pack(fill="both", expand=True)
        left = ttk.Frame(panels)
        right = ttk.Frame(panels, padding=(16, 0, 0, 0))
        panels.add(left, weight=4)
        panels.add(right, weight=6)

        self.template_list = ttk.Treeview(left, columns=("key", "preview", "state"), show="headings", selectmode="browse", height=9)
        self.template_list.heading("key", text="Клавиши")
        self.template_list.heading("preview", text="Начало текста")
        self.template_list.heading("state", text="Состояние")
        self.template_list.column("key", width=76, minwidth=72, stretch=False)
        self.template_list.column("preview", width=176, minwidth=120)
        self.template_list.column("state", width=124, minwidth=116, stretch=False)
        self.template_list.pack(fill="both", expand=True)
        for i in range(1, MAX_TEMPLATES + 1):
            self.template_list.insert("", "end", iid=str(i))
        self.template_list.selection_set("1")
        self.template_list.bind("<<TreeviewSelect>>", self._selection_changed)
        ttk.Label(left, text="Пустой шаблон освобождает сочетание клавиш.\nИзменения начинают действовать после сохранения.", wraplength=390).pack(anchor="w", pady=(10, 0))

        self.editor_title = tk.StringVar()
        ttk.Label(right, textvariable=self.editor_title, style="Heading.TLabel").pack(anchor="w", pady=(0, 8))
        text_frame = ttk.Frame(right)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, wrap="word", undo=True, font=("Segoe UI", 11),
                            background="#FFFFFF", foreground="#172033", relief="solid", borderwidth=1,
                            padx=12, pady=12, insertbackground="#172033", height=10, width=32)
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self.text.bind("<<Modified>>", self._text_changed)

        ttk.Label(right, text="Вставить подстановку:").pack(anchor="w", pady=(12, 5))
        tokens = ttk.Frame(right)
        tokens.pack(anchor="w")
        for token in ("{clip}", "{clip1}", "{clip2}", "{clip3}", "{clip4}"):
            ttk.Button(tokens, text=token, width=9,
                       command=lambda value=token: self.insert_placeholder(value)).pack(side="left", padx=(0, 4))
        ttk.Label(right, text="{clip} — текущий текст буфера; {clip1}…{clip4} — предыдущие записи.\nИстория хранится только пока программа запущена.", wraplength=540).pack(anchor="w", pady=(6, 8))
        ttk.Button(right, text="Очистить этот шаблон", command=self.clear_template).pack(anchor="w")

        self.save_status = tk.StringVar()
        ttk.Label(container, textvariable=self.save_status).pack(anchor="w", pady=(14, 8))
        buttons = ttk.Frame(container)
        buttons.pack(fill="x")
        self.save_button = ttk.Button(buttons, text="Сохранить", command=self.save_templates)
        self.save_button.pack(side="left")
        self.revert_button = ttk.Button(buttons, text="Отменить изменения", command=self.revert_changes)
        self.revert_button.pack(side="left", padx=8)
        ttk.Button(buttons, text="Свернуть в фон", command=self.hide_editor).pack(side="right")
        ttk.Button(buttons, text="Завершить программу", command=self.exit_program).pack(side="right", padx=8)
        ttk.Label(container, text="Закрытие окна оставляет программу в фоне. Открыть: Ctrl+Alt+R · Выход: Ctrl+Alt+Q", wraplength=1030).pack(anchor="w", pady=(12, 0))

    @property
    def dirty(self):
        return self._draft != self._saved

    def _commit_text(self):
        if not self._loading_text:
            self._draft[self._current] = self.text.get("1.0", "end-1c")

    def _load_text(self):
        self._loading_text = True
        self.text.delete("1.0", "end")
        self.text.insert("1.0", self._draft[self._current])
        self.text.edit_reset()
        self.text.edit_modified(False)
        self._loading_text = False
        self.editor_title.set(f"Шаблон {self._current} · Ctrl+{self._current}")

    def _text_changed(self, event=None):
        if not self.text.edit_modified():
            return
        self.text.edit_modified(False)
        if not self._loading_text:
            self._commit_text()
            self._refresh_rows()

    def _selection_changed(self, event=None):
        selected = self.template_list.selection()
        if not selected or selected[0] == self._current:
            return
        self._commit_text()
        self._current = selected[0]
        self._load_text()
        self._refresh_rows()

    def select_template(self, number):
        self.template_list.selection_set(str(number))
        self._selection_changed()
        self.text.focus_set()

    def _refresh_rows(self):
        for key, text in self._draft.items():
            preview = " ".join(text.split())[:85] or "Пустой шаблон"
            if text != self._saved[key]:
                state = "Не сохранён"
            elif not text.strip():
                state = "Свободно"
            elif int(key) in self._failed:
                state = "Клавиши заняты"
            elif int(key) in self._registered:
                state = "Активен"
            else:
                state = "Ожидание"
            self.template_list.item(key, values=(f"Ctrl+{key}", preview, state))
        suffix = " • Есть изменения" if self.dirty else ""
        self.title("Мои шаблоны — Hotkeys" + suffix)
        state = "normal" if self.dirty else "disabled"
        self.save_button.configure(state=state)
        self.revert_button.configure(state=state)
        self.save_status.set("Есть несохранённые изменения" if self.dirty else "Все изменения сохранены")

    def insert_placeholder(self, token):
        if self.text.tag_ranges("sel"):
            self.text.delete("sel.first", "sel.last")
        self.text.insert("insert", token)
        self._commit_text()
        self._refresh_rows()
        self.text.focus_set()

    def clear_template(self):
        self._commit_text()
        if self._draft[self._current] and messagebox.askyesno(
                "Очистить шаблон?", "Текст будет очищен. Сочетание освободится после сохранения.", parent=self):
            self._draft[self._current] = ""
            self._load_text()
            self._refresh_rows()

    def _save_shortcut(self, event):
        self.save_templates()
        return "break"

    def save_templates(self):
        self._commit_text()
        try:
            self.store.save(self._draft)
        except Exception as exc:
            messagebox.showerror("Не удалось сохранить", f"Шаблоны не изменены. Попробуйте ещё раз.\n\n{exc}", parent=self)
            self._refresh_rows()
            return False
        set_templates(self._draft)
        self._saved = dict(self._draft)
        self._refresh_rows()
        self.save_status.set("Сохранено. Настраиваем горячие клавиши…")
        self.service.refresh()
        return True

    def revert_changes(self):
        self._commit_text()
        if self.dirty and not messagebox.askyesno("Отменить изменения?", "Вернуть все шаблоны к последнему сохранению?", parent=self):
            return False
        self._draft = dict(self._saved)
        self._load_text()
        self._refresh_rows()
        return True

    def import_from_word(self):
        selected = filedialog.askopenfilename(parent=self, title="Выберите прежний файл шаблонов", filetypes=[("Документы Word", "*.docx")])
        if not selected:
            return
        try:
            imported = import_docx_templates(Path(selected))
            if not imported:
                raise ValueError("В документе нет непустых абзацев с текстом.")
        except Exception as exc:
            messagebox.showerror("Не удалось импортировать", str(exc), parent=self)
            return
        if not messagebox.askyesno(
                "Импорт шаблонов", f"Заменить тексты в редакторе первыми {len(imported)} шаблонами из Word?\n"
                "Текущие несохранённые правки будут заменены.\nПосле проверки нажмите «Сохранить».", parent=self):
            return
        self._draft = {str(i): imported.get(str(i), "") for i in range(1, MAX_TEMPLATES + 1)}
        self._load_text()
        self._refresh_rows()

    def _confirm_unsaved(self):
        self._commit_text()
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel("Сохранить изменения?", "Сохранить изменения в шаблонах перед закрытием окна?", parent=self)
        if answer is None:
            return False
        if answer:
            return self.save_templates()
        self._draft = dict(self._saved)
        self._load_text()
        self._refresh_rows()
        return True

    def hide_editor(self):
        if self._confirm_unsaved():
            # Keep a taskbar entry if the shortcut needed to reopen us is unavailable.
            if 100 in self._registered:
                self.withdraw()
            else:
                self.iconify()

    def show_editor(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self.text.focus_set()

    def exit_program(self):
        if not self._confirm_unsaved():
            return
        self._closing = True
        _shutdown_event.set()
        self.service.stop()
        if self._poll_id:
            self.after_cancel(self._poll_id)
            self._poll_id = None
        self.destroy()

    def _poll_events(self):
        self._poll_id = None
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "show":
                    self.show_editor()
                elif kind == "exit":
                    self.show_editor()
                    self.exit_program()
                    if self._closing:
                        return
                elif kind == "paste":
                    if own_window_is_foreground():
                        self.select_template(payload)
                    else:
                        on_hotkey_number(payload)
                elif kind == "hotkeys":
                    self._registered = payload["registered"]
                    self._failed = payload["failed"]
                    self._refresh_rows()
                    if self._failed:
                        self.runtime_status.set("Сочетания заняты другой программой: " + "; ".join(self._failed.values()))
                        self.runtime_label.configure(foreground="#A44800")
                    else:
                        self.runtime_status.set("Горячие клавиши работают · Excel: Ctrl+Alt+D · Очистка: Ctrl+Alt+L")
                        self.runtime_label.configure(foreground="#147445")
                elif kind == "error":
                    self.show_editor()
                    messagebox.showerror("Горячие клавиши", payload, parent=self)
                elif kind == "stopped" and not self._closing:
                    self.runtime_status.set("Горячие клавиши остановлены. Шаблоны можно редактировать и сохранять; затем перезапустите программу.")
                    self.runtime_label.configure(foreground="#A44800")
        except queue.Empty:
            pass
        if not self._closing:
            self._poll_id = self.after(80, self._poll_events)


def main():
    if os.name != "nt":
        raise SystemExit("Горячие клавиши и автоматизация Excel требуют Windows.")
    configure_windows_api()
    store = TemplateStore()
    try:
        set_templates(store.load())
    except Exception as exc:
        error_box(f"Не удалось прочитать сохранённые шаблоны. Файл не изменён:\n{store.path}\n\n{exc}")
        return
    _shutdown_event.clear()
    events = queue.Queue()
    service = HotkeyService(events)
    editor = TemplateEditor(store, service, events)
    _init_clip_history()
    watcher = threading.Thread(target=_clipboard_watcher, name="ClipboardHistory", daemon=True)
    watcher.start()
    service.start()
    try:
        editor.mainloop()
    finally:
        _shutdown_event.set()
        service.stop()
        if service.thread is not None:
            service.thread.join(timeout=2)
        watcher.join(timeout=1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        _shutdown_event.set()
