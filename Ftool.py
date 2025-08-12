# ftool.py — NFTOOL with centered headers, mouse+keyboard hotkeys, profiles (autosave),
import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import time
import threading
import sys
import os
import json
import ctypes
from ctypes import wintypes

# ---- Auto-Update (GitHub Releases) ----
import requests
import tempfile
import shutil
import subprocess
from packaging import version

# ====== APP / UPDATE SETTINGS ======
VERSION = "1.0.0"               # <— deine aktuelle App-Version hier pflegen
GITHUB_OWNER = "Nossigit"
GITHUB_REPO  = "NFTOOL"
RELEASE_ASSET_NAME = "NFTOOL.exe"  #
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "").strip()  # optional, um Rate Limits zu vermeiden

# ---------- THEME ----------
BG = "#36393F"; PANEL = "#2F3136"; PANEL_HOVER = "#3A3D43"
FG = "#DCDDDE"; SUBTLE = "#B9BBBE"; INPUT_BG = "#202225"; INPUT_FG = "#FFFFFF"

# ---------- LAYOUT ----------
FONT_MAIN = ("Segoe UI", 10)
FONT_SUB  = ("Segoe UI", 9)
BTN_SIZE  = 56
ROW_PADY  = 2
ROW_IPADY = 2
CELL_PADX = 6
LEFT_PAD  = 12     # linker Einzug für "Window"
TOP_PADX  = 10
HEADER_H  = 18     # Höhe der Header-Zeile
DEFAULT_INTERVAL = "1000"

KEY_OPTIONS = [str(i) for i in range(0, 10)] + [f"F{i}" for i in range(1, 9)]
VALID_KEYS = {k: (getattr(win32con, f"VK_{k}") if k.startswith("F") else ord(k)) for k in KEY_OPTIONS}
BROWSER_KEYWORDS = ["Chrome", "Firefox", "Edge", "Opera"]

PROFILES_FILE = "ftool_profiles.json"
DEFAULT_PROFILE_NAMES = [f"Profile {i}" for i in range(1, 6)]

# ---------- Windows list ----------
def enum_windows(only_filtered=True):
    items = []
    def proc(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if not title: return True
            if only_filtered:
                if ("Flyff" in title) or ("Insanity" in title):
                    if any(b in title for b in BROWSER_KEYWORDS): return True
                    items.append((hwnd, title))
            else:
                items.append((hwnd, title))
        return True
    win32gui.EnumWindows(proc, None)
    return items

# ---------- send key ----------
def send_key(hwnd, vk):
    try:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
        time.sleep(0.01)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)
    except Exception:
        pass

# ---------- Dark Titlebar ----------
def enable_dark_titlebar(tk_root):
    try:
        hwnd = tk_root.winfo_id()
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19
        value = ctypes.c_int(1)
        dwmapi = ctypes.windll.dwmapi
        res = dwmapi.DwmSetWindowAttribute(
            wintypes.HWND(hwnd),
            wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE),
            ctypes.byref(value),
            ctypes.sizeof(value)
        )
        if res != 0:
            dwmapi.DwmSetWindowAttribute(
                wintypes.HWND(hwnd),
                wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE_OLD),
                ctypes.byref(value),
                ctypes.sizeof(value)
            )
    except Exception:
        pass

# ---------- Hotkey backend (keyboard + mouse) ----------
MOD_ALT=0x0001; MOD_CONTROL=0x0002; MOD_SHIFT=0x0004; MOD_WIN=0x0008
WM_HOTKEY=0x0312
WH_MOUSE_LL = 14
WM_LBUTTONDOWN=0x0201; WM_RBUTTONDOWN=0x0204; WM_MBUTTONDOWN=0x0207
WM_XBUTTONDOWN=0x020B
XBUTTON1 = 0x0001
XBUTTON2 = 0x0002

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]

MOUSE_NAME_TO_CODE = {
    "MOUSE1": "L", "LBUTTON": "L", "LEFT": "L",
    "MOUSE2": "R", "RBUTTON": "R", "RIGHT": "R",
    "MOUSE3": "M", "MBUTTON": "M", "MIDDLE": "M",
    "MOUSE4": "X1", "MB4": "X1", "XBUTTON1": "X1", "X1": "X1",
    "MOUSE5": "X2", "MB5": "X2", "XBUTTON2": "X2", "X2": "X2",
}
MOUSE_CODE_TO_HUMAN = {"L":"Mouse1", "R":"Mouse2", "M":"Mouse3", "X1":"Mouse4", "X2":"Mouse5"}

def get_mods_state():
    user32 = ctypes.windll.user32
    def down(vk):
        return (user32.GetAsyncKeyState(vk) & 0x8000) != 0
    mods = 0
    if down(win32con.VK_CONTROL): mods |= MOD_CONTROL
    if down(win32con.VK_MENU):    mods |= MOD_ALT
    if down(win32con.VK_SHIFT):   mods |= MOD_SHIFT
    if down(win32con.VK_LWIN) or down(win32con.VK_RWIN): mods |= MOD_WIN
    return mods

def mods_to_str(mods):
    parts=[]
    if mods & MOD_CONTROL: parts.append("Ctrl")
    if mods & MOD_ALT:     parts.append("Alt")
    if mods & MOD_SHIFT:   parts.append("Shift")
    if mods & MOD_WIN:     parts.append("Win")
    return "+".join(parts)

def parse_hotkey(s: str):
    """
    Returns dict:
      {"type":"keyboard","mods":int,"vk":int}  or
      {"type":"mouse","mods":int,"button":"L|R|M|X1|X2"}
    """
    if not s: raise ValueError("Empty hotkey")
    parts = [p.strip() for p in s.split("+") if p.strip()]
    mods = 0
    key = None
    mouse_btn = None
    for p in parts:
        lp = p.lower()
        if lp in ("ctrl","control"): mods |= MOD_CONTROL
        elif lp == "alt":            mods |= MOD_ALT
        elif lp == "shift":          mods |= MOD_SHIFT
        elif lp in ("win","windows","meta"): mods |= MOD_WIN
        else:
            up = p.upper()
            if up in MOUSE_NAME_TO_CODE:
                mouse_btn = MOUSE_NAME_TO_CODE[up]
            else:
                k = p.upper()
                if k.startswith("NUMPAD"):
                    tail = k[6:]
                    if tail.isdigit(): k = "NUM"+tail
                    elif tail in ("ADD","+"): k = "NUM+"
                    elif tail in ("SUBTRACT","-"): k = "NUM-"
                    elif tail in ("MULTIPLY","*"): k = "NUM*"
                    elif tail in ("DIVIDE","/"): k = "NUM/"
                    elif tail in ("DECIMAL",".",","): k = "NUM."
                if k.startswith("F") and k[1:].isdigit():
                    key = getattr(win32con, f"VK_{k}", None)
                elif len(k)==1 and 'A'<=k<='Z':
                    key = ord(k)
                elif len(k)==1 and '0'<=k<='9':
                    key = ord(k)
                else:
                    vk_map_extra = {
                        "NUM0": win32con.VK_NUMPAD0, "NUM1": win32con.VK_NUMPAD1, "NUM2": win32con.VK_NUMPAD2,
                        "NUM3": win32con.VK_NUMPAD3, "NUM4": win32con.VK_NUMPAD4, "NUM5": win32con.VK_NUMPAD5,
                        "NUM6": win32con.VK_NUMPAD6, "NUM7": win32con.VK_NUMPAD7, "NUM8": win32con.VK_NUMPAD8,
                        "NUM9": win32con.VK_NUMPAD9, "NUM+": win32con.VK_ADD, "NUM-": win32con.VK_SUBTRACT,
                        "NUM*": win32con.VK_MULTIPLY, "NUM/": win32con.VK_DIVIDE, "NUM.": win32con.VK_DECIMAL,
                    }
                    key = vk_map_extra.get(k)
    if mouse_btn:
        return {"type":"mouse","mods":mods,"button":mouse_btn}
    if key is not None:
        return {"type":"keyboard","mods":mods,"vk":key}
    raise ValueError("Unsupported hotkey")

class KeyboardHotkeyThread(threading.Thread):
    def __init__(self, mods, vk, callback, tk_root):
        super().__init__(daemon=True)
        self.mods = mods; self.vk = vk
        self.callback = callback; self.tk_root = tk_root
        self.stop_event = threading.Event()
        self.user32 = ctypes.windll.user32
        self.id = 1
    def run(self):
        if not self.user32.RegisterHotKey(None, self.id, self.mods, self.vk):
            self.tk_root.after(0, lambda: messagebox.showwarning("Hotkey", "Failed to register hotkey (already in use?)."))
            return
        msg = wintypes.MSG()
        while not self.stop_event.is_set():
            while self.user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0, 1):
                if msg.message == WM_HOTKEY:
                    self.tk_root.after(0, self.callback)
            time.sleep(0.01)
        self.user32.UnregisterHotKey(None, self.id)
    def stop(self): self.stop_event.set()

LowLevelMouseProc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

class MouseHotkeyThread(threading.Thread):
    def __init__(self, mods_required, button_code, callback, tk_root):
        super().__init__(daemon=True)
        self.mods_required = mods_required
        self.button_code = button_code
        self.callback = callback
        self.tk_root = tk_root
        self.stop_event = threading.Event()
        self.hook = None
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        self._proc = None
    def _match_button(self, wParam, lParam):
        if self.button_code in ("L","R","M"):
            target = {"L":WM_LBUTTONDOWN,"R":WM_RBUTTONDOWN,"M":WM_MBUTTONDOWN}[self.button_code]
            return wParam == target
        if self.button_code in ("X1","X2"):
            if wParam != WM_XBUTTONDOWN: return False
            info = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
            xbtn = (info.mouseData >> 16) & 0xFFFF
            return (xbtn == XBUTTON1 and self.button_code=="X1") or (xbtn == XBUTTON2 and self.button_code=="X2")
        return False
    def _proc_func(self, nCode, wParam, lParam):
        if nCode == 0:
            if self._match_button(wParam, lParam):
                mods_now = get_mods_state()
                if mods_now == self.mods_required:
                    self.tk_root.after(0, self.callback)
        return ctypes.windll.user32.CallNextHookEx(self.hook, nCode, wParam, lParam)
    def run(self):
        self._proc = LowLevelMouseProc(self._proc_func)
        self.hook = self.user32.SetWindowsHookExW(WH_MOUSE_LL, self._proc, self.kernel32.GetModuleHandleW(None), 0)
        if not self.hook:
            self.tk_root.after(0, lambda: messagebox.showwarning("Hotkey", "Failed to install mouse hook. Try running as Admin."))
            return
        msg = wintypes.MSG()
        while not self.stop_event.is_set():
            if self.user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0, 1):
                self.user32.TranslateMessage(ctypes.byref(msg))
                self.user32.DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.01)
        if self.hook:
            self.user32.UnhookWindowsHookEx(self.hook)
            self.hook = None
    def stop(self):
        self.stop_event.set()

# ---------- storage ----------
def profiles_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), PROFILES_FILE)
    else:
        return os.path.join(os.path.dirname(__file__), PROFILES_FILE)
def load_all_profiles():
    p = profiles_path()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}
def save_all_profiles(data):
    p = profiles_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Could not save profiles:\n{e}")
        return False

# ---------- Row ----------
class WindowRow:
    def __init__(self, parent, row_index, on_change_cb, play_img, pause_img):
        self.running = False; self.thread = None; self.hwnd = None
        self.play_img = play_img; self.pause_img = pause_img
        self.on_change_cb = on_change_cb

        self.row = tk.Frame(parent, bg=PANEL)
        self.row.grid(row=row_index, column=0, sticky="ew", padx=4, pady=ROW_PADY, ipady=ROW_IPADY)
        self.row.grid_columnconfigure(0, weight=1)

        self.combo = ttk.Combobox(self.row, state="readonly", width=32)
        self.combo.grid(row=0, column=0, sticky="ew", padx=(LEFT_PAD, CELL_PADX))
        self.combo.bind("<<ComboboxSelected>>", lambda e: self.on_change_cb())

        self.key_combo = ttk.Combobox(self.row, values=KEY_OPTIONS, state="readonly", width=5)
        self.key_combo.set(KEY_OPTIONS[0])
        self.key_combo.grid(row=0, column=1, padx=(0, CELL_PADX))
        self.key_combo.bind("<<ComboboxSelected>>", lambda e: self.on_change_cb())

        self.interval = tk.Entry(self.row, width=8, bg=INPUT_BG, fg=INPUT_FG,
                                 insertbackground=INPUT_FG, relief="flat", justify="center")
        self.interval.insert(0, DEFAULT_INTERVAL)
        self.interval.grid(row=0, column=2, padx=(0, CELL_PADX))
        self.interval.bind("<Return>", lambda e: self.on_change_cb())
        self.interval.bind("<FocusOut>", lambda e: self.on_change_cb())

        self.btn = tk.Button(self.row, image=self.play_img, relief="flat", bd=0,
                             bg=PANEL, activebackground=PANEL, cursor="hand2",
                             width=BTN_SIZE, height=BTN_SIZE, command=self.toggle)
        self.btn.grid(row=0, column=3, padx=(0, 8))
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=PANEL_HOVER, activebackground=PANEL_HOVER))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=PANEL, activebackground=PANEL))

    def toggle(self):
        if self.running: self.stop()
        else: self.start()

    def start(self):
        title = self.combo.get().strip()
        key = self.key_combo.get().strip()
        int_txt = self.interval.get().strip()
        if not title or not key or not int_txt:
            messagebox.showerror("Error", "Please select a window, key and interval before starting.")
            return
        if key not in VALID_KEYS:
            messagebox.showerror("Error", "Invalid key. Only 0-9 and F1-F8 allowed.")
            return
        try:
            interval = int(int_txt)
            if interval < 1: raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Interval must be an integer ≥ 1 (ms).")
            return

        hwnd = None
        for h, t in enum_windows(only_filtered=filter_var.get()):
            if t == title:
                hwnd = h; break
        if hwnd is None:
            messagebox.showerror("Error", f"Window '{title}' not found."); return

        self.hwnd = hwnd; self.running = True
        self.btn.config(image=self.pause_img)
        self.combo.config(state="disabled"); self.key_combo.config(state="disabled"); self.interval.config(state="disabled")

        def loop():
            while self.running:
                send_key(self.hwnd, VALID_KEYS[key])
                time.sleep(interval / 1000.0)
        self.thread = threading.Thread(target=loop, daemon=True); self.thread.start()

    def stop(self):
        self.running = False
        self.btn.config(image=self.play_img)
        self.combo.config(state="readonly"); self.key_combo.config(state="readonly"); self.interval.config(state="normal")
        if self.thread is not None:
            self.thread.join(timeout=0.05); self.thread = None

    def update_windows(self, titles):
        current = self.combo.get()
        vals = [""] + (titles if current in titles or current == "" else [current] + titles)
        self.combo["values"] = vals
        if current in vals: self.combo.set(current)
        else: self.combo.set("")

    def to_dict(self):  return {"title": self.combo.get(), "key": self.key_combo.get(), "interval": self.interval.get()}
    def from_dict(self, d):
        self.combo.set(d.get("title",""))
        k = d.get("key","")
        if k in KEY_OPTIONS: self.key_combo.set(k)
        self.interval.delete(0, tk.END); self.interval.insert(0, d.get("interval", DEFAULT_INTERVAL))
    def reset(self):
        self.stop(); self.combo.set(""); self.key_combo.set(KEY_OPTIONS[0])
        self.interval.delete(0, tk.END); self.interval.insert(0, DEFAULT_INTERVAL)

# ---------- App state / autosave ----------
current_profile_name = None; autosave_pending = None
def get_state_dict(): return {"filter": bool(filter_var.get()), "rows": [r.to_dict() for r in all_rows]}
def apply_state_dict(state):
    try:
        filter_var.set(bool(state.get("filter", True)))
        rows = state.get("rows", [])
        for r, d in zip(all_rows, rows): r.from_dict(d or {})
    except Exception as e:
        messagebox.showerror("Error", f"Could not apply profile:\n{e}")
def schedule_autosave(delay_ms=120):
    global autosave_pending
    if autosave_pending: root.after_cancel(autosave_pending)
    autosave_pending = root.after(delay_ms, autosave_now)
def autosave_now():
    global autosave_pending; autosave_pending = None
    if not current_profile_name: return
    data = load_all_profiles()
    data[current_profile_name] = get_state_dict()
    data["_last_profile"] = current_profile_name
    save_all_profiles(data)
def on_any_setting_changed():
    schedule_autosave(); refresh_now()
def on_profile_selected(event=None):
    global current_profile_name
    name = profile_combo.get().strip()
    if not name: return
    data = load_all_profiles(); current_profile_name = name
    if name in data: apply_state_dict(data[name])
    else:
        data[name] = get_state_dict(); data["_last_profile"] = name; save_all_profiles(data)
    refresh_now()
def rename_profile():
    old = profile_combo.get().strip()
    if not old: return messagebox.showerror("Error","Select a profile first.")
    top = tk.Toplevel(root); top.configure(bg=BG); top.title("Rename Profile")
    tk.Label(top, text=f"New name for '{old}':", bg=BG, fg=FG).pack(padx=8, pady=4)
    e = tk.Entry(top, bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG, relief="flat"); e.pack(padx=8, pady=(0,6)); e.focus_set()
    def do_rename():
        new = e.get().strip()
        if not new: return messagebox.showerror("Error","Name cannot be empty.")
        profs = load_all_profiles()
        if old in profs:
            profs[new] = profs.pop(old, {})
            if profs.get("_last_profile")==old: profs["_last_profile"]=new
            save_all_profiles(profs)
            vals = [k for k in profs.keys() if k!="_last_profile"]
            profile_combo["values"]=vals; profile_combo.set(new); on_profile_selected()
        top.destroy()
    tk.Button(top, text="Save", command=do_rename, bg=PANEL, fg=FG, relief="flat").pack(pady=(0,8))
def delete_profile():
    name = profile_combo.get().strip()
    if not name: return
    data = load_all_profiles()
    if name not in data: return messagebox.showwarning("Not found","Profile is empty or missing.")
    if messagebox.askyesno("Delete", f"Delete '{name}'?"):
        data.pop(name, None)
        if data.get("_last_profile")==name: data["_last_profile"]=None
        save_all_profiles(data)
        vals = [k for k in data.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
        profile_combo["values"]=vals; profile_combo.set(vals[0]); on_profile_selected()
def reset_all():
    for r in all_rows: r.reset()
    schedule_autosave(); refresh_now()
    messagebox.showinfo("Reset","All slots have been reset.")

# ---------- Toggle all ----------
def toggle_profile_play_pause():
    any_run = any(r.running for r in all_rows)
    if any_run:
        for r in all_rows:
            if r.running: r.stop()
    else:
        for r in all_rows:
            t=r.combo.get().strip(); k=r.key_combo.get().strip(); it=r.interval.get().strip()
            if not (t and k and it): continue
            try: iv=int(it)
            except: continue
            if iv<1: continue
            r.start()

# ---------- Auto-Update helpers ----------
def _github_headers():
    h = {"Accept": "application/vnd.github+json"}
    if GITHUB_TOKEN:
        h["Authorization"] = f"Bearer {GITHUB_TOKEN}"
    return h

def get_latest_release_info():
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    try:
        r = requests.get(url, headers=_github_headers(), timeout=10)
        r.raise_for_status()
        data = r.json()
        tag = data.get("tag_name", "").lstrip("v")
        asset_url = None
        for asset in data.get("assets", []):
            if asset.get("name") == RELEASE_ASSET_NAME:
                asset_url = asset.get("browser_download_url")
                break
        return tag, asset_url
    except Exception:
        return None, None

def is_newer_available(current: str, latest: str) -> bool:
    try:
        return version.parse(latest) > version.parse(current)
    except Exception:
        return False

def download_file(url: str, dest_path: str):
    with requests.get(url, headers=_github_headers(), stream=True, timeout=30) as r:
        r.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

def run_replace_and_restart(downloaded_path: str):
    if not getattr(sys, 'frozen', False):
        messagebox.showinfo("Update", "Update-Replace funktioniert nur aus der EXE. Starte bitte die gebaute EXE.")
        return
    current_exe = sys.executable
    temp_bat = os.path.join(tempfile.gettempdir(), "nftool_update_replace.bat")
    bat = fr"""@echo off
setlocal
:loop
tasklist /FI "IMAGENAME eq {os.path.basename(current_exe)}" | find /I "{os.path.basename(current_exe)}" >nul
if "%ERRORLEVEL%"=="0" (
  timeout /t 1 >nul
  goto loop
)
move /Y "{downloaded_path}" "{current_exe}" >nul
start "" "{current_exe}"
del "%~f0"
"""
    with open(temp_bat, "w", encoding="utf-8") as f:
        f.write(bat)
    subprocess.Popen(['cmd', '/c', temp_bat], creationflags=subprocess.CREATE_NO_WINDOW)
    root.after(300, root.destroy)

def check_for_updates(auto=False):
    latest, asset_url = get_latest_release_info()
    if not latest or not asset_url:
        if not auto:
            messagebox.showwarning("Update", "Konnte Release-Infos nicht laden.")
        return
    if not is_newer_available(VERSION, latest):
        if not auto:
            messagebox.showinfo("Update", f"Du bist aktuell (v{VERSION}).")
        return
    if not auto:
        ok = messagebox.askyesno("Update verfügbar", f"Neue Version v{latest} gefunden.\nJetzt herunterladen und installieren?")
        if not ok:
            return
    try:
        tmp_dir = tempfile.mkdtemp(prefix="nftool_update_")
        dest = os.path.join(tmp_dir, RELEASE_ASSET_NAME)
        download_file(asset_url, dest)
    except Exception as e:
        if not auto:
            messagebox.showerror("Update", f"Download fehlgeschlagen:\n{e}")
        return
    if getattr(sys, 'frozen', False):
        run_replace_and_restart(dest)
    else:
        if not auto:
            messagebox.showinfo("Update", f"Heruntergeladen nach:\n{dest}\n\nBaue/Starte bitte die EXE, um zu ersetzen.")

# ---------- Root ----------
root = tk.Tk()
root.title("Ftool by Nossi")
root.configure(bg=BG)
enable_dark_titlebar(root)

# icon
base = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
ico_path = os.path.join(base, "Ficon_multi.ico")
if os.path.exists(ico_path):
    try: root.iconbitmap(ico_path)
    except: pass

# ttk style
s = ttk.Style()
try: s.theme_use("clam")
except: pass
s.configure(".", font=FONT_MAIN)
s.configure("TCombobox", fieldbackground=INPUT_BG, background=INPUT_BG, foreground=INPUT_FG, arrowcolor=SUBTLE)
s.map("TCombobox", fieldbackground=[("readonly", INPUT_BG)], foreground=[("readonly", INPUT_FG)], background=[("readonly", INPUT_BG)])

# Menü (Help → Check for updates)
menubar = tk.Menu(root)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Check for updates…", command=lambda: check_for_updates(auto=False))
menubar.add_cascade(label="Help", menu=help_menu)
root.config(menu=menubar)
# stiller Autocheck 3s nach Start
root.after(3000, lambda: check_for_updates(auto=True))

# --- Flyff (hoch & kompakt) ---
flyff_row = tk.Frame(root, bg=BG)
flyff_row.pack(fill="x", padx=TOP_PADX, pady=(6,0))
filter_var = tk.BooleanVar(value=True)
tk.Checkbutton(flyff_row, text="Flyff", variable=filter_var, bg=BG, fg=FG,
               activebackground=BG, selectcolor=BG, font=("Segoe UI",10,"bold"),
               command=lambda:(on_any_setting_changed())).pack(side="left")

# --- Hotkey + Profile rechts (gleiche Höhe) ---
hotkey_row = tk.Frame(root, bg=BG); hotkey_row.pack(fill="x", padx=TOP_PADX, pady=(2,2))
left_hot = tk.Frame(hotkey_row, bg=BG); left_hot.pack(side="left")
tk.Label(left_hot, text="Global Hotkey:", bg=BG, fg=FG).pack(side="left")

def hover_on(b): b.config(bg=PANEL_HOVER, activebackground=PANEL_HOVER)
def hover_off(b): b.config(bg=PANEL, activebackground=PANEL)
record_btn = tk.Button(left_hot, text="Register", bg=PANEL, fg=FG, activebackground=PANEL_HOVER,
                       relief="flat", bd=0, cursor="hand2")
record_btn.pack(side="left", padx=(6,6))
record_btn.bind("<Enter>", lambda e: hover_on(record_btn))
record_btn.bind("<Leave>", lambda e: hover_off(record_btn))

hotkey_entry = tk.Entry(left_hot, width=24, bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG, relief="flat")
hotkey_entry.pack(side="left"); hotkey_entry.insert(0,"Ctrl+Alt+P")

right_prof = tk.Frame(hotkey_row, bg=BG); right_prof.pack(side="right")
existing = load_all_profiles()
vals = [k for k in existing.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
profile_combo = ttk.Combobox(right_prof, values=vals, state="readonly", width=16)
last_choice = existing.get("_last_profile")
profile_combo.set(last_choice if (last_choice and last_choice in existing) else vals[0])
profile_combo.pack(side="left", padx=(0,6))
profile_combo.bind("<<ComboboxSelected>>", on_profile_selected)
tk.Button(right_prof, text="Rename", bg=PANEL, fg=FG, activebackground=PANEL_HOVER,
          relief="flat", bd=0, cursor="hand2", command=rename_profile).pack(side="left", padx=(0,6))
tk.Button(right_prof, text="Delete", bg=PANEL, fg=FG, activebackground=PANEL_HOVER,
          relief="flat", bd=0, cursor="hand2", command=delete_profile).pack(side="left", padx=(0,6))
tk.Button(right_prof, text="Reset",  bg=PANEL, fg=FG, activebackground=PANEL_HOVER,
          relief="flat", bd=0, cursor="hand2", command=reset_all).pack(side="left")

# --- Hotkey logic (Record & Register) ---
hotkey_thread=None
capture_active=False
capture_label=None
mouse_capture_thread=None

def stop_mouse_capture_thread():
    global mouse_capture_thread
    if mouse_capture_thread:
        mouse_capture_thread.stop()
        mouse_capture_thread = None

def unregister_hotkey():
    global hotkey_thread
    if hotkey_thread is not None:
        hotkey_thread.stop(); hotkey_thread=None

def register_hotkey_from_text(text):
    unregister_hotkey()
    try:
        spec = parse_hotkey(text)
    except Exception as e:
        messagebox.showerror("Hotkey", f"Invalid hotkey: {e}"); return False
    if spec["type"] == "keyboard":
        thread = KeyboardHotkeyThread(spec["mods"], spec["vk"], toggle_profile_play_pause, root)
        thread.start()
    else:
        thread = MouseHotkeyThread(spec["mods"], spec["button"], toggle_profile_play_pause, root)
        thread.start()
    global hotkey_thread
    hotkey_thread = thread
    return True

def start_capture():
    # Aufnahme von Tastatur ODER Maus (globale Maus via temporärem Hook)
    global capture_active, capture_label, mouse_capture_thread
    if capture_active: return
    capture_active=True
    capture_label = tk.Label(left_hot, text="Press hotkey… (mouse supported)", bg=BG, fg="#7FB2FF")
    capture_label.pack(side="left", padx=(6,0))

    pressed={"ctrl":False,"alt":False,"shift":False}
    accepted=False

    # Keyboard capture
    def key_to_symbol(ev):
        ks=ev.keysym.upper()
        if ks.startswith("CONTROL"): return "CTRL"
        if ks.startswith("ALT"): return "ALT"
        if ks.startswith("SHIFT"): return "SHIFT"
        if ks.startswith("KP_"):
            tail=ks[3:]
            if tail.isdigit(): return "NUM"+tail
            if tail=="ADD": return "NUM+"
            if tail=="SUBTRACT": return "NUM-"
            if tail=="MULTIPLY": return "NUM*"
            if tail=="DIVIDE": return "NUM/"
            if tail in ("DECIMAL","SEPARATOR"): return "NUM."
        if ks.startswith("F") and ks[1:].isdigit(): return ks
        if len(ks)==1 and ('A'<=ks<='Z' or '0'<=ks<='9'): return ks
        return None

    def commit_combo(mods_list, main_token):
        nonlocal accepted
        if accepted: return
        parts = mods_list[:]
        parts.append(main_token)
        combo = "+".join(parts)
        hotkey_entry.delete(0, tk.END)
        hotkey_entry.insert(0, combo)
        register_hotkey_from_text(combo)
        accepted = True
        stop_capture()

    def on_key_press(ev):
        nonlocal accepted
        global capture_active
        if not capture_active or accepted: return
        sym=key_to_symbol(ev)
        if sym=="CTRL": pressed["ctrl"]=True; return
        if sym=="ALT":  pressed["alt"]=True; return
        if sym=="SHIFT": pressed["shift"]=True; return
        if sym is None: return
        mods=[]
        if pressed["ctrl"]: mods.append("Ctrl")
        if pressed["alt"]:  mods.append("Alt")
        if pressed["shift"]:mods.append("Shift")
        if not mods:
            pass
        shown = "Num"+sym[3:] if sym.startswith("NUM") and len(sym)>3 else sym
        commit_combo(mods, shown)

    def on_key_release(ev):
        ks=ev.keysym.upper()
        if ks.startswith("CONTROL"): pressed["ctrl"]=False
        elif ks.startswith("ALT"):    pressed["alt"]=False
        elif ks.startswith("SHIFT"):  pressed["shift"]=False

    class _TempMouseCapture(MouseHotkeyThread):
        def __init__(self, cb):
            super().__init__(mods_required=0, button_code="X1", callback=None, tk_root=root)
            self._user_cb = cb
        def _match_any_and_commit(self, wParam, lParam):
            if wParam in (WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN):
                map_btn = {WM_LBUTTONDOWN:"L", WM_RBUTTONDOWN:"R", WM_MBUTTONDOWN:"M"}[wParam]
                return map_btn
            if wParam == WM_XBUTTONDOWN:
                info = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                xbtn = (info.mouseData >> 16) & 0xFFFF
                if xbtn == XBUTTON1: return "X1"
                if xbtn == XBUTTON2: return "X2"
            return None
        def _proc_func(self, nCode, wParam, lParam):
            if nCode == 0:
                btn = self._match_any_and_commit(wParam, lParam)
                if btn:
                    mods_now = get_mods_state()
                    parts = []
                    if mods_now & MOD_CONTROL: parts.append("Ctrl")
                    if mods_now & MOD_ALT:     parts.append("Alt")
                    if mods_now & MOD_SHIFT:   parts.append("Shift")
                    if mods_now & MOD_WIN:     parts.append("Win")
                    parts.append(MOUSE_CODE_TO_HUMAN[btn])
                    combo = "+".join(parts)
                    self._user_cb(combo)
                    root.after(0, stop_capture)
            return ctypes.windll.user32.CallNextHookEx(self.hook, nCode, wParam, lParam)

    def stop_capture(_ev=None):
        global capture_active
        root.unbind_all("<KeyPress>"); root.unbind_all("<KeyRelease>"); root.unbind("<FocusOut>")
        stop_mouse_capture_thread()
        if capture_label:
            try: capture_label.destroy()
            except: pass
        capture_active=False

    root.bind_all("<KeyPress>", on_key_press)
    root.bind_all("<KeyRelease>", on_key_release)
    root.bind("<FocusOut>", stop_capture)

    def _mouse_commit(combo_text):
        hotkey_entry.delete(0, tk.END)
        hotkey_entry.insert(0, combo_text)
        register_hotkey_from_text(combo_text)

    global mouse_capture_thread
    mouse_capture_thread = _TempMouseCapture(_mouse_commit)
    mouse_capture_thread.start()

record_btn.config(command=start_capture)

# --- Content (2 Spalten) + Headerbar (Canvas) mit auto-Alignment (CENTERED) ---
content = tk.Frame(root, bg=BG); content.pack(padx=TOP_PADX, pady=(0,2))

def build_column(parent, play_img, pause_img):
    col = tk.Frame(parent, bg=BG); col.pack(side="left", padx=4, anchor="n")

    header_bar = tk.Canvas(col, height=HEADER_H, bg=BG, highlightthickness=0, bd=0)
    header_bar.pack(fill="x", pady=(0,0))

    rows_holder = tk.Frame(col, bg=BG); rows_holder.pack(fill="x")

    rows=[]
    for i in range(5):
        rows.append(WindowRow(rows_holder, i+1, on_any_setting_changed, play_img, pause_img))

    def place_headers():
        if not rows: return
        root.update_idletasks()
        r0 = rows[0]
        def relx(w): return w.winfo_rootx() - col.winfo_rootx()
        x_window_center   = relx(r0.combo)     + r0.combo.winfo_width()     // 2
        x_key_center      = relx(r0.key_combo) + r0.key_combo.winfo_width() // 2
        x_interval_center = relx(r0.interval)  + r0.interval.winfo_width()  // 2
        x_btn_center      = relx(r0.btn)       + r0.btn.winfo_width()       // 2

        header_bar.delete("all")
        y = HEADER_H - 4
        header_bar.create_text(x_window_center,   y, text="Window",     fill=SUBTLE, font=FONT_SUB, anchor="s")
        header_bar.create_text(x_key_center,      y, text="Key",        fill=SUBTLE, font=FONT_SUB, anchor="s")
        header_bar.create_text(x_interval_center, y, text="Interval",   fill=SUBTLE, font=FONT_SUB, anchor="s")
        header_bar.create_text(x_btn_center,      y, text="Play/Pause", fill=SUBTLE, font=FONT_SUB, anchor="s")

    root.after(50, place_headers)
    col.bind("<Configure>", lambda e: place_headers())

    return rows

# load icons
if getattr(sys, 'frozen', False): base = sys._MEIPASS
else: base = os.path.dirname(__file__)
play_path = os.path.join(base, "play_darkmode_smooth.png")
pause_path = os.path.join(base, "pause_darkmode_smooth.png")
if not (os.path.exists(play_path) and os.path.exists(pause_path)):
    messagebox.showwarning("Icons missing", "Please place 'play_darkmode_smooth.png' and 'pause_darkmode_smooth.png' next to the script/EXE.")
play_img = tk.PhotoImage(file=play_path) if os.path.exists(play_path) else tk.PhotoImage(width=BTN_SIZE, height=BTN_SIZE)
pause_img = tk.PhotoImage(file=pause_path) if os.path.exists(pause_path) else tk.PhotoImage(width=BTN_SIZE, height=BTN_SIZE)

left_rows  = build_column(content, play_img, pause_img)
right_rows = build_column(content, play_img, pause_img)
all_rows = left_rows + right_rows

# --------- Window list refresh ----------
_refresh_job=None
def refresh_window_lists():
    global _refresh_job
    titles=[t for _, t in enum_windows(only_filtered=filter_var.get())]
    for r in all_rows: r.update_windows(titles)
    _refresh_job = root.after(2000, refresh_window_lists)
def refresh_now():
    global _refresh_job
    if _refresh_job:
        try: root.after_cancel(_refresh_job)
        except Exception: pass
        _refresh_job=None
    refresh_window_lists()
    root.update_idletasks()
    w = content.winfo_reqwidth() + 2*TOP_PADX
    h = flyff_row.winfo_reqheight() + hotkey_row.winfo_reqheight() + content.winfo_reqheight() + 6
    root.geometry(f"{max(760,w)}x{max(400,h)}")

# init profiles
def init_profiles_and_autoload():
    global current_profile_name
    data = load_all_profiles()
    vals = [k for k in data.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
    profile_combo["values"]=vals
    last = data.get("_last_profile")
    if last and last in data:
        profile_combo.set(last); current_profile_name=last; apply_state_dict(data[last])
    else:
        sel = profile_combo.get().strip() or vals[0]
        profile_combo.set(sel); current_profile_name=sel
        if sel not in data:
            data[sel]=get_state_dict(); data["_last_profile"]=sel; save_all_profiles(data)
    refresh_now()

# go
init_profiles_and_autoload()

def on_close():
    try: unregister_hotkey()
    except: pass
    stop_mouse_capture_thread()
    for r in all_rows: r.running=False
    time.sleep(0.05); root.destroy()
root.protocol("WM_DELETE_WINDOW", on_close)

# initial compact sizing
root.update_idletasks()
w = content.winfo_reqwidth() + 2*TOP_PADX
h = flyff_row.winfo_reqheight() + hotkey_row.winfo_reqheight() + content.winfo_reqheight() + 6
root.geometry(f"{max(760,w)}x{max(400,h)}")
root.minsize(740, 380)

# app icon again
ico_path = os.path.join(base, "Ficon_multi.ico")
if os.path.exists(ico_path):
    try: root.iconbitmap(ico_path)
    except: pass

root.mainloop()
