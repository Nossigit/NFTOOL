# -*- coding: utf-8 -*-
# NFTOOL.py - compact, DPI-safe, no-flicker UI, locked headers (fixed offsets),
# per-slot hotkeys (capture text inside field), profiles (autosave),
# GitHub updater, Discord-dark look.

import tkinter as tk
from tkinter import ttk, messagebox
import win32gui, win32con
import time, threading, sys, os, json, ctypes, requests, tempfile, subprocess
from ctypes import wintypes
from packaging import version

# ====== APP / UPDATE SETTINGS ======
VERSION = "0.9.1"
GITHUB_OWNER = "Nossigit"
GITHUB_REPO  = "NFTOOL"
RELEASE_ASSET_NAME = "NFTOOL.exe"

# ---------- THEME ----------
BG = "#36393F"; PANEL = "#2F3136"; PANEL_HOVER = "#3A3D43"
FG = "#DCDDDE"; SUBTLE = "#B9BBBE"; INPUT_BG = "#202225"; INPUT_FG = "#FFFFFF"
TOPBAR_BG = "#2F3136"; TOPBAR_HOVER = "#3A3D43"

# ---------- LAYOUT ----------
FONT_MAIN = ("Segoe UI", 10); FONT_SUB = ("Segoe UI", 9)
BTN_SIZE=56; ROW_PADY=2; ROW_IPADY=2; CELL_PADX=6
LEFT_PAD=12; LEFT_PADX=10; RIGHT_PADX=4
HEADER_H=22

DEFAULT_INTERVAL="100"          # 100 ms
INTERVAL_MAX_MS=30000           # 30 s
BROWSER_KEYWORDS=["Chrome","Firefox","Edge","Opera"]
PROFILES_FILE="ftool_profiles.json"
DEFAULT_PROFILE_NAMES=[f"Profile {i}" for i in range(1,6)]

# ---- FIXED HEADER OFFSETS (Pixel) ----
HDR_X_WINDOW   = 120
HDR_X_KEY      = 30
HDR_X_INTERVAL = 25
HDR_X_PLAY     = 33
HDR_X_HOTKEY   = 10
HDR_Y_OFFSET   = -2

# ---- DPI awareness (fix misaligned headers in EXE) ----
def make_dpi_aware():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor v2
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()   # System DPI
        except Exception:
            pass

# ---------- GitHub helper ----------
def _load_token():
    base_dir = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
    p = os.path.join(base_dir,"token.txt")
    if os.path.exists(p):
        try:
            with open(p,"r",encoding="utf-8") as f:
                t=f.read().strip()
                if t: return t
        except: pass
    return os.environ.get("GITHUB_TOKEN","").strip() or None

def _github_headers():
    h={"Accept":"application/vnd.github+json","User-Agent":"NFTOOL-Updater"}
    tok=_load_token()
    if tok: h["Authorization"]=f"Bearer {tok}"
    return h

def get_latest_release_info():
    def pick_asset(assets):
        for a in assets or []:
            if a.get("name")==RELEASE_ASSET_NAME: return a.get("browser_download_url")
        for a in assets or []:
            n=(a.get("name") or "").lower()
            if n.endswith(".exe") or n.endswith(".zip"): return a.get("browser_download_url")
        return None
    url_latest=f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    try:
        r=requests.get(url_latest,headers=_github_headers(),timeout=12)
        if r.status_code==200:
            d=r.json(); tag=(d.get("tag_name") or "").lstrip("v")
            asset=pick_asset(d.get("assets"))
            if tag and asset: return tag,asset,None
        latest_err=f"{r.status_code}: {r.text[:200]}"
    except Exception as e:
        latest_err=f"EXC: {e}"
    url_list=f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
    try:
        r=requests.get(url_list,headers=_github_headers(),timeout=12)
        if r.status_code==200:
            for rel in r.json() or []:
                if rel.get("draft"): continue
                tag=(rel.get("tag_name") or "").lstrip("v")
                asset=pick_asset(rel.get("assets"))
                if tag and asset: return tag,asset,None
        list_err=f"{r.status_code}: {r.text[:200]}"
    except Exception as e:
        list_err=f"EXC: {e}"
    return None,None,f"latest={latest_err} | list={list_err}"

def run_replace_and_restart(downloaded_path):
    if not getattr(sys,'frozen',False):
        messagebox.showinfo("Update","Replacing can only run from the EXE. Please launch the built EXE.")
        return
    current_exe=sys.executable
    temp_bat=os.path.join(tempfile.gettempdir(),"nftool_update_replace.bat")
    bat=fr"""@echo off
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
    with open(temp_bat,"w",encoding="utf-8") as f: f.write(bat)
    subprocess.Popen(['cmd','/c',temp_bat],creationflags=subprocess.CREATE_NO_WINDOW)
    root.after(300, root.destroy)

def check_for_updates(auto=False):
    latest,asset_url,raw_err=get_latest_release_info()
    if not latest or not asset_url:
        if not auto:
            msg="Could not fetch release info."
            if raw_err: msg+=f"\n\nDetails:\n{raw_err}"
            messagebox.showwarning("Update",msg)
        return
    try: newer=version.parse(latest)>version.parse(VERSION)
    except Exception: newer=(latest!=VERSION)
    if not newer:
        if not auto: messagebox.showinfo("Update",f"You are up to date (v{VERSION}).")
        return
    if not auto and not messagebox.askyesno("Update available",f"New version v{latest} found.\nInstall now?"):
        return
    try:
        tmp=tempfile.mkdtemp(prefix="nftool_update_")
        name=os.path.basename(asset_url.split("?")[0]) or RELEASE_ASSET_NAME
        dest=os.path.join(tmp,name)
        with requests.get(asset_url,headers=_github_headers(),stream=True,timeout=30) as r:
            r.raise_for_status()
            with open(dest,"wb") as f:
                for ch in r.iter_content(8192):
                    if ch: f.write(ch)
        run_replace_and_restart(dest)
    except Exception as e:
        messagebox.showerror("Update",f"Download failed:\n{e}")

# ---------- Dark Titlebar ----------
def enable_dark_titlebar(tk_root):
    try:
        hwnd=tk_root.winfo_id()
        DWMWA_USE_IMMERSIVE_DARK_MODE=20; DWMWA_USE_IMMERSIVE_DARK_MODE_OLD=19
        value=ctypes.c_int(1); dwmapi=ctypes.windll.dwmapi
        res=dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd),wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE),
                                         ctypes.byref(value), ctypes.sizeof(value))
        if res!=0:
            dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd),wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE_OLD),
                                         ctypes.byref(value), ctypes.sizeof(value))
    except Exception: pass

# ---------- Hotkeys ----------
MOD_ALT=0x0001; MOD_CONTROL=0x0002; MOD_SHIFT=0x0004; MOD_WIN=0x0008
WM_HOTKEY=0x0312

class KeyboardHotkeyThread(threading.Thread):
    def __init__(self,mods,vk,callback,tk_root):
        super().__init__(daemon=True)
        self.mods=mods; self.vk=vk; self.callback=callback; self.tk_root=tk_root
        self.stop_event=threading.Event(); self.user32=ctypes.windll.user32; self.id=id(self)&0xFFFF
    def run(self):
        if not self.user32.RegisterHotKey(None,self.id,self.mods,self.vk):
            self.tk_root.after(0, lambda: messagebox.showwarning("Hotkey","Failed to register hotkey (already in use?).")); return
        msg=wintypes.MSG()
        while not self.stop_event.is_set():
            while self.user32.PeekMessageW(ctypes.byref(msg),0,0,0,1):
                if msg.message==WM_HOTKEY: self.tk_root.after(0,self.callback)
            time.sleep(0.01)
        self.user32.UnregisterHotKey(None,self.id)
    def stop(self): self.stop_event.set()

SPECIAL_VK={
    "SPACE":win32con.VK_SPACE,"ENTER":win32con.VK_RETURN,"RETURN":win32con.VK_RETURN,"TAB":win32con.VK_TAB,
    "ESC":win32con.VK_ESCAPE,"ESCAPE":win32con.VK_ESCAPE,"BACKSPACE":win32con.VK_BACK,
    "INSERT":win32con.VK_INSERT,"INS":win32con.VK_INSERT,"DELETE":win32con.VK_DELETE,"DEL":win32con.VK_DELETE,
    "HOME":win32con.VK_HOME,"END":win32con.VK_END,"PGUP":win32con.VK_PRIOR,"PAGEUP":win32con.VK_PRIOR,
    "PGDN":win32con.VK_NEXT,"PAGEDOWN":win32con.VK_NEXT,"UP":win32con.VK_UP,"DOWN":win32con.VK_DOWN,
    "LEFT":win32con.VK_LEFT,"RIGHT":win32con.VK_RIGHT,"WIN":win32con.VK_LWIN,"APPS":win32con.VK_APPS,
    "CAPS":win32con.VK_CAPITAL,"CAPSLOCK":win32con.VK_CAPITAL,"SCROLL":win32con.VK_SCROLL,"NUMLOCK":win32con.VK_NUMLOCK,
    "PRINT":win32con.VK_SNAPSHOT,"PRINTSCREEN":win32con.VK_SNAPSHOT,"PRTSC":win32con.VK_SNAPSHOT,"PAUSE":win32con.VK_PAUSE,
    "NUM0":win32con.VK_NUMPAD0,"NUM1":win32con.VK_NUMPAD1,"NUM2":win32con.VK_NUMPAD2,"NUM3":win32con.VK_NUMPAD3,
    "NUM4":win32con.VK_NUMPAD4,"NUM5":win32con.VK_NUMPAD5,"NUM6":win32con.VK_NUMPAD6,"NUM7":win32con.VK_NUMPAD7,
    "NUM8":win32con.VK_NUMPAD8,"NUM9":win32con.VK_NUMPAD9,"NUM+":win32con.VK_ADD,"NUM-":win32con.VK_SUBTRACT,
    "NUM*":win32con.VK_MULTIPLY,"NUM/":win32con.VK_DIVIDE,"NUM.":win32con.VK_DECIMAL,
    ";":0xBA,"=":0xBB,",":0xBC,"-":0xBD,".":0xBE,"/":0xBF,"`":0xC0,"[":0xDB,"\\":0xDC,"]":0xDD,"'":0xDE
}
def parse_key_token_to_vk(token):
    if not token: return None
    t=token.strip()
    if len(t)==1:
        ch=t.upper()
        if 'A'<=ch<='Z' or '0'<=ch<='9': return ord(ch)
        return SPECIAL_VK.get(t)
    up=t.upper().replace(" ","")
    if up in SPECIAL_VK: return SPECIAL_VK[up]
    if up.startswith("F") and up[1:].isdigit():
        n=int(up[1:])
        if 1<=n<=24: return getattr(win32con,f"VK_F{n}")
    if up.startswith("NUMPAD"):
        tail=up[6:]
        if tail.isdigit(): return SPECIAL_VK.get("NUM"+tail)
        m={"ADD":"NUM+","SUBTRACT":"NUM-","MULTIPLY":"NUM*","DIVIDE":"NUM/","DECIMAL":"NUM."}
        if tail in m: return SPECIAL_VK.get(m[tail])
    if up.startswith("NUM"): return SPECIAL_VK.get(up)
    return None

def send_key(hwnd,vk):
    try:
        win32gui.PostMessage(hwnd,win32con.WM_KEYDOWN,vk,0); time.sleep(0.01)
        win32gui.PostMessage(hwnd,win32con.WM_KEYUP,vk,0)
    except Exception: pass

def enum_windows(only_filtered=True,title_query=None):
    q=(title_query or "").strip().lower()
    def match(t): return (q in t.lower()) if q else True
    items=[]
    def proc(hwnd,_):
        if win32gui.IsWindowVisible(hwnd):
            t=win32gui.GetWindowText(hwnd)
            if not t: return True
            if only_filtered:
                if ("Flyff" in t) or ("Insanity" in t):
                    if any(b in t for b in BROWSER_KEYWORDS): return True
                    if match(t): items.append((hwnd,t))
            else:
                if match(t): items.append((hwnd,t))
        return True
    win32gui.EnumWindows(proc,None)
    return items

def profiles_path():
    if getattr(sys,'frozen',False): return os.path.join(os.path.dirname(sys.executable),PROFILES_FILE)
    return os.path.join(os.path.dirname(__file__),PROFILES_FILE)

def load_all_profiles():
    p=profiles_path()
    if os.path.exists(p):
        try:
            with open(p,"r",encoding="utf-8") as f: return json.load(f)
        except Exception: return {}
    return {}

def save_all_profiles(d):
    p=profiles_path()
    try:
        with open(p,"w",encoding="utf-8") as f: json.dump(d,f,indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Error",f"Could not save profiles:\n{e}")
        return False

_active_capture_owner=None
def begin_global_capture(owner):
    global _active_capture_owner
    if _active_capture_owner is not None: return False
    _active_capture_owner=owner
    for r in all_rows:
        if r is owner: continue
        r.hk_btn.config(state="disabled")
    return True

def end_global_capture(owner):
    global _active_capture_owner
    if _active_capture_owner is owner:
        _active_capture_owner=None
        for r in all_rows: r.hk_btn.config(state="normal")

def _mods_from_state():
    user32=ctypes.windll.user32
    def down(vk): return (user32.GetAsyncKeyState(vk)&0x8000)!=0
    m=0
    if down(win32con.VK_CONTROL): m|=MOD_CONTROL
    if down(win32con.VK_MENU):    m|=MOD_ALT
    if down(win32con.VK_SHIFT):   m|=MOD_SHIFT
    if down(win32con.VK_LWIN) or down(win32con.VK_RWIN): m|=MOD_WIN
    return m

class WindowRow:
    def __init__(self,parent,row_index,on_change,play_img,pause_img,trash_img):
        self.running=False; self.thread=None; self.hwnd=None
        self.play_img=play_img; self.pause_img=pause_img; self.trash_img=trash_img
        self.on_change=on_change; self.hk_thread=None; self.capturing=False
        self.row=tk.Frame(parent,bg=PANEL); self.row.grid(row=row_index,column=0,sticky="ew",padx=4,pady=ROW_PADY,ipady=ROW_IPADY)
        self.row.grid_columnconfigure(0,weight=1)

        self.combo=ttk.Combobox(self.row,state="readonly",width=40)
        self.combo.grid(row=0,column=0,sticky="ew",padx=(LEFT_PAD,CELL_PADX))
        self.combo.bind("<<ComboboxSelected>>", lambda e: self.on_change())

        self.key_entry=tk.Entry(self.row,width=10,bg=INPUT_BG,fg=INPUT_FG,insertbackground=INPUT_FG,relief="flat",justify="center")
        self.key_entry.insert(0,"F1"); self.key_entry.grid(row=0,column=1,padx=(0,CELL_PADX))
        self.key_entry.bind("<Return>", lambda e: self.on_change()); self.key_entry.bind("<FocusOut>", lambda e: self.on_change())

        self.interval=tk.Entry(self.row,width=8,bg=INPUT_BG,fg=INPUT_FG,insertbackground=INPUT_FG,relief="flat",justify="center")
        self.interval.insert(0,DEFAULT_INTERVAL); self.interval.grid(row=0,column=2,padx=(0,CELL_PADX))
        self.interval.bind("<Return>", lambda e: self.on_change()); self.interval.bind("<FocusOut>", lambda e: self.on_change())

        self.btn=tk.Button(self.row,image=self.play_img,relief="flat",bd=0,bg=PANEL,activebackground=PANEL,cursor="hand2",
                           width=BTN_SIZE,height=BTN_SIZE,command=self.toggle)
        self.btn.grid(row=0,column=3,padx=(0,6))
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=PANEL_HOVER,activebackground=PANEL_HOVER))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=PANEL,activebackground=PANEL))

        self.hk_btn=tk.Button(self.row,text="Set",bg=PANEL,fg=FG,activebackground=PANEL_HOVER,relief="flat",bd=0,cursor="hand2",command=self.start_capture)
        self.hk_btn.grid(row=0,column=4,padx=(0,4))
        self.hk_entry=tk.Entry(self.row,width=18,bg=INPUT_BG,fg=INPUT_FG,insertbackground=INPUT_FG,relief="flat")
        self.hk_entry.grid(row=0,column=5,padx=(0,8))
        self.hk_entry.bind("<Return>", lambda e: self.register_hotkey_from_text(self.hk_entry.get()))
        self.hk_entry.bind("<FocusOut>", lambda e: self._on_hotkey_entry_changed())
        self.hk_entry.bind("<KeyRelease>", lambda e: self._on_hotkey_entry_changed())

        self.del_btn=tk.Button(self.row,image=self.trash_img,bg=PANEL,activebackground=PANEL_HOVER,relief="flat",bd=0,cursor="hand2",
                               command=lambda: remove_row(self))
        self.del_btn.grid(row=0,column=6,padx=(0,8))

    # --- persistence helpers ---
    def to_dict(self):
        return {
            "title": self.combo.get(),
            "key": self.key_entry.get(),
            "interval": self.interval.get(),
            "hotkey": self.hk_entry.get()
        }

    def from_dict(self, d):
        self.combo.set(d.get("title",""))
        self.key_entry.delete(0, tk.END); self.key_entry.insert(0, d.get("key","F1"))
        self.interval.delete(0, tk.END); self.interval.insert(0, d.get("interval", DEFAULT_INTERVAL))
        self.hk_entry.delete(0, tk.END); self.hk_entry.insert(0, d.get("hotkey",""))
        hk = d.get("hotkey","").strip()
        if hk: self.register_hotkey_from_text(hk)
        else:  self.unregister_hotkey()

    def reset(self):
        self.stop(); self.combo.set("")
        self.key_entry.delete(0, tk.END); self.key_entry.insert(0, "F1")
        self.interval.delete(0, tk.END); self.interval.insert(0, DEFAULT_INTERVAL)
        self.hk_entry.delete(0, tk.END); self.unregister_hotkey()

    # --- run loop ---
    def toggle(self):
        if self.running: self.stop()
        else: self.start()

    def start(self):
        title=self.combo.get().strip(); key_txt=self.key_entry.get().strip(); int_txt=self.interval.get().strip()
        if not title or not key_txt or not int_txt:
            messagebox.showerror("Error","Please select a window, key and interval before starting."); return
        vk=parse_key_token_to_vk(key_txt)
        if vk is None:
            messagebox.showerror("Error",f"Unknown key: '{key_txt}'. Try e.g. A, F12, Space, Enter, Left, Num1, Num+, ;, ["); return
        try:
            interval=int(int_txt); interval=max(1,min(INTERVAL_MAX_MS,interval))
            if int_txt!=str(interval):
                self.interval.delete(0,tk.END); self.interval.insert(0,str(interval))
        except ValueError:
            messagebox.showerror("Error","Interval must be an integer (ms)."); return

        hwnd=None
        for h,t in enum_windows(only_filtered=filter_var.get(),title_query=title_filter_var.get()):
            if t==title: hwnd=h; break
        if hwnd is None:
            messagebox.showerror("Error",f"Window '{title}' not found."); return
        self.hwnd=hwnd; self.running=True
        self.btn.config(image=self.pause_img)
        self.combo.config(state="disabled"); self.key_entry.config(state="disabled"); self.interval.config(state="disabled")
        def loop():
            while self.running:
                send_key(self.hwnd,vk); time.sleep(interval/1000.0)
        self.thread=threading.Thread(target=loop,daemon=True); self.thread.start()

    def stop(self):
        self.running=False; self.btn.config(image=self.play_img)
        self.combo.config(state="readonly"); self.key_entry.config(state="normal"); self.interval.config(state="normal")
        if self.thread is not None: self.thread.join(timeout=0.05); self.thread=None

    # --- hotkey register/capture ---
    def unregister_hotkey(self):
        if self.hk_thread is not None:
            try: self.hk_thread.stop()
            except: pass
            self.hk_thread=None

    def register_hotkey_from_text(self,text):
        if not text.strip():
            self.unregister_hotkey(); return False
        self.unregister_hotkey()
        parts=[p.strip() for p in (text or "").split("+") if p.strip()]
        mods=0; main=None
        for p in parts:
            up=p.upper()
            if up in ("CTRL","CONTROL"): mods|=MOD_CONTROL
            elif up=="ALT": mods|=MOD_ALT
            elif up=="SHIFT": mods|=MOD_SHIFT
            elif up in ("WIN","WINDOWS","META"): mods|=MOD_WIN
            else: main=parse_key_token_to_vk(p)
        if main is None:
            messagebox.showerror("Hotkey",f"Invalid hotkey: '{text}'"); return False
        th=KeyboardHotkeyThread(mods,main,self.toggle,root); th.start(); self.hk_thread=th; return True

    def _on_hotkey_entry_changed(self):
        if self.hk_entry.get().strip()=="": self.unregister_hotkey()

    def start_capture(self):
        # allow only one capture globally
        if not begin_global_capture(self):
            messagebox.showwarning("Hotkey", "Another hotkey is currently being set. Finish it first.")
            return
        if self.capturing:
            return
        self.capturing = True

        # show placeholder in the entry itself
        placeholder = "Press hotkey..."
        prev_text = self.hk_entry.get()
        self._hk_prev_text = prev_text
        self.hk_entry.config(state="normal")
        self.hk_entry.delete(0, tk.END)
        self.hk_entry.insert(0, placeholder)
        self.hk_entry.config(fg="#7FB2FF")

        def on_key(ev):
            ks = ev.keysym
            if any(ks.upper().startswith(p) for p in ("CONTROL", "SHIFT", "ALT", "META", "WIN")):
                return
            mods = _mods_from_state()
            parts = []
            if mods & MOD_CONTROL: parts.append("Ctrl")
            if mods & MOD_ALT:     parts.append("Alt")
            if mods & MOD_SHIFT:   parts.append("Shift")
            if mods & MOD_WIN:     parts.append("Win")
            tk_to_token = {
                "Escape":"Esc", "Return":"Enter", "Prior":"PgUp", "Next":"PgDn", "BackSpace":"Backspace",
                "space":"Space", "Left":"Left", "Right":"Right", "Up":"Up", "Down":"Down"
            }
            token = tk_to_token.get(ks, ks)
            combo = "+".join(parts + [token]) if parts else token
            self.hk_entry.config(fg=INPUT_FG)
            self.hk_entry.delete(0, tk.END)
            self.hk_entry.insert(0, combo)
            self.register_hotkey_from_text(combo)
            stop_capture()

        def stop_capture(_=None):
            root.unbind_all("<KeyPress>")
            try:
                root.unbind("<FocusOut>", stop_id)
            except Exception:
                pass
            if self.hk_entry.get().strip() == placeholder:
                self.hk_entry.config(fg=INPUT_FG)
                self.hk_entry.delete(0, tk.END)
                self.hk_entry.insert(0, getattr(self, "_hk_prev_text", ""))
            self.capturing = False
            end_global_capture(self)

        root.bind_all("<KeyPress>", on_key)
        stop_id = root.bind("<FocusOut>", stop_capture)

    # --- combobox refresh ---
    def update_windows(self, titles):
        current = self.combo.get()
        vals = [""] + sorted(titles)
        if current and current not in vals:
            vals = [""] + [current] + [t for t in sorted(titles) if t != current]
        self.combo["values"] = vals
        if current in vals:
            self.combo.set(current)
        else:
            self.combo.set("")

# ---------- Root ----------
make_dpi_aware()
root=tk.Tk(); root.withdraw()
root.title("Ftool by Nossi"); root.configure(bg=BG)

def enable_icon(win):
    base = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
    ico_path=os.path.join(base,"Ficon_multi.ico")
    if os.path.exists(ico_path):
        try: win.iconbitmap(ico_path)
        except: pass

enable_dark_titlebar(root); enable_icon(root)

s=ttk.Style()
try: s.theme_use("clam")
except: pass
s.configure(".", font=FONT_MAIN)
s.configure("TCombobox", fieldbackground=INPUT_BG, background=INPUT_BG, foreground=INPUT_FG, arrowcolor=SUBTLE)
s.map("TCombobox", fieldbackground=[("readonly", INPUT_BG)], foreground=[("readonly", INPUT_FG)], background=[("readonly", INPUT_BG)])

# ---------- Topbar ----------
topbar=tk.Frame(root,bg=TOPBAR_BG,height=30); topbar.pack(side="top",fill="x")
help_btn=tk.Button(topbar,text="Help",bg=TOPBAR_BG,fg="#FFFFFF",activebackground=TOPBAR_HOVER,activeforeground="#FFFFFF",
                   relief="flat",bd=0,cursor="hand2",font=("Segoe UI",10))
help_btn.pack(side="left",padx=8,pady=3)
help_menu=tk.Menu(root,tearoff=0,bg=TOPBAR_BG,fg=FG,activebackground=TOPBAR_HOVER,activeforeground=FG,relief="flat",bd=0)
help_menu.add_command(label="Check for updates...",command=lambda: check_for_updates(auto=False))
help_menu.add_separator()
help_menu.add_command(label="About",command=lambda: messagebox.showinfo("About",f"NFTOOL v{VERSION}\nby Nossi"))
help_btn.config(command=lambda: help_menu.tk_popup(help_btn.winfo_rootx(), help_btn.winfo_rooty()+help_btn.winfo_height()))
ver_lbl=tk.Label(topbar,text=f"v{VERSION}",bg=TOPBAR_BG,fg=SUBTLE,font=("Segoe UI",10))
ver_lbl.pack(side="right",padx=4)

# ---------- Controls row ----------
top_row=tk.Frame(root,bg=BG); top_row.pack(fill="x",padx=(LEFT_PADX,RIGHT_PADX),pady=(6,4))
left_all=tk.Frame(top_row,bg=BG); left_all.pack(side="left")
filter_var=tk.BooleanVar(value=True)
tk.Checkbutton(left_all,text="Flyff",variable=filter_var,bg=BG,fg=FG,activebackground=BG,selectcolor=BG,font=("Segoe UI",10,"bold"),
               command=lambda: request_layout()).pack(side="left")

title_filter_var=tk.StringVar(value="")
tk.Label(left_all,text="Title filter:",bg=BG,fg=SUBTLE,font=FONT_SUB).pack(side="left",padx=(12,4))
title_filter_entry=tk.Entry(left_all,textvariable=title_filter_var,width=22,bg=INPUT_BG,fg=INPUT_FG,insertbackground=INPUT_FG,relief="flat")
title_filter_entry.pack(side="left",padx=(0,12))
title_filter_var.trace_add("write", lambda *_: request_layout())

profiles_wrap=tk.Frame(left_all,bg=BG); profiles_wrap.pack(side="left",padx=(20,0))
existing=load_all_profiles()
vals=[k for k in existing.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
profile_combo=ttk.Combobox(profiles_wrap,values=vals,state="readonly",width=16)
last_choice=existing.get("_last_profile")
profile_combo.set(last_choice if (last_choice and last_choice in existing) else vals[0]); profile_combo.pack(side="left",padx=(0,6))
def on_profile_selected(_=None):
    global current_profile_name
    name=profile_combo.get().strip()
    if not name: return
    data=load_all_profiles(); current_profile_name=name
    if name in data: apply_state_dict_exact(data[name])
    else:
        data[name]=get_state_dict(); data["_last_profile"]=name; save_all_profiles(data)
    request_layout()
profile_combo.bind("<<ComboboxSelected>>", on_profile_selected)
tk.Button(profiles_wrap,text="Rename",bg=PANEL,fg=FG,activebackground=PANEL_HOVER,relief="flat",bd=0,cursor="hand2",
          command=lambda: rename_profile()).pack(side="left",padx=(0,6))
tk.Button(profiles_wrap,text="Delete",bg=PANEL,fg=FG,activebackground=PANEL_HOVER,relief="flat",bd=0,cursor="hand2",
          command=lambda: delete_profile()).pack(side="left",padx=(0,6))
tk.Button(profiles_wrap,text="Reset",bg=PANEL,fg=FG,activebackground=PANEL_HOVER,relief="flat",bd=0,cursor="hand2",
          command=lambda: reset_all()).pack(side="left")

# ---------- Content & Header ----------
content=tk.Frame(root,bg=BG); content.pack(padx=(LEFT_PADX,RIGHT_PADX),pady=(0,0),fill="x")
header_canvas=tk.Canvas(content,height=HEADER_H,bg=BG,highlightthickness=0,bd=0); header_canvas.pack(fill="x")
rows_holder=tk.Frame(content,bg=BG); rows_holder.pack(fill="x")

# Plus-Bar
add_bar=tk.Frame(root,bg=BG); add_bar.pack(fill="x",padx=(LEFT_PADX,RIGHT_PADX),pady=(4,10))

# Icons
base = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
play_path=os.path.join(base,"play_darkmode_smooth.png")
pause_path=os.path.join(base,"pause_darkmode_smooth.png")
trash_path=os.path.join(base,"trash_dark_smooth.png")
add_path=os.path.join(base,"add_darkmode_smooth.png")
play_img  = tk.PhotoImage(file=play_path)  if os.path.exists(play_path)  else tk.PhotoImage(width=BTN_SIZE,height=BTN_SIZE)
pause_img = tk.PhotoImage(file=pause_path) if os.path.exists(pause_path) else tk.PhotoImage(width=BTN_SIZE,height=BTN_SIZE)
trash_img = tk.PhotoImage(file=trash_path) if os.path.exists(trash_path) else tk.PhotoImage(width=24,height=24)
add_icon  = tk.PhotoImage(file=add_path)   if os.path.exists(add_path)   else tk.PhotoImage(width=24,height=24)

add_btn=tk.Button(add_bar,image=add_icon,bg=BG,activebackground=BG,relief="flat",bd=0,cursor="hand2",
                  highlightthickness=0,command=lambda: add_row())
add_btn.pack()

# ---------- State / autosave ----------
all_rows=[]; current_profile_name=None; autosave_pending=None
def get_state_dict():
    return {"filter":bool(filter_var.get()),"title_query":title_filter_var.get(),"rows":[r.to_dict() for r in all_rows]}
def apply_state_dict_exact(state):
    try:
        filter_var.set(bool(state.get("filter",True))); title_filter_var.set(state.get("title_query",""))
        rows=state.get("rows",[]); set_row_count_exact(len(rows) if rows is not None else 5)
        for r in all_rows: r.reset()
        for r,d in zip(all_rows,rows): r.from_dict(d or {})
    except Exception as e:
        messagebox.showerror("Error",f"Could not apply profile:\n{e}")
def schedule_autosave(delay_ms=120):
    global autosave_pending
    if autosave_pending: root.after_cancel(autosave_pending)
    autosave_pending=root.after(delay_ms, autosave_now)
def autosave_now():
    global autosave_pending; autosave_pending=None
    if not current_profile_name: return
    data=load_all_profiles(); data[current_profile_name]=get_state_dict(); data["_last_profile"]=current_profile_name; save_all_profiles(data)
def on_any_setting_changed():
    schedule_autosave(); request_layout()

def rename_profile():
    old=profile_combo.get().strip()
    if not old:
        messagebox.showerror("Error","Select a profile first."); return
    top=tk.Toplevel(root); top.configure(bg=BG); top.title("Rename Profile")
    try:
        base = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
        top.iconbitmap(os.path.join(base,"Ficon_multi.ico"))
    except: pass
    tk.Label(top,text=f"New name for '{old}':",bg=BG,fg=FG).pack(padx=8,pady=4)
    e=tk.Entry(top,bg=INPUT_BG,fg=INPUT_FG,insertbackground=INPUT_FG,relief="flat"); e.pack(padx=8,pady=(0,6)); e.focus_set()
    def do_rename():
        new=e.get().strip()
        if not new:
            messagebox.showerror("Error","Name cannot be empty."); return
        profs=load_all_profiles()
        if old in profs:
            profs[new]=profs.pop(old,{})
            if profs.get("_last_profile")==old: profs["_last_profile"]=new
            save_all_profiles(profs)
            vals=[k for k in profs.keys() if k!="_last_profile"]
            profile_combo["values"]=vals; profile_combo.set(new); on_profile_selected()
        top.destroy()
    tk.Button(top,text="Save",command=do_rename,bg=PANEL,fg=FG,relief="flat").pack(pady=(0,8))

def delete_profile():
    name=profile_combo.get().strip()
    if not name: return
    data=load_all_profiles()
    if name not in data:
        messagebox.showwarning("Not found","Profile is empty or missing."); return
    if messagebox.askyesno("Delete",f"Delete '{name}'?"):
        data.pop(name,None)
        if data.get("_last_profile")==name: data["_last_profile"]=None
        save_all_profiles(data)
        vals=[k for k in data.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
        profile_combo["values"]=vals; profile_combo.set(vals[0]); on_profile_selected()

def reset_all():
    for r in all_rows: r.reset()
    schedule_autosave(); request_layout(); messagebox.showinfo("Reset","All slots have been reset.")

# ---------- Locked headers (DPI-safe, fixed offsets, one-time) ----------
headers_locked=False
def lock_headers_now():
    global headers_locked
    if not all_rows or headers_locked: return
    root.update_idletasks()
    r0=all_rows[0]

    def x_in_canvas(w):
        x=0; cur=w
        while cur is not None and cur is not root:
            x += cur.winfo_x()
            cur = cur.master
        cx=0; cur=header_canvas
        while cur is not None and cur is not root:
            cx += cur.winfo_x()
            cur = cur.master
        return x - cx

    def center_between(widgets):
        left=min(x_in_canvas(w) for w in widgets)
        right=max(x_in_canvas(w)+w.winfo_width() for w in widgets)
        return int((left+right)/2)

    x_window   = x_in_canvas(r0.combo)     + HDR_X_WINDOW
    x_key      = x_in_canvas(r0.key_entry) + HDR_X_KEY
    x_interval = x_in_canvas(r0.interval)  + HDR_X_INTERVAL
    x_play     = x_in_canvas(r0.btn)       + HDR_X_PLAY
    x_hotkey_c = center_between([r0.hk_btn, r0.hk_entry]) + HDR_X_HOTKEY

    header_canvas.delete("all")
    y=(HEADER_H-4)+HDR_Y_OFFSET
    header_canvas.create_text(x_window,   y, text="Window",     fill=SUBTLE, font=FONT_SUB, anchor="s")
    header_canvas.create_text(x_key,      y, text="Key",        fill=SUBTLE, font=FONT_SUB, anchor="s")
    header_canvas.create_text(x_interval, y, text="Interval",   fill=SUBTLE, font=FONT_SUB, anchor="s")
    header_canvas.create_text(x_play,     y, text="Play/Pause", fill=SUBTLE, font=FONT_SUB, anchor="s")
    header_canvas.create_text(x_hotkey_c, y, text="Hotkey",     fill=SUBTLE, font=FONT_SUB, anchor="s")
    headers_locked=True

# ---------- Debounced height-only resize ----------
locked_width=None
_last_height=0
_fit_job=None
def fit_height_only_now():
    global locked_width,_last_height
    root.update_idletasks()
    if locked_width is None:
        locked_width = max(700, content.winfo_reqwidth() + LEFT_PADX + RIGHT_PADX)
    h_req = (topbar.winfo_reqheight() + top_row.winfo_reqheight() +
             header_canvas.winfo_reqheight() + rows_holder.winfo_reqheight() +
             add_bar.winfo_reqheight() + 18)
    if h_req!=_last_height:
        root.geometry(f"{locked_width}x{h_req}")
        _last_height=h_req
def request_layout():
    global _fit_job
    if _fit_job:
        try: root.after_cancel(_fit_job)
        except: pass
        _fit_job=None
    _fit_job=root.after(40, fit_height_only_now)

# ---------- Rows mgmt ----------
def add_row():
    r=WindowRow(rows_holder,len(all_rows)+1,on_any_setting_changed,play_img,pause_img,trash_img)
    all_rows.append(r); schedule_autosave(); request_layout()
def set_row_count_exact(n):
    n=max(1,int(n))
    while len(all_rows)>n: remove_row(all_rows[-1],quiet=True)
    while len(all_rows)<n: add_row()
    request_layout()
def remove_row(row,quiet=False):
    if len(all_rows)<=1 and not quiet:
        messagebox.showinfo("Info","At least one slot is required."); return
    try: row.stop(); row.unregister_hotkey()
    except: pass
    if row in all_rows: all_rows.remove(row)
    try: row.row.destroy()
    except: pass
    for i,rr in enumerate(all_rows, start=1):
        try: rr.row.grid_configure(row=i)
        except: pass
    if not quiet:
        schedule_autosave(); request_layout()

# ---------- Init ----------
for _ in range(4): add_row()

def get_all_titles():
    return [t for _,t in enum_windows(only_filtered=filter_var.get(),title_query=title_filter_var.get())]

def refresh_window_lists():
    titles = get_all_titles()
    for r in all_rows:
        r.update_windows(titles)

def init_profiles_and_autoload():
    global current_profile_name
    data=load_all_profiles()
    vals=[k for k in data.keys() if k!="_last_profile"] or DEFAULT_PROFILE_NAMES
    profile_combo["values"]=vals
    last=data.get("_last_profile")
    if last and last in data:
        profile_combo.set(last); current_profile_name=last; apply_state_dict_exact(data[last])
    else:
        sel=profile_combo.get().strip() or vals[0]
        profile_combo.set(sel); current_profile_name=sel
        if sel not in data:
            data[sel]=get_state_dict(); data["_last_profile"]=sel; save_all_profiles(data)
    refresh_window_lists()
    request_layout()
init_profiles_and_autoload()

root.update_idletasks()
fit_height_only_now()
lock_headers_now()      # once
root.resizable(False, False)
root.deiconify()

def on_close():
    for r in all_rows:
        try: r.running=False; r.unregister_hotkey()
        except: pass
    time.sleep(0.05); root.destroy()
root.protocol("WM_DELETE_WINDOW", on_close)

try:
    base = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(__file__)
    root.iconbitmap(os.path.join(base,"Ficon_multi.ico"))
except: pass

root.mainloop()
