#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Pet Assistant - Windows App Grid Launcher (Sidebar Edition)
Fixed image loading bug and added a dedicated 'Remove Game' sidebar button.
Added global hotkey (Ctrl+Alt+P) to toggle visibility.
"""

import sys
import os
import shutil
import json
import subprocess
import winreg
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from PIL import Image, ImageTk, ImageSequence

# Optional: needed to resolve shortcuts
try:
    import pythoncom
    import win32com.client
    HAVE_PYWIN32 = True
except ImportError:
    HAVE_PYWIN32 = False

try:
    import winsound
    HAVE_WINSOUND = True
except ImportError:
    HAVE_WINSOUND = False

try:
    import win32api
    import win32gui
    import win32ui
    HAVE_WIN32_ICONS = True
except ImportError:
    HAVE_WIN32_ICONS = False

# For global hotkey
try:
    from pynput import keyboard
    HAVE_PYNPUT = True
except ImportError:
    HAVE_PYNPUT = False

DEBUG = False

def get_base_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(os.path.abspath(os.path.dirname(__file__)))

BASE_PATH = get_base_path()
CONFIG_PATH = (BASE_PATH.parent if getattr(sys, "frozen", False) else BASE_PATH) / "config.json"

DEFAULT_CONFIG = {
    "gif_path": "furina_idle.gif",
    "grok_url": "https://grok.com/",
    "youtube_url": "https://youtube.com",
    "games": {},
    "search_paths": [],
    "soundboard": {},
    "launcher_sidebar_width": 140,
    "launcher_soundboard_height": 320,
}

def clean_path(raw: str) -> str:
    s = raw.strip()
    if s.lower().startswith('r"') and s.endswith('"'):
        s = s[2:-1]
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1]
    return s.replace("/", "\\")

class PetAssistant:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        
        self.card_images = []
        self.config = self.load_config()
        self.edge_path = self.find_edge_path()
        self.search_roots = self.get_search_roots()
        
        self.launcher_win = None
        self.kill_menu = None
        self.pet_visible = True
        self.launcher_should_restore = False
        self.icon_cache = {}
        self.managed_windows = []
        self.sound_buttons = []
        self.sound_volume = 100.0
        
        self.setup_styles()
        self.setup_pet_window()
        self.load_animation()
        
        # Setup hotkey
        if HAVE_PYNPUT:
            self.setup_hotkey()
        else:
            print("[Warning] pynput not installed. Global hotkey disabled.", file=sys.stderr)
            print("[Info] Install with: pip install pynput", file=sys.stderr)
        
    # --- Configuration ---
    def load_config(self) -> dict:
        if not CONFIG_PATH.is_file():
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=4, ensure_ascii=False), encoding="utf-8")
            return DEFAULT_CONFIG.copy()

        try:
            raw_cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"[config] JSON error: {exc}", file=sys.stderr)
            return DEFAULT_CONFIG.copy()

        cfg = DEFAULT_CONFIG.copy()
        cfg.update(raw_cfg)
        
        new_games = {}
        for name, data in cfg.get("games", {}).items():
            if isinstance(data, str):
                new_games[name] = {"path": clean_path(data), "image": ""}
            elif isinstance(data, dict):
                new_games[name] = {
                    "path": clean_path(data.get("path", "")),
                    "image": clean_path(data.get("image", ""))
                }
        cfg["games"] = new_games
        cfg["search_paths"] = [clean_path(p) for p in cfg.get("search_paths", []) if p]

        soundboard = {}
        for name, path in cfg.get("soundboard", {}).items():
            if isinstance(path, str) and path:
                soundboard[str(name)] = clean_path(path)
        cfg["soundboard"] = soundboard

        try:
            cfg["launcher_sidebar_width"] = max(80, min(400, int(cfg.get("launcher_sidebar_width", 140))))
        except Exception:
            cfg["launcher_sidebar_width"] = 140

        try:
            cfg["launcher_soundboard_height"] = max(80, min(500, int(cfg.get("launcher_soundboard_height", 320))))
        except Exception:
            cfg["launcher_soundboard_height"] = 320

        return cfg

    def save_config(self):
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_PATH.write_text(json.dumps(self.config, indent=4, ensure_ascii=False), encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Config save", f"Failed to write config:\n{exc}")

    # --- Utility ---
    def find_edge_path(self) -> Path | None:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe")
            p = winreg.QueryValue(key, None)
            winreg.CloseKey(key)
            return Path(p)
        except OSError:
            return None

    def get_search_roots(self) -> list[Path]:
        user_paths = [Path(p) for p in self.config.get("search_paths", []) if p]
        if user_paths:
            return user_paths

        home = Path.home()
        candidates = [
            Path(os.getenv("ProgramFiles", r"C:\Program Files")),
            Path(os.getenv("ProgramFiles(x86)", r"C:\Program Files (x86)")),
            home / "Desktop",
            home / "Documents",
            home / "Games",
            Path(r"C:\Program Files\Steam\steamapps\common"),
            Path(r"C:\Program Files (x86)\Steam\steamapps\common"),
            home / "Program Files" / "Steam" / "steamapps" / "common",
            home / "Program Files (x86)" / "Steam" / "steamapps" / "common",
            home / "AppData" / "Local" / "Programs",
            home / "AppData" / "LocalLow",
        ]

        roots = []
        seen = set()
        for candidate in candidates:
            resolved = candidate.expanduser()
            if resolved.exists() and str(resolved).lower() not in seen:
                roots.append(resolved)
                seen.add(str(resolved).lower())
        return roots

    # --- Hotkey ---
    def setup_hotkey(self):
        """Register a global hotkey to toggle the pet window."""
        self.hotkey_combo = "<ctrl>+<alt>+t"
        self.hotkey_listener = None

        try:
            self.hotkey_listener = keyboard.GlobalHotKeys({self.hotkey_combo: self.toggle_visibility})
            self.hotkey_listener.start()
            print(f"[Info] Global hotkey active: {self.hotkey_combo}", file=sys.stderr)
        except Exception as exc:
            print(f"[Warning] Failed to register hotkey {self.hotkey_combo}: {exc}", file=sys.stderr)

    def toggle_visibility(self):
        """Toggle the pet window safely from the Tk main thread."""
        if self.root and self.root.winfo_exists():
            self.root.after(0, self._toggle_visibility_now)
        else:
            self._toggle_visibility_now()

    def _toggle_visibility_now(self):
        if not hasattr(self, "pet_win") or not self.pet_win.winfo_exists():
            return

        self.close_all_menus()

        if self.pet_visible:
            self.hide_all_windows()
            self.pet_win.withdraw()
            self.pet_visible = False
        else:
            self.pet_win.deiconify()
            self.pet_win.lift()
            self.pet_win.attributes("-topmost", True)
            self.pet_visible = True
            self.restore_all_windows()
            if self.launcher_win and self.launcher_win.winfo_exists():
                self.launcher_win.deiconify()
                self.launcher_win.lift()
                self.launcher_win.attributes("-topmost", True)

    # --- UI & Animation ---
    def setup_styles(self):
        """Configure ttk styles with modern design (GitHub Dark theme)."""
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Modern color palette
        self.bg_primary = "#0d1117"
        self.bg_secondary = "#161b22"
        self.bg_tertiary = "#21262d"
        self.accent_primary = "#58a6ff"
        self.accent_secondary = "#79c0ff"
        self.text_primary = "#c9d1d9"
        self.text_secondary = "#8b949e"
        self.text_muted = "#6e7681"
        self.success = "#3fb950"
        self.danger = "#f85149"
        self.border = "#30363d"
        
        # Main Button style
        self.style.configure("TButton", font=("Segoe UI", 10, "bold"), foreground=self.bg_primary,
                             background=self.accent_primary, padding=10, borderwidth=0, focusthickness=0)
        self.style.map("TButton", 
                      background=[("active", self.accent_secondary), ("pressed", "#0969da"), ("disabled", self.text_muted)],
                      foreground=[("disabled", self.text_secondary)])
        
        # Secondary Button style
        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), foreground=self.text_primary,
                             background=self.bg_tertiary, padding=10, borderwidth=1, focusthickness=0)
        self.style.map("Secondary.TButton", 
                      background=[("active", self.bg_secondary), ("pressed", self.border)])
        
        # Success Button style (green)
        self.style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                             background=self.success, padding=10, borderwidth=0, focusthickness=0)
        self.style.map("Success.TButton", 
                      background=[("active", "#2ea043"), ("pressed", "#238636")])
        
        # Danger Button style (red)
        self.style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                             background=self.danger, padding=10, borderwidth=0, focusthickness=0)
        self.style.map("Danger.TButton", 
                      background=[("active", "#da3633"), ("pressed", "#c21c1c")])
        
        # Labels
        self.style.configure("TLabel", background=self.bg_primary, foreground=self.text_primary, font=("Segoe UI", 10))
        self.style.configure("Secondary.TLabel", background=self.bg_primary, foreground=self.text_secondary, font=("Segoe UI", 9))
        self.style.configure("Heading.TLabel", background=self.bg_primary, foreground=self.accent_primary, font=("Segoe UI", 14, "bold"))
        
        # Frames
        self.style.configure("TFrame", background=self.bg_primary, borderwidth=0)
        self.style.configure("Card.TFrame", background=self.bg_tertiary, borderwidth=0, relief="flat")
        
        # Checkbuttons
        self.style.configure("TCheckbutton", background=self.bg_primary, foreground=self.text_primary, font=("Segoe UI", 10), focusthickness=0)
        self.style.map("TCheckbutton", background=[("active", self.bg_primary)])
        
        # Combobox
        self.style.configure("TCombobox", background=self.bg_secondary, foreground=self.text_primary, fieldbackground=self.bg_secondary, borderwidth=1)
        self.style.map("TCombobox", fieldbackground=[("readonly", self.bg_secondary)])
        
        # Treeview - Modern look
        self.style.configure("Custom.Treeview", background=self.bg_tertiary, foreground=self.text_primary, fieldbackground=self.bg_tertiary,
                            borderwidth=0, relief="flat")
        self.style.configure("Custom.Treeview.Heading", background=self.bg_secondary, foreground=self.text_primary, borderwidth=0)
        self.style.map("Custom.Treeview", 
                      background=[("selected", self.accent_primary)], 
                      foreground=[("selected", self.bg_primary)])
        
        # Scrollbars
        self.style.configure("Vertical.TScrollbar", background=self.bg_secondary, troughcolor=self.bg_tertiary, 
                            arrowcolor=self.text_primary, darkcolor=self.bg_secondary, lightcolor=self.bg_secondary)
        self.style.configure("Horizontal.TScrollbar", background=self.bg_secondary, troughcolor=self.bg_tertiary,
                            arrowcolor=self.text_primary, darkcolor=self.bg_secondary, lightcolor=self.bg_secondary)

    def setup_pet_window(self):
        self.pet_win = tk.Toplevel(self.root)
        self.pet_win.overrideredirect(True)
        self.pet_win.attributes("-topmost", True)
        self.pet_win.geometry("480x720+200+50")
        transparent_color = "#ff00ff"
        self.pet_win.configure(bg=transparent_color)
        self.pet_win.wm_attributes("-transparentcolor", transparent_color)

    def load_animation(self):
        gif_file = self.config["gif_path"]
        if not Path(gif_file).is_absolute():
            gif_file = BASE_PATH / gif_file
        if not Path(gif_file).is_file():
            messagebox.showerror("GIF error", f"Could not find GIF: {gif_file}")
            sys.exit(1)

        self.pil_img = Image.open(gif_file)
        self.frames = [ImageTk.PhotoImage(frame.copy().convert("RGBA"), master=self.pet_win)
                       for frame in ImageSequence.Iterator(self.pil_img)]
        self.gif_label = tk.Label(self.pet_win, image=self.frames[0], bg="#ff00ff", borderwidth=0, highlightthickness=0)
        self.gif_label.place(relwidth=1, relheight=1)

        self.gif_label.bind("<Button-1>", self.start_drag)
        self.gif_label.bind("<B1-Motion>", self.do_drag)
        self.gif_label.bind("<Button-3>", self.show_kill_menu)
        self.gif_label.bind("<Double-Button-1>", self.show_app_launcher)
        self.animate(0)

    def animate(self, i: int):
        self.gif_label.configure(image=self.frames[i % len(self.frames)])
        self.pet_win.after(self.pil_img.info.get("duration", 100), self.animate, i + 1)

    # --- Dragging ---
    def start_drag(self, event):
        self.pet_win._drag_x = event.x_root - self.pet_win.winfo_x()
        self.pet_win._drag_y = event.y_root - self.pet_win.winfo_y()

    def do_drag(self, event):
        self.pet_win.geometry(f"+{event.x_root - self.pet_win._drag_x}+{event.y_root - self.pet_win._drag_y}")

    def make_draggable(self, win: tk.Toplevel, handle=None):
        handle = handle or win

        def on_press(event):
            widget = event.widget
            parent = widget
            while parent is not None:
                if parent is handle:
                    win._drag_x = event.x_root - win.winfo_x()
                    win._drag_y = event.y_root - win.winfo_y()
                    return
                try:
                    parent = parent.master
                except AttributeError:
                    break

        def on_motion(event):
            if not hasattr(win, "_drag_x"):
                return
            win.geometry(f"+{event.x_root - win._drag_x}+{event.y_root - win._drag_y}")

        handle.bind("<Button-1>", on_press)
        handle.bind("<B1-Motion>", on_motion)

    def register_window(self, win):
        if win and win not in self.managed_windows:
            self.managed_windows.append(win)

    def hide_all_windows(self):
        for win in list(self.managed_windows):
            if win and win.winfo_exists():
                win.withdraw()

        if self.launcher_win and self.launcher_win.winfo_exists():
            self.launcher_should_restore = True
            self.launcher_win.withdraw()

        if self.kill_menu and self.kill_menu.winfo_exists():
            self.kill_menu.withdraw()

    def restore_all_windows(self):
        for win in list(self.managed_windows):
            if win and win.winfo_exists():
                win.deiconify()
                win.lift()
                win.attributes("-topmost", True)

        if self.launcher_win and self.launcher_win.winfo_exists() and self.launcher_should_restore:
            self.launcher_win.deiconify()
            self.launcher_win.lift()
            self.launcher_win.attributes("-topmost", True)

        if self.kill_menu and self.kill_menu.winfo_exists():
            self.kill_menu.deiconify()
            self.kill_menu.lift()
            self.kill_menu.attributes("-topmost", True)

    def close_all_menus(self):
        if self.kill_menu and self.kill_menu.winfo_exists():
            self.kill_menu.destroy()
            self.kill_menu = None

        if self.launcher_win and self.launcher_win.winfo_exists():
            self.launcher_should_restore = True
            self.launcher_win.withdraw()

    def get_icon_photo(self, path: str | Path, size: int, master=None):
        if not path:
            return None

        cache_key = (str(path), size)
        if cache_key in self.icon_cache:
            return self.icon_cache[cache_key]

        target = Path(path)
        if target.exists():
            suffix = target.suffix.lower()
            if suffix in {".ico", ".png", ".jpg", ".jpeg", ".gif", ".bmp"}:
                try:
                    photo = ImageTk.PhotoImage(Image.open(target).convert("RGBA").resize((size, size), Image.Resampling.LANCZOS), master=master)
                    self.icon_cache[cache_key] = photo
                    return photo
                except Exception:
                    pass

        if HAVE_WIN32_ICONS and target.exists():
            try:
                icon = win32gui.ExtractIcon(0, str(target), 0)
                if icon:
                    icon_info = win32gui.GetIconInfo(icon)
                    hbm = icon_info[0]
                    if hbm:
                        bmp = win32ui.CreateBitmapFromHandle(hbm)
                        width = bmp.GetWidth()
                        height = bmp.GetHeight()
                        bits = bmp.GetBitmapBits(True)
                        if bits:
                            image = Image.frombuffer("RGBA", (width, height), bits, "raw", "BGRA", 0, 1).resize((size, size), Image.Resampling.LANCZOS)
                            photo = ImageTk.PhotoImage(image, master=master)
                            self.icon_cache[cache_key] = photo
                            return photo
                    win32gui.DestroyIcon(icon)
            except Exception:
                pass

        return None

    def set_sound_volume(self, percent):
        try:
            value = max(0.0, min(100.0, float(percent)))
            self.sound_volume = value

            try:
                import ctypes
                winmm = ctypes.WinDLL('winmm')
                scaled = int(max(0, min(65535, round((value / 100.0) * 65535))))
                dw_volume = (scaled << 16) | scaled
                winmm.waveOutSetVolume(0, dw_volume)
            except Exception:
                pass

            alias = getattr(self, "active_sound_alias", None)
            mci_type = getattr(self, "active_sound_mci_type", None)
            if alias:
                try:
                    import ctypes
                    winmm = ctypes.WinDLL('winmm')
                    volume = int(max(0, min(1000, int(round(value * 10)))))
                    command = f'set {alias} volume to {volume}' if mci_type == 'waveaudio' else f'setaudio {alias} volume to {volume}'
                    winmm.mciSendStringW(command, None, 0, None)
                except Exception:
                    pass

            try:
                from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
                from ctypes import cast, POINTER
                from comtypes import CLSCTX_ALL
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                volume.SetMasterVolumeLevelScalar(value / 100.0, None)
            except Exception:
                pass
        except Exception:
            pass

    def stop_sound_button(self, button):
        if not button:
            return
        button._is_playing = False
        button.config(text="🔊", bg="#31334a", fg="#cdd6f4", activebackground="#89b4fa", activeforeground="#1e1e2e")
        alias = getattr(button, '_mci_alias', None)
        if alias:
            try:
                import ctypes
                winmm = ctypes.WinDLL('winmm')
                winmm.mciSendStringW(f'stop {alias}', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception:
                pass
            button._mci_alias = None
            if getattr(self, 'active_sound_alias', None) == alias:
                self.active_sound_alias = None
                self.active_sound_mci_type = None
        proc = getattr(button, "_proc", None)
        if proc is not None:
            try:
                proc.terminate()
                proc.wait(timeout=1)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
            button._proc = None
        if HAVE_WINSOUND:
            try:
                winsound.PlaySound(None, winsound.SND_PURGE)
            except Exception:
                pass

    def toggle_sound_button(self, button, sound_path):
        if not sound_path or not Path(sound_path).is_file():
            return

        if getattr(button, "_is_playing", False):
            self.stop_sound_button(button)
            return

        for other in self.sound_buttons:
            if other is not button and getattr(other, "_is_playing", False):
                self.stop_sound_button(other)

        button._is_playing = True
        button._proc = None
        button.config(text="■", bg="#89b4fa", fg="#1e1e2e", activebackground="#89b4fa", activeforeground="#1e1e2e")

        try:
            ext = Path(sound_path).suffix.lower()
            supported = {".wav", ".mp3", ".ogg", ".m4a", ".aac", ".flac", ".wma", ".mp2", ".aiff", ".aif"}
            if ext not in supported:
                button._is_playing = False
                button.config(text="🔊", bg="#31334a", fg="#cdd6f4")
                messagebox.showinfo("Unsupported audio", f"This file type is not supported:\n{ext or 'unknown'}")
                return

            if ext == '.wav' and HAVE_WINSOUND:
                try:
                    self.active_sound_alias = None
                    self.active_sound_mci_type = None
                    button._mci_alias = None
                    winsound.PlaySound(str(sound_path), winsound.SND_FILENAME | winsound.SND_ASYNC)
                    return
                except Exception:
                    pass

            import ctypes
            winmm = ctypes.WinDLL('winmm')
            alias = f'snd_{abs(hash(str(sound_path)))}'
            button._mci_alias = alias
            self.active_sound_alias = alias
            mci_type = 'waveaudio' if ext == '.wav' else 'mpegvideo'
            self.active_sound_mci_type = mci_type
            open_cmd = f'open "{sound_path}" type {mci_type} alias {alias}'
            winmm.mciSendStringW(open_cmd, None, 0, None)
            volume = int(max(0, min(1000, int(round(self.sound_volume * 10)))))
            volume_command = f'set {alias} volume to {volume}' if mci_type == 'waveaudio' else f'setaudio {alias} volume to {volume}'
            winmm.mciSendStringW(volume_command, None, 0, None)
            winmm.mciSendStringW(f'play {alias} notify', None, 0, None)
            return
        except Exception:
            button._is_playing = False
            button.config(text="🔊", bg="#31334a", fg="#cdd6f4")
            messagebox.showerror("Sound error", f"Could not play: {sound_path}")

    def close_launcher_win(self, win=None):
        if win is None:
            win = self.launcher_win
        if win and win.winfo_exists():
            win.destroy()
        if self.launcher_win is win:
            self.launcher_win = None
        self.launcher_should_restore = False

    def show_kill_menu(self, event):
        self.close_all_menus()
        self.kill_menu = tk.Toplevel(self.pet_win)
        self.kill_menu.overrideredirect(True)
        self.kill_menu.attributes("-topmost", True, "-alpha", 0.96)
        self.kill_menu.configure(bg="#1e1e2e")
        self.kill_menu.geometry(f"+{event.x_root + 20}+{event.y_root + 20}")
        self.make_draggable(self.kill_menu)
        self.register_window(self.kill_menu)

        frm = tk.Frame(self.kill_menu, bg="#1e1e2e")
        frm.pack(padx=8, pady=8)
        ttk.Button(frm, text="Kill Pet ×", command=self.root.destroy).pack(fill="x", pady=3)
        ttk.Button(frm, text="Cancel", command=self.kill_menu.destroy).pack(fill="x", pady=3)

    # --- App Launcher ---
    def show_app_launcher(self, event=None):
        self.close_all_menus()
        self.card_images.clear()
        self.launcher_should_restore = True

        win = tk.Toplevel(self.root)
        win.title("App Launcher")
        win.geometry("620x560+250+150")
        win.configure(bg=self.bg_primary)
        win.attributes("-topmost", True, "-alpha", 0.98)
        win.overrideredirect(True)
        self.register_window(win)
        self.launcher_win = win

        title_bar = tk.Frame(win, bg=self.bg_secondary, height=36)
        title_bar.pack(fill="x", side="top")
        tk.Label(title_bar, text="  ⚙️  App Launcher", bg=self.bg_secondary, fg=self.accent_primary, font=("Segoe UI", 11, "bold")).pack(side="left", pady=6)
        ttk.Button(title_bar, text="✕", command=lambda: self.close_launcher_win(win), style="TButton").pack(side="right", padx=6, pady=4)
        self.make_draggable(win, title_bar)

        main_container = tk.Frame(win, bg=self.bg_primary)
        main_container.pack(fill="both", expand=True)

        sidebar = tk.Frame(main_container, bg=self.bg_secondary, width=self.config.get("launcher_sidebar_width", 140))
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        sidebar_drag = tk.Frame(main_container, bg=self.accent_primary, width=3, cursor="sb_h_double_arrow")
        sidebar_drag.pack(side="left", fill="y")
        sidebar_drag.pack_propagate(False)

        def resize_sidebar(event):
            delta = event.x_root - sidebar_drag._resize_start_x
            new_width = max(80, min(400, sidebar_drag._resize_start_width + delta))
            sidebar.config(width=new_width)
            self.config["launcher_sidebar_width"] = new_width
            self.save_config()

        def start_sidebar_resize(event):
            sidebar_drag._resize_start_x = event.x_root
            sidebar_drag._resize_start_width = sidebar.winfo_width()

        sidebar_drag.bind("<Button-1>", start_sidebar_resize)
        sidebar_drag.bind("<B1-Motion>", resize_sidebar)

        content_area = tk.Frame(main_container, bg=self.bg_primary)
        content_area.pack(side="left", fill="both", expand=True)

        def create_sidebar_btn(text, icon, command, close_menu=True):
            btn = tk.Frame(sidebar, bg=self.bg_secondary, cursor="hand2", height=40)
            btn.pack(fill="x", pady=2, padx=3)
            btn.pack_propagate(False)
            lbl = tk.Label(btn, text=f"{icon}  {text}", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9, "bold"), anchor="w", cursor="hand2")
            lbl.pack(fill="both", expand=True, padx=8)
            
            def on_enter(e): btn.config(bg=self.accent_primary); lbl.config(bg=self.accent_primary, fg=self.bg_primary)
            def on_leave(e): btn.config(bg=self.bg_secondary); lbl.config(bg=self.bg_secondary, fg=self.text_primary)
            def on_click(e):
                if close_menu: self.close_all_menus()
                command()
            for w in (btn, lbl):
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", on_click)

        create_sidebar_btn("Online AI", "🌐", lambda: self.open_url(self.config["grok_url"]))
        create_sidebar_btn("Offline AI", "🤖", self.open_ai_chat)
        create_sidebar_btn("YouTube", "📺", lambda: self.open_url(self.config["youtube_url"]))
        tk.Frame(sidebar, bg=self.border, height=1).pack(fill="x", padx=8, pady=10)
        create_sidebar_btn("Find Game", "🔍", self.show_find_game, close_menu=False)
        create_sidebar_btn("Add Game", "➕", self.add_new_game, close_menu=False)
        create_sidebar_btn("Remove Game", "🗑️", self.remove_game_dialog, close_menu=False)
        tk.Frame(sidebar, bg=self.border, height=1).pack(fill="x", padx=8, pady=10)
        create_sidebar_btn("Settings", "⚙️", self.show_launcher_settings, close_menu=False)

        # Games Grid
        canvas = tk.Canvas(content_area, bg=self.bg_primary, highlightthickness=0)
        grid_frame = tk.Frame(canvas, bg=self.bg_primary)

        grid_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=grid_frame, anchor="nw")

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        games = list(self.config.get("games", {}).items())
        
        if not games:
            tk.Label(grid_frame, text="No games added yet.\nUse 'Add Game' in the sidebar!", 
                     bg=self.bg_primary, fg=self.text_muted, font=("Segoe UI", 10, "italic")).grid(row=0, column=0, pady=50, padx=50)
        else:
            for idx, (name, data) in enumerate(games):
                self.create_app_card(win, grid_frame, name, data, idx // 4, idx % 4)

        bottom_section = tk.Frame(win, bg=self.bg_primary)
        bottom_section.pack(fill="x", side="bottom", padx=8, pady=(0, 8))

        soundboard_drag = tk.Frame(bottom_section, bg=self.accent_primary, height=3, cursor="sb_v_double_arrow")
        soundboard_drag.pack(fill="x", side="top")
        soundboard_drag.pack_propagate(False)

        soundboard = tk.Frame(bottom_section, bg=self.bg_secondary, height=self.config.get("launcher_soundboard_height", 320))
        soundboard.pack(fill="x", side="top")
        soundboard.pack_propagate(False)

        def resize_soundboard(event):
            delta = event.y_root - soundboard_drag._resize_start_y
            new_height = max(80, min(500, soundboard_drag._resize_start_height - delta))
            soundboard.config(height=new_height)
            self.config["launcher_soundboard_height"] = new_height
            self.save_config()

        def start_soundboard_resize(event):
            soundboard_drag._resize_start_y = event.y_root
            soundboard_drag._resize_start_height = soundboard.winfo_height()

        soundboard_drag.bind("<Button-1>", start_soundboard_resize)
        soundboard_drag.bind("<B1-Motion>", resize_soundboard)

        volume_frame = tk.Frame(soundboard, bg=self.bg_secondary)
        volume_frame.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(volume_frame, text="🔊", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))

        volume_var = tk.DoubleVar(value=100)
        volume_slider = tk.Scale(
            volume_frame,
            from_=0,
            to=100,
            orient="horizontal",
            variable=volume_var,
            bg=self.bg_secondary,
            fg=self.accent_primary,
            highlightthickness=0,
            troughcolor="#31334a",
            activebackground="#89b4fa",
            bd=0,
            length=170,
            sliderlength=14,
            showvalue=0,
            cursor="hand2"
        )
        volume_slider.pack(side="left", fill="x", expand=True)

        volume_text = tk.Label(volume_frame, text="100%", bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 8))
        volume_text.pack(side="right", padx=(6, 0))

        def on_volume_change(value):
            try:
                volume_text.config(text=f"{int(float(value))}%")
            except Exception:
                pass
            return None

        volume_slider.configure(command=lambda value: (self.set_sound_volume(value), on_volume_change(value)))

        soundboard_inner = tk.Frame(soundboard, bg=self.bg_secondary)
        soundboard_inner.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        soundboard_btns = tk.Frame(soundboard_inner, bg=self.bg_secondary)
        soundboard_btns.pack(fill="both", expand=True)

        def add_sound_tile(container, label_text, sound_path=None, is_add=False):
            tile = tk.Frame(container, bg=self.bg_secondary, width=90, height=130)
            tile.pack_propagate(False)
            tile.pack(side="left", padx=6, pady=6)

            if is_add:
                icon = tk.Button(tile, text="＋", width=5, height=2, bg=self.bg_tertiary, fg=self.text_primary,
                                activebackground=self.accent_primary, activeforeground=self.bg_primary, relief="flat", bd=0,
                                cursor="hand2", command=add_sound)
            else:
                icon = tk.Button(tile, text="🔊", width=5, height=2, bg=self.bg_tertiary, fg=self.accent_primary,
                                activebackground=self.accent_primary, activeforeground=self.bg_primary, relief="flat", bd=0,
                                cursor="hand2")
                icon._sound_path = sound_path
                icon._is_playing = False
                icon._proc = None
                icon.config(command=lambda p=sound_path: play_sound_file(p))
                self.sound_buttons.append(icon)

            icon.pack(fill="x", pady=(0, 2))
            name_label = tk.Label(tile, text=label_text, bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 8),
                                  wraplength=82, justify="center", anchor="n")
            name_label.pack(fill="both", expand=True, padx=3, pady=(0, 6))
            return tile

        def add_sound():
            path = filedialog.askopenfilename(
                title="Select sound file",
                filetypes=[
                    ("Audio files", "*.mp3;*.wav;*.ogg;*.m4a;*.aac;*.flac;*.wma"),
                    ("All files", "*.*")
                ]
            )
            if not path:
                return
            name = Path(path).stem
            key = name if name not in self.config["soundboard"] else f"{name}_{len(self.config['soundboard']) + 1}"
            self.config["soundboard"][key] = path
            self.save_config()
            self.show_app_launcher()

        def play_sound_file(sound_path):
            if not sound_path or not Path(sound_path).is_file():
                return
            ext = Path(sound_path).suffix.lower()
            if ext not in {".mp3", ".wav", ".ogg", ".m4a", ".aac", ".flac", ".wma"}:
                messagebox.showinfo("Unsupported audio", "This file type is not supported by the sound player.")
                return
            button = next((b for b in self.sound_buttons if getattr(b, "_sound_path", None) == sound_path), None)
            if button is None:
                return
            self.toggle_sound_button(button, sound_path)

        self.sound_buttons = []
        sound_items = list(sorted(self.config.get("soundboard", {}).items()))

        if sound_items:
            for name, path in sound_items:
                short_name = name
                if len(short_name) > 12:
                    short_name = short_name[:11] + "…"
                add_sound_tile(soundboard_btns, short_name, sound_path=path)

        add_sound_tile(soundboard_btns, "Add sound", is_add=True)

        if not sound_items:
            empty_label = tk.Label(soundboard_btns, text="No sounds yet", bg=self.bg_secondary, fg=self.text_muted, font=("Segoe UI", 9, "italic"))
            empty_label.pack(side="left", padx=12, pady=12)


    def create_app_card(self, win, parent, name: str, data: dict, row: int, col: int):
        card = tk.Frame(parent, bg=self.bg_tertiary, cursor="hand2", width=92, height=92, relief="flat", borderwidth=1)
        card.grid(row=row, column=col, padx=8, pady=8)
        card.grid_propagate(False)

        img_path = data.get("image", "")
        has_image = False
        icon_widget = None

        if img_path and Path(img_path).is_file():
            try:
                pil_img = Image.open(img_path).convert("RGBA").resize((44, 44), Image.Resampling.LANCZOS)
                tk_img = ImageTk.PhotoImage(pil_img, master=win)
                self.card_images.append(tk_img)
                icon_widget = tk.Label(card, image=tk_img, bg=self.bg_tertiary, cursor="hand2", bd=0, highlightthickness=0)
                has_image = True
            except Exception as e:
                if DEBUG: print(f"Image load error for {name}: {e}")
                icon_widget = tk.Label(card, text="🎮", font=("Segoe UI Emoji", 20), bg=self.bg_tertiary, fg=self.accent_primary, cursor="hand2", bd=0, highlightthickness=0)
        else:
            executable = data.get("path", "")
            icon_img = self.get_icon_photo(executable, 44, master=win)
            if icon_img:
                icon_widget = tk.Label(card, image=icon_img, bg=self.bg_tertiary, cursor="hand2", bd=0, highlightthickness=0)
                self.card_images.append(icon_img)
                has_image = True
            else:
                icon_widget = tk.Label(card, text="🎮", font=("Segoe UI Emoji", 20), bg=self.bg_tertiary, fg=self.accent_primary, cursor="hand2", bd=0, highlightthickness=0)

        icon_widget.pack(expand=True, pady=(12, 0))

        text_lbl = tk.Label(card, text=name, font=("Segoe UI", 8, "bold"), bg=self.bg_tertiary, fg=self.text_primary, 
                            wraplength=84, justify="center", cursor="hand2", bd=0, highlightthickness=0)
        text_lbl.pack(expand=True, pady=(0, 4))

        # Image setup button overlay
        img_btn = tk.Label(card, text="🖼", bg=self.bg_tertiary, fg=self.accent_secondary, font=("Segoe UI Emoji", 8), cursor="hand2", bd=0, highlightthickness=0)
        img_btn.place(x=76, y=2)

        def launch(e):
            self.close_all_menus()
            self.launch_game(name, data["path"])

        def image_cmd(e):
            path = filedialog.askopenfilename(title=f"Select Image for {name}", filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.ico;*.gif")])
            if path:
                self.config["games"][name]["image"] = path
                self.save_config()
                self.show_app_launcher()

        img_btn.bind("<Button-1>", image_cmd)
        
        hover_widgets = (card, icon_widget, text_lbl)
        for widget in hover_widgets:
            widget.bind("<Button-1>", launch)
            widget.bind("<Enter>", lambda e: self._set_card_state(card, icon_widget, text_lbl, img_btn, True, has_image))
            widget.bind("<Leave>", lambda e: self._set_card_state(card, icon_widget, text_lbl, img_btn, False, has_image))

        img_btn.bind("<Enter>", lambda e: self._set_card_state(card, icon_widget, text_lbl, img_btn, True, has_image))

    def _set_card_state(self, card, icon_lbl, text_lbl, img_btn, hover: bool, has_image: bool):
        color = self.bg_secondary if hover else self.bg_tertiary
        border_color = self.accent_primary if hover else self.border
        card.config(bg=color, relief="raised" if hover else "flat")
        text_lbl.config(bg=color)
        img_btn.config(bg=color)
        icon_lbl.config(bg=color)
        
        if not has_image:
            icon_lbl.config(font=("Segoe UI Emoji", 24 if hover else 20))
        
        text_lbl.config(font=("Segoe UI", 9, "bold") if hover else ("Segoe UI", 8, "bold"))

    # --- Action Implementations ---
    def open_url(self, url: str):
        if self.edge_path: subprocess.Popen([str(self.edge_path), "--new-window", url])
        else: subprocess.Popen(["start", url], shell=True)

    def open_ai_chat(self):
        chat = tk.Toplevel(self.root)
        chat.title("AI Chat ♡")
        chat.geometry("420x520+300+200")
        chat.configure(bg=self.bg_primary)
        chat.attributes("-topmost", True)
        self.register_window(chat)

        txt = tk.Text(chat, bg=self.bg_secondary, fg=self.text_primary, insertbackground=self.accent_primary, relief="flat", font=("Segoe UI", 10), wrap="word")
        txt.pack(expand=True, fill="both", padx=8, pady=8)
        txt.insert("end", "AI: Hello! Ask me anything.\n\n")
        txt.config(state="disabled")

        entry = tk.Entry(chat, bg=self.bg_tertiary, fg=self.text_primary, insertbackground=self.accent_primary, relief="flat", font=("Segoe UI", 10), bd=1)
        entry.pack(fill="x", padx=8, pady=(0, 8))
        entry.focus()

        def send(event=None):
            msg = entry.get().strip()
            if not msg: return
            entry.delete(0, "end")
            txt.config(state="normal")
            txt.insert("end", f"You: {msg}\nAI: thinking...\n")
            txt.see("end")
            txt.config(state="disabled")
            
            def fetch_ai():
                try:
                    result = subprocess.run(["ollama", "run", "llama3.1"], input=msg, capture_output=True, text=True)
                    reply = result.stdout.strip()
                except Exception as exc: reply = f"Error: {exc}"
                
                txt.config(state="normal")
                txt.delete("end-2l", "end-1l")
                txt.insert("end", f"AI: {reply}\n\n")
                txt.see("end")
                txt.config(state="disabled")

            threading.Thread(target=fetch_ai, daemon=True).start()
        entry.bind("<Return>", send)

    # --- Game Management ---
    def launch_game(self, name: str, path: str):
        path_obj = Path(path)
        if path_obj.is_file():
            try:
                subprocess.Popen([str(path_obj)])
            except Exception as exc: messagebox.showerror("Launch error", f"Could not start {name}:\n{exc}")
            return
        if not messagebox.askyesno("Not found", f"Cannot find:\n{path}\n\nLocate executable now?"): return

        new_path = filedialog.askopenfilename(title=f"Locate {name}", filetypes=[("Executables", "*.exe"), ("All", "*.*")])
        if new_path:
            self.config["games"][name]["path"] = new_path
            self.save_config()
            subprocess.Popen([new_path])

    def add_new_game(self):
        path = filedialog.askopenfilename(title="Select Game Executable", filetypes=[("Executables", "*.exe")])
        if path:
            game_name = Path(path).stem
            self.config["games"][game_name] = {"path": path, "image": ""}
            self.save_config()
            if self.launcher_win and self.launcher_win.winfo_exists():
                self.show_app_launcher() 

    def remove_game_dialog(self):
        games = self.config.get("games", {})
        if not games:
            messagebox.showinfo("Remove Game", "No games configured to remove.")
            return

        win = tk.Toplevel(self.root)
        win.title("Remove Game")
        win.geometry("300x320+250+200")
        win.configure(bg=self.bg_primary)
        win.attributes("-topmost", True)
        self.register_window(win)

        tk.Label(win, text="Select game to remove:", bg=self.bg_primary, fg=self.text_primary, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(12, 4))

        lb = tk.Listbox(win, bg=self.bg_secondary, fg=self.text_primary, selectbackground=self.accent_primary, selectforeground=self.bg_primary, font=("Segoe UI", 9), activestyle="none", relief="flat", bd=0)
        for name in sorted(games.keys()):
            lb.insert(tk.END, name)
        lb.pack(fill="both", expand=True, padx=12, pady=4)

        def do_remove():
            sel = lb.curselection()
            if not sel: return
            cur_name = lb.get(sel[0])
            if messagebox.askyesno("Confirm", f"Remove '{cur_name}' from launcher?"):
                del self.config["games"][cur_name]
                self.save_config()
                win.destroy()
                if self.launcher_win and self.launcher_win.winfo_exists():
                    self.show_app_launcher()

        ttk.Button(win, text="Remove Game", command=do_remove, style="Danger.TButton").pack(fill="x", padx=12, pady=8)

    def show_launcher_settings(self):
        """Show settings dialog with window size and data management."""
        win = tk.Toplevel(self.root)
        win.title("⚙️ Settings")
        win.geometry("500x650+250+150")
        win.configure(bg=self.bg_primary)
        win.attributes("-topmost", True)
        self.register_window(win)

        # Window Size Section
        size_frame = tk.Frame(win, bg=self.bg_secondary, relief="flat", bd=0)
        size_frame.pack(fill="x", padx=12, pady=(12, 8))
        
        tk.Label(size_frame, text="Launcher Window Size", bg=self.bg_secondary, fg=self.accent_primary, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
        
        # Get current geometry
        geom = self.launcher_win.geometry() if self.launcher_win and self.launcher_win.winfo_exists() else self.config.get("launcher_window_geometry", "620x560+250+150")
        try:
            size_part = geom.split("+")[0]
            w, h = size_part.split("x")
        except:
            w, h = "620", "560"
        
        # Width setting
        width_frame = tk.Frame(size_frame, bg=self.bg_secondary)
        width_frame.pack(fill="x", padx=8, pady=4)
        tk.Label(width_frame, text="Width:", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9)).pack(side="left")
        width_var = tk.StringVar(value=str(w))
        width_entry = tk.Entry(width_frame, textvariable=width_var, bg=self.bg_tertiary, fg=self.text_primary, insertbackground=self.accent_primary, width=10, relief="flat", bd=1)
        width_entry.pack(side="right", padx=(0, 4))
        
        # Height setting
        height_frame = tk.Frame(size_frame, bg=self.bg_secondary)
        height_frame.pack(fill="x", padx=8, pady=4)
        tk.Label(height_frame, text="Height:", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9)).pack(side="left")
        height_var = tk.StringVar(value=str(h))
        height_entry = tk.Entry(height_frame, textvariable=height_var, bg=self.bg_tertiary, fg=self.text_primary, insertbackground=self.accent_primary, width=10, relief="flat", bd=1)
        height_entry.pack(side="right", padx=(0, 4))

        # Sidebar Width setting
        sidebar_frame = tk.Frame(size_frame, bg=self.bg_secondary)
        sidebar_frame.pack(fill="x", padx=8, pady=4)
        tk.Label(sidebar_frame, text="Sidebar Width:", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9)).pack(side="left")
        sidebar_var = tk.StringVar(value=str(self.config.get("launcher_sidebar_width", 140)))
        sidebar_entry = tk.Entry(sidebar_frame, textvariable=sidebar_var, bg=self.bg_tertiary, fg=self.text_primary, insertbackground=self.accent_primary, width=10, relief="flat", bd=1)
        sidebar_entry.pack(side="right", padx=(0, 4))
        
        # Soundboard Height setting
        soundboard_frame = tk.Frame(size_frame, bg=self.bg_secondary)
        soundboard_frame.pack(fill="x", padx=8, pady=4)
        tk.Label(soundboard_frame, text="Soundboard Height:", bg=self.bg_secondary, fg=self.text_primary, font=("Segoe UI", 9)).pack(side="left")
        soundboard_var = tk.StringVar(value=str(self.config.get("launcher_soundboard_height", 320)))
        soundboard_entry = tk.Entry(soundboard_frame, textvariable=soundboard_var, bg=self.bg_tertiary, fg=self.text_primary, insertbackground=self.accent_primary, width=10, relief="flat", bd=1)
        soundboard_entry.pack(side="right", padx=(0, 4))

        def save_sizes():
            try:
                win_w = int(width_var.get())
                win_h = int(height_var.get())
                sidebar_w = int(sidebar_var.get())
                soundboard_h = int(soundboard_var.get())
                
                # Validate and constrain values
                win_w = max(400, min(1600, win_w))
                win_h = max(300, min(1200, win_h))
                sidebar_w = max(80, min(400, sidebar_w))
                soundboard_h = max(80, min(500, soundboard_h))
                
                self.config["launcher_window_geometry"] = f"{win_w}x{win_h}+250+150"
                self.config["launcher_sidebar_width"] = sidebar_w
                self.config["launcher_soundboard_height"] = soundboard_h
                self.save_config()
                
                # Apply window size immediately
                if self.launcher_win and self.launcher_win.winfo_exists():
                    self.launcher_win.geometry(f"{win_w}x{win_h}")
                
                # messagebox.showinfo("Success", "Size settings saved! Changes applied.")
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers")

        ttk.Button(size_frame, text="Apply Sizes", command=save_sizes, style="Success.TButton").pack(fill="x", padx=8, pady=(4, 8))

        # Games Section
        games_lbl = tk.Label(win, text="Games", bg=self.bg_primary, fg=self.accent_primary, font=("Segoe UI", 11, "bold"))
        games_lbl.pack(anchor="w", padx=12, pady=(12, 4))

        games_frame = tk.Frame(win, bg=self.bg_secondary, relief="flat", bd=0)
        games_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # Scrollable games list
        games_canvas = tk.Canvas(games_frame, bg=self.bg_secondary, highlightthickness=0)
        games_canvas.pack(side="left", fill="both", expand=True)
        
        games_scroll = ttk.Scrollbar(games_frame, orient="vertical", command=games_canvas.yview)
        games_scroll.pack(side="right", fill="y")
        games_canvas.configure(yscrollcommand=games_scroll.set)

        games_content = tk.Frame(games_canvas, bg=self.bg_secondary)
        games_canvas.create_window((0, 0), window=games_content, anchor="nw")

        games = self.config.get("games", {})
        if games:
            for game_name in sorted(games.keys()):
                game_item = tk.Frame(games_content, bg=self.bg_tertiary, relief="flat", bd=1)
                game_item.pack(fill="x", padx=4, pady=2)
                
                tk.Label(game_item, text=game_name, bg=self.bg_tertiary, fg=self.text_primary, font=("Segoe UI", 9), anchor="w").pack(side="left", fill="x", expand=True, padx=6, pady=4)
                
                def make_delete_game(name):
                    def delete_game():
                        if messagebox.askyesno("Confirm", f"Delete '{name}'?"):
                            del self.config["games"][name]
                            self.save_config()
                            messagebox.showinfo("Success", f"Deleted '{name}'")
                            self.show_launcher_settings()
                    return delete_game
                
                ttk.Button(game_item, text="Delete", command=make_delete_game(game_name), style="Danger.TButton").pack(side="right", padx=4, pady=2)
        else:
            tk.Label(games_content, text="No games configured", bg=self.bg_secondary, fg=self.text_muted, font=("Segoe UI", 9, "italic")).pack(pady=8)

        games_content.bind("<Configure>", lambda e: games_canvas.configure(scrollregion=games_canvas.bbox("all")))

        # Sounds Section
        sounds_lbl = tk.Label(win, text="Sounds", bg=self.bg_primary, fg=self.accent_primary, font=("Segoe UI", 11, "bold"))
        sounds_lbl.pack(anchor="w", padx=12, pady=(12, 4))

        sounds_frame = tk.Frame(win, bg=self.bg_secondary, relief="flat", bd=0)
        sounds_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # Scrollable sounds list
        sounds_canvas = tk.Canvas(sounds_frame, bg=self.bg_secondary, highlightthickness=0)
        sounds_canvas.pack(side="left", fill="both", expand=True)
        
        sounds_scroll = ttk.Scrollbar(sounds_frame, orient="vertical", command=sounds_canvas.yview)
        sounds_scroll.pack(side="right", fill="y")
        sounds_canvas.configure(yscrollcommand=sounds_scroll.set)

        sounds_content = tk.Frame(sounds_canvas, bg=self.bg_secondary)
        sounds_canvas.create_window((0, 0), window=sounds_content, anchor="nw")

        sounds = self.config.get("soundboard", {})
        if sounds:
            for sound_name in sorted(sounds.keys()):
                sound_item = tk.Frame(sounds_content, bg=self.bg_tertiary, relief="flat", bd=1)
                sound_item.pack(fill="x", padx=4, pady=2)
                
                tk.Label(sound_item, text=sound_name, bg=self.bg_tertiary, fg=self.text_primary, font=("Segoe UI", 9), anchor="w").pack(side="left", fill="x", expand=True, padx=6, pady=4)
                
                def make_delete_sound(name):
                    def delete_sound():
                        if messagebox.askyesno("Confirm", f"Delete '{name}'?"):
                            del self.config["soundboard"][name]
                            self.save_config()
                            messagebox.showinfo("Success", f"Deleted '{name}'")
                            self.show_launcher_settings()
                    return delete_sound
                
                ttk.Button(sound_item, text="Delete", command=make_delete_sound(sound_name), style="Danger.TButton").pack(side="right", padx=4, pady=2)
        else:
            tk.Label(sounds_content, text="No sounds configured", bg=self.bg_secondary, fg=self.text_muted, font=("Segoe UI", 9, "italic")).pack(pady=8)

        sounds_content.bind("<Configure>", lambda e: sounds_canvas.configure(scrollregion=sounds_canvas.bbox("all")))

    def show_find_game(self):
        win = tk.Toplevel(self.root)
        win.title("🔍 Find Game")
        win.geometry("460x520+250+180")
        win.configure(bg=self.bg_primary)
        win.attributes("-topmost", True, "-alpha", 0.98)
        win.overrideredirect(True)
        self.register_window(win)

        top = tk.Frame(win, bg=self.bg_secondary, height=36)
        top.pack(fill="x")
        tk.Label(top, text="  🔍 Find Game", bg=self.bg_secondary, fg=self.accent_primary, font=("Segoe UI", 11, "bold")).pack(side="left", pady=6)
        ttk.Button(top, text="✕", command=win.destroy, style="TButton").pack(side="right", padx=6, pady=4)
        self.make_draggable(win, top)

        frm_search = tk.Frame(win, bg=self.bg_primary)
        frm_search.pack(fill="x", padx=12, pady=(8, 6))

        entry = tk.Entry(frm_search, font=("Segoe UI", 10), bg=self.bg_secondary, fg=self.text_primary, insertbackground=self.accent_primary, relief="flat", bd=1)
        entry.pack(fill="x", pady=(0, 4))
        entry.focus_set()

        button_row = tk.Frame(frm_search, bg=self.bg_primary)
        button_row.pack(fill="x")
        ttk.Button(button_row, text="Search", command=lambda: start_search()).pack(side="left", padx=(0, 6))

        def add_folder():
            folder = filedialog.askdirectory(title="Select folder to include")
            if not folder:
                return
            norm = str(Path(folder))
            if norm not in {str(p) for p in self.search_roots} and norm not in self.config["search_paths"]:
                self.config["search_paths"].append(norm)
                self.save_config()
                self.search_roots = self.get_search_roots()
            folders_text.config(text="Scanning folders:\n" + "\n".join(str(p) for p in self.search_roots))
            status.config(text=f"Scanning: {folder}")

        def remove_folder():
            current = list(self.search_roots)
            if not current:
                messagebox.showinfo("No folders", "You have not added any custom search folders.")
                return

            dialog = tk.Toplevel(self.root)
            dialog.title("Remove search folder")
            dialog.geometry("360x260+250+200")
            dialog.configure(bg=self.bg_primary)
            dialog.attributes("-topmost", True)
            self.make_draggable(dialog)

            tk.Label(dialog, text="Select folder to remove:", bg=self.bg_primary, fg=self.text_primary, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
            lb = tk.Listbox(dialog, bg=self.bg_secondary, fg=self.text_primary, selectbackground=self.accent_primary, selectforeground=self.bg_primary, font=("Segoe UI", 9), width=42, height=8, activestyle="none", relief="flat", bd=0)
            for path in current:
                lb.insert(tk.END, str(path))
            lb.pack(fill="both", expand=True, padx=12, pady=(0, 8))

            def do_remove():
                sel = lb.curselection()
                if not sel:
                    return
                target = lb.get(sel[0])
                self.search_roots = [p for p in self.search_roots if str(p) != target]
                self.config["search_paths"] = [str(p) for p in self.search_roots]
                self.save_config()
                folders_text.config(text="Scanning folders:\n" + ("\n".join(str(p) for p in self.search_roots) if self.search_roots else "(none)"))
                status.config(text=f"Removed: {target}")
                dialog.destroy()

            ttk.Button(dialog, text="Remove selected", command=do_remove).pack(fill="x", padx=12, pady=(0, 8))
            dialog.transient(win)
            dialog.grab_set()

        ttk.Button(button_row, text="Add folder…", command=add_folder).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="Remove folder", command=remove_folder).pack(side="left")

        folders_text = tk.Label(win, justify="left", bg=self.bg_primary, fg=self.text_secondary, anchor="w", wraplength=420, font=("Segoe UI", 8))
        folders_text.pack(fill="x", padx=12, pady=(0, 6))
        folders_text.config(text="Scanning folders:\n" + "\n".join(str(p) for p in self.search_roots))

        frm_list = tk.Frame(win, bg=self.bg_primary)
        frm_list.pack(fill="both", expand=True, padx=12, pady=4)

        tree = ttk.Treeview(
            frm_list,
            columns=("path",),
            show="headings",
            height=10,
            style="Custom.Treeview"
        )
        tree.heading("#0", text="Game")
        tree.heading("path", text="Location")
        tree.column("#0", width=180, minwidth=120, anchor="w")
        tree.column("path", width=220, minwidth=120, anchor="w")
        tree.tag_configure("default", background=self.bg_secondary, foreground=self.text_primary)
        tree.pack(fill="both", expand=True)

        tree_scroll = ttk.Scrollbar(frm_list, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")

        self.style.configure("Custom.Treeview", background=self.bg_secondary, fieldbackground=self.bg_secondary, foreground=self.text_primary)
        self.style.map("Custom.Treeview", background=[("selected", self.accent_primary)], foreground=[("selected", self.bg_primary)])

        status = tk.Label(win, text="Type a game name and press Search", bg=self.bg_primary, fg=self.text_muted, anchor="w")
        status.pack(fill="x", padx=12, pady=(2, 0))

        actions = tk.Frame(win, bg=self.bg_primary)
        actions.pack(fill="x", padx=12, pady=8)
        launch_btn = ttk.Button(actions, text="Add & Launch", state="disabled", style="Success.TButton")
        launch_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(actions, text="Open folder", command=lambda: open_selected_folder(), style="Secondary.TButton").pack(side="right")

        tree.bind("<<TreeviewSelect>>", lambda e: launch_btn.config(state="normal" if tree.selection() else "disabled"))
        tree.bind("<Double-Button-1>", lambda e: do_launch())

        stop_evt = threading.Event()
        result_q = queue.Queue()
        current_thread = [None]
        anim_id = [None]
        anim_step = [0]

        def animate_status():
            if not win.winfo_exists():
                return
            dots = "." * (anim_step[0] % 4)
            status.config(text=f"Searching common game folders{dots}")
            anim_step[0] += 1
            anim_id[0] = win.after(250, animate_status)

        def stop_animation():
            if anim_id[0]:
                win.after_cancel(anim_id[0])
                anim_id[0] = None

        def open_selected_folder():
            sel = tree.selection()
            if not sel:
                return
            try:
                raw_path = tree.item(sel[0], "values")[0]
                target = Path(raw_path)
                if target.exists():
                    target_dir = target.parent
                    if hasattr(os, "startfile"):
                        os.startfile(str(target_dir))
                    else:
                        subprocess.Popen(["explorer", str(target_dir)])
            except Exception:
                pass

        def poll_queue():
            if not win.winfo_exists():
                return
            try:
                while True:
                    item = result_q.get_nowait()
                    if item is None:
                        stop_animation()
                        status.config(text=f"{len(tree.get_children())} result(s) found")
                        return
                    display_name, full_path = item
                    if any(tree.item(child, "text") == display_name and tree.item(child, "values")[0] == full_path for child in tree.get_children()):
                        continue
                    icon = self.get_icon_photo(full_path, 20, master=win)
                    insert_args = {"text": display_name, "values": (full_path,), "tags": ("default",)}
                    if icon is not None:
                        insert_args["image"] = icon
                    tree.insert("", tk.END, **insert_args)
            except queue.Empty:
                pass
            win.after(80, poll_queue)

        def start_search(event=None):
            nonlocal result_q
            term = entry.get().strip()
            if not term:
                messagebox.showinfo("Empty term", "Type a game name or part of a path.")
                return
            for child in tree.get_children():
                tree.delete(child)
            launch_btn.config(state="disabled")
            stop_animation()
            anim_step[0] = 0
            animate_status()
            stop_evt.set()
            if current_thread[0] and current_thread[0].is_alive():
                current_thread[0].join(timeout=0.2)
            result_q = queue.Queue()
            stop_evt.clear()
            current_thread[0] = threading.Thread(target=self.search_worker, args=(term, result_q, stop_evt), daemon=True)
            current_thread[0].start()
            poll_queue()

        entry.bind("<Return>", start_search)

        def do_launch():
            sel = tree.selection()
            if not sel:
                return
            try:
                item = sel[0]
                raw_name = tree.item(item, "text")
                raw_path = tree.item(item, "values")[0]
                path = Path(raw_path)
                if not path.exists() and path.suffix.lower() == ".lnk":
                    raise FileNotFoundError

                self.config["games"][raw_name] = {"path": str(path), "image": ""}
                self.save_config()
                if self.launcher_win and self.launcher_win.winfo_exists():
                    self.show_app_launcher()

                if path.suffix.lower() == ".lnk":
                    try:
                        os.startfile(str(path))
                    except Exception:
                        self.launch_game(path.stem, str(path))
                else:
                    self.launch_game(raw_name, str(path))
                win.destroy()
            except Exception as exc:
                if DEBUG:
                    print(f"Launch search result failed: {exc}")
                messagebox.showerror("Launch failed", f"Could not launch selected game:\n{exc}")

        launch_btn.config(command=do_launch)

        def on_close():
            stop_evt.set()
            stop_animation()
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

    def search_worker(self, term: str, result_q: queue.Queue, stop_evt: threading.Event):
        term = term.lower().strip()
        if not term:
            result_q.put(None)
            return

        search_roots = []
        seen = set()
        for root in self.search_roots:
            root = Path(root)
            if root.exists() and root.is_dir() and str(root).lower() not in seen:
                seen.add(str(root).lower())
                search_roots.append(root)

        for root in search_roots:
            if stop_evt.is_set() or not root.exists():
                continue
            try:
                for dirpath, dirnames, filenames in os.walk(root, topdown=True):
                    if stop_evt.is_set():
                        break

                    dirnames[:] = [d for d in dirnames if not d.startswith('.') and d.lower() not in {
                        'windows', 'system32', 'winsxs', 'appdata', 'cache', 'packages', 'public', 'documents and settings'
                    }]

                    full_dir = Path(dirpath)
                    dir_parts = {p.lower() for p in full_dir.parts}
                    if any(part in {'windows', 'system32', 'winsxs'} for part in dir_parts):
                        dirnames[:] = []
                        continue

                    for fname in filenames:
                        if stop_evt.is_set():
                            break
                        low = fname.lower()
                        if not (low.endswith('.exe') or low.endswith('.lnk')):
                            continue

                        if term not in low and term not in {p.lower() for p in full_dir.parts}:
                            continue

                        full_path = full_dir / fname
                        display = full_path.stem
                        if low.endswith('.lnk') and HAVE_PYWIN32:
                            try:
                                shell = win32com.client.Dispatch("WScript.Shell")
                                target = shell.CreateShortCut(str(full_path)).Targetpath
                                if target and target.lower().endswith('.exe'):
                                    display = Path(target).stem
                                    full_path = Path(target)
                            except Exception:
                                pass

                        result_q.put((display, str(full_path)))
            except (PermissionError, OSError):
                pass

        result_q.put(None)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = PetAssistant()
    app.run()