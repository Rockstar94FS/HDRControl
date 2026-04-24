import os
import sys
import psutil
import pystray
from pystray import MenuItem as item
from PIL import Image
import configparser
import threading
import time
import ctypes
from win11toast import notify
import win32com.client
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import win32api
import tkinter.messagebox as messagebox
import locale

APP_NAME, SETTINGS, TEXTS, APPS, TRAY = "HDR Control", {}, {}, [], None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HDR = ctypes.CDLL(os.path.join(BASE_DIR, "resources/HDRSwitch.dll"))
HDR_ENABLED = bool(HDR.GetGlobalHDRState())

ICON_ON_PATH = os.path.join(BASE_DIR, "resources/hdr_on.ico")
ICON_OFF_PATH = os.path.join(BASE_DIR, "resources/hdr_off.ico")
LOGO_ON_PATH = os.path.join(BASE_DIR, "resources/hdr_logo_on.ico")
LOGO_OFF_PATH = os.path.join(BASE_DIR, "resources/hdr_logo_off.ico")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings/settings.ini")
APPS_PATH = os.path.join(BASE_DIR, "settings/apps.ini")

def load_settings():
    global SETTINGS

    settings_parser = configparser.ConfigParser()
    settings_parser.optionxform = str
    settings_parser.read(SETTINGS_PATH)

    if "settings" not in settings_parser:
        settings_parser["settings"] = {}

    SETTINGS["notifications"] = settings_parser["settings"].getboolean("notifications", fallback=True)
    SETTINGS["update_time"] = settings_parser["settings"].getfloat("update_time", fallback=2)
    SETTINGS["pause"] = settings_parser["settings"].getboolean("pause", fallback=False)
    SETTINGS["primary"] = settings_parser["settings"].getboolean("primary", fallback=True)
    SETTINGS["pos_x"] = settings_parser["settings"].getint("pos_x", fallback=130)
    SETTINGS["pos_y"] = settings_parser["settings"].getint("pos_y", fallback=130)
    SETTINGS["scale_x"] = settings_parser["settings"].getint("scale_x", fallback=900)
    SETTINGS["scale_y"] = settings_parser["settings"].getint("scale_y", fallback=450)
    SETTINGS["run_minimized"] = settings_parser["settings"].getboolean("run_minimized", fallback=True)
    # SETTINGS["start_menu"] = settings_parser["settings"].getboolean("start_menu", fallback=False)

def load_texts():
    global TEXTS

    lang = locale.windows_locale.get(ctypes.windll.kernel32.GetUserDefaultUILanguage(), "en_US")
    path = os.path.join(BASE_DIR, f"language/lang_{lang}.ini")

    if not os.path.exists(path):
        path = os.path.join(BASE_DIR, "language/lang_en_US.ini")

    texts_parser = configparser.ConfigParser()
    texts_parser.optionxform = str
    texts_parser.read(path)

    TEXTS = {}
    TEXTS = dict(texts_parser["texts"])

def save_settings():
    settings_parser = configparser.ConfigParser()
    settings_parser.optionxform = str

    settings_parser["settings"] = SETTINGS.copy()

    with open(SETTINGS_PATH, "w") as f:
        settings_parser.write(f)

def load_apps():
    global APPS

    apps_parser = configparser.ConfigParser()
    apps_parser.optionxform = str
    apps_parser.read(APPS_PATH, encoding="utf-8")

    APPS = []

    for section in apps_parser.sections():
        app = {
            "path": apps_parser.get(section, "path"),
            "name": apps_parser.get(section, "name"),
            "enabled": apps_parser.getboolean(section, "enabled")
        }

        APPS.append(app)

def save_apps():
    global APPS

    apps_parser = configparser.ConfigParser()
    apps_parser.optionxform = str

    for i, app in enumerate(APPS, start=1):
        apps_parser[f"app_{i:03d}"] = {
            "path": str(app["path"]),
            "name": str(app["name"]),
            "enabled": str(app["enabled"])
        }

    with open(APPS_PATH, "w", encoding="utf-8") as f:
        apps_parser.write(f)

def check_running_apps():
    processes = set()

    for p in psutil.process_iter(['name']):
        try:
            processes.add((p.info['name'] or "").lower())
        except:
            pass

    found_app = False

    for app in APPS:
        if app.get("enabled"):
            exe_name = os.path.basename(app.get("path")).lower()

            if exe_name in processes:
                found_app = True
                break

    enable_hdr(found_app)

def monitor_apps():
    while True:
        if not SETTINGS["pause"]:
            check_running_apps()

        time.sleep(SETTINGS["update_time"])

def enable_hdr(is_enabled):
    global HDR_ENABLED

    if HDR_ENABLED != is_enabled:
        HDR_ENABLED = is_enabled

        if SETTINGS["primary"]:
            HDR.SetHDRonPrimary(HDR_ENABLED)
        else:
            HDR.SetGlobalHDRState(HDR_ENABLED)

        update_icon()

        if SETTINGS["notifications"]:
            text = TEXTS["hdr_off_notification"]
            icon = LOGO_OFF_PATH

            if HDR_ENABLED:
                text = TEXTS["hdr_on_notification"]
                icon = LOGO_ON_PATH

            notify(APP_NAME, text, icon = {'src': icon, 'placement': 'appLogoOverride'})

def add_to_autorun():
    shell = win32com.client.Dispatch("WScript.Shell")
    startup_folder = shell.SpecialFolders("Startup")
    shortcut_path = os.path.join(startup_folder, f"{APP_NAME}.lnk")

    if is_autorun_enabled():
        os.remove(shortcut_path)
    else:
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{os.path.abspath(sys.argv[0])}"'
        shortcut.WorkingDirectory = BASE_DIR
        shortcut.IconLocation = ICON_ON_PATH
        shortcut.save()

def is_autorun_enabled():
    shell = win32com.client.Dispatch("WScript.Shell")
    startup_folder = shell.SpecialFolders("Startup")
    shortcut_path = os.path.join(startup_folder, f"{APP_NAME}.lnk")

    return os.path.exists(shortcut_path)

def toggle_pause():
    SETTINGS["pause"] = not SETTINGS["pause"]
    save_settings()

def toggle_minimized():
    SETTINGS["run_minimized"] = not SETTINGS["run_minimized"]
    save_settings()

def toggle_notification():
    SETTINGS["notifications"] = not SETTINGS["notifications"]
    save_settings()

def toggle_display():
    SETTINGS["primary"] = not SETTINGS["primary"]
    save_settings()

def update_icon():
    TRAY.icon = Image.open(ICON_ON_PATH if HDR_ENABLED else ICON_OFF_PATH)

def is_running():
    ctypes.windll.kernel32.CreateMutexW(None, False, APP_NAME)

    return ctypes.GetLastError() == 183

def on_close(tray):
    tray.stop()

def _open_manage_window():
    def window_close():
        SETTINGS["pos_x"], SETTINGS["pos_y"], SETTINGS["scale_x"], SETTINGS["scale_y"] = window.winfo_x(), window.winfo_y(), window.winfo_width(), window.winfo_height()
        save_settings()
        window.destroy()

    def window_get_app_name(path):
        filename = os.path.splitext(os.path.basename(path))[0]

        try:
            language, codepage = win32api.GetFileVersionInfo(path, '\\VarFileInfo\\Translation')[0]
            file_info = u'\\StringFileInfo\\%04X%04X\\%s' % (language, codepage, "ProductName")
            product_name = win32api.GetFileVersionInfo(path, file_info)

            return product_name or filename
        except:
            pass

        return filename

    def window_double_click(event):
        selected = apps_tree.selection()

        if selected:
            for app in APPS:
                if app["path"] == selected[0]:
                    app["enabled"] = not app["enabled"]
                    window_update_list()
                    break

    def window_remove_app():
        selected = apps_tree.selection()

        if selected:
            for app in APPS:
                if app["path"] == selected[0]:
                    confirm = messagebox.askyesno(
                        TEXTS["remove_popup"],
                        TEXTS["remove_confirm_popup"].format(
                            app_name=app["name"].upper()
                        )
                    )

                    if confirm:
                        APPS.remove(app)
                        window_update_list()
                    break

    def window_add_app():
        path = filedialog.askopenfilename(
            title=TEXTS["exe_popup"],
            filetypes=[(TEXTS["exe_desc_popup"], "*.exe")]
        )

        if path and os.path.splitext(path)[1].lower() == ".exe":
            for app in APPS:
                if app.get("path") == path:
                    return

            app = {
                "name": window_get_app_name(path),
                "path": path,
                "enabled": True
            }

            APPS.append(app)
            window_update_list()

    def window_update_list(save=True):
        apps_tree.delete(*apps_tree.get_children())

        for app in sorted(APPS, key=lambda x: x.get("name").lower()):
            enabled = app.get("enabled")
            name = app.get("name")
            path = app.get("path")


            apps_tree.insert("", "end", iid=path, values=("🗹" if enabled else "☐", name, path), tags=() if os.path.exists(path) else ("missing"))

        if save:
            save_apps()

    def window_resize(event):
        total_width = apps_tree.winfo_width()

        checkbox_width = 30
        columns_width = total_width - checkbox_width

        apps_tree.column("checkbox", width=checkbox_width)
        apps_tree.column("name", width=int(columns_width * 0.33))
        apps_tree.column("path", width=int(columns_width * 0.65))

    window = tk.Tk()

    window.title(TEXTS["title_window"])
    window.iconbitmap(ICON_ON_PATH)
    window.minsize(900, 450)
    window.maxsize(1200, 900)
    window.geometry(f"{SETTINGS["scale_x"]}x{SETTINGS["scale_y"]}+{SETTINGS["pos_x"]}+{SETTINGS["pos_y"]}")

    apps_frame = ttk.Frame(window)
    apps_frame.pack(fill="both", expand=True, padx=10, pady=10)

    scrollbar = ttk.Scrollbar(apps_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")

    apps_tree = ttk.Treeview(apps_frame, columns=("checkbox", "name", "path"), show="headings", yscrollcommand=scrollbar.set)
    apps_tree.pack(fill="both", expand=True)

    apps_tree.heading("checkbox", text="")
    apps_tree.heading("name", text=TEXTS["name_window"], anchor="w")
    apps_tree.heading("path", text=TEXTS["path_window"], anchor="w")

    apps_tree.column("checkbox", anchor="center")

    apps_tree.bind("<Configure>", window_resize)
    apps_tree.bind("<Double-1>", window_double_click)

    apps_tree.tag_configure("missing", foreground="gray")

    scrollbar.config(command=apps_tree.yview)

    button_frame = ttk.Frame(window)
    button_frame.pack(fill="x", padx=10, pady=(0, 15))

    ttk.Button(button_frame, text=TEXTS["add_window"], command=window_add_app).pack(side="left", padx=(0, 10))
    ttk.Button(button_frame, text=TEXTS["remove_window"], command=window_remove_app).pack(side="left", padx=(0, 10))

    window_update_list(False)

    children = apps_tree.get_children()

    if children:
        first_child = children[0]
        apps_tree.selection_set(first_child)
        apps_tree.focus(first_child)
        apps_tree.see(first_child)

    window.protocol("WM_DELETE_WINDOW", window_close)
    window.mainloop()

def open_manage_window(icon=None, event=None):
    open_window = True

    for t in threading.enumerate():
        if "_open_manage_window" in t.name:
            open_window = False

    if open_window:
        threading.Thread(target=_open_manage_window, daemon=True).start()

def set_update_time(time):
    SETTINGS["update_time"] = time
    save_settings()

def is_start_menu_enabled():
    start_menu = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs")
    shortcut_path = os.path.join(start_menu, f"{APP_NAME}.lnk")

    return os.path.exists(shortcut_path)

def add_to_start_menu():
    shell = win32com.client.Dispatch("WScript.Shell")
    start_menu = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs")
    shortcut_path = os.path.join(start_menu, f"{APP_NAME}.lnk")

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
    else:
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{os.path.abspath(sys.argv[0])}"'
        shortcut.WorkingDirectory = BASE_DIR
        shortcut.IconLocation = ICON_ON_PATH
        shortcut.save()

class Win32PystrayIcon(pystray.Icon):
    WM_LBUTTONDBLCLK = 0x0203

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if 'on_double_click' in kwargs:
            self.on_double_click = kwargs['on_double_click']

    def _on_notify(self, wparam, lparam):
        super()._on_notify(wparam, lparam)
        if lparam == self.WM_LBUTTONDBLCLK:
            self.on_double_click(self, None)

def start_tray():
    global TRAY

    TRAY = Win32PystrayIcon(
        APP_NAME,
        None,
        APP_NAME,
        menu=pystray.Menu(
            item(TEXTS["manage_apps_menu"], open_manage_window),
            pystray.Menu.SEPARATOR,
            item(TEXTS["update_speed_menu"], pystray.Menu(
                item(TEXTS["update_speed_fast_menu"],
                    lambda: set_update_time(1),
                    checked=lambda item: SETTINGS["update_time"] == 1),
                item(TEXTS["update_speed_normal_menu"],
                    lambda: set_update_time(2),
                    checked=lambda item: SETTINGS["update_time"] == 2),
                item(TEXTS["update_speed_slow_menu"],
                    lambda: set_update_time(3),
                    checked=lambda item: SETTINGS["update_time"] == 3),
            )),
            pystray.Menu.SEPARATOR,
            item(TEXTS["primary_display_menu"], toggle_display, checked=lambda item: SETTINGS["primary"]),
            item(TEXTS["notifications_menu"], toggle_notification, checked=lambda item: SETTINGS["notifications"]),
            item(TEXTS["pause_menu"], toggle_pause, checked=lambda item: SETTINGS["pause"]),



            item(TEXTS["start_menu"], add_to_start_menu, checked=lambda item: is_start_menu_enabled()),
            item(TEXTS["run_minimized_menu"], toggle_minimized, checked=lambda item: SETTINGS["run_minimized"]),
            item(TEXTS["autorun_menu"], add_to_autorun, checked=lambda item: is_autorun_enabled()),
            pystray.Menu.SEPARATOR,
            item(TEXTS["close_menu"], on_close)
        ),
        on_double_click = open_manage_window
    )

    update_icon()
    TRAY.run()

if not is_running():
    load_settings()
    load_texts()

    load_apps()

    threading.Thread(target=monitor_apps, daemon=True).start()

    if not SETTINGS["run_minimized"]:
        open_manage_window()

    start_tray()