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
from win11toast import toast
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_NAME = "HDR Control"
SETTINGS = {}
TEXTS = {}
TRAY = None
APPS = {}

HDR = ctypes.CDLL(os.path.join(BASE_DIR, "resources/HDRSwitch.dll"))
HDR_ENABLED = bool(HDR.GetGlobalHDRState())

ICON_ON_PATH = os.path.join(BASE_DIR, "resources/hdr_on.ico")
ICON_OFF_PATH = os.path.join(BASE_DIR, "resources/hdr_off.ico")
LOGO_PATH = os.path.join(BASE_DIR, "resources/hdr_logo.ico")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings/settings.ini")
APPS_PATH = os.path.join(BASE_DIR, "settings/apps.ini")

def load_settings():
    global SETTINGS, TEXTS

    settings_parser = configparser.ConfigParser()
    settings_parser.read(SETTINGS_PATH)

    SETTINGS = {}
    SETTINGS["LANGUAGE"] = settings_parser["UI"]["LANGUAGE"]
    SETTINGS["NOTIFICATIONS"] = settings_parser["UI"].getboolean("NOTIFICATIONS")
    SETTINGS["UPDATE_TIME"] = settings_parser["GENERAL"].getfloat("UPDATE_TIME")
    SETTINGS["PAUSE"] = settings_parser["GENERAL"].getboolean("PAUSE")
    SETTINGS["PRIMARY"] = settings_parser["GENERAL"].getboolean("PRIMARY")

    texts_parser = configparser.ConfigParser()
    texts_parser.read(os.path.join(BASE_DIR, f"language/lang_{SETTINGS['LANGUAGE']}.ini"))

    TEXTS = {}
    TEXTS["CLOSE"] = texts_parser["TEXTS"]["CLOSE"]
    TEXTS["HDR_ON"] = texts_parser["TEXTS"]["HDR_ON"]
    TEXTS["HDR_OFF"] = texts_parser["TEXTS"]["HDR_OFF"]
    TEXTS["AUTORUN"] = texts_parser["TEXTS"]["AUTORUN"]
    TEXTS["PAUSE"] = texts_parser["TEXTS"]["PAUSE"]
    TEXTS["NOTIFICATIONS"] = texts_parser["TEXTS"]["NOTIFICATIONS"]
    # TEXTS["RUN"] = texts_parser["TEXTS"]["RUN"]
    TEXTS["PRIMARY_DISPLAY"] = texts_parser["TEXTS"]["PRIMARY_DISPLAY"]
    TEXTS["SETTINGS_DIR"] = texts_parser["TEXTS"]["SETTINGS_DIR"]
    TEXTS["SETTINGS_FILE"] = texts_parser["TEXTS"]["SETTINGS_FILE"]

def save_settings():
    settings_parser = configparser.ConfigParser()
    settings_parser.read(SETTINGS_PATH)

    settings_parser["GENERAL"]["PAUSE"] = str(SETTINGS["PAUSE"]).lower()
    settings_parser["GENERAL"]["PRIMARY"] = str(SETTINGS["PRIMARY"]).lower()
    settings_parser["UI"]["NOTIFICATIONS"] = str(SETTINGS["NOTIFICATIONS"]).lower()

    with open(SETTINGS_PATH, "w") as f:
        settings_parser.write(f)

def load_apps():
    global APPS

    apps_parser = configparser.ConfigParser()
    apps_parser.optionxform = str
    apps_parser.read(APPS_PATH)

    APPS = dict(apps_parser["APPS"])

def check_running_apps():
    processes = set()

    for p in psutil.process_iter(['name']):
        try:
            processes.add(p.info['name'])
        except:
            pass

    found_app = False

    for name, path in APPS.items():
        exe_name = os.path.basename(path)

        if exe_name in processes:
            found_app = True
            break

    enable_hdr(found_app)

def monitor_apps():
    while True:
        if not SETTINGS["PAUSE"]:
            check_running_apps()

        time.sleep(SETTINGS["UPDATE_TIME"])

def enable_hdr(is_enabled):
    global HDR_ENABLED

    if HDR_ENABLED != is_enabled:
        HDR_ENABLED = is_enabled

        if SETTINGS["PRIMARY"]:
            HDR.SetHDRonPrimary(HDR_ENABLED)
        else:
            HDR.SetGlobalHDRState(HDR_ENABLED)

        update_icon()

        if SETTINGS["NOTIFICATIONS"]:
            toast_text = TEXTS["HDR_OFF"]

            if HDR_ENABLED:
                toast_text = TEXTS["HDR_ON"]

            toast(APP_NAME, toast_text, icon=LOGO_PATH)

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
    SETTINGS["PAUSE"] = not SETTINGS["PAUSE"]
    save_settings()

def is_paused():
    return SETTINGS["PAUSE"]

def toggle_notification():
    SETTINGS["NOTIFICATIONS"] = not SETTINGS["NOTIFICATIONS"]
    save_settings()

def is_notification_enabled():
    return SETTINGS["NOTIFICATIONS"]

def toggle_display():
    SETTINGS["PRIMARY"] = not SETTINGS["PRIMARY"]
    save_settings()

def is_primary_display_enabled():
    return SETTINGS["PRIMARY"]

def update_icon():
    icon = ICON_OFF_PATH

    if HDR_ENABLED:
        icon = ICON_ON_PATH

    TRAY.icon = Image.open(icon)

def is_running():
    arg = os.path.basename(sys.argv[0])

    i = 0

    for q in psutil.process_iter():
        if 'python' in q.name():
            if len(q.cmdline()) > 1 and arg in q.cmdline()[1]:
                i += 1

    return i > 1

def on_close(tray):
    tray.stop()

def open_settings_folder():
    os.startfile(os.path.join(BASE_DIR, "settings"))

def open_settings_file():
    os.startfile(os.path.join(BASE_DIR, APPS_PATH))

def start_tray():
    global TRAY

    # def run_app(path):
        # def _run():
            # os.startfile(path)
        # return _run

    # def add_apps_submenu():
        # items = []

        # for name, path in sorted(APPS.items(), key=lambda x: x[0].lower()):
            # items.append(
                # item(name, run_app(path))
            # )

        # return items

    TRAY = pystray.Icon(
        APP_NAME,
        None,
        APP_NAME,
        menu=pystray.Menu(
            # item(TEXTS["RUN"], pystray.Menu(*add_apps_submenu())),
            # pystray.Menu.SEPARATOR,
            item(TEXTS["SETTINGS_DIR"], open_settings_folder),
            item(TEXTS["SETTINGS_FILE"], open_settings_file),
            pystray.Menu.SEPARATOR,
            item(TEXTS["PRIMARY_DISPLAY"], toggle_display, checked=lambda item: is_primary_display_enabled()),
            item(TEXTS["NOTIFICATIONS"], toggle_notification, checked=lambda item: is_notification_enabled()),
            item(TEXTS["PAUSE"], toggle_pause, checked=lambda item: is_paused()),
            item(TEXTS["AUTORUN"], add_to_autorun, checked=lambda item: is_autorun_enabled()),
            pystray.Menu.SEPARATOR,
            item(TEXTS["CLOSE"], on_close)
        )
    )

    update_icon()

    TRAY.run()

if not is_running():
    load_settings()
    load_apps()
    thread = threading.Thread(target=monitor_apps, daemon=True)
    thread.start()
    start_tray()