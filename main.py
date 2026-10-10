import tkinter as tk
from tkinter import simpledialog, messagebox, ttk, colorchooser
import tkinter.font as tkfont
import json
import os
import threading
import queue
import urllib.request
import urllib.parse
import time
import math
import subprocess
import shutil
import tempfile
import hashlib
import re
from datetime import date as date_cls, timedelta
from pathlib import Path
from collections import deque
import http.server
import socketserver
import webbrowser
import sys
import ssl
from PIL import Image, ImageDraw
import pystray
import ctypes
from ctypes import wintypes
import pythoncom
import win32com.client

try:
    from lunar_python import Solar as LunarSolar
    _LUNAR_AVAILABLE = True
except Exception:
    LunarSolar = None
    _LUNAR_AVAILABLE = False


SOLAR_FESTIVALS = {
    (1, 1): "元旦",
    (2, 14): "情人节",
    (3, 8): "妇女节",
    (3, 12): "植树节",
    (4, 1): "愚人节",
    (5, 1): "劳动节",
    (5, 4): "青年节",
    (6, 1): "儿童节",
    (7, 1): "建党节",
    (8, 1): "建军节",
    (9, 10): "教师节",
    (10, 1): "国庆节",
    (12, 24): "平安夜",
    (12, 25): "圣诞节",
}


def _nth_weekday_of_month(year, month, weekday, n):
    count = 0
    for day in range(1, 32):
        try:
            d = date_cls(year, month, day)
        except ValueError:
            break
        if d.weekday() == weekday:
            count += 1
            if count == n:
                return d
    return None


def _enable_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_enable_dpi_awareness()

APP_VERSION = "1.0.9"
GITHUB_REPO = "BigSteve3838292/FloatingWeatherClock"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_API_RELEASES = f"https://api.github.com/repos/{GITHUB_REPO}/releases"
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases"

MUTEX_NAME = "FloatingWeatherClock_SingleInstance_Mutex"

# ===== Beta 版开关（出厂默认，每次打包按需修改） =====
BETA_EDITION = False
BETA_VERSION = 1

HOLIDAY_API = "https://holiday.dreace.top"

PERIODIC_CHECK_INTERVAL_MS = 15 * 60 * 1000

DEFAULT_FONT_FAMILY = "Segoe UI"

# ===== 愚人节彩蛋配置 =====
APRIL_FOOLS_ENABLED = True
APRIL_FOOLS_URL = "https://www.bilibili.com/video/BV1GJ411x7h7/"
APRIL_FOOLS_LOADING_TEXT_1 = "正在连接 Rockstar 服务器…"
APRIL_FOOLS_LOADING_TEXT_2 = "请稍候，正在缓冲 4K 超清预告片…"
APRIL_FOOLS_BUTTON_TEXT = ">>  观看 GTA7 的预告片"
APRIL_FOOLS_TITLE_1 = "*** 愚人节限定首发 ***"
APRIL_FOOLS_TITLE_2 = "Rockstar 官方授权 · 全球独家 · 4K 超清"
APRIL_FOOLS_FOOTER = "（仅 4 月 1 日可见，明天就会消失）"

user32 = ctypes.windll.user32
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020

user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.GetWindowLongPtrW.restype = ctypes.c_void_p
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
user32.SetWindowLongPtrW.restype = ctypes.c_void_p
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
user32.SetWindowPos.restype = wintypes.BOOL


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]


CLSID_TaskbarList = _GUID(
    0x56FDF344, 0xFD6D, 0x11d0,
    (ctypes.c_ubyte * 8)(0x95, 0x8A, 0x00, 0x60, 0x97, 0xC9, 0xA0, 0x90)
)
IID_ITaskbarList = _GUID(
    0x56FDF342, 0xFD6D, 0x11d0,
    (ctypes.c_ubyte * 8)(0x95, 0x8A, 0x00, 0x60, 0x97, 0xC9, 0xA0, 0x90)
)
CLSCTX_INPROC_SERVER = 0x1

ole32 = ctypes.windll.ole32


class ITaskbarListVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("HrInit", ctypes.c_void_p),
        ("AddTab", ctypes.c_void_p),
        ("DeleteTab", ctypes.c_void_p),
        ("ActivateTab", ctypes.c_void_p),
        ("SetActiveAlt", ctypes.c_void_p),
    ]


def _create_taskbar_list():
    p = ctypes.c_void_p()
    try:
        hr = ole32.CoCreateInstance(
            ctypes.byref(CLSID_TaskbarList),
            None,
            CLSCTX_INPROC_SERVER,
            ctypes.byref(IID_ITaskbarList),
            ctypes.byref(p),
        )
        if hr != 0 or not p:
            return None
        vtbl_ptr = ctypes.cast(p, ctypes.POINTER(ctypes.c_void_p))[0]
        vtbl = ctypes.cast(vtbl_ptr, ctypes.POINTER(ITaskbarListVtbl)).contents
        HrInit = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)(vtbl.HrInit)
        HrInit(p)
        return p, vtbl
    except Exception:
        return None


def delete_taskbar_tab(hwnd):
    try:
        result = _create_taskbar_list()
        if not result:
            return False
        p, vtbl = result
        DeleteTab = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, wintypes.HWND)(vtbl.DeleteTab)
        DeleteTab(p, hwnd)
        Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtbl.Release)
        Release(p)
        return True
    except Exception:
        return False


def remove_taskbar_entry(hwnd):
    try:
        ex_style = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        ex_style = ctypes.c_ulonglong(ex_style).value
        ex_style = ex_style & ~WS_EX_APPWINDOW
        ex_style = ex_style | WS_EX_TOOLWINDOW
        user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ctypes.c_void_p(ex_style))
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        delete_taskbar_tab(hwnd)
    except Exception:
        pass


def hide_toplevel_taskbar(tk_toplevel_win):
    tk_toplevel_win.update_idletasks()
    hwnd = tk_toplevel_win.winfo_id()
    try:
        ex_style = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        ex_style = ctypes.c_ulonglong(ex_style).value
        ex_style = ex_style & ~WS_EX_APPWINDOW
        ex_style = ex_style | WS_EX_TOOLWINDOW
        user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ctypes.c_void_p(ex_style))
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        delete_taskbar_tab(hwnd)
    except Exception:
        pass


def _get_current_exe_path():
    try:
        if getattr(sys, "frozen", False):
            return os.path.abspath(sys.executable)
        return os.path.abspath(sys.argv[0])
    except Exception:
        return os.path.abspath(sys.argv[0])


def _parse_version(v_str):
    if not v_str:
        return (0,)
    s = str(v_str).strip().lstrip("vV")
    parts = []
    for p in s.split("."):
        num = ""
        for ch in p:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    return tuple(parts) if parts else (0,)


def _is_newer(remote_ver, local_ver):
    r = _parse_version(remote_ver)
    l = _parse_version(local_ver)
    n = max(len(r), len(l))
    r = r + (0,) * (n - len(r))
    l = l + (0,) * (n - len(l))
    return r > l


def _fmt_size(num_bytes):
    try:
        b = float(num_bytes)
    except Exception:
        return "--"
    if b < 0:
        return "--"
    for unit in ("B", "KB", "MB", "GB"):
        if b < 1024.0:
            return f"{b:.1f} {unit}" if unit != "B" else f"{int(b)} B"
        b /= 1024.0
    return f"{b:.1f} TB"


def _fmt_duration(seconds):
    try:
        s = int(round(seconds))
    except Exception:
        return "--"
    if s < 0:
        return "--"
    if s < 60:
        return f"{s} 秒"
    if s < 3600:
        return f"{s // 60} 分 {s % 60} 秒"
    return f"{s // 3600} 小时 {(s % 3600) // 60} 分"


def _lerp_color(c1, c2, t):
    r1, g1, b1 = c1
    r2, g2, b2 = c2
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


def _rgb255_to_hex(rgb):
    r, g, b = rgb
    return f"#{r:02x}{g:02x}{b:02x}"


# ===== 风速等级 / 风向辅助 =====
def _wind_scale_kmh(speed_kmh):
    try:
        s = float(speed_kmh)
    except Exception:
        return None, ""
    if s < 0:
        s = 0
    thresholds = [
        (2, 0, "无风"),
        (6, 1, "软风"),
        (12, 2, "轻风"),
        (19, 3, "微风"),
        (30, 4, "和风"),
        (40, 5, "清风"),
        (51, 6, "强风"),
        (62, 7, "疾风"),
        (75, 8, "大风"),
        (87, 9, "烈风"),
        (103, 10, "狂风"),
        (117, 11, "暴风"),
    ]
    for limit, level, desc in thresholds:
        if s < limit:
            return level, desc
    return 12, "飓风"


def _wind_dir_to_chinese(deg):
    try:
        d = float(deg) % 360
    except Exception:
        return ""
    dirs = ["北", "东北偏北", "东北", "东北偏东",
            "东", "东南偏东", "东南", "东南偏南",
            "南", "西南偏南", "西南", "西南偏西",
            "西", "西北偏西", "西北", "西北偏北"]
    idx = int((d + 11.25) % 360 / 22.5)
    if idx < 0:
        idx = 0
    if idx >= len(dirs):
        idx = len(dirs) - 1
    return dirs[idx] + "风"


def _c_to_f(c):
    try:
        return float(c) * 9.0 / 5.0 + 32.0
    except Exception:
        return None


DOC_PATH = Path(os.path.expanduser("~")) / "Documents"
CONFIG_FILE = str(DOC_PATH / "weather_clock_config.json")
RUNTIME_LOG_FILE = os.path.join(tempfile.gettempdir(), "fwc_runtime.log")
UPDATE_CACHE_DIR = os.path.join(tempfile.gettempdir(), "fwc_update_cache")

CITY_LIST = [
    ["合肥", 31.86, 117.28],
    ["北京", 39.90, 116.40],
    ["上海", 31.23, 121.47],
    ["广州", 23.13, 113.27],
    ["深圳", 22.54, 114.06],
    ["杭州", 30.27, 120.15],
    ["南京", 32.06, 118.78],
    ["成都", 30.67, 104.07],
    ["重庆", 29.56, 106.58],
    ["武汉", 30.59, 114.31],
    ["西安", 34.27, 108.95],
    ["苏州", 31.32, 120.62],
    ["长沙", 28.23, 112.93],
    ["宁波", 29.87, 121.55],
    ["郑州", 34.76, 113.65],
    ["济南", 36.67, 117.00],
    ["青岛", 36.07, 120.38],
    ["福州", 26.08, 119.28],
    ["厦门", 24.48, 118.09]
]
CITY_DICT = {name: (lat, lon) for name, lat, lon in CITY_LIST}

WEATHER_CODE_MAP = {
    0: {"desc": "晴朗", "tag": "clear"},
    1: {"desc": "大部晴朗", "tag": "clear"},
    2: {"desc": "多云", "tag": "cloudy"},
    3: {"desc": "阴天", "tag": "overcast"},
    45: {"desc": "雾", "tag": "fog"},
    48: {"desc": "雾凇", "tag": "fog"},
    51: {"desc": "毛毛雨", "tag": "drizzle"},
    53: {"desc": "毛毛雨", "tag": "drizzle"},
    55: {"desc": "毛毛雨", "tag": "drizzle"},
    61: {"desc": "小雨", "tag": "rain"},
    63: {"desc": "中雨", "tag": "rain"},
    65: {"desc": "大雨", "tag": "rain"},
    80: {"desc": "阵雨", "tag": "rain"},
    81: {"desc": "阵雨", "tag": "rain"},
    82: {"desc": "强阵雨", "tag": "rain"},
    71: {"desc": "小雪", "tag": "snow"},
    73: {"desc": "中雪", "tag": "snow"},
    75: {"desc": "大雪", "tag": "snow"},
    85: {"desc": "阵雪", "tag": "snow"},
    86: {"desc": "阵雪", "tag": "snow"},
    95: {"desc": "雷阵雨", "tag": "thunder"},
    96: {"desc": "雷阵雨伴冰雹", "tag": "thunder"},
    99: {"desc": "雷阵雨伴冰雹", "tag": "thunder"}
}

TEST_WEATHER_TAGS = [
    ("晴天", "clear"),
    ("多云", "cloudy"),
    ("阴天", "overcast"),
    ("毛毛雨", "drizzle"),
    ("雨", "rain"),
    ("雷阵雨", "thunder"),
    ("雪", "snow"),
    ("雾", "fog"),
]

GPS_PORT = 18080
SIM_HTTP_PORT = 18923
GPS_QUEUE_TIMEOUT = 60
gps_result_queue = queue.Queue()
MAX_ERROR_LOG = 200

REFRESH_TIMEOUT_SEC = 20

STYLE_SIZE = {
    "brief": (350, 150),
    "normal": (380, 260),
    "detailed": (500, 400)
}

SUNRISE_BOTTOM = (255, 170, 90)
SUNRISE_TOP = (60, 70, 120)
SUNSET_BOTTOM = (255, 130, 80)
SUNSET_TOP = (50, 45, 90)

STARTUP_FOLDER = Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs\Startup"
LINK_NAME = "悬浮时钟天气.lnk"
LINK_PATH = STARTUP_FOLDER / LINK_NAME


def is_startup_enabled():
    if sys.platform != "win32":
        return False
    return LINK_PATH.exists()


def set_startup(enable: bool):
    if sys.platform != "win32":
        return False
    try:
        if enable:
            target_path = _get_current_exe_path()
            work_dir = os.path.dirname(target_path)
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(str(LINK_PATH))
            shortcut.TargetPath = target_path
            shortcut.WorkingDirectory = work_dir
            shortcut.WindowStyle = 1
            shortcut.Description = "悬浮时钟天气 - 桌面悬浮天气时钟"
            shortcut.Save()
            pythoncom.CoUninitialize()
            return True
        else:
            if LINK_PATH.exists():
                os.remove(LINK_PATH)
            return True
    except Exception as e:
        print(f"开机自启操作失败: {e}")
        return False


def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return r, g, b


def get_text_color_by_bg(bg_hex):
    r, g, b = hex_to_rgb(bg_hex)
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    if luminance > 0.5:
        return {"main": "#000000", "sub": "#444444"}
    return {"main": "#ffffff", "sub": "#d0d0d0"}


def _hex_to_rgb_tuple(hx):
    try:
        if isinstance(hx, str) and hx.startswith("#") and len(hx) == 7:
            return (int(hx[1:3], 16), int(hx[3:5], 16), int(hx[5:7], 16))
    except Exception:
        pass
    return (255, 255, 255)


def make_scrollable_toplevel(parent, title, geometry="640x850",
                              minsize=(600, 800), dpi_scale=1.0):
    win = tk.Toplevel(parent)
    win.title(title)

    try:
        if dpi_scale and dpi_scale > 1.0:
            m = re.match(r"(\d+)x(\d+)", geometry or "")
            if m:
                gw, gh = int(m.group(1)), int(m.group(2))
                geometry = f"{int(gw * dpi_scale)}x{int(gh * dpi_scale)}"
            if minsize and len(minsize) == 2:
                minsize = (int(minsize[0] * dpi_scale),
                           int(minsize[1] * dpi_scale))
    except Exception:
        pass

    win.geometry(geometry)
    win.minsize(*minsize)
    win.attributes("-topmost", True)
    win.transient(parent)
    hide_toplevel_taskbar(win)
    win.grab_set()

    try:
        if dpi_scale and dpi_scale > 1.0:
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                         "TkHeadingFont"):
                try:
                    f = tkfont.nametofont(name)
                    cur = f.cget("size")
                    if cur == 0:
                        continue
                    new_size = max(6, int(round(abs(cur) * dpi_scale)))
                    f.configure(size=new_size if cur > 0 else -new_size)
                except Exception:
                    pass
            try:
                style = ttk.Style(win)
                f = tkfont.nametofont("TkDefaultFont")
                for sname in ("TLabel", "TButton", "TCheckbutton",
                              "TRadiobutton", "TEntry", "TCombobox",
                              "TSpinbox", "TLabelframe", "TLabelframe.Label"):
                    try:
                        style.configure(sname, font=f)
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                win.option_add("*Font", tkfont.nametofont("TkDefaultFont"))
            except Exception:
                pass
    except Exception:
        pass

    outer_frame = ttk.Frame(win)
    outer_frame.pack(fill=tk.BOTH, expand=True)

    canvas_scroll = tk.Canvas(outer_frame, highlightthickness=0)
    scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas_scroll.yview)
    scroll_frame = ttk.Frame(canvas_scroll)
    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
    )
    canvas_scroll.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas_scroll.configure(yscrollcommand=scrollbar.set)
    canvas_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_canvas_configure(event):
        canvas_scroll.itemconfigure(canvas_scroll.find_all()[0], width=event.width)
    canvas_scroll.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event):
        try:
            canvas_scroll.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
    canvas_scroll.bind_all("<MouseWheel>", _on_mousewheel)

    main_frame = ttk.Frame(scroll_frame, padding=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    fixed_btn_frame = ttk.Frame(win, padding=(10, 8))
    fixed_btn_frame.pack(side=tk.BOTTOM, fill=tk.X)

    def on_close():
        try:
            canvas_scroll.unbind_all("<MouseWheel>")
        except Exception:
            pass
        try:
            win.grab_release()
        except Exception:
            pass
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", on_close)
    return win, main_frame, fixed_btn_frame, on_close


class GPSHttpHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def do_GET(self):
        if self.path.startswith("/gps?"):
            from urllib.parse import urlparse, parse_qs
            query = urlparse(self.path).query
            params = parse_qs(query)
            lat = params.get("lat", [None])[0]
            lon = params.get("lon", [None])[0]
            city = params.get("city", ["未知位置"])[0]
            if lat and lon:
                gps_result_queue.put({"lat": float(lat), "lon": float(lon), "city": city})
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"OK")
            return
        html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>获取GPS定位</title>
</head>
<body>
<script>
if(navigator.geolocation){
    navigator.geolocation.getCurrentPosition(pos=>{
        const lat = pos.coords.latitude;
        const lon = pos.coords.longitude;
        fetch(`/gps?lat=${lat}&lon=${lon}&city=当前定位`).then(()=>{
            document.body.innerHTML = `<h2>定位成功</h2><p>纬度:${lat}<br>经度:${lon}</p><p>可以关闭本页面</p>`;
        })
    },err=>{
        document.body.innerHTML = `<h2>定位失败</h2><p>请允许位置权限，或使用WiFi定位</p>`;
    })
}else{
    document.body.innerHTML = "<h2>浏览器不支持定位</h2>"
}
</script>
</body>
</html>
"""
        self.send_response(200)
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))


def create_tray_icon():
    img = Image.new("RGB", (64, 64), color=(34, 40, 54))
    draw = ImageDraw.Draw(img)
    draw.ellipse((8, 8, 56, 56), fill="#4299e1")
    draw.text((22, 20), "时", fill="white", font_size=22)
    return img


class DownloadState:
    def __init__(self):
        self.pause_event = threading.Event()
        self.cancel_event = threading.Event()
        self.tmp_path = None
        self.final_path = None
        self.total_size = 0
        self.downloaded = 0

    def request_cancel(self):
        self.cancel_event.set()
        self.pause_event.clear()


class _ReusableTCPServer(socketserver.TCPServer):
    allow_reuse_address = True


class FloatingWeatherClock:
    def __init__(self, root):
        self.root = root
        self.root.title("悬浮时钟天气")
        self.root.overrideredirect(True)
        try:
            self.root.wm_attributes("-toolwindow", True)
        except Exception:
            pass

        try:
            dpi = self.root.winfo_fpixels("1i")
            self.dpi_scale = dpi / 96.0
        except Exception:
            self.dpi_scale = 1.0

        try:
            if self.dpi_scale > 1.0:
                self.root.tk.call("tk", "scaling", self.dpi_scale * 1.0)
        except Exception:
            pass

        self.is_quitting = False
        self.error_log = deque(maxlen=MAX_ERROR_LOG)
        self.tray_icon = None
        self.cached_w = 0
        self.cached_h = 0
        self.last_time_str = ""
        self.need_redraw = True
        self.resize_throttle_id = None
        self.config = self.load_config()

        self.city = self.config.get("city", "合肥")
        self.city_lat = self.config.get("city_lat", 31.86)
        self.city_lon = self.config.get("city_lon", 117.28)
        self.font_size = self.config.get("font_size", 11)
        self.font_family = self.config.get("font_family", DEFAULT_FONT_FAMILY)
        self.mem_x = self.config.get("win_x", 200)
        self.mem_y = self.config.get("win_y", 200)
        self.win_style = self.config.get("win_style", "normal")
        self.win_alpha = self.config.get("win_alpha", 0.86)
        self.bg_color = self.config.get("bg_color", "#222836")
        self.always_on_top = self.config.get("always_on_top", True)
        self.lock_window = self.config.get("lock_window", False)
        self.developer_mode = self.config.get("developer_mode", False)
        self.custom_bottom_text = self.config.get("custom_bottom_text", "")
        self.show_auto_refresh_hint = self.config.get("show_auto_refresh_hint", True)
        self.show_last_refresh_hint = self.config.get("show_last_refresh_hint", True)
        self.download_multithread = self.config.get("download_multithread", False)
        self.download_resume = self.config.get("download_resume", False)
        self.enable_refresh_animation = self.config.get("enable_refresh_animation", True)
        self.auto_check_update = self.config.get("auto_check_update", True)
        self.periodic_check_update = self.config.get("periodic_check_update", True)

        self.exp_rpg_bg = self.config.get("exp_rpg_bg", False)
        self.weather_bg_enabled = self.config.get("weather_bg_enabled", False)

        # 温度单位：C / F
        self.temp_unit = self.config.get("temp_unit", "C")
        if self.temp_unit not in ("C", "F"):
            self.temp_unit = "C"

        self.debug_offline = self.config.get("debug_offline", False)

        self.test_weather_enabled = False
        self.test_weather_data = None
        self.test_weather_mode = "auto"
        self.sim_date = None

        self.is_refresh_hint = False
        self.auto_refresh_interval = 300
        self.refresh_countdown = self.auto_refresh_interval
        self.last_weather_refresh_time = None

        self.refresh_start_time = 0

        self.is_checking_update = False
        self._last_notified_version = None
        self._settings_lock_var_ref = None

        self._singleton_windows = {}
        self._dev_options_opening = False

        self.beta_edition = BETA_EDITION
        self.beta_version = BETA_VERSION

        self._log_file_handle = None
        self._init_log_file()

        self._gps_server = None
        self._gps_cancel_event = threading.Event()
        self._gps_lock = threading.Lock()

        self._sim_http_server = None
        self._sim_http_thread = None
        self._sim_fake_exe_path = None

        self.date_cache = None
        self.holiday_cache = None
        self._holiday_fetching = False

        self._flash_anim_id = None
        self._text_rgb_main = (255, 255, 255)
        self._text_rgb_date = (208, 208, 208)
        self._text_rgb_weather = (255, 255, 255)

        self._cleanup_old_files()

        self.canvas = tk.Canvas(root, bg=self.bg_color, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.recalc_window_size_by_font()
        self.root.geometry(f"+{self.mem_x}+{self.mem_y}")
        self.root.update_idletasks()

        try:
            hwnd0 = self.root.winfo_id()
            remove_taskbar_entry(hwnd0)
        except Exception:
            pass

        self._hide_taskbar_timer()

        self.root.bind("<Configure>", self.on_window_configure)
        self.root.bind("<FocusIn>", self._fix_taskbar_style)
        self.drag_x = 0
        self.drag_y = 0
        self.bind_mouse_events()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_window)

        self._lock_menu_var = tk.BooleanVar(value=self.lock_window)

        self.right_menu = tk.Menu(root, tearoff=0)
        self._build_right_menu()

        self.weather_queue = queue.Queue()
        self.weather_data = None
        self.is_fetching_weather = False

        self.root.after_idle(self.apply_window_attr)
        self.start_tray()

        self.root.after(100, self.update_clock_loop)
        self.root.after(200, self.fetch_all_async)
        self.root.after(800, self.check_weather_queue)
        self.root.after(1000, self.auto_refresh_all_timer)
        self.root.after(2000, self.auto_save_pos_timer)
        self.root.after(5000, self.check_for_update_silent)
        self.root.after(200, self._check_simulated_update_marker)
        self.root.after(PERIODIC_CHECK_INTERVAL_MS, self._periodic_check_timer)

        self.log(f"软件启动完成，版本 {self.get_display_version()}，DPI 缩放系数 {self.dpi_scale:.2f}")
        if self.beta_edition:
            self.log(f"当前为 Beta 版（beta{self.beta_version}），已禁用所有更新功能")
        if not _LUNAR_AVAILABLE:
            self.log("警告：lunar_python 未安装，农历/节气功能不可用")

        try:
            self.refresh_beta_ui()
        except Exception as e:
            self.log(f"初始化 beta UI 失败：{e}")

        if not BETA_EDITION and self.developer_mode:
            self.developer_mode = False
            try:
                self.save_config()
                self.log("[安全] 非 Beta 版检测到残留的开发者模式，已清除")
            except Exception as e:
                self.log(f"[安全] 清除残留开发者模式失败：{e}")

    # ================= 版本号输出 =================
    def get_display_version(self):
        if self.beta_edition:
            return f"v{APP_VERSION} - beta{self.beta_version}"
        return f"v{APP_VERSION}"

    def get_compare_version(self):
        if self.beta_edition:
            return f"{APP_VERSION}-beta{self.beta_version}"
        return APP_VERSION

    # ================= 温度单位 =================
    def _temp_symbol(self):
        return "℉" if self.temp_unit == "F" else "℃"

    def _format_temp(self, c_value):
        try:
            c = float(c_value)
        except Exception:
            return f"{c_value}{self._temp_symbol()}"
        if self.temp_unit == "F":
            f = _c_to_f(c)
            if f is None:
                return f"{c_value}℉"
            return f"{int(round(f))}℉"
        return f"{int(round(c))}℃"

    # ================= 日志 =================
    def _init_log_file(self):
        try:
            os.makedirs(os.path.dirname(RUNTIME_LOG_FILE), exist_ok=True)
            self._log_file_handle = open(RUNTIME_LOG_FILE, "a", encoding="utf-8", buffering=1)
            self._log_file_handle.write(
                f"\n===== {time.strftime('%Y-%m-%d %H:%M:%S')} 启动 v{APP_VERSION} =====\n"
            )
        except Exception:
            self._log_file_handle = None

    def log(self, msg):
        t = time.strftime("%H:%M:%S", time.localtime())
        line = f"[{t}] {msg}"
        try:
            print(line, flush=True)
        except Exception:
            pass
        if self._log_file_handle is not None:
            try:
                self._log_file_handle.write(line + "\n")
            except Exception:
                pass
        try:
            self.error_log.append(line)
        except Exception:
            pass

    def debug_open_terminal_log(self):
        if sys.platform != "win32":
            messagebox.showinfo("提示", "该功能仅支持 Windows。", parent=self.root)
            return
        try:
            os.makedirs(os.path.dirname(RUNTIME_LOG_FILE), exist_ok=True)
            if not os.path.exists(RUNTIME_LOG_FILE):
                open(RUNTIME_LOG_FILE, "a", encoding="utf-8").close()
            ps_cmd = (
                "$host.UI.RawUI.WindowTitle='悬浮时钟天气 - 实时日志'; "
                f"Write-Host '日志文件: {RUNTIME_LOG_FILE}' -ForegroundColor Green; "
                f"Get-Content -Path '{RUNTIME_LOG_FILE}' -Wait -Encoding UTF8 -Tail 50"
            )
            subprocess.Popen(
                ["powershell", "-NoExit", "-Command", ps_cmd],
                creationflags=0x00000010,
            )
            self.log("已打开实时日志终端")
        except Exception as e:
            self.log(f"打开日志终端失败: {e}")
            messagebox.showerror("打开失败", f"无法打开终端：\n\n{e}", parent=self.root)

    def add_error_log(self, msg):
        self.log(msg)

    # ============= 日期 / 农历 / 节气 / 节假日 =============
    def _get_today(self):
        if self.sim_date:
            try:
                y, m, d = map(int, self.sim_date.split("-"))
                return date_cls(y, m, d)
            except Exception:
                pass
        return date_cls.today()

    def _get_date_info(self):
        if not _LUNAR_AVAILABLE:
            return None
        try:
            today = self._get_today()
            solar = LunarSolar.fromYmd(today.year, today.month, today.day)
            lunar = solar.getLunar()

            lunar_day = lunar.getMonthInChinese() + "月" + lunar.getDayInChinese()
            gan_zhi = lunar.getYearInGanZhi()
            zodiac = lunar.getYearShengXiao()

            jie_qi_today = lunar.getJieQi() or ""

            jie_qi_tomorrow = ""
            try:
                tomorrow = today + timedelta(days=1)
                s2 = LunarSolar.fromYmd(tomorrow.year, tomorrow.month, tomorrow.day)
                jie_qi_tomorrow = s2.getLunar().getJieQi() or ""
            except Exception:
                pass

            festivals = list(lunar.getFestivals() or [])

            other_list = []
            key = (today.month, today.day)
            if key in SOLAR_FESTIVALS:
                other_list.append(SOLAR_FESTIVALS[key])

            try:
                if _nth_weekday_of_month(today.year, 5, 6, 2) == today:
                    other_list.append("母亲节")
                if _nth_weekday_of_month(today.year, 6, 6, 3) == today:
                    other_list.append("父亲节")
                if _nth_weekday_of_month(today.year, 11, 3, 4) == today:
                    other_list.append("感恩节")
            except Exception:
                pass

            for f in (lunar.getOtherFestivals() or []):
                if f not in other_list:
                    other_list.append(f)

            return {
                "date": today.strftime("%Y-%m-%d"),
                "lunar_day": lunar_day,
                "gan_zhi": gan_zhi,
                "zodiac": zodiac,
                "jie_qi_today": jie_qi_today,
                "jie_qi_tomorrow": jie_qi_tomorrow,
                "festivals": festivals,
                "other_festivals": other_list,
            }
        except Exception as e:
            self.log(f"日期信息获取失败: {e}")
            return None

    def _fetch_holiday_info_async(self, force=False):
        if self._holiday_fetching:
            return
        self._holiday_fetching = True

        def worker():
            try:
                today = self._get_today()
                date_str = today.strftime("%Y-%m-%d")

                if (not force
                        and self.holiday_cache is not None
                        and self.holiday_cache.get("date") == date_str
                        and self.holiday_cache.get("info") != "failed"):
                    return

                if self.debug_offline:
                    self.holiday_cache = {"date": date_str, "info": "failed"}
                    return

                url = f"{HOLIDAY_API}?date={date_str}"
                req = urllib.request.Request(url, headers={
                    "User-Agent": f"FloatingWeatherClock/{APP_VERSION}"
                })
                with urllib.request.urlopen(req, timeout=REFRESH_TIMEOUT_SEC) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                self.holiday_cache = {"date": date_str, "info": data}
                self.log(f"节假日查询成功：{date_str} → {data}")
            except Exception as e:
                self.log(f"节假日查询失败：{e}")
                try:
                    today = self._get_today()
                    self.holiday_cache = {"date": today.strftime("%Y-%m-%d"), "info": "failed"}
                except Exception:
                    pass
            finally:
                self._holiday_fetching = False
                try:
                    self.root.after(0, self._mark_redraw)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _mark_redraw(self):
        self.need_redraw = True
        try:
            self.update_clock()
        except Exception:
            pass

    def _get_holiday_display(self):
        if self.holiday_cache is None:
            return "loading"
        info = self.holiday_cache.get("info")
        if info == "failed":
            return "failed"
        return info

    def _is_date_ready(self):
        return (self.date_cache is not None
                and self.holiday_cache is not None)

    def _is_date_timed_out(self):
        if self.refresh_start_time <= 0:
            return False
        return (time.time() - self.refresh_start_time) > REFRESH_TIMEOUT_SEC

    def _is_weather_timed_out(self):
        if self.refresh_start_time <= 0:
            return False
        return (time.time() - self.refresh_start_time) > REFRESH_TIMEOUT_SEC

    # ================= 刷新动画 =================
    def _flash_refresh_animation(self):
        if not self.enable_refresh_animation:
            return

        if self._flash_anim_id is not None:
            try:
                self.root.after_cancel(self._flash_anim_id)
            except Exception:
                pass
            self._flash_anim_id = None

        if not self.canvas.winfo_exists():
            return

        start = (120, 190, 255)
        main_end = tuple(self._text_rgb_main)
        date_end = tuple(self._text_rgb_date)
        weather_end = tuple(self._text_rgb_weather)

        tracks = [
            ("main_text",    start, main_end,    0),
            ("date_text",    start, date_end,    3),
            ("weather_text", start, weather_end, 6),
        ]

        steps = 10
        interval = 40
        total_frames = steps + 6

        sweep_id = None
        try:
            if self.cached_h > 20 and self.cached_w > 20:
                sweep_id = self.canvas.create_line(
                    0, 0, self.cached_w, 0,
                    fill="#78beff", width=2,
                    tags=("refresh_sweep",),
                )
        except Exception:
            sweep_id = None

        def step(frame=0):
            if self.is_quitting or not self.canvas.winfo_exists():
                self._flash_anim_id = None
                try:
                    self.canvas.delete("refresh_sweep")
                except Exception:
                    pass
                return

            if frame > total_frames:
                self._flash_anim_id = None
                for tag, _, end_col, _ in tracks:
                    try:
                        self.canvas.itemconfig(tag, fill=_rgb255_to_hex(end_col))
                    except Exception:
                        pass
                try:
                    self.canvas.delete("refresh_sweep")
                except Exception:
                    pass
                return

            for tag, s_col, e_col, delay in tracks:
                local = frame - delay
                if local < 0:
                    continue
                if local >= steps:
                    rgb = e_col
                else:
                    t = local / steps
                    rgb = _lerp_color(s_col, e_col, t)
                try:
                    self.canvas.itemconfig(tag, fill=_rgb255_to_hex(rgb))
                except Exception:
                    pass

            if sweep_id is not None:
                try:
                    y = int(self.cached_h * frame / max(total_frames, 1))
                    self.canvas.coords("refresh_sweep", 0, y, self.cached_w, y)
                except Exception:
                    pass

            self._flash_anim_id = self.root.after(interval, lambda: step(frame + 1))

        step(0)

    # ================= 模拟更新标记 =================
    def _get_simulated_update_marker_path(self):
        return os.path.join(tempfile.gettempdir(), "fwc_simulated_update.marker")

    def _check_simulated_update_marker(self):
        marker = self._get_simulated_update_marker_path()
        if not os.path.exists(marker):
            return
        try:
            with open(marker, "r", encoding="utf-8") as f:
                info = json.load(f)
            try:
                os.remove(marker)
            except Exception:
                pass
            self.root.after(1300, lambda: self._show_simulated_update_notice(info))
        except Exception as e:
            self.log(f"读取模拟更新标记失败: {e}")
            try:
                os.remove(marker)
            except Exception:
                pass

    def _show_simulated_update_notice(self, info):
        try:
            old_version = info.get("old_version", "未知")
            ts = info.get("timestamp", "")
            messagebox.showinfo(
                "模拟更新已启动",
                f"当前进程由「模拟更新」启动，流程正常。\n\n"
                f"启动时的版本号：v{old_version}\n"
                f"当前进程版本号：{self.get_display_version()}\n"
                f"模拟时间：{ts}\n\n"
                f"（这是一次测试，程序内容没有实际变化）",
                parent=self.root
            )
        except Exception:
            pass

    def _cleanup_old_files(self):
        try:
            backup_exe = _get_current_exe_path() + ".old"
            if os.path.exists(backup_exe):
                def _try_delete():
                    for _ in range(10):
                        try:
                            os.remove(backup_exe)
                            return
                        except Exception:
                            time.sleep(0.5)
                threading.Thread(target=_try_delete, daemon=True).start()
        except Exception:
            pass

    # ================= 任务栏修复 =================
    def _fix_taskbar_style(self, event=None):
        try:
            try:
                self.root.wm_attributes("-toolwindow", True)
            except Exception:
                pass
            hwnd = self.root.winfo_id()
            ex_style = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
            ex_style = ctypes.c_ulonglong(ex_style).value
            ex_style = (ex_style & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
            user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ctypes.c_void_p(ex_style))
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
            delete_taskbar_tab(hwnd)
        except Exception:
            pass

    def _show_window_and_fix(self):
        try:
            self.root.deiconify()
            self.root.update_idletasks()
            self._fix_taskbar_style()
        except Exception:
            pass

    # ================= DPI 辅助 =================
    def _apply_dpi_to_toplevel(self, win):
        try:
            scale = getattr(self, "dpi_scale", 1.0)
            if scale <= 1.0:
                return

            win.update_idletasks()
            cur_w = win.winfo_width()
            cur_h = win.winfo_height()
            if cur_w > 1 and cur_h > 1:
                win.geometry(f"{int(cur_w * scale)}x{int(cur_h * scale)}")

            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                         "TkHeadingFont", "TkCaptionFont", "TkSmallCaptionFont",
                         "TkIconFont", "TkTooltipFont"):
                try:
                    f = tkfont.nametofont(name)
                    cur = f.cget("size")
                    if cur == 0:
                        continue
                    new_size = max(6, int(round(abs(cur) * scale)))
                    f.configure(size=new_size if cur > 0 else -new_size)
                except Exception:
                    pass

            try:
                style = ttk.Style(win)
                for sname in ("TLabel", "TButton", "TCheckbutton",
                              "TRadiobutton", "TEntry", "TCombobox",
                              "TSpinbox", "TNotebook", "TNotebook.Tab",
                              "TFrame", "TLabelframe", "TLabelframe.Label"):
                    try:
                        f = tkfont.nametofont("TkDefaultFont")
                        style.configure(sname, font=f)
                    except Exception:
                        pass
            except Exception:
                pass

            try:
                win.option_add("*Font", tkfont.nametofont("TkDefaultFont"))
            except Exception:
                pass

            try:
                sw = win.winfo_screenwidth()
                sh = win.winfo_screenheight()
                new_w = min(int(cur_w * scale), int(sw * 0.95))
                new_h = min(int(cur_h * scale), int(sh * 0.9))
                win.geometry(f"{new_w}x{new_h}")
                win.minsize(int(new_w * 0.7), int(new_h * 0.7))
            except Exception:
                pass

            try:
                self._center_on_parent(win)
            except Exception:
                pass

        except Exception as e:
            try:
                self.log(f"[DPI] 应用 DPI 到 Toplevel 失败：{e}")
            except Exception:
                pass

    def _center_on_parent(self, win):
        try:
            win.update_idletasks()
            w = win.winfo_width()
            h = win.winfo_height()
            try:
                px = self.root.winfo_x()
                py = self.root.winfo_y()
                pw = self.root.winfo_width()
                ph = self.root.winfo_height()
            except Exception:
                px = py = 0
                pw = ph = 0

            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()

            if pw > 0 and ph > 0:
                x = px + (pw - w) // 2
                y = py + (ph - h) // 2
            else:
                x = (sw - w) // 2
                y = (sh - h) // 2

            x = max(0, min(x, sw - w))
            y = max(0, min(y, sh - h))
            win.geometry(f"+{x}+{y}")
        except Exception:
            pass

    # ================= 单例窗口管理 =================
    def _open_singleton(self, key, creator):
        existing = self._singleton_windows.get(key)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return None
            except Exception:
                pass
            self._singleton_windows.pop(key, None)

        win = creator()
        if win is None:
            return None
        self._singleton_windows[key] = win

        def _on_destroy(event, k=key):
            if event.widget is win:
                if self._singleton_windows.get(k) is win:
                    self._singleton_windows.pop(k, None)

        try:
            win.bind("<Destroy>", _on_destroy, add="+")
        except Exception:
            pass

        return win

    # ================= 字体预览对话框 =================
    def preview_font_dialog(self, new_font):
        try:
            result_holder = {"accept": False}

            win = tk.Toplevel(self.root)
            win.title("字体预览")
            win.geometry("500x280")
            win.resizable(False, False)
            win.attributes("-topmost", True)
            win.transient(self.root)
            hide_toplevel_taskbar(win)
            win.grab_set()

            self._apply_dpi_to_toplevel(win)

            main_frame = ttk.Frame(win, padding=16)
            main_frame.pack(fill=tk.BOTH, expand=True)

            tk.Label(main_frame,
                     text="字体看起来正常吗？",
                     font=(DEFAULT_FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 4))
            tk.Label(main_frame,
                     text="如果正常请点击「是」，反之请点击「否」。",
                     font=(DEFAULT_FONT_FAMILY, 9),
                     fg="#666666").pack(anchor="w", pady=(0, 12))

            preview_frame = tk.Frame(main_frame, bg="#f5f5f5",
                                     bd=1, relief="solid", height=110)
            preview_frame.pack(fill=tk.X, pady=(0, 6))
            preview_frame.pack_propagate(False)

            preview_inner = tk.Frame(preview_frame, bg="#f5f5f5")
            preview_inner.pack(expand=True, fill=tk.BOTH, padx=12, pady=10)

            try:
                tk.Label(preview_inner,
                         text="12:34:56  2026年10月7日  星期三",
                         font=(new_font, 14, "bold"),
                         bg="#f5f5f5", fg="#222222").pack(anchor="w")
                tk.Label(preview_inner,
                         text="悬浮时钟天气 · 字体预览",
                         font=(new_font, 11),
                         bg="#f5f5f5", fg="#222222").pack(anchor="w", pady=(6, 0))
                tk.Label(preview_inner,
                         text="Sample Text 0123456789",
                         font=(new_font, 10),
                         bg="#f5f5f5", fg="#555555").pack(anchor="w", pady=(4, 0))
            except Exception:
                tk.Label(preview_inner,
                         text="（该字体无法渲染，请点击「否」）",
                         font=(DEFAULT_FONT_FAMILY, 10),
                         bg="#f5f5f5", fg="#cc0000").pack(anchor="w")

            tk.Label(main_frame,
                     text=f"字体名称：{new_font}",
                     font=(DEFAULT_FONT_FAMILY, 8),
                     fg="#888888").pack(anchor="w", pady=(0, 12))

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(side=tk.BOTTOM)

            def on_yes():
                result_holder["accept"] = True
                try:
                    win.grab_release()
                except Exception:
                    pass
                win.destroy()

            def on_no():
                result_holder["accept"] = False
                try:
                    win.grab_release()
                except Exception:
                    pass
                win.destroy()

            win.protocol("WM_DELETE_WINDOW", on_no)

            ttk.Button(btn_frame, text="是", command=on_yes, width=12).grid(row=0, column=0, padx=8)
            ttk.Button(btn_frame, text="否", command=on_no, width=12).grid(row=0, column=1, padx=8)

            self.root.wait_window(win)
            return result_holder["accept"]
        except Exception as e:
            self.log(f"字体预览对话框异常: {e}")
            return False

    # ================= 愚人节彩蛋 =================
    def _is_april_fools_today(self):
        if not APRIL_FOOLS_ENABLED:
            return False
        try:
            today = self._get_today()
            return (today.month == 4 and today.day == 1)
        except Exception:
            return False

    def _show_april_fools_egg(self, parent_frame, parent_win):
        """在设置面板最顶部绘制愚人节彩蛋。
        为了保证稳定：外层 Frame 只做容器，内部逐控件加 try/except，
        按钮用 Label + bind 模拟，避免 tk.Button 在部分 Windows 主题下尺寸塌缩。
        """
        try:
            egg_frame = tk.Frame(parent_frame, bg="#ff2b2b", bd=3, relief="solid")
            egg_frame.pack(fill=tk.X, pady=(0, 12))
        except Exception as e:
            self.log(f"[愚人节彩蛋] 创建外框失败：{e}")
            return

        # ---- 标题 1 ----
        try:
            tk.Label(egg_frame,
                     text=APRIL_FOOLS_TITLE_1,
                     bg="#ff2b2b", fg="#ffff00",
                     font=(self.font_family, 14, "bold"),
                     padx=8, pady=8).pack(fill=tk.X)
        except Exception as e:
            self.log(f"[愚人节彩蛋] 标题1失败：{e}")

        # ---- 标题 2 ----
        try:
            tk.Label(egg_frame,
                     text=APRIL_FOOLS_TITLE_2,
                     bg="#ff2b2b", fg="#ffffff",
                     font=(self.font_family, 9),
                     padx=8).pack(fill=tk.X)
        except Exception as e:
            self.log(f"[愚人节彩蛋] 标题2失败：{e}")

        # ---- 按钮点击处理 ----
        def _open_trailer(_event=None):
            try:
                loading = tk.Toplevel(parent_win)
                loading.title("正在加载预告片")
                loading.geometry("360x140")
                loading.resizable(False, False)
                loading.attributes("-topmost", True)
                loading.transient(parent_win)
                try:
                    hide_toplevel_taskbar(loading)
                except Exception:
                    pass

                tk.Label(loading,
                         text=APRIL_FOOLS_LOADING_TEXT_1,
                         font=(self.font_family, 11, "bold"),
                         fg="#2f80ed").pack(pady=(20, 6))
                tk.Label(loading,
                         text=APRIL_FOOLS_LOADING_TEXT_2,
                         font=(self.font_family, 9),
                         fg="#666666").pack()

                prog = ttk.Progressbar(loading, mode="indeterminate", length=280)
                prog.pack(pady=12)
                prog.start(15)

                def _really_open():
                    try:
                        loading.destroy()
                    except Exception:
                        pass
                    try:
                        webbrowser.open(APRIL_FOOLS_URL)
                    except Exception as e:
                        self.log(f"打开愚人节彩蛋失败：{e}")

                loading.after(2500, _really_open)
            except Exception as e:
                self.log(f"[愚人节彩蛋] 打开失败：{e}")

        # ---- 黄底红字大按钮（用 Label + bind，避免 Button 主题问题）----
        try:
            btn = tk.Label(
                egg_frame,
                text=APRIL_FOOLS_BUTTON_TEXT,
                bg="#ffff00", fg="#cc0000",
                font=(self.font_family, 16, "bold"),
                cursor="hand2",
                padx=20, pady=14,
            )
            btn.pack(fill=tk.X, padx=10, pady=(10, 10))
            btn.bind("<Button-1>", _open_trailer)

            # 鼠标悬停变色（增强"可点击"感）
            def _on_enter(_e):
                try:
                    btn.config(bg="#ffdd00")
                except Exception:
                    pass

            def _on_leave(_e):
                try:
                    btn.config(bg="#ffff00")
                except Exception:
                    pass

            btn.bind("<Enter>", _on_enter)
            btn.bind("<Leave>", _on_leave)
        except Exception as e:
            self.log(f"[愚人节彩蛋] 按钮失败：{e}")

        # ---- 底部说明 ----
        try:
            tk.Label(egg_frame,
                     text=APRIL_FOOLS_FOOTER,
                     bg="#ff2b2b", fg="#ffdddd",
                     font=(self.font_family, 8),
                     padx=8, pady=(0, 6)).pack(fill=tk.X)
        except Exception as e:
            self.log(f"[愚人节彩蛋] 底部说明失败：{e}")


    # ================= 菜单构建 =================
    def _build_right_menu(self):
        try:
            self.right_menu.delete(0, tk.END)
        except Exception:
            pass
        self.right_menu.add_command(label="显示窗口", command=self._show_window_and_fix)
        self.right_menu.add_command(label="刷新全局", command=self.menu_refresh_all)
        self.right_menu.add_checkbutton(label="锁定窗口位置",
                                        variable=self._lock_menu_var,
                                        command=self._on_lock_menu_clicked)
        self.right_menu.add_command(label="打开设置", command=self.open_settings)
        if self.beta_edition:
            self.right_menu.add_command(label="检查更新",
                                        command=self.menu_check_update,
                                        state="disabled")
        else:
            self.right_menu.add_command(label="检查更新",
                                        command=self.menu_check_update)
        self.right_menu.add_command(label="关于本软件", command=self.show_about)
        self.right_menu.add_separator()
        self.right_menu.add_command(label="开发者选项", command=self.open_dev_options)
        self.right_menu.add_separator()
        self.right_menu.add_command(label="退出", command=self.quit_app)

    def _build_tray_menu(self):
        def tray_open_settings():
            self.root.after(0, self.open_settings)

        def tray_show_window():
            self.root.after(0, self._show_window_and_fix)

        def tray_show_about():
            self.root.after(0, self.show_about)

        def tray_dev_options():
            self.root.after(0, self.open_dev_options)

        def tray_refresh_all():
            self.root.after(0, self.menu_refresh_all)

        def tray_check_update():
            self.root.after(0, self.menu_check_update)

        def tray_toggle_lock(icon, item):
            self.root.after(0, self._toggle_lock_from_tray)

        def tray_quit():
            self.root.after(0, self.quit_app)

        check_item = pystray.MenuItem(
            "检查更新",
            lambda: tray_check_update(),
            enabled=(not self.beta_edition),
        )

        menu = pystray.Menu(
            pystray.MenuItem("显示窗口", lambda: tray_show_window()),
            pystray.MenuItem("刷新全局", lambda: tray_refresh_all()),
            pystray.MenuItem("打开设置", lambda: tray_open_settings()),
            check_item,
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("锁定窗口位置", tray_toggle_lock,
                             checked=lambda item: self.lock_window),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("关于本软件", lambda: tray_show_about()),
            pystray.MenuItem("开发者选项", lambda: tray_dev_options()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出", lambda: tray_quit())
        )
        return menu

    def refresh_beta_ui(self):
        try:
            if self.beta_edition:
                self.root.title(f"悬浮时钟天气 {self.get_display_version()}")
            else:
                self.root.title("悬浮时钟天气")
        except Exception:
            pass

        try:
            self._build_right_menu()
        except Exception as e:
            self.log(f"重建右键菜单失败：{e}")

        try:
            if self.tray_icon is not None:
                self.tray_icon.menu = self._build_tray_menu()
                self.tray_icon.title = self.get_display_version()
                self.tray_icon.update_menu()
        except Exception as e:
            self.log(f"更新托盘失败：{e}")

        self.need_redraw = True
        try:
            self.update_clock()
        except Exception:
            pass

    # ================= 更新相关 =================
    def check_for_update_silent(self):
        if self.beta_edition:
            return
        if not self.auto_check_update:
            self.log("启动时自动检查更新已关闭，跳过")
            return
        self.check_for_update_async(manual=False)

    def _periodic_check_timer(self):
        if self.is_quitting:
            return
        try:
            if (not self.beta_edition
                    and self.periodic_check_update
                    and not self.is_checking_update
                    and not self.debug_offline):
                self.log("后台定时检查更新：开始")
                self.check_for_update_async(manual=False)
        except Exception as e:
            self.log(f"后台定时检查更新异常：{e}")
        self.root.after(PERIODIC_CHECK_INTERVAL_MS, self._periodic_check_timer)

    def menu_check_update(self):
        if self.beta_edition:
            return
        self.check_for_update_async(manual=True)

    def check_for_update_async(self, manual=False):
        if self.is_checking_update:
            if manual:
                messagebox.showinfo("检查更新", "正在检查更新，请稍候…", parent=self.root)
            return
        if self.debug_offline:
            if manual:
                messagebox.showwarning("检查更新", "调试模式：已关闭软件联网", parent=self.root)
            return
        self.is_checking_update = True
        self.log(f"开始检查更新（{'手动' if manual else '自动'}）…")

        def worker():
            try:
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(
                    GITHUB_API_LATEST,
                    headers={
                        "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                        "Accept": "application/vnd.github+json",
                    },
                )
                with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                tag = data.get("tag_name") or data.get("name") or ""
                download_url = None
                for asset in data.get("assets", []) or []:
                    name = (asset.get("name") or "").lower()
                    if name.endswith(".exe"):
                        download_url = asset.get("browser_download_url")
                        break
                result = {
                    "tag": tag,
                    "body": data.get("body") or "",
                    "html_url": data.get("html_url") or GITHUB_RELEASES_PAGE,
                    "download_url": download_url,
                }
                self.log(f"检查更新完成，远程版本：{tag}")
                self.root.after(0, lambda: self._on_update_checked(result, manual))
            except Exception as e:
                self.log(f"检查更新失败: {e}")
                if manual:
                    self.root.after(0, lambda err=e: messagebox.showerror(
                        "检查更新",
                        f"检查更新失败，请检查网络或稍后再试。\n\n错误：{err}",
                        parent=self.root))
            finally:
                self.is_checking_update = False

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_checked(self, result, manual):
        if self.is_quitting:
            return
        tag = result.get("tag", "")
        if not tag:
            if manual:
                messagebox.showinfo("检查更新", "未找到可用版本信息。", parent=self.root)
            return
        if _is_newer(tag, APP_VERSION):
            if manual:
                self.show_update_dialog(result)
            else:
                if tag != self._last_notified_version:
                    self._last_notified_version = tag
                    self.log(f"自动检查发现新版本 {tag}，弹出提示")
                    self.show_update_dialog(result)
                else:
                    self.log(f"自动检查发现新版本 {tag}，已提示过，静默跳过")
        else:
            if manual:
                messagebox.showinfo("检查更新", f"已是最新版本 {self.get_display_version()}。",
                                    parent=self.root)

    def show_update_dialog(self, result):
        tag = result.get("tag", "")
        body = result.get("body", "") or "（无更新说明）"
        html_url = result.get("html_url", GITHUB_RELEASES_PAGE)
        download_url = result.get("download_url")

        win = tk.Toplevel(self.root)
        win.title("发现新版本")
        win.geometry("520x420")
        win.minsize(480, 380)
        win.attributes("-topmost", True)
        win.transient(self.root)
        hide_toplevel_taskbar(win)
        win.grab_set()

        self._apply_dpi_to_toplevel(win)

        main_frame = ttk.Frame(win, padding=14)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text=f"发现新版本 {tag}",
                 font=(self.font_family, 13, "bold"), fg="#2f80ed").pack(anchor="w")
        tk.Label(main_frame, text=f"当前版本：{self.get_display_version()}",
                 font=(self.font_family, 9), fg="#666666").pack(anchor="w", pady=(2, 8))

        tk.Label(main_frame, text="更新内容：", font=(self.font_family, 10, "bold")).pack(anchor="w")
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 8))
        txt = tk.Text(text_frame, wrap="word", height=10, font=(self.font_family, 9))
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.config(yscrollcommand=sb.set, state="normal")
        txt.insert("1.0", body)
        txt.config(state="disabled")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=6)

        def open_release_page():
            webbrowser.open(html_url)
            win.destroy()

        def do_download_install():
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
            self.download_and_install_update(download_url)

        if download_url:
            ttk.Button(btn_frame, text="下载并安装", command=do_download_install).grid(row=0, column=0, padx=6)
            ttk.Button(btn_frame, text="前往下载页", command=open_release_page).grid(row=0, column=1, padx=6)
        else:
            ttk.Button(btn_frame, text="查看发布页", command=open_release_page).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="稍后", command=win.destroy).grid(row=0, column=2, padx=6)

        self.root.wait_window(win)

    def prompt_upgrade_to_beta(self):
        if self.beta_edition:
            return
        if self.debug_offline:
            messagebox.showwarning("升级至 Beta 版", "调试模式：已关闭软件联网", parent=self.root)
            return

        def worker():
            try:
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(
                    GITHUB_API_RELEASES,
                    headers={
                        "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                        "Accept": "application/vnd.github+json",
                    },
                )
                with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                    releases = json.loads(resp.read().decode("utf-8"))
                latest_beta = None
                for r in releases:
                    if r.get("prerelease"):
                        latest_beta = r
                        break
                if latest_beta is None:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "升级至 Beta 版", "当前没有可用的 Beta 版。", parent=self.root))
                    return

                tag = latest_beta.get("tag_name") or latest_beta.get("name") or ""
                body = latest_beta.get("body") or "（无更新说明）"
                html_url = latest_beta.get("html_url") or GITHUB_RELEASES_PAGE
                download_url = None
                for asset in latest_beta.get("assets", []) or []:
                    name = (asset.get("name") or "").lower()
                    if name.endswith(".exe"):
                        download_url = asset.get("browser_download_url")
                        break

                result = {
                    "tag": tag,
                    "body": body,
                    "html_url": html_url,
                    "download_url": download_url,
                    "is_beta": True,
                }
                self.root.after(0, lambda: self.show_beta_upgrade_dialog(result))
            except Exception as e:
                self.log(f"升级至 Beta 版失败：{e}")
                self.root.after(0, lambda err=e: messagebox.showerror(
                    "升级至 Beta 版", f"获取 Beta 版信息失败：\n\n{err}", parent=self.root))

        threading.Thread(target=worker, daemon=True).start()

    def show_beta_upgrade_dialog(self, result):
        tag = result.get("tag", "")
        body = result.get("body", "") or "（无更新说明）"
        html_url = result.get("html_url", GITHUB_RELEASES_PAGE)
        download_url = result.get("download_url")

        win = tk.Toplevel(self.root)
        win.title("发现 Beta 版")
        win.geometry("540x480")
        win.minsize(500, 420)
        win.attributes("-topmost", True)
        win.transient(self.root)
        hide_toplevel_taskbar(win)
        win.grab_set()

        self._apply_dpi_to_toplevel(win)

        main_frame = ttk.Frame(win, padding=14)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text=f"发现 Beta 版 {tag}",
                 font=(self.font_family, 13, "bold"), fg="#e08a00").pack(anchor="w")
        tk.Label(main_frame, text=f"当前版本：{self.get_display_version()}",
                 font=(self.font_family, 9), fg="#666666").pack(anchor="w", pady=(2, 8))

        warn_frame = tk.Frame(main_frame, bg="#fff4e5", bd=1, relief="solid")
        warn_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(warn_frame,
                 text="⚠ 升级到 Beta 版后，软件将不再自动接收任何更新。\n"
                      "想回到正式版，需要手动下载正式版安装包覆盖。",
                 bg="#fff4e5", fg="#c05000", justify="left",
                 font=(self.font_family, 9, "bold"), padx=8, pady=6).pack(anchor="w")

        tk.Label(main_frame, text="更新内容：", font=(self.font_family, 10, "bold")).pack(anchor="w")
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 8))
        txt = tk.Text(text_frame, wrap="word", height=10, font=(self.font_family, 9))
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.config(yscrollcommand=sb.set, state="normal")
        txt.insert("1.0", body)
        txt.config(state="disabled")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=6)

        def open_release_page():
            webbrowser.open(html_url)
            win.destroy()

        def do_download():
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
            self.download_and_install_update(download_url)

        if download_url:
            ttk.Button(btn_frame, text="下载并安装 Beta 版",
                       command=do_download).grid(row=0, column=0, padx=6)
            ttk.Button(btn_frame, text="前往下载页",
                       command=open_release_page).grid(row=0, column=1, padx=6)
        else:
            ttk.Button(btn_frame, text="查看发布页",
                       command=open_release_page).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="稍后", command=win.destroy).grid(row=0, column=2, padx=6)

        self.root.wait_window(win)

    # ========== 下载相关：辅助 ==========
    def _get_cache_dir_for_url(self, download_url, mode="single"):
        try:
            os.makedirs(UPDATE_CACHE_DIR, exist_ok=True)
            h = hashlib.md5(download_url.encode("utf-8")).hexdigest()[:16]
            d = os.path.join(UPDATE_CACHE_DIR, f"{h}_{mode}")
            os.makedirs(d, exist_ok=True)
            return d
        except Exception:
            return None

    def _read_cache_meta(self, cache_dir):
        try:
            meta_path = os.path.join(cache_dir, "meta.json")
            if not os.path.exists(meta_path):
                return None
            with open(meta_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _write_cache_meta(self, cache_dir, meta):
        try:
            meta_path = os.path.join(cache_dir, "meta.json")
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _write_chunks(self, cache_dir, chunks):
        try:
            path = os.path.join(cache_dir, "chunks.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"chunks": chunks}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _clear_cache_dir(self, cache_dir):
        try:
            if cache_dir and os.path.isdir(cache_dir):
                shutil.rmtree(cache_dir, ignore_errors=True)
        except Exception:
            pass

    def _clean_update_cache_older_than(self, days=7):
        try:
            if not os.path.isdir(UPDATE_CACHE_DIR):
                return
            now = time.time()
            for name in os.listdir(UPDATE_CACHE_DIR):
                p = os.path.join(UPDATE_CACHE_DIR, name)
                if os.path.isdir(p):
                    try:
                        mtime = os.path.getmtime(p)
                        if now - mtime > days * 86400:
                            shutil.rmtree(p, ignore_errors=True)
                    except Exception:
                        pass
        except Exception:
            pass

    # ========== 下载并安装 ==========
    def download_and_install_update(self, download_url, on_close_callback=None):
        if not download_url:
            messagebox.showerror("更新失败", "没有可用的下载地址。", parent=self.root)
            if on_close_callback is not None:
                try:
                    on_close_callback()
                except Exception:
                    pass
            return

        self._clean_update_cache_older_than(7)

        state = DownloadState()
        ts = time.strftime("%Y%m%d_%H%M%S")
        state.final_path = os.path.join(tempfile.gettempdir(), f"fwc_new_{ts}.exe")

        use_multi = bool(self.download_multithread)
        use_resume = bool(self.download_resume)
        multi_threads = 4

        cache_dir = None
        if use_resume:
            cache_dir = self._get_cache_dir_for_url(
                download_url, "multi" if use_multi else "single"
            )

        if use_resume and cache_dir and not use_multi:
            state.tmp_path = os.path.join(cache_dir, "installer.part")
        else:
            state.tmp_path = os.path.join(tempfile.gettempdir(), f"fwc_new_{ts}.exe.part")

        prog_win = tk.Toplevel(self.root)
        prog_win.title("正在下载更新")
        prog_win.geometry("440x330")
        prog_win.resizable(False, False)
        prog_win.attributes("-topmost", True)
        prog_win.transient(self.root)
        hide_toplevel_taskbar(prog_win)
        prog_win.grab_set()

        self._apply_dpi_to_toplevel(prog_win)

        main_frame = ttk.Frame(prog_win, padding=14)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="正在下载新版本…",
                 font=(self.font_family, 11, "bold")).pack(anchor="w")

        mode_str = "多线程" if use_multi else "单线程"
        if use_resume:
            mode_str += "（可续传）"
        tk.Label(main_frame, text=f"下载模式：{mode_str}",
                 font=(self.font_family, 8), fg="#888888", anchor="w").pack(fill=tk.X, pady=(2, 6))

        progress_var = tk.DoubleVar(value=0.0)
        pb = ttk.Progressbar(main_frame, variable=progress_var, maximum=100, length=400)
        pb.pack(fill=tk.X, pady=(4, 4))

        speed_pct_var = tk.StringVar(value="计算中… (0%)")
        tk.Label(main_frame, textvariable=speed_pct_var,
                 font=(self.font_family, 9), fg="#444444", anchor="w").pack(fill=tk.X)

        size_var = tk.StringVar(value="已下载：0 B / --")
        tk.Label(main_frame, textvariable=size_var,
                 font=(self.font_family, 9), fg="#444444", anchor="w").pack(fill=tk.X, pady=(2, 0))

        eta_var = tk.StringVar(value="剩余时间：--")
        tk.Label(main_frame, textvariable=eta_var,
                 font=(self.font_family, 9), fg="#444444", anchor="w").pack(fill=tk.X, pady=(2, 0))

        status_var = tk.StringVar(value="状态：下载中")
        status_lbl = tk.Label(main_frame, textvariable=status_var,
                              font=(self.font_family, 9, "bold"), fg="#2f80ed", anchor="w")
        status_lbl.pack(fill=tk.X, pady=(6, 0))

        ui_refs = {
            "prog_win": prog_win,
            "progress_var": progress_var,
            "speed_pct_var": speed_pct_var,
            "size_var": size_var,
            "eta_var": eta_var,
            "status_var": status_var,
            "status_lbl": status_lbl,
            "phase": {"value": "downloading"},
        }

        if use_multi:
            seg_frame = ttk.Frame(main_frame)
            seg_frame.pack(fill=tk.X, pady=(6, 0))
            seg_vars = []
            for i in range(multi_threads):
                v = tk.DoubleVar(value=0.0)
                b = ttk.Progressbar(seg_frame, variable=v, maximum=100, length=90)
                b.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
                seg_vars.append(v)
            ui_refs["seg_vars"] = seg_vars
            ui_refs["seg_count"] = multi_threads

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        btn_pause = ttk.Button(btn_frame, text="暂停")
        btn_cancel = ttk.Button(btn_frame, text="取消")
        btn_install = ttk.Button(btn_frame, text="立即安装")

        ui_refs["btn_pause"] = btn_pause
        ui_refs["btn_cancel"] = btn_cancel
        ui_refs["btn_install"] = btn_install
        ui_refs["btn_frame"] = btn_frame

        phase = ui_refs["phase"]
        closed_flag = {"value": False}

        def _show_download_buttons():
            btn_install.pack_forget()
            btn_pause.pack(side=tk.LEFT, padx=6)
            btn_cancel.pack(side=tk.LEFT, padx=6)

        _show_download_buttons()

        def close_window():
            if closed_flag["value"]:
                return
            closed_flag["value"] = True
            try:
                prog_win.grab_release()
            except Exception:
                pass
            try:
                prog_win.destroy()
            except Exception:
                pass
            if on_close_callback is not None:
                try:
                    on_close_callback()
                except Exception:
                    pass

        def cleanup_tmp():
            try:
                if state.tmp_path and os.path.exists(state.tmp_path):
                    os.remove(state.tmp_path)
            except Exception:
                pass

        def toggle_pause():
            if phase["value"] in ("finished", "installing"):
                return
            if state.pause_event.is_set():
                state.pause_event.clear()
                phase["value"] = "downloading"
                btn_pause.config(text="暂停")
                status_var.set("状态：下载中")
                status_lbl.config(fg="#2f80ed")
                ui_refs["resume_reset"] = True
            else:
                state.pause_event.set()
                phase["value"] = "paused"
                btn_pause.config(text="继续")
                status_var.set("状态：已暂停")
                status_lbl.config(fg="#e08a00")
                try:
                    speed_pct_var.set("-- (已暂停)")
                    eta_var.set("剩余时间：--")
                except Exception:
                    pass

        btn_pause.config(command=toggle_pause)

        def do_cancel():
            if phase["value"] == "installing":
                return
            if phase["value"] == "finished":
                if not messagebox.askyesno("取消更新", "已下载完成，确定要丢弃本次更新吗？",
                                            parent=prog_win):
                    return
                state.request_cancel()
                cleanup_tmp()
                close_window()
                self.log("[下载] 用户取消安装，已删除下载文件")
                return
            if use_resume:
                keep = messagebox.askyesno(
                    "取消下载",
                    "下载未完成。是否保留已下载的部分，以便下次继续？\n\n"
                    "选「是」：保留缓存（可续传）\n"
                    "选「否」：删除已下载的部分",
                    parent=prog_win,
                )
                if keep:
                    state.request_cancel()
                    close_window()
                    self.log("[续传] 用户取消并保留缓存")
                    return
            state.request_cancel()
            cleanup_tmp()
            close_window()
            self.log("[下载] 用户取消并删除临时文件")

        btn_cancel.config(command=do_cancel)

        def do_install():
            if phase["value"] != "finished":
                self.log(f"[安装] 忽略点击：当前阶段={phase['value']}")
                return
            self.log("[安装] 用户点击「立即安装」，开始替换")
            phase["value"] = "installing"
            state.pause_event.clear()
            state.cancel_event.clear()
            try:
                btn_pause.pack_forget()
                btn_install.pack_forget()
                btn_cancel.pack_forget()
            except Exception:
                pass
            status_var.set("状态：正在安装…")
            status_lbl.config(fg="#2f80ed")
            close_window()
            self.root.after(100, lambda: self._do_replace_and_restart(state.final_path, None))

        btn_install.config(command=do_install)

        resume_offset = 0
        multi_resume_info = None

        if use_resume and cache_dir:
            if use_multi:
                meta = self._read_cache_meta(cache_dir)
                part_path = os.path.join(cache_dir, "installer.part")
                chunks_path = os.path.join(cache_dir, "chunks.json")

                if meta and os.path.exists(part_path) and os.path.exists(chunks_path):
                    try:
                        remote_size = self._probe_content_length(download_url)
                        url_ok = (meta.get("url") == download_url)
                        saved_total = meta.get("total_size", 0)
                        saved_n = meta.get("n_threads", 0)

                        if (url_ok and saved_total > 0 and saved_total == remote_size
                                and saved_n == multi_threads):
                            try:
                                with open(chunks_path, "r", encoding="utf-8") as f:
                                    cj = json.load(f)
                                chunks_done = cj.get("chunks", [0] * multi_threads)
                                if len(chunks_done) == multi_threads:
                                    multi_resume_info = {
                                        "chunks": chunks_done,
                                        "total": saved_total,
                                        "part_path": part_path,
                                    }
                                    self.log(f"[续传] 多线程缓存命中，各块进度：{chunks_done}")
                            except Exception as e:
                                self.log(f"[续传] 读 chunks.json 失败：{e}")
                        else:
                            self.log("[续传] 多线程缓存不匹配（URL/大小/线程数变了），清理重下")
                            self._clear_cache_dir(cache_dir)
                    except Exception as e:
                        self.log(f"[续传] 检查多线程缓存异常：{e}")
                else:
                    self.log("[续传] 多线程无缓存记录，将重新下载")
            else:
                meta = self._read_cache_meta(cache_dir)
                part_path = os.path.join(cache_dir, "installer.part")
                if meta and os.path.exists(part_path):
                    try:
                        url_ok = (meta.get("url") == download_url)
                        remote_size = self._probe_content_length(download_url)
                        local_size = os.path.getsize(part_path)
                        saved_total = meta.get("total_size", 0)

                        self.log(f"[续传] 单线程缓存目录 {cache_dir}")
                        self.log(f"[续传] URL匹配={url_ok}, 本地part={local_size}B, "
                                 f"记录总大小={saved_total}, 远程总大小={remote_size}")

                        if (url_ok and remote_size > 0
                                and remote_size == saved_total
                                and 0 < local_size < remote_size):
                            resume_offset = local_size
                            state.tmp_path = part_path
                            self.log(f"[续传] 单线程命中，从 {resume_offset} 字节继续")
                        else:
                            self.log("[续传] 单线程未命中")
                            if not url_ok:
                                self.log("[续传]   · URL 不匹配")
                            if remote_size <= 0:
                                self.log("[续传]   · 远程大小获取失败")
                            elif remote_size != saved_total:
                                self.log("[续传]   · 远程大小已变化，清理旧缓存")
                                self._clear_cache_dir(cache_dir)
                            if local_size <= 0 or local_size >= (remote_size or 0):
                                self.log("[续传]   · 本地 part 大小异常")
                    except Exception as e:
                        self.log(f"[续传] 检查单线程缓存异常：{e}")
                else:
                    self.log("[续传] 单线程无缓存记录，将重新下载")

        def worker():
            try:
                if use_multi:
                    self._download_multi(download_url, state, ui_refs, closed_flag,
                                          multi_threads, cache_dir, use_resume,
                                          multi_resume_info)
                else:
                    self._download_single(download_url, state, ui_refs, closed_flag,
                                           resume_offset, cache_dir, use_resume)
            except Exception as e:
                if state.cancel_event.is_set() or closed_flag["value"]:
                    return
                self.log(f"[下载] 下载更新失败: {e}")
                if use_resume and cache_dir:
                    self.log("[续传] 异常时保留缓存，下次可继续")
                else:
                    try:
                        self.root.after(0, cleanup_tmp)
                    except Exception:
                        pass
                try:
                    self.root.after(0, lambda err=e: self._on_update_download_failed(err, prog_win))
                except Exception:
                    pass

        def on_window_close():
            if phase["value"] == "installing":
                return
            do_cancel()

        prog_win.protocol("WM_DELETE_WINDOW", on_window_close)

        threading.Thread(target=worker, daemon=True).start()

    def _probe_content_length(self, url):
        try:
            ssl_ctx = ssl.create_default_context()
            ssl_ctx.check_hostname = False
            ssl_ctx.verify_mode = ssl.CERT_NONE
            req = urllib.request.Request(
                url, method="HEAD",
                headers={"User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}"}
            )
            with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                length = int(resp.headers.get("Content-Length", 0) or 0)
                if length == 0:
                    try:
                        req2 = urllib.request.Request(
                            url,
                            headers={
                                "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                                "Range": "bytes=0-0",
                            }
                        )
                        with urllib.request.urlopen(req2, timeout=15, context=ssl_ctx) as resp2:
                            cr = resp2.headers.get("Content-Range", "")
                            if "/" in cr:
                                return int(cr.split("/")[-1])
                    except Exception:
                        pass
                return length
        except Exception:
            return 0

    def _download_single(self, url, state, ui_refs, closed_flag,
                          resume_offset, cache_dir, use_resume):
        ssl_ctx = ssl.create_default_context()
        ssl_ctx.check_hostname = False
        ssl_ctx.verify_mode = ssl.CERT_NONE

        headers = {"User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}"}
        mode = "wb"
        if resume_offset > 0:
            headers["Range"] = f"bytes={resume_offset}-"
            mode = "ab"

        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=60, context=ssl_ctx) as resp:
            status = getattr(resp, "status", 200)
            if resume_offset > 0 and status != 206:
                self.log("[续传] 服务端不支持 Range，重新下载")
                resume_offset = 0
                mode = "wb"
                if state.tmp_path and os.path.exists(state.tmp_path):
                    try:
                        os.remove(state.tmp_path)
                    except Exception:
                        pass
                req = urllib.request.Request(
                    url,
                    headers={"User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}"}
                )
                resp = urllib.request.urlopen(req, timeout=60, context=ssl_ctx)

            total_in_resp = int(resp.headers.get("Content-Length", 0) or 0)
            total = total_in_resp + resume_offset
            state.total_size = total
            state.downloaded = resume_offset

            if use_resume and cache_dir:
                self._write_cache_meta(cache_dir, {
                    "url": url,
                    "total_size": total,
                    "ts": time.time(),
                })

            last_time = time.time()
            last_downloaded = resume_offset
            speed = 0.0
            last_ui = 0.0

            with open(state.tmp_path, mode) as f:
                while True:
                    if state.cancel_event.is_set() or closed_flag["value"]:
                        self.log("[下载] 收到取消信号，停止下载")
                        return
                    was_paused = False
                    while state.pause_event.is_set():
                        was_paused = True
                        if state.cancel_event.is_set() or closed_flag["value"]:
                            return
                        time.sleep(0.1)
                    if was_paused or ui_refs.get("resume_reset"):
                        last_time = time.time()
                        last_downloaded = state.downloaded
                        speed = 0.0
                        ui_refs["resume_reset"] = False

                    chunk = resp.read(64 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    state.downloaded += len(chunk)

                    now = time.time()
                    if now - last_ui >= 0.5 or state.downloaded == total:
                        dt = now - last_time
                        if dt > 0:
                            inst = (state.downloaded - last_downloaded) / dt
                            speed = inst if speed <= 0 else speed * 0.6 + inst * 0.4
                            last_time = now
                            last_downloaded = state.downloaded
                        last_ui = now

                        pct = (state.downloaded * 100.0 / total) if total > 0 else 0.0
                        if pct > 100:
                            pct = 100
                        spd_str = f"{_fmt_size(int(speed))}/s" if speed > 0 else "计算中…"
                        self._push_progress(
                            ui_refs, pct, spd_str, state.downloaded, total, speed
                        )

            if state.cancel_event.is_set() or closed_flag["value"]:
                return
            try:
                os.replace(state.tmp_path, state.final_path)
            except Exception as e:
                self.log(f"[下载] 重命名失败: {e}")
                self.root.after(0, lambda err=e: self._on_update_download_failed(err, ui_refs["prog_win"]))
                return
            if cache_dir:
                self._clear_cache_dir(cache_dir)

            self._mark_finished(ui_refs, state)
            self.log(f"[下载] 更新下载完成：{state.final_path}")

    def _download_multi(self, url, state, ui_refs, closed_flag, n_threads,
                         cache_dir, use_resume, resume_info):
        total = self._probe_content_length(url)
        if total <= 0:
            self.log("[下载] 多线程：无法获取文件大小，回退单线程")
            self._download_single(url, state, ui_refs, closed_flag, 0, None, False)
            return

        block_size = total // n_threads
        ranges = []
        for i in range(n_threads):
            start = i * block_size
            end = total - 1 if i == n_threads - 1 else (start + block_size - 1)
            ranges.append((start, end))

        part_path = state.tmp_path
        chunks_done = [0] * n_threads

        if use_resume and cache_dir:
            part_path = os.path.join(cache_dir, "installer.part")
            state.tmp_path = part_path

            if resume_info is not None:
                chunks_done = list(resume_info.get("chunks", [0] * n_threads))
                if len(chunks_done) != n_threads:
                    chunks_done = [0] * n_threads
                self.log(f"[续传] 多线程续传，各块已下：{chunks_done}")

            self._write_cache_meta(cache_dir, {
                "url": url,
                "total_size": total,
                "n_threads": n_threads,
                "block_size": block_size,
                "ts": time.time(),
            })
            self._write_chunks(cache_dir, chunks_done)

        need_create = True
        if os.path.exists(part_path):
            try:
                if os.path.getsize(part_path) == total:
                    need_create = False
            except Exception:
                pass
        if need_create:
            with open(part_path, "wb") as f:
                f.truncate(total)

        state.total_size = total
        state.downloaded = sum(chunks_done)

        progress_lock = threading.Lock()
        state_ref = {"last_time": time.time(), "last_total": state.downloaded,
                     "speed": 0.0, "last_ui": 0.0}

        def worker_seg(idx, start, end):
            seg_size = end - start + 1
            already = chunks_done[idx]
            if already >= seg_size:
                self.log(f"[续传] 分块 {idx} 已完成，跳过")
                return

            req_start = start + already
            req_end = end

            try:
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(
                    url,
                    headers={
                        "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                        "Range": f"bytes={req_start}-{req_end}",
                    },
                )
                with urllib.request.urlopen(req, timeout=60, context=ssl_ctx) as resp:
                    with open(part_path, "r+b") as f:
                        f.seek(req_start)
                        pos = req_start
                        while pos <= end:
                            if state.cancel_event.is_set() or closed_flag["value"]:
                                return
                            was_paused = False
                            while state.pause_event.is_set():
                                was_paused = True
                                if state.cancel_event.is_set() or closed_flag["value"]:
                                    return
                                time.sleep(0.1)
                            if was_paused or ui_refs.get("resume_reset"):
                                state_ref["last_time"] = time.time()
                                state_ref["last_total"] = state.downloaded
                                state_ref["speed"] = 0.0
                                ui_refs["resume_reset"] = False

                            chunk = resp.read(64 * 1024)
                            if not chunk:
                                break
                            if pos + len(chunk) - 1 > end:
                                chunk = chunk[:end - pos + 1]
                            f.write(chunk)
                            pos += len(chunk)
                            with progress_lock:
                                chunks_done[idx] = pos - start
                                state.downloaded += len(chunk)
                                if use_resume and cache_dir:
                                    self._write_chunks(cache_dir, list(chunks_done))
                            now = time.time()
                            if now - state_ref["last_ui"] >= 0.5:
                                with progress_lock:
                                    cur_total = state.downloaded
                                    cur_chunks = list(chunks_done)
                                dt = now - state_ref["last_time"]
                                if dt > 0:
                                    inst = (cur_total - state_ref["last_total"]) / dt
                                    if state_ref["speed"] <= 0:
                                        state_ref["speed"] = inst
                                    else:
                                        state_ref["speed"] = state_ref["speed"] * 0.6 + inst * 0.4
                                    state_ref["last_time"] = now
                                    state_ref["last_total"] = cur_total
                                state_ref["last_ui"] = now
                                pct = cur_total * 100.0 / total
                                if pct > 100:
                                    pct = 100
                                spd_str = (f"{_fmt_size(int(state_ref['speed']))}/s"
                                           if state_ref["speed"] > 0 else "计算中…")
                                self._push_progress(
                                    ui_refs, pct, spd_str, cur_total, total, state_ref["speed"]
                                )
                                self._push_segments(ui_refs, cur_chunks, ranges)
            except Exception as e:
                self.log(f"[下载] 分块 {idx} 下载失败: {e}")

        threads = []
        for i, (s, e) in enumerate(ranges):
            t = threading.Thread(target=worker_seg, args=(i, s, e), daemon=True)
            t.start()
            threads.append(t)
        for t in threads:
            t.join()

        if state.cancel_event.is_set() or closed_flag["value"]:
            if use_resume and cache_dir:
                with progress_lock:
                    self._write_chunks(cache_dir, list(chunks_done))
                self.log(f"[续传] 多线程取消，已保存进度：{chunks_done}")
            return

        all_done = True
        for i, (s, e) in enumerate(ranges):
            if chunks_done[i] < (e - s + 1):
                all_done = False
                break
        if not all_done:
            self.log(f"[下载] 多线程未全部完成，各块：{chunks_done}，保留缓存")
            if use_resume and cache_dir:
                with progress_lock:
                    self._write_chunks(cache_dir, list(chunks_done))
            self.root.after(0, lambda err="部分分块未完成": self._on_update_download_failed(
                err, ui_refs["prog_win"]))
            return

        try:
            if state.tmp_path != state.final_path:
                os.replace(state.tmp_path, state.final_path)
        except Exception as e:
            self.log(f"[下载] 重命名失败: {e}")
            self.root.after(0, lambda err=e: self._on_update_download_failed(
                err, ui_refs["prog_win"]))
            return

        if cache_dir:
            self._clear_cache_dir(cache_dir)

        self._mark_finished(ui_refs, state)
        self.log(f"[下载] 更新下载完成（多线程）：{state.final_path}")

    def _push_segments(self, ui_refs, chunks_done, ranges):
        seg_vars = ui_refs.get("seg_vars")
        if not seg_vars:
            return
        n = ui_refs.get("seg_count", len(seg_vars))
        vals = []
        for i in range(min(n, len(chunks_done), len(ranges))):
            s, e = ranges[i]
            seg_size = e - s + 1
            if seg_size <= 0:
                vals.append(0.0)
            else:
                vals.append(min(100.0, max(0.0, chunks_done[i] * 100.0 / seg_size)))

        def _ui():
            try:
                if not ui_refs["prog_win"].winfo_exists():
                    return
                for i, v in enumerate(vals):
                    if i < len(seg_vars):
                        seg_vars[i].set(v)
            except Exception:
                pass

        try:
            self.root.after(0, _ui)
        except Exception:
            pass

    def _push_progress(self, ui_refs, pct, spd_str, downloaded, total, speed):
        if total > 0:
            size_str = f"{_fmt_size(downloaded)} / {_fmt_size(total)}"
            if speed > 0:
                remain = max(0.0, (total - downloaded) / speed)
                eta_str = "约 " + _fmt_duration(remain)
            else:
                eta_str = "计算中…"
        else:
            size_str = f"{_fmt_size(downloaded)} / 未知"
            eta_str = "计算中…"

        spd_pct = f"{spd_str} ({int(pct)}%)"

        def _ui():
            try:
                if not ui_refs["prog_win"].winfo_exists():
                    return
                ui_refs["progress_var"].set(pct)
                ui_refs["speed_pct_var"].set(spd_pct)
                ui_refs["size_var"].set(f"已下载：{size_str}")
                ui_refs["eta_var"].set(f"剩余时间：{eta_str}")
            except Exception:
                pass

        try:
            self.root.after(0, _ui)
        except Exception:
            pass

    def _mark_finished(self, ui_refs, state):
        def _ui():
            try:
                if not ui_refs["prog_win"].winfo_exists():
                    return
                ui_refs["progress_var"].set(100)
                ui_refs["speed_pct_var"].set("-- (100%)")
                ui_refs["size_var"].set(
                    f"已下载：{_fmt_size(state.downloaded)} / "
                    f"{_fmt_size(state.total_size or state.downloaded)}"
                )
                ui_refs["eta_var"].set("剩余时间：--")
                ui_refs["status_var"].set("状态：下载完成，等待确认安装")
                ui_refs["status_lbl"].config(fg="#2f80ed")
                try:
                    ui_refs["btn_pause"].pack_forget()
                except Exception:
                    pass
                try:
                    ui_refs["btn_install"].pack_forget()
                except Exception:
                    pass
                try:
                    ui_refs["btn_cancel"].pack_forget()
                except Exception:
                    pass
                try:
                    ui_refs["btn_install"].pack(side=tk.LEFT, padx=6)
                except Exception:
                    pass
                try:
                    ui_refs["btn_cancel"].pack(side=tk.LEFT, padx=6)
                except Exception:
                    pass
                ui_refs["phase"]["value"] = "finished"
                self.log("[下载] 状态切换：finished，[立即安装] 已可用")
            except Exception as e:
                self.log(f"[下载] 切换完成状态异常：{e}")

        try:
            self.root.after(0, _ui)
        except Exception:
            pass

    def _on_update_download_failed(self, err, prog_win):
        try:
            if prog_win is not None:
                prog_win.grab_release()
        except Exception:
            pass
        try:
            if prog_win is not None:
                prog_win.destroy()
        except Exception:
            pass
        messagebox.showerror(
            "更新失败",
            f"下载或安装过程中出错：\n\n{err}\n\n"
            "你可以点击「前往下载页」手动下载新版。",
            parent=self.root,
        )

    def _do_replace_and_restart(self, new_exe_path, prog_win):
        try:
            self.log("开始替换 exe 并重启…")
            current_exe = _get_current_exe_path()
            backup_exe = current_exe + ".old"

            if os.path.exists(backup_exe):
                try:
                    os.remove(backup_exe)
                except Exception:
                    pass

            os.rename(current_exe, backup_exe)
            shutil.move(new_exe_path, current_exe)

            time.sleep(0.5)

            clean_env = os.environ.copy()
            for key in list(clean_env.keys()):
                if key.startswith("_MEI") or key.startswith("_PYI"):
                    del clean_env[key]
            clean_env.pop("TCL_LIBRARY", None)
            clean_env.pop("TK_LIBRARY", None)

            subprocess.Popen(
                [current_exe],
                close_fds=True,
                env=clean_env,
                creationflags=0x00000008 | 0x08000000,
            )

            try:
                if prog_win is not None:
                    prog_win.grab_release()
            except Exception:
                pass
            try:
                if prog_win is not None:
                    prog_win.destroy()
            except Exception:
                pass

            self.quit_app()

        except Exception as e:
            try:
                if os.path.exists(backup_exe) and not os.path.exists(current_exe):
                    os.rename(backup_exe, current_exe)
            except Exception:
                pass
            self.log(f"替换更新失败: {e}")
            try:
                if prog_win is not None:
                    prog_win.grab_release()
            except Exception:
                pass
            try:
                if prog_win is not None:
                    prog_win.destroy()
            except Exception:
                pass
            messagebox.showerror(
                "更新失败",
                f"替换文件时出错：\n\n{e}\n\n"
                "你可以点击「前往下载页」手动下载新版。",
                parent=self.root,
            )

    # ================= 原有逻辑 =================
    def _hide_taskbar_timer(self):
        if self.is_quitting:
            return
        remove_taskbar_entry(self.root.winfo_id())
        self.root.after(2000, self._hide_taskbar_timer)

    def auto_save_pos_timer(self):
        if self.is_quitting:
            return
        self.save_config()
        self.root.after(2000, self.auto_save_pos_timer)

    def on_window_configure(self, event):
        if event.widget == self.root:
            self.mem_x = self.root.winfo_x()
            self.mem_y = self.root.winfo_y()
        if self.resize_throttle_id is not None:
            self.root.after_cancel(self.resize_throttle_id)
        self.resize_throttle_id = self.root.after(120, self._real_resize)

    def show_about(self):
        def _creator():
            win = tk.Toplevel(self.root)
            win.title("关于本软件")
            win.geometry("560x480")
            win.minsize(520, 440)
            win.attributes("-topmost", True)
            win.transient(self.root)
            hide_toplevel_taskbar(win)
            win.grab_set()

            self._apply_dpi_to_toplevel(win)

            main_frame = ttk.Frame(win, padding=16)
            main_frame.pack(fill=tk.BOTH, expand=True)

            header = ttk.Frame(main_frame)
            header.pack(fill=tk.X, pady=(0, 8))
            tk.Label(header, text=f"悬浮时钟天气  {self.get_display_version()}",
                     font=(self.font_family, 13, "bold"), fg="#2f80ed").pack(anchor="w")
            tk.Label(header, text=f"仓库：{GITHUB_REPO}",
                     font=(self.font_family, 8), fg="#888888").pack(anchor="w", pady=(2, 0))

            tk.Label(main_frame, text="版本介绍：", font=(self.font_family, 10, "bold")).pack(anchor="w")
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill=tk.X, expand=False, pady=(4, 10))

            txt = tk.Text(text_frame, wrap="word", font=(self.font_family, 9),
                          height=10, width=1,
                          relief="solid", borderwidth=1, padx=8, pady=6)
            txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            txt.config(yscrollcommand=sb.set, state="normal")

            loading_hint = "（正在从 GitHub 获取最新版本介绍…）"
            credit_line = "开发者：BigSteve3838292 和 Deepseek"
            txt.insert("1.0", f"{loading_hint}\n{credit_line}")
            txt.config(state="disabled")

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

            def open_bilibili():
                webbrowser.open("https://space.bilibili.com/3546580546226675")

            def open_github():
                webbrowser.open(f"https://github.com/{GITHUB_REPO}")

            ttk.Button(btn_frame, text="访问我的B站主页", command=open_bilibili).pack(side=tk.LEFT, padx=4)
            ttk.Button(btn_frame, text="打开我的GitHub仓库", command=open_github).pack(side=tk.LEFT, padx=4)
            ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.RIGHT, padx=4)

            def fetch_release_body():
                if self.debug_offline:
                    self.root.after(0, lambda: self._update_about_text(
                        txt, f"调试模式：已关闭软件联网\n\n{credit_line}"))
                    return
                try:
                    ssl_ctx = ssl.create_default_context()
                    ssl_ctx.check_hostname = False
                    ssl_ctx.verify_mode = ssl.CERT_NONE

                    if self.beta_edition:
                        req = urllib.request.Request(
                            GITHUB_API_RELEASES,
                            headers={
                                "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                                "Accept": "application/vnd.github+json",
                            },
                        )
                        with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                            releases = json.loads(resp.read().decode("utf-8"))
                        latest_beta = None
                        for r in releases:
                            if r.get("prerelease"):
                                latest_beta = r
                                break
                        if latest_beta is None:
                            new_text = f"暂无 Beta 版发布信息\n\n{credit_line}"
                            self.root.after(0, lambda: self._update_about_text(txt, new_text))
                            return
                        tag = latest_beta.get("tag_name") or latest_beta.get("name") or ""
                        body = latest_beta.get("body") or ""
                        html_url = latest_beta.get("html_url") or GITHUB_RELEASES_PAGE
                    else:
                        req = urllib.request.Request(
                            GITHUB_API_LATEST,
                            headers={
                                "User-Agent": f"FloatingWeatherClock/{self.get_compare_version()}",
                                "Accept": "application/vnd.github+json",
                            },
                        )
                        with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                            data = json.loads(resp.read().decode("utf-8"))
                        tag = data.get("tag_name") or data.get("name") or ""
                        body = data.get("body") or ""
                        html_url = data.get("html_url") or GITHUB_RELEASES_PAGE

                    if body.strip():
                        new_text = (
                            f"最新发布版本：{tag}\n"
                            f"发布页：{html_url}\n"
                            f"{'─' * 40}\n\n"
                            f"{body.strip()}\n\n"
                            f"{'─' * 40}\n"
                            f"{credit_line}\n"
                        )
                    else:
                        new_text = (
                            f"最新发布版本：{tag}（暂无版本说明）\n"
                            f"发布页：{html_url}\n\n"
                            f"{credit_line}"
                        )
                    self.root.after(0, lambda: self._update_about_text(txt, new_text))
                except Exception as e:
                    self.log(f"关于页获取最新介绍失败: {e}")
                    self.root.after(0, lambda err=e: self._update_about_text(
                        txt, f"（获取最新介绍失败：{err}）\n\n{credit_line}"))

            threading.Thread(target=fetch_release_body, daemon=True).start()
            return win

        win = self._open_singleton("about", _creator)
        if win is not None:
            self.root.wait_window(win)

    def _update_about_text(self, txt_widget, new_text):
        try:
            if not txt_widget.winfo_exists():
                return
            txt_widget.config(state="normal")
            txt_widget.delete("1.0", tk.END)
            txt_widget.insert("1.0", new_text)
            txt_widget.config(state="disabled")
        except Exception:
            pass

    # ================= 开发者选项 =================
    def open_dev_options(self):
        if self._dev_options_opening:
            return
        if not BETA_EDITION and not self.developer_mode:
            self._dev_options_opening = True
            try:
                pwd = simpledialog.askstring("开发者验证", "请输入开发者密码：", show="*", parent=self.root)
                if pwd is None:
                    return
                if pwd != "zhusihao20120824":
                    messagebox.showerror("密码错误", "开发者密码不正确！")
                    return
                self.developer_mode = True
                self.save_config()
            finally:
                self._dev_options_opening = False

        def _creator():
            win, main_frame, fixed_btn_frame, on_close = make_scrollable_toplevel(
                self.root,
                "悬浮时钟天气 - 开发者设置",
                geometry="620x900",
                minsize=(580, 820),
                dpi_scale=self.dpi_scale
            )

            group_dev = ttk.LabelFrame(main_frame, text="开发者功能", padding=10)
            group_dev.pack(fill=tk.X, pady=6)

            dev_mode_var = tk.BooleanVar(value=self.developer_mode)
            ttk.Checkbutton(group_dev,
                            text="启用开发者模式（关闭后下次进入需要密码）",
                            variable=dev_mode_var).pack(anchor="w")

            frame_btn_row1 = ttk.Frame(group_dev)
            frame_btn_row1.pack(pady=6, fill=tk.X)
            ttk.Button(frame_btn_row1, text="查看信息",
                       command=self.show_debug_info).pack(side=tk.LEFT, padx=4)
            ttk.Button(frame_btn_row1, text="立即检查更新",
                       command=self.menu_check_update).pack(side=tk.LEFT, padx=4)

            frame_btn_row2 = ttk.Frame(group_dev)
            frame_btn_row2.pack(pady=6, fill=tk.X)
            ttk.Button(frame_btn_row2, text="打开实时日志终端",
                       command=self.debug_open_terminal_log).pack(side=tk.LEFT, padx=4)

            tk.Label(group_dev,
                     text=f"实时日志文件：{RUNTIME_LOG_FILE}",
                     fg="#888888", font=(self.font_family, 8), justify="left"
                     ).pack(anchor="w", padx=(4, 0), pady=(2, 0))

            group_beta = ttk.LabelFrame(main_frame, text="模拟 Beta 版（仅内存，不保存）", padding=10)
            group_beta.pack(fill=tk.X, pady=6)

            sim_beta_var = tk.BooleanVar(value=self.beta_edition)
            ttk.Checkbutton(group_beta, text="模拟为 Beta 版",
                            variable=sim_beta_var).pack(anchor="w")

            frame_beta_num = ttk.Frame(group_beta)
            frame_beta_num.pack(fill=tk.X, pady=(4, 0))
            tk.Label(frame_beta_num, text="Beta 编号：").pack(side=tk.LEFT)
            beta_num_var = tk.StringVar(value=str(self.beta_version))
            ttk.Entry(frame_beta_num, textvariable=beta_num_var, width=8).pack(side=tk.LEFT, padx=4)

            tk.Label(group_beta,
                     text="勾选后软件立即表现为 Beta 版，重启后恢复真实状态。\n"
                          "此操作仅写入内存，不保存到配置文件。",
                     fg="#888888", font=(self.font_family, 8), justify="left"
                     ).pack(anchor="w", pady=(4, 0))

            def apply_sim_beta():
                try:
                    n = int(beta_num_var.get().strip() or "1")
                    if n < 1:
                        n = 1
                except Exception:
                    n = 1
                new_edition = bool(sim_beta_var.get())
                self.beta_edition = new_edition
                self.beta_version = n
                try:
                    self.refresh_beta_ui()
                except Exception as e:
                    self.log(f"应用模拟 Beta 失败：{e}")
                self.log(f"[模拟Beta] beta_edition={self.beta_edition}, beta_version={self.beta_version}")

            ttk.Button(group_beta, text="立即应用（不保存）",
                       command=apply_sim_beta).pack(anchor="w", pady=(6, 0))

            group_sim_date = ttk.LabelFrame(main_frame, text="模拟日期（独立功能，不依赖测试天气）", padding=10)
            group_sim_date.pack(fill=tk.X, pady=6)

            tk.Label(group_sim_date,
                     text="格式 YYYY-MM-DD，留空表示恢复真实日期。\n"
                          "影响：农历/节气/节假日显示 + 昼夜判定。",
                     fg="#888888", font=(self.font_family, 8), justify="left").pack(anchor="w")

            sim_date_var = tk.StringVar(value=self.sim_date or "")
            frame_sd = ttk.Frame(group_sim_date)
            frame_sd.pack(fill=tk.X, pady=(4, 0))
            ttk.Entry(frame_sd, textvariable=sim_date_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

            def apply_sim_date():
                sd = sim_date_var.get().strip()
                if sd:
                    try:
                        date_cls(*map(int, sd.split("-")))
                    except Exception:
                        messagebox.showerror("错误", "日期格式不对，应为 YYYY-MM-DD", parent=win)
                        return
                    self.sim_date = sd
                else:
                    self.sim_date = None
                self.date_cache = None
                self.holiday_cache = None
                self.refresh_start_time = time.time()
                self._fetch_all_dates_async()
                self.need_redraw = True
                self.log(f"[模拟日期] 已设置为：{self.sim_date or '（真实日期）'}")

            ttk.Button(frame_sd, text="应用", command=apply_sim_date).pack(side=tk.LEFT, padx=4)

            group_dbg = ttk.LabelFrame(main_frame, text="调试选项", padding=10)
            group_dbg.pack(fill=tk.X, pady=6)

            dbg_offline_var = tk.BooleanVar(value=self.debug_offline)
            ttk.Checkbutton(group_dbg, text="关闭软件联网（不请求天气/定位/更新）",
                            variable=dbg_offline_var).pack(anchor="w")

            ttk.Button(group_dbg, text="将自动刷新倒计时快进至最后 5 秒",
                       command=self.debug_skip_countdown).pack(anchor="w", pady=6)

            ttk.Separator(group_dbg, orient="horizontal").pack(fill=tk.X, pady=6)

            tk.Label(group_dbg, text="模拟更新（验证自动更新功能，不联网）",
                     font=(self.font_family, 9, "bold")).pack(anchor="w")

            frame_sim = ttk.Frame(group_dbg)
            frame_sim.pack(fill=tk.X, pady=(6, 0))

            ttk.Button(frame_sim, text="模拟替换重启",
                       command=self.debug_simulate_update).pack(side=tk.LEFT, padx=4)
            ttk.Button(frame_sim, text="模拟完整更新流程",
                       command=self.debug_simulate_full_update).pack(side=tk.LEFT, padx=4)

            tk.Label(group_dbg,
                     text="· 模拟替换重启：跳过下载，直接复制当前 exe 替换后重启。\n"
                          "· 模拟完整更新流程：本地起 HTTP 服务挂当前 exe，\n"
                          "  限速到约 5 秒完成，便于测试多线程与断点续传。\n"
                          "· 模拟源文件固定为 fwc_sim_new.exe，便于验证断点续传。",
                     fg="#888888", font=(self.font_family, 8), justify="left"
                     ).pack(anchor="w", padx=(4, 0), pady=(6, 0))

            group_test = ttk.LabelFrame(main_frame, text="测试天气模式", padding=10)
            group_test.pack(fill=tk.X, pady=6)

            test_var = tk.BooleanVar(value=self.test_weather_enabled)
            ttk.Checkbutton(group_test,
                            text="启用测试天气（关闭后立即恢复默认，继续接收真实天气）",
                            variable=test_var).pack(anchor="w")

            tk.Label(group_test, text="说明：启用后所有联网天气请求暂停，显示你手动填写的值。",
                     fg="#888888", font=(self.font_family, 8)).pack(anchor="w", pady=(2, 6))

            frame_test_btn = ttk.Frame(group_test)
            frame_test_btn.pack(fill=tk.X)
            ttk.Button(frame_test_btn, text="打开测试天气面板",
                       command=self.open_test_weather_panel).pack(side=tk.LEFT, padx=4)

            def save_dev_setting():
                old_test = self.test_weather_enabled
                new_test = test_var.get()

                self.developer_mode = dev_mode_var.get()
                self.debug_offline = dbg_offline_var.get()

                if new_test != old_test:
                    if new_test:
                        self.test_weather_enabled = True
                        if self.test_weather_data is None:
                            self.test_weather_data = self._make_default_test_weather()
                    else:
                        self.test_weather_enabled = False
                        self.test_weather_data = None
                        self.test_weather_mode = "auto"
                        self.weather_data = None
                        self.is_refresh_hint = True
                        self.reset_refresh_countdown()
                        self.fetch_all_async()

                self.apply_window_attr()
                self.save_config()
                self.need_redraw = True
                self.update_clock()
                on_close()

            ttk.Button(fixed_btn_frame, text="保存设置", command=save_dev_setting).pack(side=tk.LEFT, padx=6)
            ttk.Button(fixed_btn_frame, text="取消", command=on_close).pack(side=tk.LEFT, padx=6)

            return win

        self._open_singleton("dev_options", _creator)

    def _make_default_test_weather(self):
        return {
            "tag": "clear",
            "city": self.city,
            "temp": "25",
            "desc": "测试晴朗",
            "humidity": "50",
            "wind": "10",
            "wind_dir": "0",
            "precipitation": "0",
            "precip_prob": "0",
            "sunrise": "06:00",
            "sunset": "18:00",
        }

    def debug_simulate_update(self):
        current_exe = _get_current_exe_path()
        if not os.path.exists(current_exe):
            messagebox.showerror("模拟更新失败",
                                 f"找不到当前 exe：\n{current_exe}",
                                 parent=self.root)
            return
        if not messagebox.askyesno(
            "模拟替换重启",
            "将进行以下操作：\n\n"
            "  1. 复制当前 exe 到临时目录\n"
            "  2. 直接替换、重启（不经过下载 UI）\n"
            "  3. 当前进程将退出，新进程启动\n\n"
            "继续吗？",
            parent=self.root,
        ):
            return
        try:
            fd, tmp_path = tempfile.mkstemp(suffix=".exe", prefix="fwc_sim_")
            os.close(fd)
            shutil.copy2(current_exe, tmp_path)
            self.log(f"模拟替换重启：已复制到 {tmp_path}")
        except Exception as e:
            self.log(f"模拟替换重启：复制失败 {e}")
            messagebox.showerror("模拟更新失败", f"复制 exe 时出错：\n\n{e}", parent=self.root)
            return
        try:
            marker_path = self._get_simulated_update_marker_path()
            with open(marker_path, "w", encoding="utf-8") as f:
                json.dump({
                    "old_version": APP_VERSION,
                    "simulated_version": APP_VERSION,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"模拟替换重启：写标记失败 {e}")

        prog_win = tk.Toplevel(self.root)
        prog_win.title("模拟替换重启")
        prog_win.geometry("380x170")
        prog_win.resizable(False, False)
        prog_win.attributes("-topmost", True)
        prog_win.transient(self.root)
        hide_toplevel_taskbar(prog_win)
        prog_win.grab_set()

        self._apply_dpi_to_toplevel(prog_win)

        tk.Label(prog_win, text="正在模拟替换重启…", font=(self.font_family, 10)).pack(pady=(16, 4))
        tk.Label(prog_win, text=f"模拟版本号：{self.get_display_version()}",
                 font=(self.font_family, 9), fg="#2f80ed").pack()
        progress_var = tk.DoubleVar(value=0.0)
        ttk.Progressbar(prog_win, variable=progress_var, maximum=100, length=320).pack(pady=8)
        status_lbl = tk.Label(prog_win, text="0%", font=(self.font_family, 9), fg="#666666")
        status_lbl.pack()

        def animate(i=0):
            if i > 100:
                try:
                    self._do_replace_and_restart(tmp_path, prog_win)
                except Exception as e:
                    self.log(f"模拟替换重启：替换失败 {e}")
                    try:
                        prog_win.destroy()
                    except Exception:
                        pass
                    messagebox.showerror("模拟更新失败", f"替换流程出错：\n\n{e}", parent=self.root)
                return
            progress_var.set(i)
            status_lbl.config(text=f"{i}%")
            prog_win.after(10, lambda: animate(i + 2))

        animate(0)

    def debug_simulate_full_update(self):
        if self._sim_http_server is not None:
            messagebox.showinfo("提示", "模拟更新的本地服务已在运行，请先完成或取消当前模拟。",
                                parent=self.root)
            return
        current_exe = _get_current_exe_path()
        if not os.path.exists(current_exe):
            messagebox.showerror("模拟更新失败", f"找不到当前 exe：\n{current_exe}", parent=self.root)
            return
        if not messagebox.askyesno(
            "模拟完整更新流程",
            "将进行以下操作：\n\n"
            "  1. 把当前 exe 复制到临时目录，当作「新版安装包」\n"
            "  2. 本地启动一个限速 HTTP 服务（约 5 秒下完）\n"
            "  3. 调用真正的下载流程（进度/速度/暂停/取消/确认安装）\n"
            "  4. 点「立即安装」后替换、重启\n\n"
            "用途：验证多线程、断点续传、暂停/取消。\n"
            "模拟源文件固定为 fwc_sim_new.exe，便于续传命中。\n\n"
            "继续吗？",
            parent=self.root,
        ):
            return
        fake_new_exe = os.path.join(tempfile.gettempdir(), "fwc_sim_new.exe")
        if os.path.exists(fake_new_exe):
            try:
                os.remove(fake_new_exe)
            except Exception as e:
                self.log(f"模拟完整更新：旧模拟文件无法删除 {e}")
                messagebox.showerror(
                    "模拟更新失败",
                    f"旧的模拟文件被占用，无法删除：\n{fake_new_exe}\n\n{e}\n\n"
                    f"请关闭正在使用该文件的程序后重试。",
                    parent=self.root,
                )
                return
        try:
            shutil.copy2(current_exe, fake_new_exe)
            self.log(f"模拟完整更新：已复制到 {fake_new_exe}")
        except Exception as e:
            self.log(f"模拟完整更新：复制失败 {e}")
            messagebox.showerror("模拟更新失败", f"复制 exe 时出错：\n\n{e}", parent=self.root)
            return

        try:
            sim_dir = tempfile.gettempdir()

            class _ThrottledHandler(http.server.SimpleHTTPRequestHandler):
                def log_message(self, format, *args):
                    return

                def do_GET(self):
                    try:
                        path = self.translate_path(self.path)
                        if not os.path.isfile(path):
                            self.send_error(404, "File not found")
                            return
                        total = os.path.getsize(path)

                        range_header = self.headers.get("Range")
                        start = 0
                        end = total - 1
                        status = 200
                        if range_header and range_header.startswith("bytes="):
                            r = range_header[len("bytes="):].strip()
                            parts = r.split("-")
                            try:
                                if parts[0]:
                                    start = int(parts[0])
                                if len(parts) > 1 and parts[1]:
                                    end = int(parts[1])
                            except Exception:
                                start = 0
                                end = total - 1
                            if start < 0:
                                start = 0
                            if end >= total:
                                end = total - 1
                            if start > end:
                                self.send_error(416, "Requested Range Not Satisfiable")
                                return
                            status = 206

                        length = end - start + 1

                        self.send_response(status)
                        self.send_header("Content-Type", "application/octet-stream")
                        self.send_header("Accept-Ranges", "bytes")
                        self.send_header("Content-Length", str(length))
                        if status == 206:
                            self.send_header("Content-Range",
                                             f"bytes {start}-{end}/{total}")
                        self.end_headers()

                        duration = 5.0
                        bytes_per_sec = max(1.0, length / duration)
                        chunk_size = 64 * 1024
                        interval = chunk_size / bytes_per_sec

                        with open(path, "rb") as f:
                            f.seek(start)
                            remaining = length
                            next_time = time.time()
                            while remaining > 0:
                                to_read = min(chunk_size, remaining)
                                buf = f.read(to_read)
                                if not buf:
                                    break
                                try:
                                    self.wfile.write(buf)
                                except (BrokenPipeError, ConnectionResetError):
                                    return
                                remaining -= len(buf)
                                next_time += interval
                                sleep_time = next_time - time.time()
                                if sleep_time > 0:
                                    time.sleep(sleep_time)
                            try:
                                self.wfile.flush()
                            except Exception:
                                pass
                    except Exception:
                        pass

            def _handler_factory(*args, **kwargs):
                kwargs["directory"] = sim_dir
                return _ThrottledHandler(*args, **kwargs)

            server = _ReusableTCPServer(("127.0.0.1", SIM_HTTP_PORT), _handler_factory)
            server.daemon_threads = True
            th = threading.Thread(target=server.serve_forever, daemon=True)
            th.start()
            self._sim_http_server = server
            self._sim_http_thread = th
            self._sim_fake_exe_path = fake_new_exe
        except OSError as e:
            self.log(f"模拟完整更新：HTTP 端口被占用 {e}")
            messagebox.showerror("模拟更新失败",
                                 f"本地 HTTP 端口 {SIM_HTTP_PORT} 被占用：\n\n{e}\n\n"
                                 f"请关闭占用该端口的程序后重试。",
                                 parent=self.root)
            try:
                os.remove(fake_new_exe)
            except Exception:
                pass
            return

        filename = os.path.basename(fake_new_exe)
        url = f"http://127.0.0.1:{SIM_HTTP_PORT}/{urllib.parse.quote(filename)}"
        self.log(f"模拟完整更新：下载 URL = {url}（限速：约 5 秒）")

        def _cleanup_after_sim():
            self._stop_sim_http_server()
            try:
                if self._sim_fake_exe_path and os.path.exists(self._sim_fake_exe_path):
                    os.remove(self._sim_fake_exe_path)
                    self.log(f"模拟完整更新：已清理 {self._sim_fake_exe_path}")
            except Exception as e:
                self.log(f"模拟完整更新：清理模拟 exe 失败 {e}")
            self._sim_fake_exe_path = None

        self.download_and_install_update(url, on_close_callback=_cleanup_after_sim)

    def _stop_sim_http_server(self):
        try:
            if self._sim_http_server is not None:
                self._sim_http_server.shutdown()
                self._sim_http_server.server_close()
                self.log("模拟完整更新：本地 HTTP 服务已关闭")
        except Exception:
            pass
        self._sim_http_server = None
        self._sim_http_thread = None

    def open_test_weather_panel(self):
        if self.test_weather_data is None:
            self.test_weather_data = self._make_default_test_weather()

        data = self.test_weather_data

        def _creator():
            win, main_frame, fixed_btn_frame, on_close = make_scrollable_toplevel(
                self.root,
                "测试天气数据",
                geometry="500x760",
                minsize=(480, 700),
                dpi_scale=self.dpi_scale
            )

            tk.Label(main_frame, text="手动设置天气信息（仅用于 UI 测试）",
                     font=(self.font_family, 11, "bold")).pack(anchor="w", pady=(0, 10))

            frame_tag = ttk.LabelFrame(main_frame, text="天气图标", padding=8)
            frame_tag.pack(fill=tk.X, pady=6)
            tag_var = tk.StringVar(value=data.get("tag", "clear"))
            frame_tags = ttk.Frame(frame_tag)
            frame_tags.pack(fill=tk.X)
            for i, (label, tag) in enumerate(TEST_WEATHER_TAGS):
                ttk.Radiobutton(frame_tags, text=label, variable=tag_var,
                                value=tag).grid(row=i // 4, column=i % 4, sticky="w", padx=4, pady=2)

            mode_var = tk.StringVar(value=self.test_weather_mode)

            def on_mode_change():
                self.test_weather_mode = mode_var.get()
                self.need_redraw = True
                try:
                    self.update_clock()
                except Exception:
                    pass

            mode_frame = ttk.Frame(frame_tag)
            mode_frame.pack(fill=tk.X, pady=(6, 0))
            tk.Label(mode_frame, text="时段预览：", font=(self.font_family, 9)).pack(side=tk.LEFT)
            for txt, val in [("自动", "auto"), ("白天", "day"), ("夜晚", "night"),
                             ("日出", "sunrise"), ("日落", "sunset")]:
                ttk.Radiobutton(mode_frame, text=txt, variable=mode_var, value=val,
                                command=on_mode_change).pack(side=tk.LEFT, padx=3)

            frame_city = ttk.LabelFrame(main_frame, text="城市名称", padding=8)
            frame_city.pack(fill=tk.X, pady=6)
            city_var = tk.StringVar(value=data.get("city", self.city))
            ttk.Entry(frame_city, textvariable=city_var).pack(fill=tk.X)

            frame_temp = ttk.LabelFrame(main_frame, text="温度 / 天气描述", padding=8)
            frame_temp.pack(fill=tk.X, pady=6)
            row1 = ttk.Frame(frame_temp)
            row1.pack(fill=tk.X)
            tk.Label(row1, text="温度：", width=6).pack(side=tk.LEFT)
            temp_var = tk.StringVar(value=str(data.get("temp", "25")))
            ttk.Entry(row1, textvariable=temp_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row1, text="℃", width=3).pack(side=tk.LEFT)
            tk.Label(row1, text="描述：", width=6).pack(side=tk.LEFT, padx=(10, 0))
            desc_var = tk.StringVar(value=str(data.get("desc", "测试晴朗")))
            ttk.Entry(row1, textvariable=desc_var, width=14).pack(side=tk.LEFT, padx=4)

            frame_env = ttk.LabelFrame(main_frame, text="湿度 / 风速 / 风向 / 降水", padding=8)
            frame_env.pack(fill=tk.X, pady=6)
            row2 = ttk.Frame(frame_env)
            row2.pack(fill=tk.X)
            tk.Label(row2, text="湿度：", width=6).pack(side=tk.LEFT)
            humidity_var = tk.StringVar(value=str(data.get("humidity", "50")))
            ttk.Entry(row2, textvariable=humidity_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row2, text="%", width=3).pack(side=tk.LEFT)
            tk.Label(row2, text="风速：", width=6).pack(side=tk.LEFT, padx=(10, 0))
            wind_var = tk.StringVar(value=str(data.get("wind", "10")))
            ttk.Entry(row2, textvariable=wind_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row2, text="km/h", width=6).pack(side=tk.LEFT)

            row2c = ttk.Frame(frame_env)
            row2c.pack(fill=tk.X, pady=(6, 0))
            tk.Label(row2c, text="风向：", width=6).pack(side=tk.LEFT)
            wind_dir_var = tk.StringVar(value=str(data.get("wind_dir", "0")))
            ttk.Entry(row2c, textvariable=wind_dir_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row2c, text="° (0-359)", width=12).pack(side=tk.LEFT)

            row2b = ttk.Frame(frame_env)
            row2b.pack(fill=tk.X, pady=(6, 0))
            tk.Label(row2b, text="降水量：", width=6).pack(side=tk.LEFT)
            precip_var = tk.StringVar(value=str(data.get("precipitation", "0")))
            ttk.Entry(row2b, textvariable=precip_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row2b, text="mm", width=3).pack(side=tk.LEFT)
            tk.Label(row2b, text="降水概率：", width=8).pack(side=tk.LEFT, padx=(10, 0))
            precip_prob_var = tk.StringVar(value=str(data.get("precip_prob", "0")))
            ttk.Entry(row2b, textvariable=precip_prob_var, width=8).pack(side=tk.LEFT, padx=4)
            tk.Label(row2b, text="%", width=3).pack(side=tk.LEFT)

            frame_sun = ttk.LabelFrame(main_frame, text="日出 / 日落", padding=8)
            frame_sun.pack(fill=tk.X, pady=6)
            row3 = ttk.Frame(frame_sun)
            row3.pack(fill=tk.X)
            tk.Label(row3, text="日出：", width=6).pack(side=tk.LEFT)
            sunrise_var = tk.StringVar(value=str(data.get("sunrise", "06:00")))
            ttk.Entry(row3, textvariable=sunrise_var, width=10).pack(side=tk.LEFT, padx=4)
            tk.Label(row3, text="日落：", width=6).pack(side=tk.LEFT, padx=(10, 0))
            sunset_var = tk.StringVar(value=str(data.get("sunset", "18:00")))
            ttk.Entry(row3, textvariable=sunset_var, width=10).pack(side=tk.LEFT, padx=4)

            def apply_test():
                self.test_weather_data = {
                    "tag": tag_var.get(),
                    "city": city_var.get().strip() or self.city,
                    "temp": temp_var.get().strip() or "--",
                    "desc": desc_var.get().strip() or "",
                    "humidity": humidity_var.get().strip() or "--",
                    "wind": wind_var.get().strip() or "--",
                    "wind_dir": wind_dir_var.get().strip() or "0",
                    "precipitation": precip_var.get().strip() or "--",
                    "precip_prob": precip_prob_var.get().strip() or "--",
                    "sunrise": sunrise_var.get().strip() or "--:--",
                    "sunset": sunset_var.get().strip() or "--:--",
                }
                if self.test_weather_enabled:
                    self.need_redraw = True
                    try:
                        self.update_clock()
                    except Exception:
                        pass
                on_close()

            def reset_default():
                self.test_weather_data = self._make_default_test_weather()
                d = self.test_weather_data
                tag_var.set(d["tag"])
                city_var.set(d["city"])
                temp_var.set(d["temp"])
                desc_var.set(d["desc"])
                humidity_var.set(d["humidity"])
                wind_var.set(d["wind"])
                wind_dir_var.set(d.get("wind_dir", "0"))
                precip_var.set(d["precipitation"])
                precip_prob_var.set(d["precip_prob"])
                sunrise_var.set(d["sunrise"])
                sunset_var.set(d["sunset"])

            ttk.Button(fixed_btn_frame, text="应用", command=apply_test).pack(side=tk.LEFT, padx=6)
            ttk.Button(fixed_btn_frame, text="恢复默认值", command=reset_default).pack(side=tk.LEFT, padx=6)
            ttk.Button(fixed_btn_frame, text="取消", command=on_close).pack(side=tk.LEFT, padx=6)

            return win

        self._open_singleton("test_weather", _creator)

    def _fetch_all_dates_async(self):
        self.date_cache = self._get_date_info()
        self._fetch_holiday_info_async(force=True)

    def recalc_window_size_by_font(self):
        base_w, base_h = STYLE_SIZE[self.win_style]
        base_font = 11
        scale = (self.font_size / base_font) * self.dpi_scale
        new_w = max(int(base_w * scale), 120)
        new_h = max(int(base_h * scale), 60)
        self.root.geometry(f"{new_w}x{new_h}+{self.mem_x}+{self.mem_y}")
        self.cached_w = new_w
        self.cached_h = new_h
        self.root.update_idletasks()
        self.need_redraw = True

    def debug_skip_countdown(self):
        self.refresh_countdown = 5
        self.need_redraw = True
        try:
            self.update_clock()
        except Exception:
            pass

    def _real_resize(self):
        self.resize_throttle_id = None
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        if w != self.cached_w or h != self.cached_h:
            self.cached_w = w
            self.cached_h = h
            self.need_redraw = True
            self.update_clock()

    def start_tray(self):
        icon_img = create_tray_icon()
        self.tray_icon = pystray.Icon("weather_clock", icon_img,
                                      self.get_display_version(),
                                      self._build_tray_menu())
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _set_lock_window(self, value):
        self.lock_window = bool(value)
        try:
            self._lock_menu_var.set(self.lock_window)
        except Exception:
            pass
        ref = getattr(self, "_settings_lock_var_ref", None)
        if callable(ref):
            try:
                ref()
            except Exception:
                pass
        self.bind_mouse_events()
        self.save_config()
        try:
            if self.tray_icon is not None:
                self.tray_icon.update_menu()
        except Exception:
            pass

    def _on_lock_menu_clicked(self):
        try:
            target = bool(self._lock_menu_var.get())
        except Exception:
            target = not self.lock_window
        self._set_lock_window(target)

    def _toggle_lock_from_tray(self):
        self._set_lock_window(not self.lock_window)

    def on_close_window(self):
        self.save_config()
        self.root.withdraw()

    def bind_mouse_events(self):
        if self.lock_window:
            self.canvas.unbind("<ButtonPress-1>")
            self.canvas.unbind("<B1-Motion>")
            self.root.unbind("<ButtonRelease-1>")
            self.canvas.unbind("<Button-3>")
        else:
            self.canvas.bind("<ButtonPress-1>", self.on_drag_start)
            self.canvas.bind("<B1-Motion>", self.on_drag_move)
            self.root.bind("<ButtonRelease-1>", self.on_drag_end)
            self.canvas.bind("<Button-3>", self.show_right_menu)

    def apply_window_attr(self):
        try:
            self.root.attributes("-alpha", self.win_alpha)
        except Exception:
            pass

        try:
            self.root.attributes("-topmost", bool(self.always_on_top))
        except Exception:
            pass

        self.bind_mouse_events()
        self.root.update_idletasks()
        hwnd = self.root.winfo_id()

        remove_taskbar_entry(hwnd)

        self.need_redraw = True
        try:
            self.canvas.update_idletasks()
            self.update_clock()
        except Exception:
            pass

    # ================= 设置面板 =================
    def open_settings(self):
        def _creator():
            win, main_frame, fixed_btn_frame, on_close = make_scrollable_toplevel(
                self.root,
                "悬浮时钟天气 - 设置面板",
                geometry="640x900",
                minsize=(600, 840),
                dpi_scale=self.dpi_scale
            )

            # ============ 愚人节彩蛋（仅 4 月 1 日显示，置顶；异常不影响后续控件）============
            if self._is_april_fools_today():
                self._show_april_fools_egg(main_frame, win)
            # ============ 愚人节彩蛋结束 ============

            group_win = ttk.LabelFrame(main_frame, text="窗口基础设置", padding=8)
            group_win.pack(fill=tk.X, pady=6)
            tk.Label(group_win, text="窗口透明度（0%完全透明 ~ 100%不透明）").pack(anchor="w")
            alpha_pct_var = tk.IntVar(value=int(self.win_alpha * 100))
            ttk.Scale(group_win, from_=0, to=100, variable=alpha_pct_var,
                      orient=tk.HORIZONTAL).pack(fill=tk.X)
            pct_label = tk.Label(group_win, text=f"当前：{alpha_pct_var.get()}%")
            pct_label.pack()

            def update_pct_label(*args):
                pct_label.config(text=f"当前：{alpha_pct_var.get()}%")

            alpha_pct_var.trace_add("write", update_pct_label)
            frame_color = ttk.Frame(group_win)
            frame_color.pack(fill=tk.X, pady=4)
            tk.Label(frame_color, text="窗口底色：").pack(side=tk.LEFT)
            color_preview = tk.Label(frame_color, text="　　", bg=self.bg_color, relief="solid")
            color_preview.pack(side=tk.LEFT, padx=6)

            def pick_color():
                c = colorchooser.askcolor(title="选择背景色", initialcolor=self.bg_color)
                if c[1]:
                    self.bg_color = c[1]
                    color_preview.config(bg=self.bg_color)
                    try:
                        self.canvas.config(bg=self.bg_color)
                    except Exception:
                        pass

            btn_pick_color = tk.Button(frame_color, text="选择颜色", command=pick_color)
            btn_pick_color.pack(side=tk.RIGHT)

            top_var = tk.BooleanVar(value=self.always_on_top)
            lock_var = tk.BooleanVar(value=self.lock_window)

            ttk.Checkbutton(group_win, text="窗口置顶", variable=top_var).pack(anchor="w")
            ttk.Checkbutton(group_win,
                            text="锁定窗口位置（锁定后无法用鼠标拖动窗口，从托盘菜单解锁）",
                            variable=lock_var).pack(anchor="w")

            weather_bg_var = tk.BooleanVar(value=self.weather_bg_enabled)

            def on_weather_bg_toggle():
                enabled = weather_bg_var.get()
                if enabled:
                    self.bg_color = "#222836"
                    color_preview.config(bg=self.bg_color)
                    try:
                        self.canvas.config(bg=self.bg_color)
                    except Exception:
                        pass
                    btn_pick_color.config(state="disabled")
                else:
                    btn_pick_color.config(state="normal")

            ttk.Checkbutton(
                group_win,
                text="在背景显示静态效果（根据天气显示日光/明月/云雨等，勾选后背景色固定为默认值）",
                variable=weather_bg_var,
                command=on_weather_bg_toggle
            ).pack(anchor="w")

            if self.weather_bg_enabled:
                btn_pick_color.config(state="disabled")

            def refresh_lock_var_state():
                try:
                    lock_var.set(self.lock_window)
                except Exception:
                    pass

            self._settings_lock_var_ref = refresh_lock_var_state

            group_units = ttk.LabelFrame(main_frame, text="单位设置", padding=8)
            group_units.pack(fill=tk.X, pady=6)
            tk.Label(group_units, text="温度单位：").pack(anchor="w")
            temp_unit_var = tk.StringVar(value=self.temp_unit)
            frame_temp_unit = ttk.Frame(group_units)
            frame_temp_unit.pack(anchor="w", pady=(2, 0))
            ttk.Radiobutton(frame_temp_unit, text="摄氏度（℃）",
                            variable=temp_unit_var, value="C").pack(side=tk.LEFT, padx=4)
            ttk.Radiobutton(frame_temp_unit, text="华氏度（℉）",
                            variable=temp_unit_var, value="F").pack(side=tk.LEFT, padx=4)

            group_display = ttk.LabelFrame(main_frame, text="内容显示设置", padding=8)
            group_display.pack(fill=tk.X, pady=6)

            frame_font_family = ttk.Frame(group_display)
            frame_font_family.pack(fill=tk.X, pady=(0, 4))
            tk.Label(frame_font_family, text="字体：").pack(side=tk.LEFT)
            try:
                all_fonts = sorted(set(tkfont.families()))
            except Exception:
                all_fonts = [DEFAULT_FONT_FAMILY]
            if self.font_family not in all_fonts:
                all_fonts = [self.font_family] + all_fonts
            font_family_var = tk.StringVar(value=self.font_family)
            cb_font = ttk.Combobox(frame_font_family, textvariable=font_family_var,
                                   values=all_fonts, state="readonly", width=30)
            cb_font.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

            def on_font_selected(event=None):
                new_font = font_family_var.get().strip()
                if not new_font or new_font == self.font_family:
                    return
                accepted = self.preview_font_dialog(new_font)
                if accepted:
                    self.font_family = new_font
                    self.save_config()
                    self.need_redraw = True
                    try:
                        self.update_clock()
                    except Exception:
                        pass
                    self.log(f"[字体] 已切换到 {new_font}")
                else:
                    try:
                        font_family_var.set(self.font_family)
                    except Exception:
                        pass
                    self.log(f"[字体] 用户拒绝字体 {new_font}，保持 {self.font_family}")

            cb_font.bind("<<ComboboxSelected>>", on_font_selected)

            frame_font = ttk.Frame(group_display)
            frame_font.pack(fill=tk.X)
            tk.Label(frame_font, text="字号：").pack(side=tk.LEFT)
            fontsize_var = tk.IntVar(value=self.font_size)
            ttk.Spinbox(frame_font, from_=6, to=40, textvariable=fontsize_var, width=6).pack(side=tk.LEFT, padx=5)

            frame_style = ttk.Frame(group_display)
            frame_style.pack(fill=tk.X, pady=4)
            style_var = tk.StringVar(value=self.win_style)
            ttk.Radiobutton(frame_style, text="简略模式", variable=style_var, value="brief").pack(side=tk.LEFT, padx=4)
            ttk.Radiobutton(frame_style, text="普通模式", variable=style_var, value="normal").pack(side=tk.LEFT, padx=4)
            ttk.Radiobutton(frame_style, text="详细模式", variable=style_var, value="detailed").pack(side=tk.LEFT, padx=4)

            rpg_var = tk.BooleanVar(value=self.exp_rpg_bg)
            ttk.Checkbutton(group_display, text="RPG 背景（像素风渐变 + 装饰边框）",
                            variable=rpg_var).pack(anchor="w", pady=(4, 0))

            frame_hint = ttk.Frame(group_display)
            frame_hint.pack(fill=tk.X, pady=4)
            show_auto_hint_var = tk.BooleanVar(value=self.show_auto_refresh_hint)
            show_last_hint_var = tk.BooleanVar(value=self.show_last_refresh_hint)
            ttk.Checkbutton(frame_hint, text="显示自动刷新倒计时",
                            variable=show_auto_hint_var).pack(side=tk.LEFT, padx=4)
            ttk.Checkbutton(frame_hint, text="显示上次刷新时间",
                            variable=show_last_hint_var).pack(side=tk.LEFT, padx=4)

            anim_var = tk.BooleanVar(value=self.enable_refresh_animation)
            ttk.Checkbutton(group_display,
                            text="启用刷新动画（浅蓝过渡 + 扫描线）",
                            variable=anim_var).pack(anchor="w", pady=(4, 0))

            group_dl = ttk.LabelFrame(main_frame, text="下载设置", padding=8)
            group_dl.pack(fill=tk.X, pady=6)
            dl_multi_var = tk.BooleanVar(value=self.download_multithread)
            ttk.Checkbutton(group_dl, text="启用多线程下载（4 线程，加速下载，可能不稳定）",
                            variable=dl_multi_var).pack(anchor="w")
            dl_resume_var = tk.BooleanVar(value=self.download_resume)
            ttk.Checkbutton(group_dl, text="启用断点续传（下载失败后保留已下载部分，下次继续）",
                            variable=dl_resume_var).pack(anchor="w")

            group_custom = ttk.LabelFrame(main_frame, text="个性化设置", padding=8)
            group_custom.pack(fill=tk.X, pady=6)
            tk.Label(group_custom, text="悬浮窗底部自定义文字（留空不显示）：").pack(anchor="w")
            custom_text_var = tk.StringVar(value=self.custom_bottom_text)
            ttk.Entry(group_custom, textvariable=custom_text_var).pack(fill=tk.X, pady=4)

            group_weather = ttk.LabelFrame(main_frame, text="城市天气设置", padding=8)
            group_weather.pack(fill=tk.X, pady=6)
            tk.Label(group_weather, text=f"当前城市：{self.city}").pack(anchor="w")

            def change_city():
                on_close()
                self.menu_change_city()

            tk.Button(group_weather, text="更换城市 / GPS定位", command=change_city).pack(anchor="w")

            group_update = ttk.LabelFrame(main_frame, text="软件更新", padding=8)
            group_update.pack(fill=tk.X, pady=6)
            tk.Label(group_update, text=f"当前版本：{self.get_display_version()}").pack(anchor="w")
            tk.Label(group_update, text=f"仓库：{GITHUB_REPO}",
                     fg="#666666", font=(self.font_family, 8)).pack(anchor="w", pady=(0, 4))

            auto_chk_var = tk.BooleanVar(value=self.auto_check_update)
            periodic_chk_var = tk.BooleanVar(value=self.periodic_check_update)

            chk_auto = ttk.Checkbutton(group_update,
                                       text="启动时自动检查更新",
                                       variable=auto_chk_var)
            chk_auto.pack(anchor="w", pady=(0, 2))

            chk_periodic = ttk.Checkbutton(group_update,
                                           text="后台每隔 15 分钟检查更新（有新版才提示）",
                                           variable=periodic_chk_var)
            chk_periodic.pack(anchor="w", pady=(0, 4))

            btn_check_update = ttk.Button(group_update, text="立即检查更新",
                                          command=self.menu_check_update)
            btn_check_update.pack(anchor="w", pady=(0, 4))

            btn_upgrade_beta = ttk.Button(group_update, text="升级至 Beta 版",
                                          command=self.prompt_upgrade_to_beta)
            btn_upgrade_beta.pack(anchor="w")

            if self.beta_edition:
                try:
                    chk_auto.config(state="disabled")
                    chk_periodic.config(state="disabled")
                    btn_check_update.config(state="disabled")
                    btn_upgrade_beta.config(state="disabled")
                except Exception:
                    pass

            group_system = ttk.LabelFrame(main_frame, text="系统选项", padding=8)
            group_system.pack(fill=tk.X, pady=6)
            startup_var = tk.BooleanVar(value=is_startup_enabled())
            ttk.Checkbutton(group_system, text="开机自动启动", variable=startup_var).pack(anchor="w")

            def reset_all_config():
                if messagebox.askyesno("重置确认", "确定重置全部设置为默认值吗？"):
                    self.reset_config()
                    on_close()

            tk.Button(group_system, text="一键重置全部设置",
                      command=reset_all_config, fg="#d62828").pack(anchor="w", pady=4)

            def clean_all_files():
                if not messagebox.askyesno(
                    "清理确认",
                    "将执行以下操作：\n\n"
                    "  1. 重置所有设置为默认值\n"
                    "  2. 删除配置文件\n"
                    "  3. 删除开机自启快捷方式（如果开启了）\n"
                    "  4. 删除临时目录下的 fwc_* 文件（包括更新缓存）\n"
                    "  5. 退出软件\n\n"
                    "注意：exe 本体需要你手动删除。\n\n"
                    "确定清理吗？",
                    parent=win,
                ):
                    return
                cleaned = []
                failed = []
                self.is_quitting = True
                try:
                    if self._log_file_handle is not None:
                        self._log_file_handle.close()
                        self._log_file_handle = None
                except Exception:
                    pass
                try:
                    self._stop_sim_http_server()
                except Exception:
                    pass
                try:
                    if os.path.exists(CONFIG_FILE):
                        os.remove(CONFIG_FILE)
                        cleaned.append(f"配置文件：{CONFIG_FILE}")
                except Exception as e:
                    failed.append(f"配置文件：{e}")
                try:
                    if LINK_PATH.exists():
                        os.remove(LINK_PATH)
                        cleaned.append(f"开机自启：{LINK_PATH}")
                except Exception as e:
                    failed.append(f"开机自启：{e}")
                try:
                    if os.path.isdir(UPDATE_CACHE_DIR):
                        shutil.rmtree(UPDATE_CACHE_DIR, ignore_errors=True)
                        cleaned.append(f"更新缓存目录：{UPDATE_CACHE_DIR}")
                except Exception as e:
                    failed.append(f"更新缓存目录：{e}")
                try:
                    tmp_dir = tempfile.gettempdir()
                    count = 0
                    for name in os.listdir(tmp_dir):
                        if name.startswith("fwc_"):
                            path = os.path.join(tmp_dir, name)
                            try:
                                if os.path.isfile(path):
                                    os.remove(path)
                                    count += 1
                            except Exception as e:
                                failed.append(f"{name}：{e}")
                    if count > 0:
                        cleaned.append(f"临时文件：{count} 个（{tmp_dir}）")
                except Exception as e:
                    failed.append(f"临时目录：{e}")
                exe_path = _get_current_exe_path()
                lines = []
                if cleaned:
                    lines.append("已清理：")
                    lines.extend("  ✓ " + x for x in cleaned)
                else:
                    lines.append("（没有找到可清理的文件）")
                if failed:
                    lines.append("")
                    lines.append("清理失败（可能被占用）：")
                    lines.extend("  ✗ " + x for x in failed)
                lines.append("")
                lines.append(f"请手动删除 exe 本体：\n{exe_path}")
                lines.append("（如果同目录下有 .old 文件，也一起删掉）")
                lines.append("")
                lines.append("点击「确定」后软件将退出。")
                messagebox.showinfo("清理完成", "\n".join(lines), parent=win)
                try:
                    on_close()
                except Exception:
                    pass
                try:
                    self.quit_app()
                except Exception:
                    pass

            tk.Button(group_system, text="一键清理全部文件",
                      command=clean_all_files, fg="#d62828").pack(anchor="w", pady=4)
            tk.Label(group_system,
                     text="（清理后会退出软件，exe 需手动删除）",
                     fg="#888888", font=(self.font_family, 8)).pack(anchor="w", padx=(20, 0))

            def save_setting():
                old_font = self.font_size
                self.win_alpha = alpha_pct_var.get() / 100.0
                self.always_on_top = top_var.get()
                new_lock = lock_var.get()
                self.font_size = fontsize_var.get()
                self.exp_rpg_bg = rpg_var.get()
                self.custom_bottom_text = custom_text_var.get().strip()
                self.show_auto_refresh_hint = show_auto_hint_var.get()
                self.show_last_refresh_hint = show_last_hint_var.get()
                self.download_multithread = bool(dl_multi_var.get())
                self.download_resume = bool(dl_resume_var.get())
                self.enable_refresh_animation = bool(anim_var.get())
                self.auto_check_update = bool(auto_chk_var.get())
                self.periodic_check_update = bool(periodic_chk_var.get())

                new_unit = temp_unit_var.get()
                if new_unit not in ("C", "F"):
                    new_unit = "C"
                self.temp_unit = new_unit

                self.weather_bg_enabled = bool(weather_bg_var.get())
                if self.weather_bg_enabled:
                    self.bg_color = "#222836"
                    self.exp_rpg_bg = False

                new_style = style_var.get()
                if new_style != self.win_style:
                    self.set_style(new_style)
                if old_font != self.font_size:
                    self.recalc_window_size_by_font()
                self._set_lock_window(new_lock)
                self.apply_window_attr()
                set_startup(startup_var.get())
                self.save_config()
                self.need_redraw = True
                self.update_clock()
                on_close()

            ttk.Button(fixed_btn_frame, text="保存设置", command=save_setting).pack(side=tk.LEFT, padx=6)
            ttk.Button(fixed_btn_frame, text="取消", command=on_close).pack(side=tk.LEFT, padx=6)

            return win

        self._open_singleton("settings", _creator)

    def reset_config(self):
        default = {
            "city": "合肥", "city_lat": 31.86, "city_lon": 117.28,
            "font_size": 11, "font_family": DEFAULT_FONT_FAMILY,
            "win_x": 200, "win_y": 200,
            "win_style": "normal", "win_alpha": 0.86, "bg_color": "#222836",
            "always_on_top": True, "lock_window": False, "developer_mode": False,
            "custom_bottom_text": "", "show_auto_refresh_hint": True,
            "show_last_refresh_hint": True, "exp_rpg_bg": False,
            "weather_bg_enabled": False,
            "temp_unit": "C",
            "debug_offline": False,
            "download_multithread": False, "download_resume": False,
            "enable_refresh_animation": True,
            "auto_check_update": True,
            "periodic_check_update": True,
        }
        for k, v in default.items():
            setattr(self, k, v)
        self.mem_x = default["win_x"]
        self.mem_y = default["win_y"]
        self.config = default
        self._settings_lock_var_ref = None
        self._last_notified_version = None
        self.last_weather_refresh_time = None
        self.test_weather_enabled = False
        self.test_weather_data = None
        self.test_weather_mode = "auto"
        self.sim_date = None
        self.date_cache = None
        self.holiday_cache = None
        self.refresh_start_time = 0
        self.weather_bg_enabled = False
        self.temp_unit = "C"
        try:
            self._lock_menu_var.set(self.lock_window)
        except Exception:
            pass
        self.set_style("normal")
        self.apply_window_attr()
        self.save_config()

    def set_style(self, style_name):
        if style_name not in STYLE_SIZE:
            return
        self.win_style = style_name
        self.recalc_window_size_by_font()
        self.save_config()

    def auto_refresh_all_timer(self):
        if self.is_quitting:
            return
        if self.test_weather_enabled:
            self.need_redraw = True
            self.root.after(1000, self.auto_refresh_all_timer)
            return
        self.refresh_countdown -= 1
        if self.refresh_countdown <= 0:
            self.refresh_countdown = self.auto_refresh_interval
            self.is_refresh_hint = True
            self.fetch_all_async()
        self.need_redraw = True
        self.root.after(1000, self.auto_refresh_all_timer)

    def reset_refresh_countdown(self):
        self.refresh_countdown = self.auto_refresh_interval
        self.need_redraw = True

    def load_config(self):
        d = {
            "city": "合肥", "city_lat": 31.86, "city_lon": 117.28,
            "font_size": 11, "font_family": DEFAULT_FONT_FAMILY,
            "win_x": 200, "win_y": 200,
            "win_style": "normal", "win_alpha": 0.86, "bg_color": "#222836",
            "always_on_top": True, "lock_window": False, "developer_mode": False,
            "custom_bottom_text": "", "show_auto_refresh_hint": True,
            "show_last_refresh_hint": True, "exp_rpg_bg": False,
            "weather_bg_enabled": False,
            "temp_unit": "C",
            "debug_offline": False,
            "download_multithread": False, "download_resume": False,
            "enable_refresh_animation": True,
            "auto_check_update": True,
            "periodic_check_update": True,
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    d.update(json.load(f))
            except Exception as e:
                print(f"[配置] 读取失败: {e}")
        return d

    def save_config(self):
        if self.is_quitting:
            return
        try:
            os.makedirs(DOC_PATH, exist_ok=True)
            self.config.update({
                "win_x": self.mem_x, "win_y": self.mem_y,
                "city": self.city, "city_lat": self.city_lat, "city_lon": self.city_lon,
                "font_size": self.font_size, "font_family": self.font_family,
                "win_style": self.win_style,
                "win_alpha": self.win_alpha, "bg_color": self.bg_color,
                "always_on_top": self.always_on_top, "lock_window": self.lock_window,
                "developer_mode": self.developer_mode,
                "custom_bottom_text": self.custom_bottom_text,
                "show_auto_refresh_hint": self.show_auto_refresh_hint,
                "show_last_refresh_hint": self.show_last_refresh_hint,
                "exp_rpg_bg": self.exp_rpg_bg,
                "weather_bg_enabled": self.weather_bg_enabled,
                "temp_unit": self.temp_unit,
                "debug_offline": self.debug_offline,
                "download_multithread": self.download_multithread,
                "download_resume": self.download_resume,
                "enable_refresh_animation": self.enable_refresh_animation,
                "auto_check_update": self.auto_check_update,
                "periodic_check_update": self.periodic_check_update,
            })
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存配置失败: {e}")

    def on_drag_start(self, e):
        self.drag_x = e.x
        self.drag_y = e.y

    def on_drag_move(self, e):
        dx = e.x - self.drag_x
        dy = e.y - self.drag_y
        new_x = self.root.winfo_x() + dx
        new_y = self.root.winfo_y() + dy
        self.root.geometry(f"+{new_x}+{new_y}")
        self.mem_x = new_x
        self.mem_y = new_y

    def on_drag_end(self, e):
        self.save_config()

    def show_right_menu(self, event):
        if self.is_quitting or self.lock_window:
            return
        try:
            self._lock_menu_var.set(self.lock_window)
        except Exception:
            pass
        try:
            self.right_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.right_menu.grab_release()

    def show_debug_info(self):
        last_refresh_str = (
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(self.last_weather_refresh_time))
            if self.last_weather_refresh_time is not None else "尚未成功刷新"
        )
        messagebox.showinfo(
            "调试信息",
            f"版本: {self.get_display_version()}\n"
            f"仓库: {GITHUB_REPO}\n"
            f"DPI缩放: {self.dpi_scale:.2f}\n"
            f"lunar_python: {'已加载' if _LUNAR_AVAILABLE else '未加载'}\n"
            f"窗口样式:{self.win_style}\n"
            f"字体:{self.font_family}\n"
            f"字号:{self.font_size}\n"
            f"温度单位:{self.temp_unit}\n"
            f"城市:{self.city}\n"
            f"纬度:{self.city_lat}\n"
            f"经度:{self.city_lon}\n"
            f"透明度:{self.win_alpha:.2f} ({int(self.win_alpha * 100)}%)\n"
            f"锁定位置:{self.lock_window}\n"
            f"置顶:{self.always_on_top}\n"
            f"RPG背景:{self.exp_rpg_bg}\n"
            f"天气背景:{self.weather_bg_enabled}\n"
            f"关闭联网:{self.debug_offline}\n"
            f"测试天气:{self.test_weather_enabled}\n"
            f"测试模式:{self.test_weather_mode}\n"
            f"模拟日期:{self.sim_date or '无'}\n"
            f"多线程下载:{self.download_multithread}\n"
            f"断点续传:{self.download_resume}\n"
            f"刷新动画:{self.enable_refresh_animation}\n"
            f"启动检查更新:{self.auto_check_update}\n"
            f"定时检查更新:{self.periodic_check_update}\n"
            f"已提示版本:{self._last_notified_version or '无'}\n"
            f"Beta 版:{self.beta_edition} (beta{self.beta_version})\n"
            f"真Beta常量:{BETA_EDITION}\n"
            f"开发者模式:{self.developer_mode}\n"
            f"单例窗口数:{len(self._singleton_windows)}\n"
            f"mem_x={self.mem_x}, mem_y={self.mem_y}\n"
            f"倒计时={self.refresh_countdown}秒\n"
            f"上次刷新={last_refresh_str}\n"
            f"日志文件={RUNTIME_LOG_FILE}",
            parent=self.root
        )

    def menu_refresh_all(self):
        if self.is_quitting or self.is_fetching_weather:
            return
        if self.test_weather_enabled:
            self.need_redraw = True
            try:
                self.refresh_start_time = time.time()
                self._fetch_all_dates_async()
                self.update_clock()
            except Exception:
                pass
            return
        self.is_refresh_hint = True
        self.reset_refresh_countdown()
        self.need_redraw = True
        self.fetch_all_async()

    def get_gps_location(self):
        with self._gps_lock:
            if self._gps_server is not None:
                self.log("GPS 定位：检测到上一次定位仍在进行，先取消上一次")
                self._gps_cancel_event.set()
                for _ in range(20):
                    if self._gps_server is None:
                        break
                    time.sleep(0.1)

        def work():
            server = None
            cancel_event = threading.Event()
            try:
                with self._gps_lock:
                    self._gps_cancel_event = cancel_event
                    server = _ReusableTCPServer(("127.0.0.1", GPS_PORT), GPSHttpHandler)
                    server.daemon_threads = True
                    self._gps_server = server
                    threading.Thread(target=server.serve_forever, daemon=True).start()

                webbrowser.open(f"http://127.0.0.1:{GPS_PORT}")

                deadline = time.time() + GPS_QUEUE_TIMEOUT
                res = None
                while time.time() < deadline:
                    if cancel_event.is_set():
                        self.log("GPS 定位：已被新的定位请求打断")
                        return
                    try:
                        res = gps_result_queue.get(timeout=0.2)
                        break
                    except queue.Empty:
                        continue

                if res is None:
                    self.log(f"GPS定位超时({GPS_QUEUE_TIMEOUT}秒)")
                    self.weather_queue.put(("error", f"GPS定位超时({GPS_QUEUE_TIMEOUT}秒)，浏览器定位较慢，请重试"))
                    return
                if cancel_event.is_set():
                    return

                self.city_lat = res["lat"]
                self.city_lon = res["lon"]
                self.city = "定位坐标"
                try:
                    ssl_ctx = ssl.create_default_context()
                    ssl_ctx.check_hostname = False
                    ssl_ctx.verify_mode = ssl.CERT_NONE
                    url = (f"https://geocoding-api.open-meteo.com/v1/reverse?"
                           f"latitude={self.city_lat}&longitude={self.city_lon}&language=zh")
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=20, context=ssl_ctx) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                        if data.get("results"):
                            self.city = data["results"][0]["name"]
                except Exception as geo_err:
                    self.log(f"反向地理编码失败，已保留经纬度: {geo_err}")
                if cancel_event.is_set():
                    return
                self.save_config()
                self.is_refresh_hint = True
                self.reset_refresh_countdown()
                self.need_redraw = True
                self.fetch_all_async()
            except OSError as port_err:
                self.log(f"GPS端口占用: {port_err}")
                self.weather_queue.put(("error", "GPS端口被占用，请重启软件再试"))
            except Exception as err:
                self.log(f"GPS定位失败：{err}")
                self.weather_queue.put(("error", str(err)))
            finally:
                try:
                    if server is not None:
                        server.shutdown()
                        server.server_close()
                except Exception:
                    pass
                with self._gps_lock:
                    if self._gps_server is server:
                        self._gps_server = None

        threading.Thread(target=work, daemon=True).start()

    def menu_change_city(self):
        if self.is_quitting:
            return

        def _creator():
            win = tk.Toplevel(self.root)
            win.title("选择/搜索城市")
            win.geometry("380x220")
            win.attributes("-topmost", True)
            win.transient(self.root)
            hide_toplevel_taskbar(win)
            win.grab_set()

            self._apply_dpi_to_toplevel(win)

            tk.Label(win, text="从列表选择城市：").pack(anchor="w", padx=10)
            cb = ttk.Combobox(win, values=[item[0] for item in CITY_LIST], state="readonly")
            cb.set(self.city)
            cb.pack(pady=4, padx=10, fill=tk.X)
            tk.Label(win, text="或直接输入地名自动查询（如：芜湖）：").pack(anchor="w", padx=10, pady=(8, 0))
            entry = tk.Entry(win)
            entry.pack(pady=4, padx=10, fill=tk.X)

            def use_gps_loc():
                win.destroy()
                self.get_gps_location()

            def ok():
                text = entry.get().strip()
                if text:
                    self.geocode_and_set_city(text)
                    win.destroy()
                    return
                sel = cb.get().strip()
                if sel and sel in CITY_DICT:
                    self.city = sel
                    self.city_lat, self.city_lon = CITY_DICT[sel]
                    self.save_config()
                    self.is_refresh_hint = True
                    self.reset_refresh_countdown()
                    self.need_redraw = True
                    self.fetch_all_async()
                win.destroy()

            frame_btn = tk.Frame(win)
            frame_btn.pack(pady=8)
            tk.Button(frame_btn, text="确定", command=ok).grid(row=0, column=0, padx=4)
            tk.Button(frame_btn, text="以我的GPS定位获取天气", command=use_gps_loc).grid(row=0, column=1, padx=4)
            self.root.wait_window(win)
            return win

        self._open_singleton("change_city", _creator)

    def geocode_and_set_city(self, location_name):
        def work():
            if self.debug_offline:
                self.weather_queue.put(("error", "调试模式：已关闭软件联网"))
                return
            try:
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                params = urllib.parse.urlencode({"name": location_name, "language": "zh", "count": 1})
                url = f"https://geocoding-api.open-meteo.com/v1/search?{params}"
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=20, context=ssl_ctx) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    if not data.get("results"):
                        msg = f"地理编码：找不到【{location_name}】"
                        self.log(msg)
                        self.weather_queue.put(("error", msg))
                        return
                    res = data["results"][0]
                    self.city = res["name"]
                    self.city_lat = res["latitude"]
                    self.city_lon = res["longitude"]
                    self.save_config()
                    self.is_refresh_hint = True
                    self.reset_refresh_countdown()
                    self.need_redraw = True
                    self.fetch_all_async()
            except Exception as err:
                msg = f"地理编码失败：{err}"
                self.log(msg)
                self.weather_queue.put(("error", str(err)))
        threading.Thread(target=work, daemon=True).start()

    def font_bigger(self):
        if self.is_quitting:
            return
        self.font_size += 1
        self.recalc_window_size_by_font()
        self.save_config()
        self.need_redraw = True

    def font_smaller(self):
        if self.is_quitting or self.font_size <= 6:
            return
        self.font_size -= 1
        self.recalc_window_size_by_font()
        self.save_config()
        self.need_redraw = True

    def quit_app(self):
        if self.is_quitting:
            try:
                self._stop_sim_http_server()
            except Exception:
                pass
            try:
                if self.tray_icon:
                    self.tray_icon.stop()
            except Exception:
                pass
            try:
                self.root.quit()
                self.root.destroy()
            except Exception:
                pass
            return
        self.log("软件正在退出…")
        self.is_quitting = True
        self.save_config()
        try:
            self._stop_sim_http_server()
        except Exception:
            pass
        try:
            self._gps_cancel_event.set()
        except Exception:
            pass
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        try:
            if self._log_file_handle is not None:
                self._log_file_handle.close()
                self._log_file_handle = None
        except Exception:
            pass
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass

    def fetch_all_async(self):
        if self.is_quitting or self.is_fetching_weather:
            return

        self.refresh_start_time = time.time()

        self.date_cache = None
        self.holiday_cache = None
        self.weather_data = None

        self.date_cache = self._get_date_info()
        self._fetch_holiday_info_async(force=True)

        if self.test_weather_enabled:
            self.need_redraw = True
            return
        if self.debug_offline:
            self.weather_queue.put(("error", "调试模式：已关闭软件联网"))
            return
        lat, lon = self.city_lat, self.city_lon
        self.is_fetching_weather = True
        self.log(f"刷新全局：请求天气 {self.city} ({lat}, {lon})")

        def worker():
            try:
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                url = ("https://api.open-meteo.com/v1/forecast?"
                       f"latitude={lat}&longitude={lon}"
                       "&current=temperature_2m,relative_humidity_2m,weather_code,"
                       "wind_speed_10m,wind_direction_10m,precipitation"
                       "&daily=sunrise,sunset,precipitation_probability_max"
                       "&timezone=auto")
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=REFRESH_TIMEOUT_SEC, context=ssl_ctx) as resp:
                    self.weather_queue.put(("ok", json.loads(resp.read().decode("utf-8"))))
            except Exception as err:
                self.log(f"天气接口请求异常: {err}")
                self.weather_queue.put(("error", str(err)))
            finally:
                self.is_fetching_weather = False

        threading.Thread(target=worker, daemon=True).start()

    def check_weather_queue(self):
        if self.is_quitting:
            return
        try:
            status, payload = self.weather_queue.get_nowait()
            if not self.test_weather_enabled:
                if status == "ok":
                    self.weather_data = payload
                    self.last_weather_refresh_time = time.time()
                    t = payload.get("current", {}).get("temperature_2m", "?")
                    self.log(f"天气更新成功：{self.city} {t}℃")
                    if self.enable_refresh_animation:
                        try:
                            self._flash_refresh_animation()
                        except Exception:
                            pass
                else:
                    self.weather_data = None
                    self.log(f"天气更新失败：{payload}")
                self.is_refresh_hint = False
                self.need_redraw = True
        except queue.Empty:
            pass
        self.root.after(800, self.check_weather_queue)

    def _classify_period(self, effective):
        if self.test_weather_enabled:
            m = self.test_weather_mode
            if m in ("day", "night", "sunrise", "sunset"):
                return m
        try:
            now = time.localtime()
            now_min = now.tm_hour * 60 + now.tm_min
            sr = effective.get("sunrise", "")
            ss = effective.get("sunset", "")
            sr_h, sr_m = map(int, sr.split(":"))
            ss_h, ss_m = map(int, ss.split(":"))
            sr_min = sr_h * 60 + sr_m
            ss_min = ss_h * 60 + ss_m
            if ss_min <= sr_min:
                raise ValueError("invalid sunrise/sunset")

            if sr_min - 2 <= now_min <= sr_min + 2:
                return "sunrise"
            if ss_min - 2 <= now_min <= ss_min + 2:
                return "sunset"
            if sr_min <= now_min < ss_min:
                return "day"
            return "night"
        except Exception:
            h = time.localtime().tm_hour
            return "day" if 7 <= h <= 19 else "night"

    def _is_daytime(self, effective):
        return self._classify_period(effective) in ("day", "sunrise")

    def _get_effective_weather(self):
        if self.test_weather_enabled and self.test_weather_data is not None:
            d = self.test_weather_data

            def _f(v, default):
                try:
                    return float(v)
                except Exception:
                    return default

            return {
                "city": d.get("city", self.city),
                "temp": _f(d.get("temp", "25"), d.get("temp", "25")),
                "humidity": _f(d.get("humidity", "50"), d.get("humidity", "50")),
                "wind": _f(d.get("wind", "10"), d.get("wind", "10")),
                "wind_dir": _f(d.get("wind_dir", "0"), 0),
                "precipitation": _f(d.get("precipitation", "0"), d.get("precipitation", "0")),
                "precip_prob": _f(d.get("precip_prob", "0"), d.get("precip_prob", "0")),
                "desc": d.get("desc", ""),
                "tag": d.get("tag", "clear"),
                "sunrise": d.get("sunrise", "--:--"),
                "sunset": d.get("sunset", "--:--"),
                "is_test": True,
            }
        if self.weather_data is None:
            return None
        try:
            curr = self.weather_data["current"]
            winfo = WEATHER_CODE_MAP.get(curr.get("weather_code", -1),
                                         {"desc": "未知天气", "tag": "cloudy"})
            precip_prob = "--"
            try:
                pp_arr = self.weather_data.get("daily", {}).get("precipitation_probability_max", [])
                if pp_arr and pp_arr[0] is not None:
                    precip_prob = pp_arr[0]
            except Exception:
                pass
            return {
                "city": self.city,
                "temp": curr.get("temperature_2m", "--"),
                "humidity": curr.get("relative_humidity_2m", "--"),
                "wind": curr.get("wind_speed_10m", "--"),
                "wind_dir": curr.get("wind_direction_10m", "--"),
                "precipitation": curr.get("precipitation", "--"),
                "precip_prob": precip_prob,
                "desc": winfo["desc"],
                "tag": winfo["tag"],
                "sunrise": self.weather_data["daily"]["sunrise"][0].split("T")[-1],
                "sunset": self.weather_data["daily"]["sunset"][0].split("T")[-1],
                "is_test": False,
            }
        except Exception:
            return None

    def _format_wind_line(self, effective):
        try:
            wind_speed = effective.get("wind", "--")
            wind_dir = effective.get("wind_dir", "--")
            level, level_desc = _wind_scale_kmh(wind_speed)
            dir_str = _wind_dir_to_chinese(wind_dir)
            parts = []
            try:
                s = float(wind_speed)
                parts.append(f"风速 {int(round(s))} km/h")
            except Exception:
                parts.append(f"风速 {wind_speed} km/h")
            if level is not None:
                parts.append(f"{level}级{level_desc}")
            if dir_str:
                parts.append(dir_str)
            return " ｜ ".join(parts)
        except Exception:
            return "风速 --"

    def draw_weather_icon(self, x, y, is_day, tag, period="day"):
        c = self.canvas
        c.delete("icon")
        scale = (self.font_size / 11) * self.dpi_scale
        r = max(8, min(64, int(14 * scale)))

        def ss(v):
            return int(round(v * scale))

        if period == "sunrise":
            c.create_line(x - r * 1.2, y + int(r * 0.3),
                          x + r * 1.2, y + int(r * 0.3),
                          fill="#8a6d2f", width=max(1, ss(2)),
                          capstyle=tk.ROUND, tags="icon")
            c.create_arc(x - r, y - r, x + r, y + r,
                         start=0, extent=180,
                         fill="#ffd54a", outline="#ffaa00",
                         width=max(1, ss(1)), tags="icon")
            for i in range(5):
                ang = math.pi * (0.1 + i * 0.2)
                x1 = x + int(r * 1.0 * math.cos(-ang))
                y1 = y - int(r * 1.0 * math.sin(ang))
                x2 = x + int(r * 1.4 * math.cos(-ang))
                y2 = y - int(r * 1.4 * math.sin(ang))
                c.create_line(x1, y1, x2, y2, fill="#ffb700",
                              width=max(1, ss(2)), capstyle=tk.ROUND, tags="icon")
            return
        if period == "sunset":
            c.create_line(x - r * 1.2, y + int(r * 0.3),
                          x + r * 1.2, y + int(r * 0.3),
                          fill="#8a3a1f", width=max(1, ss(2)),
                          capstyle=tk.ROUND, tags="icon")
            c.create_arc(x - r, y - r, x + r, y + r,
                         start=180, extent=180,
                         fill="#ff7a2a", outline="#cc4400",
                         width=max(1, ss(1)), tags="icon")
            for i in range(5):
                ang = math.pi * (0.1 + i * 0.2)
                x1 = x + int(r * 1.0 * math.cos(ang))
                y1 = y + int(r * 0.3)
                x2 = x + int(r * 1.3 * math.cos(ang))
                y2 = y + int(r * 0.3) + int(r * 0.15)
                c.create_line(x1, y1, x2, y2, fill="#ff9a4a",
                              width=max(1, ss(2)), capstyle=tk.ROUND, tags="icon")
            return

        if tag == "clear":
            if is_day:
                for i in range(12):
                    ang = math.pi * 2 * i / 12
                    long_ray = (i % 2 == 0)
                    r1 = r + ss(3 if long_ray else 1)
                    r2 = r + ss(8 if long_ray else 5)
                    c.create_line(
                        x + r1 * math.cos(ang), y + r1 * math.sin(ang),
                        x + r2 * math.cos(ang), y + r2 * math.sin(ang),
                        fill="#ffb700", width=max(1, ss(2)),
                        capstyle=tk.ROUND, tags="icon")
                c.create_oval(x - r - ss(1), y - r - ss(1), x + r + ss(1), y + r + ss(1),
                              fill="#ffdd55", outline="", tags="icon")
                c.create_oval(x - r, y - r, x + r, y + r,
                              fill="#ffcc00", outline="#ff9900",
                              width=max(1, ss(1)), tags="icon")
                hl = int(r * 0.4)
                c.create_oval(x - int(r * 0.55), y - int(r * 0.55),
                              x - int(r * 0.55) + hl, y - int(r * 0.55) + hl,
                              fill="#fff5b8", outline="", tags="icon")
            else:
                bg = self.bg_color
                c.create_oval(x - r, y - r, x + r, y + r, fill="#f5e6a8", outline="", tags="icon")
                c.create_oval(x - r + int(r * 0.55), y - r - int(r * 0.15),
                              x + r + int(r * 0.55), y + r - int(r * 0.15),
                              fill=bg, outline="", tags="icon")
                star_r = max(1, ss(1))
                for (sx, sy) in [(x + int(r * 0.5), y - int(r * 0.6)),
                                 (x + int(r * 0.2), y + int(r * 0.3)),
                                 (x - int(r * 0.5), y - int(r * 0.4))]:
                    c.create_oval(sx - star_r, sy - star_r, sx + star_r, sy + star_r,
                                  fill="#ffffff", outline="", tags="icon")

        elif tag == "cloudy":
            if is_day:
                c.create_oval(x - r, y - int(r * 0.1), x + int(r * 0.2), y + int(r * 0.7),
                              fill="#e8eef5", outline="", tags="icon")
                c.create_oval(x - int(r * 0.2), y - int(r * 0.8), x + int(r * 0.7), y + int(r * 0.3),
                              fill="#f5f9ff", outline="", tags="icon")
                c.create_oval(x - int(r * 0.8), y - int(r * 0.4), x + int(r * 0.2), y + int(r * 0.5),
                              fill="#eef4fa", outline="", tags="icon")
                c.create_oval(x - int(r * 0.15), y - int(r * 0.7), x + int(r * 0.55), y - int(r * 0.05),
                              fill="#ffffff", outline="", tags="icon")
                sun_r = int(r * 0.4)
                sun_x = x + int(r * 0.75)
                sun_y = y - int(r * 0.55)
                c.create_oval(sun_x - sun_r, sun_y - sun_r, sun_x + sun_r, sun_y + sun_r,
                              fill="#ffd54a", outline="#ffb300",
                              width=max(1, ss(1)), tags="icon")
            else:
                c.create_oval(x - r, y - int(r * 0.1), x + int(r * 0.2), y + int(r * 0.7),
                              fill="#9aa6b4", outline="", tags="icon")
                c.create_oval(x - int(r * 0.2), y - int(r * 0.8), x + int(r * 0.7), y + int(r * 0.3),
                              fill="#aab6c4", outline="", tags="icon")
                c.create_oval(x - int(r * 0.8), y - int(r * 0.4), x + int(r * 0.2), y + int(r * 0.5),
                              fill="#a0acba", outline="", tags="icon")
                c.create_oval(x - int(r * 0.15), y - int(r * 0.7), x + int(r * 0.55), y - int(r * 0.05),
                              fill="#b8c4d0", outline="", tags="icon")
                moon_r = int(r * 0.4)
                moon_x = x + int(r * 0.75)
                moon_y = y - int(r * 0.55)
                c.create_oval(moon_x - moon_r, moon_y - moon_r,
                              moon_x + moon_r, moon_y + moon_r,
                              fill="#f0ecd0", outline="#c8c098",
                              width=max(1, ss(1)), tags="icon")
                shadow_off = int(moon_r * 0.55)
                c.create_oval(moon_x - moon_r + shadow_off, moon_y - moon_r - int(moon_r * 0.1),
                              moon_x + moon_r + shadow_off, moon_y + moon_r - int(moon_r * 0.1),
                              fill=self.bg_color, outline="", tags="icon")
                star_r = max(1, ss(1))
                for (sx, sy) in [(x + int(r * 0.95), y - int(r * 0.85)),
                                 (x + int(r * 0.60), y - int(r * 0.95))]:
                    c.create_oval(sx - star_r, sy - star_r, sx + star_r, sy + star_r,
                                  fill="#ffffff", outline="", tags="icon")

        elif tag == "overcast":
            c.create_oval(x - r, y - int(r * 0.1), x + int(r * 0.3), y + int(r * 0.7),
                          fill="#a8b3c0", outline="", tags="icon")
            c.create_oval(x - int(r * 0.6), y - int(r * 0.5), x + int(r * 0.4), y + int(r * 0.4),
                          fill="#b8c2ce", outline="", tags="icon")
            c.create_oval(x - int(r * 0.2), y - int(r * 0.7), x + int(r * 0.7), y + int(r * 0.2),
                          fill="#c5ced9", outline="", tags="icon")
            c.create_oval(x - int(r * 0.3), y - int(r * 0.85), x + int(r * 0.5), y - int(r * 0.05),
                          fill="#99a5b3", outline="", tags="icon")

        elif tag in ("rain", "drizzle"):
            c.create_oval(x - int(r * 0.9), y - int(r * 0.2), x + int(r * 0.4), y + int(r * 0.5),
                          fill="#b8c2ce", outline="", tags="icon")
            c.create_oval(x - int(r * 0.4), y - int(r * 0.7), x + int(r * 0.5), y + int(r * 0.2),
                          fill="#cdd6e0", outline="", tags="icon")
            c.create_oval(x - int(r * 0.1), y - int(r * 0.55), x + int(r * 0.85), y + int(r * 0.15),
                          fill="#d8e0e8", outline="", tags="icon")
            drop_len = int(r * 0.55)
            start_y = y + int(r * 0.3)
            for i in range(3 if tag == "rain" else 2):
                dx = x - int(r * 0.5) + i * int(r * 0.55)
                c.create_line(dx, start_y, dx - int(r * 0.12), start_y + drop_len,
                              fill="#4aa3e8", width=max(1, ss(2)),
                              capstyle=tk.ROUND, tags="icon")
                c.create_line(dx, start_y, dx - int(r * 0.12), start_y + int(drop_len * 0.6),
                              fill="#7ec0f0", width=max(1, ss(2)),
                              capstyle=tk.ROUND, tags="icon")

        elif tag == "thunder":
            c.create_oval(x - int(r * 0.9), y - int(r * 0.3), x + int(r * 0.4), y + int(r * 0.4),
                          fill="#8f9aa8", outline="", tags="icon")
            c.create_oval(x - int(r * 0.4), y - int(r * 0.8), x + int(r * 0.5), y + int(r * 0.1),
                          fill="#a3aebb", outline="", tags="icon")
            c.create_oval(x - int(r * 0.1), y - int(r * 0.6), x + int(r * 0.85), y + int(r * 0.1),
                          fill="#aeb8c4", outline="", tags="icon")
            zx = x - int(r * 0.05)
            c.create_polygon(
                zx + int(r * 0.15), y + int(r * 0.15),
                zx - int(r * 0.1), y + int(r * 0.55),
                zx + int(r * 0.05), y + int(r * 0.55),
                zx - int(r * 0.15), y + int(r * 0.75),
                zx + int(r * 0.25), y + int(r * 0.75),
                zx + int(r * 0.05), y + int(r * 0.55),
                zx + int(r * 0.2), y + int(r * 0.55),
                zx + int(r * 0.3), y + int(r * 0.15),
                fill="#ffd54a", outline="#ffaa00",
                width=max(1, ss(1)), tags="icon")

        elif tag == "snow":
            c.create_oval(x - int(r * 0.9), y - int(r * 0.3), x + int(r * 0.4), y + int(r * 0.4),
                          fill="#cdd6e0", outline="", tags="icon")
            c.create_oval(x - int(r * 0.4), y - int(r * 0.75), x + int(r * 0.5), y + int(r * 0.15),
                          fill="#dde5ec", outline="", tags="icon")
            c.create_oval(x - int(r * 0.1), y - int(r * 0.55), x + int(r * 0.85), y + int(r * 0.15),
                          fill="#e5ecf3", outline="", tags="icon")
            flake_r = max(2, int(r * 0.2))
            for fx in [x - int(r * 0.4), x + int(r * 0.1), x + int(r * 0.5)]:
                fy = y + int(r * 0.55)
                for i in range(6):
                    ang = math.pi * 2 * i / 6
                    c.create_line(fx, fy,
                                  fx + flake_r * math.cos(ang),
                                  fy + flake_r * math.sin(ang),
                                  fill="#ffffff", width=max(1, ss(2)),
                                  capstyle=tk.ROUND, tags="icon")
                c.create_oval(fx - 1, fy - 1, fx + 1, fy + 1,
                              fill="#ffffff", outline="", tags="icon")

        elif tag == "fog":
            for dy, color in [(-int(r * 0.3), "#b8c2ce"),
                              (0, "#cdd6e0"),
                              (int(r * 0.35), "#dde5ec")]:
                c.create_line(x - r, y + dy, x + r, y + dy,
                              fill=color, width=max(1, ss(2)),
                              capstyle=tk.ROUND, tags="icon")
            c.create_line(x - int(r * 0.8), y + int(r * 0.15),
                          x - int(r * 0.3), y + int(r * 0.15),
                          fill="#e5ecf3", width=max(1, ss(2)),
                          capstyle=tk.ROUND, tags="icon")
            c.create_line(x + int(r * 0.1), y + int(r * 0.15),
                          x + int(r * 0.8), y + int(r * 0.15),
                          fill="#e5ecf3", width=max(1, ss(2)),
                          capstyle=tk.ROUND, tags="icon")

        else:
            c.create_oval(x - int(r * 0.8), y - int(r * 0.3), x + int(r * 0.8), y + int(r * 0.5),
                          fill="#cdd6e0", outline="", tags="icon")

    def update_clock_loop(self):
        if self.is_quitting:
            return
        time_str = time.strftime("%H:%M:%S", time.localtime())
        if time_str != self.last_time_str:
            self.last_time_str = time_str
            self.need_redraw = True
        if self.need_redraw:
            self.update_clock()
            self.need_redraw = False
        self.root.after(1000, self.update_clock_loop)

    def _draw_rpg_bg(self, canvas, w, h):
        base = self.bg_color
        try:
            r0, g0, b0 = hex_to_rgb(base)
        except Exception:
            r0, g0, b0 = (0.13, 0.16, 0.21)

        luminance = 0.2126 * r0 + 0.7152 * g0 + 0.0722 * b0
        is_light_bg = luminance > 0.5

        bands = max(12, h // 8)
        if is_light_bg:
            for i in range(bands):
                t = i / max(bands - 1, 1)
                lighten = 0.15 + 0.15 * (1 - t)
                rr = max(0, min(255, int((r0 + (1 - r0) * lighten) * 255)))
                gg = max(0, min(255, int((g0 + (1 - g0) * lighten) * 255)))
                bb = max(0, min(255, int((b0 + (1 - b0) * lighten) * 255)))
                y1 = int(i * h / bands)
                y2 = int((i + 1) * h / bands) + 1
                canvas.create_rectangle(0, y1, w, y2,
                                        fill=f"#{rr:02x}{gg:02x}{bb:02x}",
                                        outline="", tags="rpgbg")
            grid_color = "#c8c8c8"
            border_outer = "#8a6d2f"
            border_inner = "#a68a4a"
            corner_fill = "#a68a4a"
        else:
            for i in range(bands):
                t = i / max(bands - 1, 1)
                rr = max(0, min(255, int((r0 * (0.6 + 0.6 * (1 - t))) * 255)))
                gg = max(0, min(255, int((g0 * (0.6 + 0.6 * (1 - t))) * 255)))
                bb = max(0, min(255, int((b0 * (0.6 + 0.6 * (1 - t))) * 255)))
                y1 = int(i * h / bands)
                y2 = int((i + 1) * h / bands) + 1
                canvas.create_rectangle(0, y1, w, y2,
                                        fill=f"#{rr:02x}{gg:02x}{bb:02x}",
                                        outline="", tags="rpgbg")
            grid_color = "#000000"
            border_outer = "#c8a24a"
            border_inner = "#8a6d2f"
            corner_fill = "#c8a24a"

        for x in range(0, w, 16):
            canvas.create_line(x, 0, x, h, fill=grid_color, width=1, tags="rpgbg")
        for y in range(0, h, 16):
            canvas.create_line(0, y, w, y, fill=grid_color, width=1, tags="rpgbg")

        canvas.create_rectangle(4, 4, w - 4, h - 4,
                                outline=border_outer, width=2, tags="rpgbg")
        canvas.create_rectangle(9, 9, w - 9, h - 9,
                                outline=border_inner, width=1, tags="rpgbg")

        corner = 10
        for (cx, cy) in [(6, 6), (w - 6 - corner, 6),
                         (6, h - 6 - corner), (w - 6 - corner, h - 6 - corner)]:
            canvas.create_rectangle(cx, cy, cx + corner, cy + corner,
                                    fill=corner_fill, outline=border_inner, tags="rpgbg")

    def _draw_sunrise_gradient(self, canvas, w, h, kind):
        if kind == "sunrise":
            bottom = SUNRISE_BOTTOM
            top = SUNRISE_TOP
        else:
            bottom = SUNSET_BOTTOM
            top = SUNSET_TOP
        bands = max(20, h // 4)
        for i in range(bands):
            t = i / max(bands - 1, 1)
            rgb = _lerp_color(top, bottom, t)
            color = _rgb255_to_hex(rgb)
            y1 = int(i * h / bands)
            y2 = int((i + 1) * h / bands) + 1
            canvas.create_rectangle(0, y1, w, y2, fill=color, outline="", tags="sunbg")

    def _soft_glow(self, canvas, cx, cy, radius, core_rgb, bg_rgb,
                   layers=10, outline_w=1, tags="weatherbg"):
        if radius < 2:
            return
        steps = [
            (1.00, 0.20),
            (0.72, 0.55),
            (0.48, 1.00),
        ]
        for ratio, blend in steps:
            rr = max(1, int(radius * ratio))
            rgb = _lerp_color(bg_rgb, core_rgb, blend)
            try:
                canvas.create_oval(cx - rr, cy - rr, cx + rr, cy + rr,
                                   fill=_rgb255_to_hex(rgb), outline="", tags=tags)
            except Exception:
                pass

    def _soft_cloud(self, canvas, cx, cy, r, top_rgb, bot_rgb,
                    bg_rgb, tags="weatherbg"):
        canvas.create_oval(cx - r, cy - int(r * 0.3), cx + int(r * 0.3), cy + int(r * 0.5),
                           fill=_rgb255_to_hex(bot_rgb), outline="", tags=tags)
        canvas.create_oval(cx - int(r * 0.6), cy - int(r * 0.6), cx + int(r * 0.6), cy + int(r * 0.2),
                           fill=_rgb255_to_hex(top_rgb), outline="", tags=tags)
        canvas.create_oval(cx - int(r * 0.2), cy - int(r * 0.4), cx + int(r * 0.8), cy + int(r * 0.3),
                           fill=_rgb255_to_hex(top_rgb), outline="", tags=tags)

    def _draw_weather_bg(self, canvas, w, h, tag, period):
        is_day = period in ("day", "sunrise")
        is_night = not is_day

        base = self.bg_color
        try:
            r0, g0, b0 = hex_to_rgb(base)
        except Exception:
            r0, g0, b0 = (0.13, 0.16, 0.21)

        # ============================================================
        # 背景渐变（顶部到底部）
        # ============================================================
        if tag == "clear" and is_day:
            top_rgb = (int(r0*255*0.55+50), int(g0*255*0.55+75), int(b0*255*0.55+120))
            bot_rgb = (int(r0*255*0.78+38), int(g0*255*0.78+58), int(b0*255*0.78+90))
        elif tag == "clear" and is_night:
            top_rgb = (int(r0*255*0.30), int(g0*255*0.30), int(b0*255*0.50+22))
            bot_rgb = (int(r0*255*0.48), int(g0*255*0.48), int(b0*255*0.68+16))
        elif tag == "cloudy":
            top_rgb = (int(r0*255*0.68+18), int(g0*255*0.68+22), int(b0*255*0.68+30))
            bot_rgb = (int(r0*255*0.84+10), int(g0*255*0.84+14), int(b0*255*0.84+20))
        elif tag == "overcast":
            top_rgb = (int(r0*255*0.58+14), int(g0*255*0.58+18), int(b0*255*0.58+24))
            bot_rgb = (int(r0*255*0.76+8),  int(g0*255*0.76+10), int(b0*255*0.76+16))
        elif tag in ("rain", "drizzle", "thunder"):
            top_rgb = (int(r0*255*0.44), int(g0*255*0.44+6), int(b0*255*0.44+18))
            bot_rgb = (int(r0*255*0.60), int(g0*255*0.60+3), int(b0*255*0.60+10))
        elif tag == "snow":
            top_rgb = (int(r0*255*0.58+20), int(g0*255*0.58+24), int(b0*255*0.58+32))
            bot_rgb = (int(r0*255*0.74+12), int(g0*255*0.74+16), int(b0*255*0.74+24))
        elif tag == "fog":
            top_rgb = (int(r0*255*0.64+30), int(g0*255*0.64+32), int(b0*255*0.64+36))
            bot_rgb = (int(r0*255*0.80+22), int(g0*255*0.80+24), int(b0*255*0.80+28))
        else:
            top_rgb = (int(r0*255*0.64), int(g0*255*0.64), int(b0*255*0.64))
            bot_rgb = (int(r0*255*0.84), int(g0*255*0.84), int(b0*255*0.84))

        top_rgb = tuple(max(0, min(255, v)) for v in top_rgb)
        bot_rgb = tuple(max(0, min(255, v)) for v in bot_rgb)
        mid_rgb = _lerp_color(top_rgb, bot_rgb, 0.5)

        # 背景渐变
        bands = max(24, h // 3)
        for i in range(bands):
            t = i / max(bands - 1, 1)
            rgb = _lerp_color(top_rgb, bot_rgb, t)
            color = _rgb255_to_hex(rgb)
            y1 = int(i * h / bands)
            y2 = int((i + 1) * h / bands) + 1
            canvas.create_rectangle(0, y1, w, y2, fill=color, outline="", tags="weatherbg")

        k = (self.font_size / 11) * self.dpi_scale

        # 面板底
        panel_rgb = tuple(max(0, min(255, int(c * 0.80))) for c in mid_rgb)
        panel_color = _rgb255_to_hex(panel_rgb)
        try:
            margin = int(5 * k)
            radius = int(14 * k)
            self._draw_round_rect(canvas, margin, margin, w - margin, h - margin,
                                  radius, fill=panel_color, outline="", tags="weatherbg")
        except Exception:
            pass

        bg_for_glow = panel_rgb

        # ============================================================
        # 晴朗白天
        # ============================================================
        if tag == "clear" and is_day:
            sun_x = int(w * 0.84)
            sun_y = int(h * 0.16)
            sun_r = int(13 * k)

            self._soft_glow(canvas, sun_x, sun_y, int(sun_r * 2.6),
                            (255, 235, 160), bg_for_glow)
            canvas.create_oval(sun_x - sun_r, sun_y - sun_r,
                               sun_x + sun_r, sun_y + sun_r,
                               fill="#ffe066", outline="", tags="weatherbg")
            hl = int(sun_r * 0.5)
            canvas.create_oval(sun_x - int(sun_r*0.5), sun_y - int(sun_r*0.5),
                               sun_x - int(sun_r*0.5) + hl, sun_y - int(sun_r*0.5) + hl,
                               fill="#fff8c0", outline="", tags="weatherbg")

        # ============================================================
        # 晴朗夜晚
        # ============================================================
        elif tag == "clear" and is_night:
            moon_x = int(w * 0.84)
            moon_y = int(h * 0.16)
            moon_r = int(11 * k)

            self._soft_glow(canvas, moon_x, moon_y, int(moon_r * 2.4),
                            (235, 245, 255), bg_for_glow)
            canvas.create_oval(moon_x - moon_r, moon_y - moon_r,
                               moon_x + moon_r, moon_y + moon_r,
                               fill="#f8f8e0", outline="", tags="weatherbg")
            shadow_offset = int(moon_r * 0.55)
            canvas.create_oval(moon_x - moon_r + shadow_offset, moon_y - moon_r - int(moon_r*0.1),
                               moon_x + moon_r + shadow_offset, moon_y + moon_r - int(moon_r*0.1),
                               fill=panel_color, outline="", tags="weatherbg")

            import random
            random.seed(42)
            for _ in range(16):
                sx = random.randint(int(w * 0.05), int(w * 0.90))
                sy = random.randint(int(h * 0.05), int(h * 0.55))
                dx_moon = sx - moon_x
                dy_moon = sy - moon_y
                if dx_moon*dx_moon + dy_moon*dy_moon < (moon_r * 3.5) ** 2:
                    continue
                sr = max(1, int(random.uniform(0.8, 1.4) * k))
                canvas.create_oval(sx - sr, sy - sr, sx + sr, sy + sr,
                                   fill="#ffffff", outline="", tags="weatherbg")

        # ============================================================
        # 多云
        # ============================================================
        elif tag == "cloudy":
            if is_day:
                sun_x = int(w * 0.84)
                sun_y = int(h * 0.16)
                sun_r = int(11 * k)
                self._soft_glow(canvas, sun_x, sun_y, int(sun_r * 2.4),
                                (255, 235, 160), bg_for_glow)
                canvas.create_oval(sun_x - sun_r, sun_y - sun_r,
                                   sun_x + sun_r, sun_y + sun_r,
                                   fill="#ffe066", outline="", tags="weatherbg")
            else:
                moon_x = int(w * 0.84)
                moon_y = int(h * 0.16)
                moon_r = int(10 * k)
                self._soft_glow(canvas, moon_x, moon_y, int(moon_r * 2.2),
                                (235, 245, 255), bg_for_glow)
                canvas.create_oval(moon_x - moon_r, moon_y - moon_r,
                                   moon_x + moon_r, moon_y + moon_r,
                                   fill="#f8f8e0", outline="", tags="weatherbg")
                shadow_offset = int(moon_r * 0.55)
                canvas.create_oval(moon_x - moon_r + shadow_offset, moon_y - moon_r,
                                   moon_x + moon_r + shadow_offset, moon_y + moon_r,
                                   fill=panel_color, outline="", tags="weatherbg")

            cloud_top = (int(mid_rgb[0]*0.5+115), int(mid_rgb[1]*0.5+122), int(mid_rgb[2]*0.5+134))
            cloud_bot = (int(mid_rgb[0]*0.5+80),  int(mid_rgb[1]*0.5+88),  int(mid_rgb[2]*0.5+100))
            self._soft_cloud(canvas, int(w*0.16), int(h*0.13), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.62), int(h*0.10), int(12*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.88), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.24), int(h*0.90), int(12*k),
                             cloud_top, cloud_bot, bg_for_glow)

        # ============================================================
        # 阴天
        # ============================================================
        elif tag == "overcast":
            cloud_top = (int(mid_rgb[0]*0.5+100), int(mid_rgb[1]*0.5+108), int(mid_rgb[2]*0.5+120))
            cloud_bot = (int(mid_rgb[0]*0.5+72),  int(mid_rgb[1]*0.5+80),  int(mid_rgb[2]*0.5+92))
            self._soft_cloud(canvas, int(w*0.18), int(h*0.12), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.52), int(h*0.10), int(15*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.13), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.28), int(h*0.90), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.70), int(h*0.92), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)

        # ============================================================
        # 雨
        # ============================================================
        elif tag == "rain":
            cloud_top = (int(mid_rgb[0]*0.5+72), int(mid_rgb[1]*0.5+80), int(mid_rgb[2]*0.5+94))
            cloud_bot = (int(mid_rgb[0]*0.5+48), int(mid_rgb[1]*0.5+56), int(mid_rgb[2]*0.5+72))
            self._soft_cloud(canvas, int(w*0.16), int(h*0.11), int(15*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.62), int(h*0.09), int(16*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.90), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)

            import random
            random.seed(123)
            for _ in range(34):
                rx = random.randint(int(w*0.05), int(w*0.95))
                ry = random.randint(int(h*0.08), int(h*0.92))
                if 0.18*w < rx < 0.82*w and 0.20*h < ry < 0.80*h:
                    continue
                dl = int(random.uniform(4, 7) * k)
                canvas.create_line(rx, ry, rx - int(1*k), ry + dl,
                                   fill="#8ec0e8", width=1,
                                   capstyle=tk.ROUND, tags="weatherbg")

        # ============================================================
        # 毛毛雨
        # ============================================================
        elif tag == "drizzle":
            cloud_top = (int(mid_rgb[0]*0.5+88), int(mid_rgb[1]*0.5+98), int(mid_rgb[2]*0.5+114))
            cloud_bot = (int(mid_rgb[0]*0.5+64), int(mid_rgb[1]*0.5+74),  int(mid_rgb[2]*0.5+90))
            self._soft_cloud(canvas, int(w*0.16), int(h*0.11), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.62), int(h*0.09), int(15*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.90), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)

            import random
            random.seed(222)
            for _ in range(45):
                rx = random.randint(int(w*0.05), int(w*0.95))
                ry = random.randint(int(h*0.08), int(h*0.92))
                if 0.18*w < rx < 0.82*w and 0.20*h < ry < 0.80*h:
                    continue
                dl = int(random.uniform(2, 3) * k)
                canvas.create_line(rx, ry, rx, ry + dl,
                                   fill="#a8c8e0", width=1,
                                   capstyle=tk.ROUND, tags="weatherbg")

        # ============================================================
        # 雷阵雨
        # ============================================================
        elif tag == "thunder":
            cloud_top = (int(mid_rgb[0]*0.5+58), int(mid_rgb[1]*0.5+66), int(mid_rgb[2]*0.5+86))
            cloud_bot = (int(mid_rgb[0]*0.5+36), int(mid_rgb[1]*0.5+42), int(mid_rgb[2]*0.5+62))
            self._soft_cloud(canvas, int(w*0.16), int(h*0.10), int(16*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.62), int(h*0.08), int(17*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.90), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)

            import random
            random.seed(456)
            for _ in range(30):
                rx = random.randint(int(w*0.05), int(w*0.95))
                ry = random.randint(int(h*0.10), int(h*0.92))
                if 0.18*w < rx < 0.82*w and 0.22*h < ry < 0.78*h:
                    continue
                dl = int(random.uniform(4, 8) * k)
                canvas.create_line(rx, ry, rx - int(1*k), ry + dl,
                                   fill="#7cb0d8", width=1,
                                   capstyle=tk.ROUND, tags="weatherbg")

            # ---- 闪电（大型增强版：多层辉光 + 长折线 + 分叉，位于文字区外的边角）----
            # 位置往左上取，落在左侧留白与右上云层之间的区域
            lx = int(w * 0.62)
            ly = int(h * 0.18)
            ls = int(26 * k)

            # 折线路径：整体更长、锯齿更明显，一直延伸到画面下缘附近
            bolt_pts = [
                (lx + int(ls * 0.25), ly),
                (lx - int(ls * 0.10), ly + int(ls * 0.55)),
                (lx + int(ls * 0.22), ly + int(ls * 0.55)),
                (lx - int(ls * 0.30), ly + int(ls * 1.30)),
                (lx + int(ls * 0.05), ly + int(ls * 1.30)),
                (lx - int(ls * 0.40), ly + int(ls * 2.00)),
                (lx - int(ls * 0.05), ly + int(ls * 2.00)),
                (lx - int(ls * 0.55), ly + int(ls * 2.70)),
            ]

            # 1) 外辉光：由粗到细、由暗到亮的多层光晕（整体调暗，更含蓄）
            glow_layers = [
                (max(10, int(5.0 * k)), "#1c2438"),   # 最外层暗蓝（比原来更暗）
                (max(8,  int(3.8 * k)), "#2e3a5c"),   # 次外层蓝（压暗）
                (max(6,  int(2.8 * k)), "#4e5a80"),   # 中层蓝灰（压暗）
                (max(4,  int(2.0 * k)), "#7e88a8"),   # 内层浅蓝（压暗，不再发白）
            ]
            for gw, gc in glow_layers:
                canvas.create_line(*[c for p in bolt_pts for c in p],
                                   fill=gc, width=gw,
                                   capstyle=tk.ROUND, joinstyle=tk.ROUND,
                                   tags="weatherbg")

            # 2) 核心：由亮黄白改为柔和米黄，降低对比度
            canvas.create_line(*[c for p in bolt_pts for c in p],
                               fill="#d8d0a0", width=max(3, int(1.8 * k)),
                               capstyle=tk.ROUND, joinstyle=tk.ROUND,
                               tags="weatherbg")

            # 3) 主分叉：同步调暗
            bx0 = lx + int(ls * 0.22)
            by0 = ly + int(ls * 0.55)
            canvas.create_line(bx0, by0,
                               bx0 + int(ls * 0.45), by0 + int(ls * 0.60),
                               fill="#b8b088", width=max(2, int(1.2 * k)),
                               capstyle=tk.ROUND, tags="weatherbg")

            # 4) 次分叉：再暗一档，几乎融进背景
            cx0 = lx - int(ls * 0.30)
            cy0 = ly + int(ls * 1.30)
            canvas.create_line(cx0, cy0,
                               cx0 - int(ls * 0.35), cy0 + int(ls * 0.75),
                               fill="#9a9478", width=max(1, int(1.0 * k)),
                               capstyle=tk.ROUND, tags="weatherbg")

        # ============================================================
        # 雪
        # ============================================================
        elif tag == "snow":
            cloud_top = (int(mid_rgb[0]*0.5+112), int(mid_rgb[1]*0.5+120), int(mid_rgb[2]*0.5+132))
            cloud_bot = (int(mid_rgb[0]*0.5+82),  int(mid_rgb[1]*0.5+90),  int(mid_rgb[2]*0.5+102))
            self._soft_cloud(canvas, int(w*0.16), int(h*0.11), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.62), int(h*0.09), int(15*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.84), int(h*0.90), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)

            import random
            random.seed(789)
            for _ in range(28):
                sx = random.randint(int(w*0.05), int(w*0.95))
                sy = random.randint(int(h*0.08), int(h*0.92))
                if 0.18*w < sx < 0.82*w and 0.22*h < sy < 0.80*h:
                    continue
                sr = max(1, int(random.uniform(1.0, 1.8) * k))
                canvas.create_oval(sx - sr, sy - sr, sx + sr, sy + sr,
                                   fill="#ffffff", outline="", tags="weatherbg")

        # ============================================================
        # 雾
        # ============================================================
        elif tag == "fog":
            fog_rgb = (int(mid_rgb[0]*0.5+55), int(mid_rgb[1]*0.5+60), int(mid_rgb[2]*0.5+68))
            fog_color = _rgb255_to_hex(fog_rgb)
            canvas.create_oval(-int(w*0.3), -int(h*0.3), int(w*1.3), int(h*1.3),
                               fill=fog_color, outline="", tags="weatherbg")

            cloud_top = (int(mid_rgb[0]*0.5+135), int(mid_rgb[1]*0.5+140), int(mid_rgb[2]*0.5+148))
            cloud_bot = (int(mid_rgb[0]*0.5+115), int(mid_rgb[1]*0.5+120), int(mid_rgb[2]*0.5+128))
            self._soft_cloud(canvas, int(w*0.20), int(h*0.16), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.72), int(h*0.20), int(12*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.30), int(h*0.86), int(14*k),
                             cloud_top, cloud_bot, bg_for_glow)
            self._soft_cloud(canvas, int(w*0.80), int(h*0.90), int(13*k),
                             cloud_top, cloud_bot, bg_for_glow)

    def _draw_bg_cloud(self, canvas, cx, cy, r, color1, color2):
        canvas.create_oval(cx - r, cy - int(r * 0.3), cx + int(r * 0.3), cy + int(r * 0.6),
                           fill=color2, outline="", tags="weatherbg")
        canvas.create_oval(cx - int(r * 0.5), cy - int(r * 0.7), cx + int(r * 0.5), cy + int(r * 0.2),
                           fill=color1, outline="", tags="weatherbg")
        canvas.create_oval(cx - int(r * 0.2), cy - int(r * 0.5), cx + int(r * 0.8), cy + int(r * 0.3),
                           fill=color1, outline="", tags="weatherbg")
        canvas.create_oval(cx - int(r * 0.7), cy - int(r * 0.1), cx + int(r * 0.2), cy + int(r * 0.5),
                           fill=color2, outline="", tags="weatherbg")

    def _build_date_lines(self):
        di = self.date_cache
        today = self._get_today()
        week_map = {"0": "星期日", "1": "星期一", "2": "星期二", "3": "星期三",
                    "4": "星期四", "5": "星期五", "6": "星期六"}
        try:
            week_str = week_map[today.strftime("%w")]
        except Exception:
            week_str = ""

        date_main = today.strftime("%Y年%m月%d日 ") + week_str

        lunar_part = ""
        if di:
            lunar_part = f"农历{di.get('lunar_day', '')}"

        normal_line1 = date_main
        if lunar_part:
            normal_line1 += " · " + lunar_part

        line2_parts = []
        if di:
            for f in di.get("festivals", []):
                line2_parts.append(f)
            for f in di.get("other_festivals", []):
                line2_parts.append(f)
            jq_today = di.get("jie_qi_today", "")
            jq_tomorrow = di.get("jie_qi_tomorrow", "")
            if jq_today:
                line2_parts.append(f"{jq_today}（今天）")
            elif jq_tomorrow:
                line2_parts.append(f"{jq_tomorrow}（明天）")

        holiday_disp = self._get_holiday_display()
        holiday_str = ""
        if holiday_disp == "loading":
            holiday_str = ""
        elif holiday_disp == "failed":
            holiday_str = "？"
        elif isinstance(holiday_disp, dict):
            info_type = holiday_disp.get("type", "")
            note = holiday_disp.get("note", "")
            if info_type == "假日":
                if note and note != "周末":
                    holiday_str = f"[休]{note}"
                else:
                    holiday_str = "[休]"
            elif info_type == "工作日":
                if note and ("补班" in note or "调休" in note):
                    holiday_str = f"[班]{note}"
                else:
                    holiday_str = ""

        if holiday_str:
            line2_parts.append(holiday_str)

        normal_line2 = " · ".join(line2_parts)

        detail_lunar = ""
        if di:
            detail_lunar = (f"农历：{di.get('lunar_day', '')} · "
                            f"{di.get('gan_zhi', '')}{di.get('zodiac', '')}年")

        detail_jieqi = ""
        if di:
            jq_today = di.get("jie_qi_today", "")
            jq_tomorrow = di.get("jie_qi_tomorrow", "")
            if jq_today:
                detail_jieqi = f"节气：{jq_today}（今天）"
            elif jq_tomorrow:
                detail_jieqi = f"节气：{jq_tomorrow}（明天）"

        detail_festivals = ""
        if di:
            all_fest = []
            for f in di.get("festivals", []):
                all_fest.append(f)
            for f in di.get("other_festivals", []):
                all_fest.append(f)
            if all_fest:
                detail_festivals = "节日：" + "、".join(all_fest)

        detail_holiday = ""
        if holiday_disp == "loading":
            detail_holiday = "节假日：获取中…"
        elif holiday_disp == "failed":
            detail_holiday = "节假日：获取失败"
        elif isinstance(holiday_disp, dict):
            info_type = holiday_disp.get("type", "")
            note = holiday_disp.get("note", "")
            if info_type == "假日":
                if note and note != "周末":
                    detail_holiday = f"节假日：{note} [休]"
                else:
                    detail_holiday = "节假日：周末 [休]"
            elif info_type == "工作日":
                if note and ("补班" in note or "调休" in note):
                    detail_holiday = f"节假日：{note} [班]"
                else:
                    detail_holiday = f"节假日：{note or '普通工作日'}"

        return {
            "normal_line1": normal_line1,
            "normal_line2": normal_line2,
            "detail_lunar": detail_lunar,
            "detail_jieqi": detail_jieqi,
            "detail_festivals": detail_festivals,
            "detail_holiday": detail_holiday,
        }

    def update_clock(self):
        if self.is_quitting:
            return
        c = self.canvas
        c.delete("all")
        w = self.cached_w
        h = self.cached_h

        effective = self._get_effective_weather()

        period = "day"
        if effective is not None:
            period = self._classify_period(effective)
        else:
            period = "day" if 7 <= time.localtime().tm_hour <= 19 else "night"

        is_sunrise = (period == "sunrise")
        is_sunset = (period == "sunset")

        if is_sunrise or is_sunset:
            self._draw_sunrise_gradient(c, w, h, period)
            main_color = "#ffffff"
            sub_color = "#ffe8d0"
            border_color = "#fff0d0"
        elif self.exp_rpg_bg:
            color_cfg = get_text_color_by_bg(self.bg_color)
            main_color = color_cfg["main"]
            sub_color = color_cfg["sub"]
            border_color = "#444c60"
            self._draw_rpg_bg(c, w, h)
        elif self.weather_bg_enabled:
            weather_tag = effective["tag"] if effective is not None else "cloudy"
            if self.test_weather_enabled and self.test_weather_data is not None:
                weather_tag = self.test_weather_data.get("tag", "clear")
            self._draw_weather_bg(c, w, h, weather_tag, period)

            main_color = "#ffffff"
            if weather_tag in ("snow", "fog", "cloudy", "overcast"):
                sub_color = "#d8e0e8"
            elif weather_tag in ("rain", "drizzle", "thunder"):
                sub_color = "#c8d2dc"
            else:
                sub_color = "#d0d8e0"
            border_color = "#444c60"

            k0 = (self.font_size / 11) * self.dpi_scale
            self._draw_round_rect(c, int(round(2 * k0)), int(round(2 * k0)),
                                  w - int(round(2 * k0)), h - int(round(2 * k0)),
                                  int(round(16 * k0)),
                                  fill="", outline=border_color,
                                  width=max(1, int(1.4 * k0)))
        else:
            color_cfg = get_text_color_by_bg(self.bg_color)
            main_color = color_cfg["main"]
            sub_color = color_cfg["sub"]
            border_color = "#444c60"
            k0 = (self.font_size / 11) * self.dpi_scale
            self._draw_round_rect(c, int(round(2 * k0)), int(round(2 * k0)),
                                  w - int(round(2 * k0)), h - int(round(2 * k0)),
                                  int(round(16 * k0)),
                                  fill=self.bg_color, outline=border_color,
                                  width=max(1, int(1.4 * k0)))

        self._text_rgb_main = _hex_to_rgb_tuple(main_color)
        self._text_rgb_weather = _hex_to_rgb_tuple(main_color)

        fs = self.font_size
        k = (fs / 11) * self.dpi_scale
        ff = self.font_family

        def s(v):
            return int(round(v * k))

        time_str = time.strftime("%H:%M:%S", time.localtime())

        date_info = self._build_date_lines()
        date_ready = self._is_date_ready()
        date_timed_out = self._is_date_timed_out()

        if date_ready:
            date_main_text = date_info["normal_line1"]
            date_main_color = sub_color
            date_sub_text = date_info["normal_line2"]
            date_sub_color = sub_color
        else:
            if date_timed_out:
                date_main_text = "日期获取失败"
                date_main_color = "#ff5555"
            else:
                date_main_text = "正在获取日期…"
                date_main_color = sub_color
            date_sub_text = ""
            date_sub_color = sub_color

        self._text_rgb_date = _hex_to_rgb_tuple(date_main_color)

        weather_ready = (effective is not None)
        weather_timed_out = self._is_weather_timed_out()

        bottom_y = h - s(12)
        if self.show_auto_refresh_hint:
            if self.test_weather_enabled:
                c.create_text(w // 2, bottom_y, anchor="s", text="测试天气模式",
                              fill="#ffb86c", font=(ff, max(6, fs - 1)))
            else:
                mm = self.refresh_countdown // 60
                ss2 = self.refresh_countdown % 60
                c.create_text(w // 2, bottom_y, anchor="s",
                              text=f"自动刷新 {mm:02d}:{ss2:02d}",
                              fill=sub_color, font=(ff, max(6, fs - 1)))
            bottom_y -= s(16)
        if (self.show_last_refresh_hint and self.last_weather_refresh_time is not None
                and not self.test_weather_enabled):
            c.create_text(w // 2, bottom_y, anchor="s",
                          text="上次刷新 " + time.strftime(
                              "%H:%M:%S", time.localtime(self.last_weather_refresh_time)),
                          fill=sub_color, font=(ff, max(6, fs - 1)))
            bottom_y -= s(16)
        if self.custom_bottom_text.strip():
            c.create_text(w // 2, bottom_y, anchor="s", text=self.custom_bottom_text,
                          fill=sub_color, font=(ff, max(6, fs - 1)))

        suffix = ""
        if is_sunrise:
            suffix = "（日出）"
        elif is_sunset:
            suffix = "（日落）"

        if self.win_style == "brief":
            c.create_text(s(16), s(14), anchor="nw", text=time_str,
                          fill=main_color, font=(ff, fs + 8, "bold"),
                          tags=("main_text",))
            c.create_text(s(16), s(48), anchor="nw", text=date_main_text,
                          fill=date_main_color, font=(ff, fs),
                          tags=("date_text",))
            if date_sub_text:
                c.create_text(s(16), s(66), anchor="nw", text=date_sub_text,
                              fill=date_sub_color, font=(ff, fs - 1),
                              tags=("date_text",))
            return

        if self.win_style == "normal":
            c.create_text(s(20), s(14), anchor="nw", text=time_str,
                          fill=main_color, font=(ff, fs + 8, "bold"),
                          tags=("main_text",))
            c.create_text(s(20), s(48), anchor="nw", text=date_main_text,
                          fill=date_main_color, font=(ff, fs),
                          tags=("date_text",))
            if date_sub_text:
                c.create_text(s(20), s(66), anchor="nw", text=date_sub_text,
                              fill=date_sub_color, font=(ff, max(6, fs - 1)),
                              tags=("date_text",))
            weather_y = s(88)
            icon_cx, icon_cy, text_x = s(28), weather_y + s(12), s(60)
            if weather_ready:
                try:
                    is_day = period in ("day", "sunrise")
                    self.draw_weather_icon(icon_cx, icon_cy, is_day, effective["tag"], period)
                    desc_with_suffix = f"{effective['desc']}{suffix}"
                    temp_text = self._format_temp(effective['temp'])
                    c.create_text(text_x, weather_y, anchor="nw",
                                  text=f"{effective['city']}  {temp_text}  {desc_with_suffix}",
                                  fill=main_color, font=(ff, fs, "bold"),
                                  tags=("weather_text",))
                    c.create_text(text_x, weather_y + s(28), anchor="nw",
                                  text=f"湿度 {effective['humidity']}% ｜ {self._format_wind_line(effective)}",
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    c.create_text(text_x, weather_y + s(48), anchor="nw",
                                  text=f"降水 {effective['precipitation']} mm ｜ 概率 {effective['precip_prob']}%",
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    c.create_text(text_x, weather_y + s(68), anchor="nw",
                                  text=f"日出 {effective['sunrise']} ｜ 日落 {effective['sunset']}",
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                except Exception as e:
                    self.log(f"天气UI解析异常: {e}")
                    c.create_text(text_x, weather_y, anchor="nw", text="天气解析出错",
                                  fill="#ff8888", font=(ff, fs))
            else:
                if weather_timed_out:
                    c.create_text(text_x, weather_y, anchor="nw", text="天气获取失败",
                                  fill="#ff5555", font=(ff, fs))
                else:
                    c.create_text(text_x, weather_y, anchor="nw", text="正在获取天气…",
                                  fill=sub_color, font=(ff, fs))
            return

        if self.win_style == "detailed":
            c.create_text(s(24), s(14), anchor="nw", text=time_str,
                          fill=main_color, font=(ff, fs + 8, "bold"),
                          tags=("main_text",))
            c.create_text(s(24), s(48), anchor="nw", text=date_main_text,
                          fill=date_main_color, font=(ff, fs),
                          tags=("date_text",))

            if date_ready:
                info_y = s(70)
                if date_info["detail_lunar"]:
                    c.create_text(s(24), info_y, anchor="nw", text=date_info["detail_lunar"],
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    info_y += s(18)
                if date_info["detail_jieqi"]:
                    c.create_text(s(24), info_y, anchor="nw", text=date_info["detail_jieqi"],
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    info_y += s(18)
                if date_info["detail_festivals"]:
                    c.create_text(s(24), info_y, anchor="nw", text=date_info["detail_festivals"],
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    info_y += s(18)
                if date_info["detail_holiday"]:
                    c.create_text(s(24), info_y, anchor="nw", text=date_info["detail_holiday"],
                                  fill=sub_color, font=(ff, max(6, fs - 1)))
                    info_y += s(18)

            line_y = s(150)
            c.create_line(s(24), line_y, w - s(24), line_y,
                          fill=border_color, width=max(1, int(1.2 * k)))
            weather_y = line_y + s(14)
            icon_cx, icon_cy, text_x = s(32), weather_y + s(12), s(64)
            if weather_ready:
                try:
                    is_day = period in ("day", "sunrise")
                    self.draw_weather_icon(icon_cx, icon_cy, is_day, effective["tag"], period)
                    desc_with_suffix = f"{effective['desc']}{suffix}"
                    temp_text = self._format_temp(effective['temp'])
                    c.create_text(text_x, weather_y, anchor="nw", text="当前天气",
                                  fill="#a0b4d8",
                                  font=(ff, fs, "bold"),
                                  tags=("weather_text",))
                    c.create_text(text_x, weather_y + s(24), anchor="nw",
                                  text=f"{effective['city']} · {temp_text} · {desc_with_suffix}",
                                  fill=main_color, font=(ff, fs + 1))
                    c.create_text(text_x, weather_y + s(56), anchor="nw", text="环境参数",
                                  fill="#a0b4d8",
                                  font=(ff, fs, "bold"))
                    c.create_text(text_x, weather_y + s(78), anchor="nw",
                                  text=f"相对湿度：{effective['humidity']} %",
                                  fill=sub_color, font=(ff, fs))
                    c.create_text(text_x, weather_y + s(100), anchor="nw",
                                  text=self._format_wind_line(effective),
                                  fill=sub_color, font=(ff, fs))
                    c.create_text(text_x, weather_y + s(122), anchor="nw",
                                  text=f"降水量：{effective['precipitation']} mm",
                                  fill=sub_color, font=(ff, fs))
                    c.create_text(text_x, weather_y + s(144), anchor="nw",
                                  text=f"降水概率：{effective['precip_prob']} %",
                                  fill=sub_color, font=(ff, fs))
                    c.create_text(text_x, weather_y + s(172), anchor="nw", text="日照时间",
                                  fill="#a0b4d8",
                                  font=(ff, fs, "bold"))
                    c.create_text(text_x, weather_y + s(194), anchor="nw",
                                  text=f"日出：{effective['sunrise']}",
                                  fill=sub_color, font=(ff, fs))
                    c.create_text(text_x, weather_y + s(216), anchor="nw",
                                  text=f"日落：{effective['sunset']}",
                                  fill=sub_color, font=(ff, fs))
                except Exception as e:
                    self.log(f"天气UI解析异常: {e}")
                    c.create_text(text_x, weather_y, anchor="nw", text="天气解析出错",
                                  fill="#ff8888", font=(ff, fs))
            else:
                if weather_timed_out:
                    c.create_text(text_x, weather_y, anchor="nw", text="天气获取失败",
                                  fill="#ff5555", font=(ff, fs))
                else:
                    c.create_text(text_x, weather_y, anchor="nw", text="正在获取天气…",
                                  fill=sub_color, font=(ff, fs))

    def _draw_round_rect(self, canvas, x1, y1, x2, y2, r, **kw):
        pts = [
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1
        ]
        return canvas.create_polygon(pts, smooth=True, **kw)


if __name__ == "__main__":
    mutex_handle = None
    try:
        mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
        last_error = ctypes.windll.kernel32.GetLastError()
        ERROR_ALREADY_EXISTS = 183
        if last_error == ERROR_ALREADY_EXISTS:
            try:
                ctypes.windll.user32.MessageBoxW(
                    0,
                    "悬浮时钟天气已经在运行中。\n\n"
                    "请在任务栏托盘区查找图标，或将鼠标移到屏幕右侧寻找悬浮窗。",
                    "软件已在运行",
                    0x00000040,
                )
            except Exception:
                pass
            try:
                ctypes.windll.kernel32.CloseHandle(mutex_handle)
            except Exception:
                pass
            sys.exit(0)
    except Exception as e:
        try:
            print(f"[启动] 单实例检测失败: {e}")
        except Exception:
            pass

    win = tk.Tk()
    app = FloatingWeatherClock(win)
    win.mainloop()

    try:
        if mutex_handle:
            ctypes.windll.kernel32.ReleaseMutex(mutex_handle)
            ctypes.windll.kernel32.CloseHandle(mutex_handle)
    except Exception:
        pass