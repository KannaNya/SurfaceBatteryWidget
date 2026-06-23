import ctypes
import os
import sys
import threading
import time
from ctypes import wintypes
from pathlib import Path

import win32api
import win32com.client
import win32con
import win32gui
import winreg
from PIL import Image, ImageDraw, ImageFont


BASE_DIR = Path(__file__).resolve().parent
LOG_PATH = BASE_DIR / "SurfaceBatteryWidgetV10.log"
START_CMD = BASE_DIR / "Start_SurfaceBatteryWidgetV10.cmd"
APP_NAME = "SurfaceBatteryWidgetV10"
MUTEX_NAME = "Global\\SurfaceBatteryWidgetV10"

LOGICAL_WIDTH = 72
LOGICAL_HEIGHT = 72
RIGHT_MARGIN = 6
BOTTOM_MARGIN = 104

TIMER_ID = 1
TIMER_MS = 1000
WM_APP_UPDATE = win32con.WM_APP + 10

MENU_TOGGLE_STARTUP = 1001
MENU_TOGGLE_DRAG = 1002
MENU_RELOCK = 1003
MENU_EXIT = 1004

ULW_ALPHA = 0x00000002
AC_SRC_OVER = 0x00
AC_SRC_ALPHA = 0x01
DIB_RGB_COLORS = 0
BI_RGB = 0

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
shcore = ctypes.WinDLL("shcore", use_last_error=True)

mutex = None


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_ubyte),
        ("BlendFlags", ctypes.c_ubyte),
        ("SourceConstantAlpha", ctypes.c_ubyte),
        ("AlphaFormat", ctypes.c_ubyte),
    ]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC,
    ctypes.POINTER(BITMAPINFO),
    wintypes.UINT,
    ctypes.POINTER(ctypes.c_void_p),
    wintypes.HANDLE,
    wintypes.DWORD,
]
gdi32.CreateDIBSection.restype = wintypes.HBITMAP
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL

user32.UpdateLayeredWindow.argtypes = [
    wintypes.HWND,
    wintypes.HDC,
    ctypes.POINTER(POINT),
    ctypes.POINTER(SIZE),
    wintypes.HDC,
    ctypes.POINTER(POINT),
    wintypes.COLORREF,
    ctypes.POINTER(BLENDFUNCTION),
    wintypes.DWORD,
]
user32.UpdateLayeredWindow.restype = wintypes.BOOL
UINT_PTR = ctypes.c_size_t
user32.SetTimer.argtypes = [wintypes.HWND, UINT_PTR, wintypes.UINT, wintypes.LPVOID]
user32.SetTimer.restype = UINT_PTR
user32.KillTimer.argtypes = [wintypes.HWND, UINT_PTR]
user32.KillTimer.restype = wintypes.BOOL


def log(message: str) -> None:
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except Exception:
        pass


def enable_dpi_awareness() -> None:
    try:
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass


def acquire_single_instance() -> bool:
    global mutex
    mutex = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    if not mutex:
        return False
    return ctypes.get_last_error() != 183


class PdhPowerMeter:
    PDH_FMT_DOUBLE = 0x00000200

    class PDH_FMT_COUNTERVALUE(ctypes.Structure):
        _fields_ = [("CStatus", wintypes.DWORD), ("doubleValue", ctypes.c_double)]

    def __init__(self) -> None:
        self.pdh = ctypes.WinDLL("pdh.dll")
        self.query = wintypes.HANDLE()
        self.counter = wintypes.HANDLE()
        self.ready = False
        self._init_counter()

    def _init_counter(self) -> None:
        for path in (r"\Power Meter(_Total)\Power", r"\Power Meter(power meter (0))\Power"):
            query = wintypes.HANDLE()
            counter = wintypes.HANDLE()
            if self.pdh.PdhOpenQueryW(None, 0, ctypes.byref(query)) != 0:
                continue
            add_counter = getattr(self.pdh, "PdhAddEnglishCounterW", self.pdh.PdhAddCounterW)
            if add_counter(query, path, 0, ctypes.byref(counter)) == 0:
                self.query = query
                self.counter = counter
                self.ready = True
                self.pdh.PdhCollectQueryData(self.query)
                log(f"PDH initialized: {path}")
                return
            self.pdh.PdhCloseQuery(query)
        log("PDH initialization failed")

    def read_mw(self) -> float | None:
        if not self.ready:
            self._init_counter()
            if not self.ready:
                return None
        try:
            self.pdh.PdhCollectQueryData(self.query)
            value = self.PDH_FMT_COUNTERVALUE()
            value_type = wintypes.DWORD()
            status = self.pdh.PdhGetFormattedCounterValue(
                self.counter,
                self.PDH_FMT_DOUBLE,
                ctypes.byref(value_type),
                ctypes.byref(value),
            )
            if status == 0 and value.CStatus == 0 and value.doubleValue > 0:
                return float(value.doubleValue)
        except Exception as exc:
            log(f"PDH read failed: {exc}")
            self.close()
        return None

    def close(self) -> None:
        try:
            if self.query:
                self.pdh.PdhCloseQuery(self.query)
        except Exception:
            pass
        self.query = wintypes.HANDLE()
        self.counter = wintypes.HANDLE()
        self.ready = False


class BatteryReader:
    def __init__(self) -> None:
        self.wmi_default = win32com.client.GetObject("winmgmts:")
        self.wmi_battery = win32com.client.GetObject(r"winmgmts:\\.\root\wmi")
        self.snapshot = {"percent": None, "remaining_wh": None, "online": False}

    def refresh(self) -> dict:
        try:
            battery = list(self.wmi_default.InstancesOf("Win32_Battery"))
            statuses = list(self.wmi_battery.InstancesOf("BatteryStatus"))
            b = battery[0] if battery else None
            s = statuses[0] if statuses else None
            self.snapshot = {
                "percent": int(b.EstimatedChargeRemaining) if b is not None else None,
                "remaining_wh": float(s.RemainingCapacity) / 1000.0 if s is not None else None,
                "online": bool(s.PowerOnline) if s is not None else False,
            }
        except Exception as exc:
            log(f"WMI battery read failed: {exc}")
        return self.snapshot


def startup_dir() -> Path:
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def cleanup_old_startup() -> None:
    for name in (
        "Surface Battery Widget.lnk",
        "Surface Battery Widget V2.lnk",
        "Surface Battery Widget V3.lnk",
        "Surface Battery Widget V4.lnk",
        "Surface Battery Widget V5.lnk",
        "Surface Battery Widget V6.lnk",
        "Surface Battery Widget V7.lnk",
    ):
        try:
            (startup_dir() / name).unlink(missing_ok=True)
        except Exception:
            pass
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            for name in (
                "SurfaceBatteryWidget",
                "SurfaceBatteryWidgetV2",
                "SurfaceBatteryWidgetV3",
                "SurfaceBatteryWidgetV4",
                "SurfaceBatteryWidgetV5",
                "SurfaceBatteryWidgetV6",
                "SurfaceBatteryWidgetV7",
            ):
                try:
                    winreg.DeleteValue(key, name)
                except FileNotFoundError:
                    pass
    except Exception as exc:
        log(f"Old startup cleanup failed: {exc}")


def enable_startup() -> None:
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    exe = pythonw if pythonw.exists() else Path(sys.executable)
    command = f'"{exe}" "{Path(__file__).resolve()}"'
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        log("HKCU Run enabled")
    except Exception as exc:
        log(f"HKCU Run failed: {exc}")

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(startup_dir() / "Surface Battery Widget V10.lnk"))
        shortcut.TargetPath = str(START_CMD)
        shortcut.WorkingDirectory = str(BASE_DIR)
        shortcut.Description = "Surface Battery Widget V10"
        shortcut.Save()
        log("Startup shortcut enabled")
    except Exception as exc:
        log(f"Startup shortcut failed: {exc}")


def disable_startup() -> None:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass
    except Exception:
        pass
    try:
        (startup_dir() / "Surface Battery Widget V10.lnk").unlink(missing_ok=True)
    except Exception:
        pass


def startup_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as key:
            winreg.QueryValueEx(key, APP_NAME)
            return True
    except Exception:
        return (startup_dir() / "Surface Battery Widget V10.lnk").exists()


def format_eta(hours: float | None) -> str:
    if hours is None or hours <= 0:
        return "--"
    minutes = round(hours * 60)
    if minutes >= 600:
        return "10h+"
    return f"{minutes}m"


def eta_lines(eta: str) -> list[str]:
    if eta.endswith("m") and eta[:-1].isdigit():
        return [eta[:-1], "min"]
    return [eta]


FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"


def load_font(name: str, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for filename in (name, "seguisb.ttf", "segoeuib.ttf", "segoeui.ttf", "msyh.ttc", "arial.ttf"):
        path = FONT_DIR / filename
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> int:
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0]


def fit_font(draw: ImageDraw.ImageDraw, text: str, font_name: str, start_size: int, min_size: int, max_width: int):
    font = load_font(font_name, start_size)
    size = start_size
    while text_width(draw, text, font) > max_width and size > min_size:
        size -= 1
        font = load_font(font_name, size)
    return font


def fit_shared_font(
    draw: ImageDraw.ImageDraw,
    texts: list[str],
    font_name: str,
    start_size: int,
    min_size: int,
    max_width: int,
):
    font = load_font(font_name, start_size)
    size = start_size
    while any(text_width(draw, text, font) > max_width for text in texts) and size > min_size:
        size -= 1
        font = load_font(font_name, size)
    return font


def centered_text(draw: ImageDraw.ImageDraw, width: int, y: int, text: str, font, fill: str) -> None:
    draw.text(((width - text_width(draw, text, font)) // 2, y), text, font=font, fill=fill)


def render_widget_image(eta: str, watts: float | None, dpi: int = 96) -> Image.Image:
    scale = dpi / 96.0
    width = max(1, round(LOGICAL_WIDTH * scale))
    height = max(1, round(LOGICAL_HEIGHT * scale))
    ss = 4

    def p(value: float) -> int:
        return round(value * scale)

    def ps(value: float) -> int:
        return round(value * scale * ss)

    bg = Image.new("RGBA", (width * ss, height * ss), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bg)
    bd.rounded_rectangle(
        (ps(0), ps(0), ps(LOGICAL_WIDTH - 1), ps(LOGICAL_HEIGHT - 3)),
        radius=ps(15),
        fill="#20252b",
        outline=(255, 255, 255, 24),
        width=max(1, ps(1)),
    )
    img = bg.resize((width, height), Image.Resampling.LANCZOS)

    draw = ImageDraw.Draw(img)
    lines = eta_lines(eta)
    if len(lines) == 2 and lines[1] == "min":
        eta_text = f"{lines[0]} min"
        eta_font = fit_font(draw, eta_text, "seguisb.ttf", p(16), p(12), p(60))
        centered_text(draw, width, p(18), eta_text, eta_font, "#f7fbff")
    else:
        eta_font = fit_font(draw, eta, "seguisb.ttf", p(17), p(10), p(54))
        centered_text(draw, width, p(18), eta, eta_font, "#f7fbff")

    draw.line((p(22), p(46), p(LOGICAL_WIDTH - 22), p(46)), fill=(255, 255, 255, 96), width=max(1, p(1)))

    power_text = "--" if watts is None else f"{watts:.1f}W"
    power_font = fit_font(draw, power_text, "seguisb.ttf", p(10), p(8), p(52))
    centered_text(draw, width, p(51), power_text, power_font, "#f1f4f7")
    return img


def premultiply_bgra(image: Image.Image) -> bytes:
    rgba = image.convert("RGBA").tobytes()
    out = bytearray(len(rgba))
    for i in range(0, len(rgba), 4):
        r = rgba[i]
        g = rgba[i + 1]
        b = rgba[i + 2]
        a = rgba[i + 3]
        out[i] = (b * a + 127) // 255
        out[i + 1] = (g * a + 127) // 255
        out[i + 2] = (r * a + 127) // 255
        out[i + 3] = a
    return bytes(out)


class SurfaceBatteryWidget:
    def __init__(self) -> None:
        cleanup_old_startup()
        enable_startup()
        self.power = PdhPowerMeter()
        self.battery = BatteryReader()
        self.tick = 0
        self.allow_drag = False
        self.dpi = 96
        self.width = LOGICAL_WIDTH
        self.height = LOGICAL_HEIGHT
        self.class_name = "SurfaceBatteryWidgetV10Window"
        self.hinst = win32api.GetModuleHandle(None)
        self.hwnd = None
        self.last_eta = "--"
        self.last_watts = None
        self.running = False
        self.last_menu_time = 0.0
        self._register_window()
        self._create_window()

    def _register_window(self) -> None:
        wc = win32gui.WNDCLASS()
        wc.hInstance = self.hinst
        wc.lpszClassName = self.class_name
        wc.lpfnWndProc = self.wnd_proc
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        try:
            win32gui.RegisterClass(wc)
        except win32gui.error as exc:
            if exc.winerror != win32con.ERROR_CLASS_ALREADY_EXISTS:
                raise

    def _create_window(self) -> None:
        ex_style = win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW
        self.hwnd = win32gui.CreateWindowEx(
            ex_style,
            self.class_name,
            "Battery V10",
            win32con.WS_POPUP,
            0,
            0,
            self.width,
            self.height,
            0,
            0,
            self.hinst,
            None,
        )
        self.refresh_dpi()
        self.lock_position()
        self.update_data()
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOWNOACTIVATE)
        self.running = True
        self.start_update_worker()
        log("Widget shown")

    def start_update_worker(self) -> None:
        thread = threading.Thread(target=self.update_worker, name="SurfaceBatteryWidgetV10Update", daemon=True)
        thread.start()

    def update_worker(self) -> None:
        while self.running:
            time.sleep(TIMER_MS / 1000)
            try:
                if self.hwnd and win32gui.IsWindow(self.hwnd):
                    win32gui.PostMessage(self.hwnd, WM_APP_UPDATE, 0, 0)
                else:
                    break
            except Exception as exc:
                log(f"Update worker failed: {exc!r}")
                break

    def refresh_dpi(self) -> None:
        try:
            self.dpi = int(user32.GetDpiForWindow(self.hwnd))
        except Exception:
            try:
                self.dpi = int(user32.GetDpiForSystem())
            except Exception:
                self.dpi = 96
        scale = self.dpi / 96.0
        self.width = max(1, round(LOGICAL_WIDTH * scale))
        self.height = max(1, round(LOGICAL_HEIGHT * scale))

    def px(self, value: float) -> int:
        return round(value * self.dpi / 96.0)

    def lock_position(self) -> None:
        monitor = win32api.MonitorFromWindow(self.hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        work = win32api.GetMonitorInfo(monitor)["Work"]
        left, top, right, bottom = work
        x = right - self.width - self.px(RIGHT_MARGIN)
        y = bottom - self.height - self.px(BOTTOM_MARGIN)
        if y < top + self.px(24):
            y = top + max(0, (bottom - top - self.height) // 2)
        win32gui.SetWindowPos(
            self.hwnd,
            win32con.HWND_TOPMOST,
            x,
            y,
            self.width,
            self.height,
            win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW,
        )
        log(f"Position locked: {x},{y},{self.width},{self.height},dpi={self.dpi}")

    def update_data(self) -> None:
        self.tick += 1
        if self.tick == 1 or self.tick % 5 == 0:
            snap = self.battery.refresh()
        else:
            snap = self.battery.snapshot
        mw = self.power.read_mw()
        watts = mw / 1000.0 if mw else None
        remaining = snap.get("remaining_wh")
        if snap.get("online"):
            eta = "AC"
        elif watts and remaining:
            eta = format_eta(remaining / watts)
        else:
            eta = "--"
        self.last_eta = eta
        self.last_watts = watts
        self.render()
        if self.tick <= 5 or self.tick % 30 == 0:
            watts_text = "--" if watts is None else f"{watts:.1f}W"
            log(f"Tick {self.tick}: eta={eta}, watts={watts_text}")

    def render(self) -> None:
        image = render_widget_image(self.last_eta, self.last_watts, self.dpi)
        self.width, self.height = image.size
        self.update_layered_window(image)

    def update_layered_window(self, image: Image.Image) -> None:
        width, height = image.size
        bgra = premultiply_bgra(image)
        screen_dc = user32.GetDC(None)
        mem_dc = gdi32.CreateCompatibleDC(screen_dc)
        bits = ctypes.c_void_p()
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB
        bitmap = gdi32.CreateDIBSection(screen_dc, ctypes.byref(bmi), DIB_RGB_COLORS, ctypes.byref(bits), None, 0)
        if not bitmap:
            gdi32.DeleteDC(mem_dc)
            user32.ReleaseDC(None, screen_dc)
            raise ctypes.WinError(ctypes.get_last_error())
        old_bitmap = gdi32.SelectObject(mem_dc, bitmap)
        try:
            ctypes.memmove(bits, bgra, len(bgra))
            x, y, _, _ = win32gui.GetWindowRect(self.hwnd)
            dst = POINT(x, y)
            size = SIZE(width, height)
            src = POINT(0, 0)
            blend = BLENDFUNCTION(AC_SRC_OVER, 0, 255, AC_SRC_ALPHA)
            if not user32.UpdateLayeredWindow(
                self.hwnd,
                screen_dc,
                ctypes.byref(dst),
                ctypes.byref(size),
                mem_dc,
                ctypes.byref(src),
                0,
                ctypes.byref(blend),
                ULW_ALPHA,
            ):
                raise ctypes.WinError(ctypes.get_last_error())
        finally:
            gdi32.SelectObject(mem_dc, old_bitmap)
            gdi32.DeleteObject(bitmap)
            gdi32.DeleteDC(mem_dc)
            user32.ReleaseDC(None, screen_dc)

    def show_context_menu(self) -> None:
        now = time.monotonic()
        if now - self.last_menu_time < 0.35:
            return
        self.last_menu_time = now
        log("Context menu requested")
        menu = win32gui.CreatePopupMenu()
        try:
            win32gui.AppendMenu(
                menu,
                win32con.MF_STRING,
                MENU_TOGGLE_STARTUP,
                "Disable startup" if startup_enabled() else "Enable startup",
            )
            win32gui.AppendMenu(
                menu,
                win32con.MF_STRING,
                MENU_TOGGLE_DRAG,
                "Lock to right edge" if self.allow_drag else "Unlock drag",
            )
            win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_RELOCK, "Re-lock to right edge")
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
            win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_EXIT, "Exit")
            x, y = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.hwnd)
            command = win32gui.TrackPopupMenu(
                menu,
                win32con.TPM_RETURNCMD | win32con.TPM_RIGHTBUTTON | win32con.TPM_LEFTALIGN,
                x,
                y,
                0,
                self.hwnd,
                None,
            )
        finally:
            win32gui.DestroyMenu(menu)
            try:
                win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
            except Exception:
                pass
        log(f"Context menu command: {command}")
        if command == MENU_TOGGLE_STARTUP:
            if startup_enabled():
                disable_startup()
            else:
                enable_startup()
        elif command == MENU_TOGGLE_DRAG:
            self.allow_drag = not self.allow_drag
            if not self.allow_drag:
                self.lock_position()
        elif command == MENU_RELOCK:
            self.allow_drag = False
            self.lock_position()
        elif command == MENU_EXIT:
            win32gui.DestroyWindow(self.hwnd)

    def close(self) -> None:
        self.running = False
        try:
            user32.KillTimer(self.hwnd, TIMER_ID)
        except Exception:
            pass
        try:
            self.power.close()
        except Exception:
            pass
        try:
            kernel32.ReleaseMutex(mutex)
            kernel32.CloseHandle(mutex)
        except Exception:
            pass
        log("Widget closed")

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_APP_UPDATE or (msg == win32con.WM_TIMER and wparam == TIMER_ID):
            self.update_data()
            return 0
        if msg == win32con.WM_NCHITTEST:
            return win32con.HTCLIENT
        if msg in (win32con.WM_CONTEXTMENU, win32con.WM_RBUTTONDOWN, win32con.WM_RBUTTONUP):
            self.show_context_menu()
            return 0
        if msg == win32con.WM_LBUTTONDOWN:
            if self.allow_drag:
                win32gui.ReleaseCapture()
                win32gui.SendMessage(hwnd, win32con.WM_NCLBUTTONDOWN, win32con.HTCAPTION, 0)
            return 0
        if msg == win32con.WM_DPICHANGED:
            self.refresh_dpi()
            self.lock_position()
            self.render()
            return 0
        if msg == win32con.WM_DISPLAYCHANGE:
            self.refresh_dpi()
            self.lock_position()
            self.render()
            return 0
        if msg == win32con.WM_DESTROY:
            self.close()
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def run(self) -> None:
        win32gui.PumpMessages()


def main() -> None:
    enable_dpi_awareness()
    if not acquire_single_instance():
        return
    SurfaceBatteryWidget().run()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"Fatal error: {exc!r}")
        raise
