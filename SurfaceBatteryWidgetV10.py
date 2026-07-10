import ctypes
import csv
import os
import sys
import threading
import time
from collections import defaultdict
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
DIARY_DIR = BASE_DIR / "power-diary"
START_CMD = BASE_DIR / "Start_SurfaceBatteryWidgetV10.cmd"
APP_NAME = "SurfaceBatteryWidgetV10"
MUTEX_NAME = "Global\\SurfaceBatteryWidgetV10"
CONFIG_KEY = r"Software\SurfaceBatteryWidget"
STARTUP_PREFERENCE_VALUE = "StartupEnabled"

LOGICAL_WIDTH = 95
LOGICAL_HEIGHT = 30
RIGHT_MARGIN = 89
EDGE_PAD = 1
BOTTOM_MARGIN = -1
CARD_RADIUS = 4
# Match Win11 taskbar typography: Segoe UI Variable Text, 12 DIP, Regular.
UI_FONT_SIZE_DIP = 12.0
UI_FONT_WEIGHT = 400
UI_FONT_OPTICAL_SIZE = 12

TIMER_MS = 1000
WM_APP_UPDATE = win32con.WM_APP + 10
WM_DPICHANGED = 0x02E0
WM_POWERBROADCAST = 0x0218
WM_SETTINGCHANGE = 0x001A
PBT_APMRESUMEAUTOMATIC = 0x0012
PBT_APMRESUMESUSPEND = 0x0007
PBT_POWERSETTINGCHANGE = 0x8013
RESUME_RELOCK_SECONDS = 8.0
POWER_DIARY_INTERVAL = 10.0
POWER_DIARY_SUMMARY_INTERVAL = 60.0
POWER_DIARY_TOP_N = 12
MAX_REASONABLE_BATTERY_WATTS = 180.0
ENERGY_SAVER_STATUS_INTERVAL = 60.0
MAX_LOG_BYTES = 2 * 1024 * 1024
MAX_DIARY_BYTES = 8 * 1024 * 1024

MENU_TOGGLE_STARTUP = 1001
MENU_TOGGLE_DRAG = 1002
MENU_RELOCK = 1003
MENU_EXIT = 1004
MENU_OPEN_DIARY_SUMMARY = 1005
MENU_OPEN_DIARY_FOLDER = 1006

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


class SYSTEM_POWER_STATUS(ctypes.Structure):
    _fields_ = [
        ("ACLineStatus", ctypes.c_ubyte),
        ("BatteryFlag", ctypes.c_ubyte),
        ("BatteryLifePercent", ctypes.c_ubyte),
        ("SystemStatusFlag", ctypes.c_ubyte),
        ("BatteryLifeTime", wintypes.DWORD),
        ("BatteryFullLifeTime", wintypes.DWORD),
    ]


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
kernel32.GetSystemPowerStatus.argtypes = [ctypes.POINTER(SYSTEM_POWER_STATUS)]
kernel32.GetSystemPowerStatus.restype = wintypes.BOOL


def rotate_file(path: Path, max_bytes: int) -> None:
    if not path.exists() or path.stat().st_size < max_bytes:
        return
    backup = path.with_name(path.name + ".1")
    backup.unlink(missing_ok=True)
    path.replace(backup)


def log(message: str) -> None:
    try:
        rotate_file(LOG_PATH, MAX_LOG_BYTES)
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


def valid_battery_watts(value) -> float | None:
    try:
        watts = float(value) / 1000.0
    except Exception:
        return None
    if 0 < watts <= MAX_REASONABLE_BATTERY_WATTS:
        return watts
    return None


class BatteryReader:
    def __init__(self) -> None:
        self.wmi_default = win32com.client.GetObject("winmgmts:")
        self.wmi_battery = win32com.client.GetObject(r"winmgmts:\\.\root\wmi")
        self.snapshot = {
            "percent": None,
            "remaining_wh": None,
            "online": False,
            "charging": False,
            "charge_rate_w": None,
            "full_charge_wh": None,
        }

    def refresh(self) -> dict:
        try:
            battery = list(self.wmi_default.InstancesOf("Win32_Battery"))
            statuses = list(self.wmi_battery.InstancesOf("BatteryStatus"))
            full = list(self.wmi_battery.InstancesOf("BatteryFullChargedCapacity"))
            b = battery[0] if battery else None
            s = statuses[0] if statuses else None
            full_capacity = full[0] if full else None
            charge_rate = valid_battery_watts(s.ChargeRate) if s is not None else None
            self.snapshot = {
                "percent": int(b.EstimatedChargeRemaining) if b is not None else None,
                "remaining_wh": float(s.RemainingCapacity) / 1000.0 if s is not None else None,
                "online": bool(s.PowerOnline) if s is not None else False,
                "charging": bool(s.Charging) if s is not None else False,
                "charge_rate_w": charge_rate,
                "full_charge_wh": (
                    float(full_capacity.FullChargedCapacity) / 1000.0 if full_capacity is not None else None
                ),
            }
        except Exception as exc:
            log(f"WMI battery read failed: {exc}")
        return self.snapshot


class ThermalReader:
    def __init__(self) -> None:
        self.wmi = win32com.client.GetObject("winmgmts:")
        self.temperature_c: float | None = None

    def refresh(self) -> float | None:
        try:
            rows = self.wmi.ExecQuery(
                "SELECT Temperature,HighPrecisionTemperature "
                "FROM Win32_PerfFormattedData_Counters_ThermalZoneInformation"
            )
            temperatures = []
            for row in rows:
                high_precision = float(row.HighPrecisionTemperature or 0)
                kelvin = high_precision / 10.0 if high_precision >= 1000 else float(row.Temperature or 0)
                temperature_c = kelvin - 273.15
                if 0.0 < temperature_c < 120.0:
                    temperatures.append(temperature_c)
            if temperatures:
                self.temperature_c = max(temperatures)
        except Exception as exc:
            log(f"Thermal-zone read failed: {exc}")
        return self.temperature_c


def startup_dir() -> Path:
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def startup_shortcut_path() -> Path:
    return startup_dir() / "Surface Battery Widget V10.lnk"


def cleanup_old_startup() -> None:
    for name in (
        "Surface Battery Widget.lnk",
        "Surface Battery Widget V2.lnk",
        "Surface Battery Widget V3.lnk",
        "Surface Battery Widget V4.lnk",
        "Surface Battery Widget V5.lnk",
        "Surface Battery Widget V6.lnk",
        "Surface Battery Widget V7.lnk",
        "Surface Battery Widget V8.lnk",
        "Surface Battery Widget V9.lnk",
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
                "SurfaceBatteryWidgetV8",
                "SurfaceBatteryWidgetV9",
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
    run_enabled = False
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        run_enabled = True
        log("HKCU Run enabled")
    except Exception as exc:
        log(f"HKCU Run failed: {exc}")

    if run_enabled:
        try:
            startup_shortcut_path().unlink(missing_ok=True)
        except Exception as exc:
            log(f"Startup shortcut cleanup failed: {exc}")
        return

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(startup_shortcut_path()))
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
        startup_shortcut_path().unlink(missing_ok=True)
    except Exception:
        pass


def startup_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as key:
            winreg.QueryValueEx(key, APP_NAME)
            return True
    except Exception:
        return startup_shortcut_path().exists()


def read_startup_preference() -> bool | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CONFIG_KEY) as key:
            value, _ = winreg.QueryValueEx(key, STARTUP_PREFERENCE_VALUE)
            return bool(value)
    except FileNotFoundError:
        return None
    except Exception as exc:
        log(f"Startup preference read failed: {exc}")
        return None


def write_startup_preference(enabled: bool) -> None:
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, CONFIG_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, STARTUP_PREFERENCE_VALUE, 0, winreg.REG_DWORD, int(enabled))
    except Exception as exc:
        log(f"Startup preference write failed: {exc}")


def set_startup_enabled(enabled: bool) -> None:
    write_startup_preference(enabled)
    if enabled:
        enable_startup()
    else:
        disable_startup()


def initialize_startup() -> None:
    preference = read_startup_preference()
    if preference is None:
        preference = True
        write_startup_preference(preference)
    if preference:
        enable_startup()
    else:
        disable_startup()


def format_eta(hours: float | None) -> str:
    if hours is None or hours <= 0:
        return "--"
    minutes = round(hours * 60)
    if minutes >= 600:
        return "600+"
    return str(minutes)


def format_charge_eta(remaining_wh: float | None, full_charge_wh: float | None, charge_rate_w: float | None) -> str:
    if not remaining_wh or not full_charge_wh or not charge_rate_w or charge_rate_w <= 0:
        return "AC"
    needed_wh = max(0.0, full_charge_wh - remaining_wh)
    if needed_wh <= 0.15:
        return "FULL"
    return format_eta(needed_wh / charge_rate_w)


FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"


def load_font(
    name: str,
    size: int,
    weight: int = 400,
    optical_size: int = 13,
) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for filename in (name, "seguisb.ttf", "segoeuib.ttf", "segoeui.ttf", "msyh.ttc", "arial.ttf"):
        path = FONT_DIR / filename
        if path.exists():
            font = ImageFont.truetype(str(path), size)
            if filename.lower() == "seguivar.ttf":
                try:
                    font.set_variation_by_axes([weight, optical_size])
                except (AttributeError, OSError):
                    pass
            return font
    return ImageFont.load_default()


def system_uses_light_theme() -> bool:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "SystemUsesLightTheme")
            return bool(value)
    except Exception:
        return False


def render_widget_image(
    eta_text: str,
    watts_text: str,
    dpi: int = 96,
    charging: bool = False,
    temperature_text: str | None = None,
) -> Image.Image:
    scale = dpi / 96.0
    width = max(1, round(LOGICAL_WIDTH * scale))
    height = max(1, round(LOGICAL_HEIGHT * scale))
    pad = max(1, round(EDGE_PAD * scale))
    card_w = width - pad * 2
    card_h = height - pad * 2
    radius = max(1, round(CARD_RADIUS * scale))
    ss = 4

    light_theme = system_uses_light_theme()
    if light_theme:
        card_fill = (243, 243, 243, 255)
        border_fill = (0, 0, 0, 24)
        primary = (26, 26, 26, 255)
        secondary = (92, 92, 92, 255)
    else:
        # Sampled from this machine's translucent Win11 taskbar surface.
        card_fill = (40, 41, 44, 255)
        border_fill = (255, 255, 255, 18)
        primary = (255, 255, 255, 255)
        secondary = (166, 166, 166, 255)

    # Match Win11 taskbar surfaces: one translucent layer, a hairline border,
    # and no simulated bevel, painted gradient, or heavy shadow.
    img = Image.new("RGBA", (width * ss, height * ss), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    cr = (pad * ss, pad * ss, (pad + card_w) * ss - 1, (pad + card_h) * ss - 1)
    draw.rounded_rectangle(
        cr,
        radius=radius * ss,
        fill=card_fill,
        outline=border_fill,
        width=max(1, round(0.5 * scale * ss)),
    )
    
    # Parse remaining minutes for status coloring
    minutes = None
    if eta_text not in ("AC", "--"):
        try:
            if eta_text.endswith("+"):
                minutes = int(eta_text[:-1])
            else:
                minutes = int(eta_text)
        except ValueError:
            pass

    # Status color logic. Charging ETA is positive information, so warning
    # colors only apply while the battery is discharging.
    if charging or eta_text == "AC":
        status_color = primary
    elif minutes is not None:
        if minutes < 20:
            status_color = "#ff6b72"  # Fluent Soft Red
        elif minutes < 45:
            status_color = "#ff9d5c"  # Fluent Soft Orange
        else:
            status_color = primary
    else:
        status_color = secondary

    temperature_color = secondary
    if temperature_text and temperature_text != "--":
        try:
            temperature_value = float(temperature_text.rstrip("°C"))
            if temperature_value >= 60:
                temperature_color = "#ff6b72"
            elif temperature_value >= 50:
                temperature_color = "#ff9d5c"
            else:
                temperature_color = primary
        except ValueError:
            pass

    # Match the visual weight of Win11 taskbar status text, with tighter labels.
    val_size = round(UI_FONT_SIZE_DIP * scale * ss)
    unit_size = val_size
    max_width = card_w * ss - round(4 * scale * ss)

    # Scale down sizes if text runs too wide (adaptive fitting)
    for size_reduce in range(0, 5):
        font_val = load_font(
            "SegUIVar.ttf",
            val_size - size_reduce * ss,
            UI_FONT_WEIGHT,
            UI_FONT_OPTICAL_SIZE,
        )
        font_unit = font_val
        
        segments = []
        
        # 1. ETA / AC Segment
        if eta_text == "AC":
            # Icon dimensions
            icon_w = round(6 * scale * ss)
            icon_h = round(10 * scale * ss)
            
            def make_draw_lightning(w, h):
                return lambda d, x, y: d.polygon([
                    (x + w * 0.65, y),
                    (x + w * 0.15, y + h * 0.55),
                    (x + w * 0.55, y + h * 0.55),
                    (x + w * 0.35, y + h),
                    (x + w * 0.85, y + h * 0.45),
                    (x + w * 0.45, y + h * 0.45)
                ], fill=primary)

            segments.append({
                "type": "icon",
                "width": icon_w + round(3 * scale * ss),
                "height": icon_h,
                "draw_fn": make_draw_lightning(icon_w, icon_h)
            })
            segments.append({
                "type": "text",
                "text": "AC",
                "font": font_val,
                "color": primary
            })
        elif eta_text in ("--", "FULL"):
            segments.append({
                "type": "text",
                "text": eta_text,
                "font": font_val,
                "color": secondary
            })
        else:
            clock_size = round(8 * scale * ss)
            clock_stroke = max(1, round(0.75 * scale * ss))

            def make_draw_clock(size, stroke, color):
                def draw_clock(d, x, y):
                    inset = max(1, stroke // 2)
                    center_x = x + size // 2
                    center_y = y + size // 2
                    d.ellipse(
                        (x + inset, y + inset, x + size - inset, y + size - inset),
                        outline=color,
                        width=stroke,
                    )
                    d.line(
                        (center_x, center_y, center_x, y + round(size * 0.27)),
                        fill=color,
                        width=stroke,
                    )
                    d.line(
                        (center_x, center_y, x + round(size * 0.72), y + round(size * 0.58)),
                        fill=color,
                        width=stroke,
                    )

                return draw_clock

            segments.append({
                "type": "icon",
                "width": clock_size + round(2.5 * scale * ss),
                "height": clock_size,
                "draw_fn": make_draw_clock(clock_size, clock_stroke, status_color),
            })
            segments.append({
                "type": "text",
                "text": eta_text,
                "font": font_val,
                "color": status_color
            })

        # 2. Native taskbar-style whitespace between the two values.
        segments.append({
            "type": "spacer",
            "width": round(3 * scale * ss),
        })

        if temperature_text:
            segments.append({
                "type": "text",
                "text": temperature_text,
                "font": font_val,
                "color": temperature_color,
            })
            segments.append({
                "type": "spacer",
                "width": round(3 * scale * ss),
            })

        # 3. Wattage Segment
        if watts_text == "--":
            segments.append({
                "type": "text",
                "text": "--",
                "font": font_val,
                "color": secondary
            })
        else:
            val = watts_text[:-1] if watts_text.endswith("W") else watts_text
            segments.append({
                "type": "text",
                "text": val,
                "font": font_val,
                "color": primary
            })
            segments.append({
                "type": "text",
                "text": "W",
                "font": font_unit,
                "color": primary
            })

        # Calculate width of all segments
        total_w = 0
        for seg in segments:
            if seg["type"] == "text":
                box = draw.textbbox((0, 0), seg["text"], font=seg["font"])
                seg["width"] = box[2] - box[0]
                seg["box_top_offset"] = box[1]
                seg["box_height"] = box[3] - box[1]
            total_w += seg["width"]
            
        if total_w <= max_width or size_reduce == 4:
            break

    # Draw segments aligned to center
    tx = pad * ss + (card_w * ss - total_w) // 2
    for seg in segments:
        if seg["type"] == "text":
            ty = pad * ss + (card_h * ss - seg["box_height"]) // 2 - seg["box_top_offset"]
            draw.text((tx, ty), seg["text"], font=seg["font"], fill=seg["color"])
        elif seg["type"] == "icon":
            iy = pad * ss + (card_h * ss - seg["height"]) // 2
            seg["draw_fn"](draw, tx, iy)
        tx += seg["width"]

    # Downsample everything together
    img = img.resize((width, height), Image.Resampling.LANCZOS)

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


def process_base_name(name: str) -> str:
    return name.split("#", 1)[0].strip() or name


def open_path(path: Path) -> None:
    try:
        if not path.exists():
            if path.suffix:
                path.parent.mkdir(exist_ok=True)
                path.write_text("No data yet. Keep the widget running for a minute.\n", encoding="utf-8")
            else:
                path.mkdir(exist_ok=True)
        os.startfile(str(path))
    except Exception as exc:
        log(f"Open path failed: {path}: {exc}")


def power_saving_recommendations(top_names: list[str]) -> list[str]:
    names = {name.lower() for name in top_names}
    tips = []
    if "chrome" in names or "msedge" in names:
        tips.append("Browser is active: close video/chat tabs, enable sleeping tabs, or use fewer tabs in class.")
    if "voicerecorder" in names or "audiodg" in names:
        tips.append("Audio stack is active: stop recording/playback when you do not need it.")
    if "msmpeng" in names:
        tips.append("Defender is active: avoid large downloads/unzips on battery; let scans finish while plugged in.")
    if "dwm" in names:
        tips.append("Desktop compositor is active: lower brightness, avoid full-screen animations, and use 60 Hz in class mode.")
    if "codex" in names:
        tips.append("Codex is active: pause heavy agent work when you need maximum battery life.")
    if "telegram" in names:
        tips.append("Messaging apps are active: mute/quit background messengers during lectures.")
    if not tips:
        tips.append("No clear app culprit yet; collect at least 30 minutes of samples during a real class session.")
    return tips


def read_energy_saver_threshold() -> int | None:
    try:
        schemes_key = r"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, schemes_key) as key:
            active_scheme = str(winreg.QueryValueEx(key, "ActivePowerScheme")[0])
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            schemes_key
            + rf"\{active_scheme}"
            + r"\de830923-a562-41af-a086-e3a2c6bad2da\e69653ca-cf7f-4f05-aa73-cb833fa90ad4",
        ) as key:
            return int(winreg.QueryValueEx(key, "DCSettingIndex")[0])
    except Exception:
        return None


def read_energy_saver_status() -> str:
    status = SYSTEM_POWER_STATUS()
    if kernel32.GetSystemPowerStatus(ctypes.byref(status)):
        return "On" if status.SystemStatusFlag else "Off"
    log(f"Energy saver status read failed: {ctypes.get_last_error()}")
    return "Unknown"


class PowerDiary:
    def __init__(self) -> None:
        self.ready = False
        self.wmi_perf = None
        self.own_pid = os.getpid()
        self.cpu_count = max(1, os.cpu_count() or 1)
        self.last_sample_time = 0.0
        self.last_summary_time = 0.0
        self.last_raw: dict[tuple[int, str], tuple[int, int, str]] = {}
        self.process_score = defaultdict(float)
        self.process_peak = defaultdict(float)
        self.power_seconds = 0.0
        self.power_sample_seconds = 0.0
        self.sample_count = 0
        self.energy_saver_status = "Unknown"
        self.energy_saver_threshold = read_energy_saver_threshold()
        self.last_energy_saver_check = 0.0
        self.power_path = DIARY_DIR / "power_samples_v2.csv"
        self.process_path = DIARY_DIR / "process_activity.csv"
        self.summary_path = DIARY_DIR / "summary.txt"
        try:
            DIARY_DIR.mkdir(exist_ok=True)
            self.wmi_perf = win32com.client.GetObject("winmgmts:")
            self.ready = True
            log("Power diary initialized")
        except Exception as exc:
            log(f"Power diary initialization failed: {exc}")

    def current_energy_saver_status(self) -> str:
        now = time.time()
        if now - self.last_energy_saver_check >= ENERGY_SAVER_STATUS_INTERVAL:
            self.energy_saver_status = read_energy_saver_status()
            self.energy_saver_threshold = read_energy_saver_threshold()
            self.last_energy_saver_check = now
        return self.energy_saver_status

    def append_csv(self, path: Path, header: list[str], row: list) -> None:
        self.append_csv_rows(path, header, [row])

    def append_csv_rows(self, path: Path, header: list[str], rows: list[list]) -> None:
        if not rows:
            return
        try:
            rotate_file(path, MAX_DIARY_BYTES)
            needs_header = not path.exists()
            with path.open("a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if needs_header:
                    writer.writerow(header)
                writer.writerows(rows)
        except Exception as exc:
            log(f"Power diary write failed: {exc}")

    def read_process_cpu(self, elapsed: float | None) -> list[dict]:
        if not self.ready or self.wmi_perf is None:
            return []
        try:
            rows = self.wmi_perf.ExecQuery(
                "SELECT IDProcess,Name,PercentProcessorTime,TimeStamp_Sys100NS "
                "FROM Win32_PerfRawData_PerfProc_Process"
            )
        except Exception as exc:
            log(f"Power diary process query failed: {exc}")
            return []

        current: dict[tuple[int, str], tuple[int, int, str]] = {}
        activity = []
        for row in rows:
            try:
                pid = int(row.IDProcess)
                instance_name = str(row.Name)
                if pid == 0 or pid == self.own_pid or instance_name in ("Idle", "_Total"):
                    continue
                base_name = process_base_name(instance_name)
                raw = int(row.PercentProcessorTime)
                timestamp = int(row.TimeStamp_Sys100NS)
            except Exception:
                continue

            key = (pid, instance_name)
            current[key] = (raw, timestamp, base_name)
            previous = self.last_raw.get(key)
            if not previous:
                continue
            prev_raw, prev_timestamp, _ = previous
            delta_raw = raw - prev_raw
            delta_timestamp = timestamp - prev_timestamp
            if delta_raw < 0 or delta_timestamp <= 0:
                continue

            cpu_percent = (delta_raw / delta_timestamp) * 100.0 / self.cpu_count
            if cpu_percent < 0.1:
                continue
            activity.append({"pid": pid, "name": base_name, "cpu_percent": cpu_percent})
            if elapsed:
                self.process_score[base_name] += cpu_percent * elapsed
                self.process_peak[base_name] = max(self.process_peak[base_name], cpu_percent)

        self.last_raw = current
        activity.sort(key=lambda item: item["cpu_percent"], reverse=True)
        return activity

    def sample(
        self,
        snap: dict,
        eta: str,
        system_watts: float | None,
        display_watts: float | None,
        charging: bool,
    ) -> None:
        now = time.time()
        if now - self.last_sample_time < POWER_DIARY_INTERVAL:
            return
        elapsed = now - self.last_sample_time if self.last_sample_time else None
        self.last_sample_time = now

        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        activity = self.read_process_cpu(elapsed)
        top_activity = activity[:POWER_DIARY_TOP_N]
        energy_saver_status = self.current_energy_saver_status()
        threshold = self.energy_saver_threshold
        energy_saver_expected = (
            bool(not snap.get("online") and threshold is not None and snap.get("percent") is not None)
            and int(snap.get("percent")) <= threshold
        )
        if elapsed and system_watts:
            self.power_seconds += system_watts * elapsed
            self.power_sample_seconds += elapsed
        self.sample_count += 1

        top_text = "; ".join(f"{item['name']}#{item['pid']}={item['cpu_percent']:.1f}%" for item in top_activity)
        self.append_csv(
            self.power_path,
            [
                "timestamp",
                "online",
                "charging",
                "battery_percent",
                "remaining_wh",
                "eta",
                "system_watts",
                "display_watts",
                "energy_saver_status",
                "energy_saver_threshold",
                "energy_saver_auto_expected",
                "top_processes_by_cpu",
            ],
            [
                timestamp,
                bool(snap.get("online")),
                charging,
                snap.get("percent"),
                snap.get("remaining_wh"),
                eta,
                "" if system_watts is None else f"{system_watts:.2f}",
                "" if display_watts is None else f"{display_watts:.2f}",
                energy_saver_status,
                "" if threshold is None else threshold,
                energy_saver_expected,
                top_text,
            ],
        )

        process_rows = []
        for item in top_activity:
            cpu_guess_w = ""
            if system_watts:
                cpu_guess_w = f"{system_watts * min(item['cpu_percent'], 100.0) / 100.0:.2f}"
            process_rows.append(
                [timestamp, item["pid"], item["name"], f"{item['cpu_percent']:.2f}", cpu_guess_w]
            )
        self.append_csv_rows(
            self.process_path,
            ["timestamp", "pid", "name", "cpu_percent", "cpu_weighted_w_guess"],
            process_rows,
        )

        if self.process_score and (
            self.last_summary_time == 0.0 or now - self.last_summary_time >= POWER_DIARY_SUMMARY_INTERVAL
        ):
            self.last_summary_time = now
            self.write_summary(timestamp)

    def write_summary(self, timestamp: str) -> None:
        top = sorted(self.process_score.items(), key=lambda item: item[1], reverse=True)[:15]
        avg_power = None
        if self.power_sample_seconds > 0:
            avg_power = self.power_seconds / self.power_sample_seconds

        lines = [
            "Surface Battery Widget power diary",
            f"Updated: {timestamp}",
            f"Samples: {self.sample_count}",
            f"Average system power: {'--' if avg_power is None else f'{avg_power:.1f} W'}",
            f"Energy saver status: {self.energy_saver_status}",
            f"Energy saver auto threshold: {'--' if self.energy_saver_threshold is None else str(self.energy_saver_threshold) + '%'}",
            "",
            "Top process activity since widget start:",
        ]
        if not top:
            lines.append("No process activity baseline yet; wait for another sample.")
        else:
            for index, (name, score) in enumerate(top, 1):
                peak = self.process_peak.get(name, 0.0)
                lines.append(f"{index:2d}. {name:<28} score={score:8.1f} peak_cpu={peak:5.1f}%")
        lines.extend(["", "Suggested power-saving actions:"])
        for tip in power_saving_recommendations([name for name, _ in top]):
            lines.append(f"- {tip}")
        lines.extend(
            [
                "",
                "Reading guide:",
                "Process scores are CPU activity over time, not exact per-app watts.",
                "Use this to find apps that correlate with high total power, then limit or close them.",
            ]
        )
        try:
            self.summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        except Exception as exc:
            log(f"Power diary summary write failed: {exc}")


class SurfaceBatteryWidget:
    def __init__(self) -> None:
        cleanup_old_startup()
        initialize_startup()
        self.power = PdhPowerMeter()
        self.battery = BatteryReader()
        self.thermal = ThermalReader()
        self.tick = 0
        self.allow_drag = False
        self.dpi = 96
        self.width = LOGICAL_WIDTH
        self.height = LOGICAL_HEIGHT
        self.diary = PowerDiary()
        self.class_name = "SurfaceBatteryWidgetV10Window"
        self.hinst = win32api.GetModuleHandle(None)
        self.hwnd = None
        self.last_eta = "--"
        self.last_watts = None
        self.last_temperature_c = None
        self.last_charging = False
        self.relock_until = 0.0
        self.last_update_wall = time.time()
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
        _, top, right, bottom = work
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

    def request_relock(self, reason: str, duration: float = RESUME_RELOCK_SECONDS) -> None:
        if self.allow_drag:
            return
        self.relock_until = max(self.relock_until, time.monotonic() + duration)
        log(f"Relock requested: {reason}")
        self.refresh_dpi()
        self.lock_position()
        self.render()

    def maybe_relock(self) -> None:
        if self.allow_drag or time.monotonic() > self.relock_until:
            return
        self.refresh_dpi()
        self.lock_position()

    def update_data(self) -> None:
        now_wall = time.time()
        if now_wall - self.last_update_wall > 10:
            self.request_relock(f"update gap {now_wall - self.last_update_wall:.1f}s", RESUME_RELOCK_SECONDS)
        self.last_update_wall = now_wall
        self.tick += 1
        if self.tick == 1 or self.tick % 5 == 0:
            snap = self.battery.refresh()
            temperature_c = self.thermal.refresh()
        else:
            snap = self.battery.snapshot
            temperature_c = self.thermal.temperature_c
        mw = self.power.read_mw()
        system_watts = mw / 1000.0 if mw else None
        watts = system_watts
        remaining = snap.get("remaining_wh")
        charging = False
        if snap.get("online"):
            charging = True
            eta = format_charge_eta(
                snap.get("remaining_wh"),
                snap.get("full_charge_wh"),
                snap.get("charge_rate_w"),
            )
            if snap.get("charging") and snap.get("charge_rate_w"):
                watts = snap.get("charge_rate_w")
        elif watts and remaining:
            eta = format_eta(remaining / watts)
        else:
            eta = "--"
        self.last_eta = eta
        self.last_watts = watts
        self.last_temperature_c = temperature_c
        self.last_charging = charging
        self.render()
        self.maybe_relock()
        try:
            self.diary.sample(snap, eta, system_watts, watts, charging)
        except Exception as exc:
            log(f"Power diary sample failed: {exc}")
        if self.tick <= 5 or self.tick % 30 == 0:
            watts_text = "--" if watts is None else f"{watts:.1f}W"
            temperature_text = "--" if temperature_c is None else f"{temperature_c:.1f}C"
            log(f"Tick {self.tick}: eta={eta}, temp={temperature_text}, watts={watts_text}")

    def render(self) -> None:
        watts_text = "--" if self.last_watts is None else f"{self.last_watts:.1f}W"
        temperature_text = "--" if self.last_temperature_c is None else f"{round(self.last_temperature_c):.0f}°"
        image = render_widget_image(
            self.last_eta,
            watts_text,
            self.dpi,
            self.last_charging,
            temperature_text,
        )
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
            win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_OPEN_DIARY_SUMMARY, "Open power summary")
            win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_OPEN_DIARY_FOLDER, "Open power diary folder")
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
            set_startup_enabled(not startup_enabled())
        elif command == MENU_TOGGLE_DRAG:
            self.allow_drag = not self.allow_drag
            if not self.allow_drag:
                self.lock_position()
        elif command == MENU_RELOCK:
            self.allow_drag = False
            self.lock_position()
        elif command == MENU_OPEN_DIARY_SUMMARY:
            self.diary.write_summary(time.strftime("%Y-%m-%d %H:%M:%S"))
            open_path(self.diary.summary_path)
        elif command == MENU_OPEN_DIARY_FOLDER:
            open_path(DIARY_DIR)
        elif command == MENU_EXIT:
            win32gui.DestroyWindow(self.hwnd)

    def close(self) -> None:
        self.running = False
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
        if msg == WM_APP_UPDATE:
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
        if msg == WM_DPICHANGED:
            self.request_relock("dpi changed", 3.0)
            return 0
        if msg == win32con.WM_DISPLAYCHANGE:
            self.request_relock("display changed", 5.0)
            return 0
        if msg == WM_SETTINGCHANGE:
            self.request_relock("system setting changed", 5.0)
            return 0
        if msg == WM_POWERBROADCAST:
            if wparam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND, PBT_POWERSETTINGCHANGE):
                self.request_relock(f"power broadcast {wparam}", RESUME_RELOCK_SECONDS)
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
