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
from PIL import Image, ImageDraw, ImageFilter, ImageFont


BASE_DIR = Path(__file__).resolve().parent
LOG_PATH = BASE_DIR / "SurfaceBatteryWidgetV10.log"
START_CMD = BASE_DIR / "Start_SurfaceBatteryWidgetV10.cmd"
APP_NAME = "SurfaceBatteryWidgetV10"
MUTEX_NAME = "Global\\SurfaceBatteryWidgetV10"

LOGICAL_WIDTH = 82
LOGICAL_HEIGHT = 30
RIGHT_MARGIN = 78
SHADOW_PAD = 1
BOTTOM_MARGIN = -2
CARD_RADIUS = 4

TIMER_ID = 1
TIMER_MS = 1000
WM_APP_UPDATE = win32con.WM_APP + 10
WM_DPICHANGED = 0x02E0

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
            charge_rate = float(s.ChargeRate) / 1000.0 if s is not None and int(s.ChargeRate) > 0 else None
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


def render_widget_image(eta_text: str, watts_text: str, dpi: int = 96, charging: bool = False) -> Image.Image:
    scale = dpi / 96.0
    width = max(1, round(LOGICAL_WIDTH * scale))
    height = max(1, round(LOGICAL_HEIGHT * scale))
    pad = max(1, round(SHADOW_PAD * scale))
    card_w = width - pad * 2
    card_h = height - pad * 2
    radius = max(1, round(CARD_RADIUS * scale))
    ss = 4

    # Shadow layer (offset 1px down for natural light direction)
    shadow = Image.new("RGBA", (width * ss, height * ss), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    shadow_offset = max(1, round(scale * ss))
    sr = (pad * ss, pad * ss + shadow_offset,
          (pad + card_w) * ss - 1, (pad + card_h) * ss - 1 + shadow_offset)
    sd.rounded_rectangle(sr, radius=radius * ss, fill=(0, 0, 0, 40))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=max(1, round(2.5 * scale * ss))))

    # Card layer: Mica dark gray theme with vertical gradient
    card = Image.new("RGBA", (width * ss, height * ss), (0, 0, 0, 0))
    
    # Create gradient fill image for the card area (Mica dark gray theme)
    gradient = Image.new("RGBA", (1, card_h * ss))
    for y in range(card_h * ss):
        # Subtle top-to-bottom dark theme gradient:
        # Top: (32, 32, 32, 225), Bottom: (24, 24, 24, 215)
        factor = y / max(1, card_h * ss - 1)
        r = int(32 - factor * 8)
        g = int(32 - factor * 8)
        b = int(32 - factor * 8)
        a = int(225 - factor * 10)
        gradient.putpixel((0, y), (r, g, b, a))
    gradient = gradient.resize((card_w * ss, card_h * ss), Image.Resampling.BILINEAR)

    # Rounded rectangle mask for the card shape
    mask = Image.new("L", (card_w * ss, card_h * ss), 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle((0, 0, card_w * ss - 1, card_h * ss - 1), radius=radius * ss, fill=255)

    # Paste gradient onto card
    card.paste(gradient, (pad * ss, pad * ss), mask)

    # Draw border and 3D highlights
    cd = ImageDraw.Draw(card)
    cr = (pad * ss, pad * ss, (pad + card_w) * ss - 1, (pad + card_h) * ss - 1)
    
    # Subtle all-around thin border (Mica style)
    cd.rounded_rectangle(cr, radius=radius * ss,
                         outline=(255, 255, 255, 10), width=max(1, round(scale * ss)))
                         
    # Top highlight line inside the rounded rect flat top part
    line_w = max(1, round(scale * ss))
    cd.line((pad * ss + radius * ss, pad * ss, (pad + card_w) * ss - radius * ss, pad * ss),
            fill=(255, 255, 255, 30), width=line_w)
            
    # Bottom shadow line
    cd.line((pad * ss + radius * ss, (pad + card_h) * ss - 1, (pad + card_w) * ss - radius * ss, (pad + card_h) * ss - 1),
            fill=(0, 0, 0, 30), width=line_w)

    # Composite shadow + card
    img = Image.alpha_composite(shadow, card)

    # Draw text at 4x supersampled resolution for crisp rendering
    draw = ImageDraw.Draw(img)
    
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
        status_color = "#f0f6fc"  # Standard white
    elif minutes is not None:
        if minutes < 20:
            status_color = "#ff6b72"  # Fluent Soft Red
        elif minutes < 45:
            status_color = "#ff9d5c"  # Fluent Soft Orange
        else:
            status_color = "#f0f6fc"  # Normal White
    else:
        status_color = "#8b949e"  # Normal Gray

    # Match the visual weight of Win11 taskbar status text, with tighter labels.
    val_size = round(11.5 * scale * ss)
    unit_size = round(11.5 * scale * ss)
    max_width = card_w * ss - round(1 * scale * ss)

    # Scale down sizes if text runs too wide (adaptive fitting)
    for size_reduce in range(0, 5):
        font_val = load_font("seguisb.ttf", val_size - size_reduce * ss)
        font_unit = load_font("segoeui.ttf", unit_size - size_reduce * ss)
        
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
                ], fill=(240, 246, 252, 255))  # Clean white/gray lightning
                
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
                "color": "#f0f6fc"
            })
        elif eta_text in ("--", "FULL"):
            segments.append({
                "type": "text",
                "text": eta_text,
                "font": font_val,
                "color": "#8b949e"
            })
        else:
            segments.append({
                "type": "text",
                "text": eta_text,
                "font": font_val,
                "color": status_color
            })
            segments.append({
                "type": "text",
                "text": "m",
                "font": font_unit,
                "color": "#8b949e"
            })
            
        # 2. Separator
        segments.append({
            "type": "text",
            "text": " ",
            "font": font_unit,
            "color": "#8b949e"
        })
        
        # 3. Wattage Segment
        if watts_text == "--":
            segments.append({
                "type": "text",
                "text": "--",
                "font": font_val,
                "color": "#8b949e"
            })
        else:
            val = watts_text[:-1] if watts_text.endswith("W") else watts_text
            segments.append({
                "type": "text",
                "text": val,
                "font": font_val,
                "color": "#f0f6fc"
            })
            segments.append({
                "type": "text",
                "text": "W",
                "font": font_unit,
                "color": "#8b949e"
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
        self.last_charging = False
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
        self.last_charging = charging
        self.render()
        if self.tick <= 5 or self.tick % 30 == 0:
            watts_text = "--" if watts is None else f"{watts:.1f}W"
            log(f"Tick {self.tick}: eta={eta}, watts={watts_text}")

    def render(self) -> None:
        watts_text = "--" if self.last_watts is None else f"{self.last_watts:.1f}W"
        image = render_widget_image(self.last_eta, watts_text, self.dpi, self.last_charging)
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
        if msg == WM_DPICHANGED:
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
