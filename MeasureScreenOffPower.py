import argparse
import csv
import ctypes
import statistics
import time
from ctypes import wintypes
from pathlib import Path

import win32com.client


BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "power-diary"

HWND_BROADCAST = 0xFFFF
WM_SYSCOMMAND = 0x0112
SC_MONITORPOWER = 0xF170
MONITOR_OFF = 2

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
MAX_REASONABLE_BATTERY_WATTS = 180.0

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)


class PdhPowerMeter:
    PDH_FMT_DOUBLE = 0x00000200

    class PDH_FMT_COUNTERVALUE(ctypes.Structure):
        _fields_ = [("CStatus", wintypes.DWORD), ("doubleValue", ctypes.c_double)]

    def __init__(self) -> None:
        self.pdh = ctypes.WinDLL("pdh.dll")
        self.query = wintypes.HANDLE()
        self.counter = wintypes.HANDLE()
        self.ready = False
        self.init_counter()

    def init_counter(self) -> None:
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
                return
            self.pdh.PdhCloseQuery(query)

    def read_watts(self) -> float | None:
        if not self.ready:
            self.init_counter()
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
                return float(value.doubleValue) / 1000.0
        except Exception:
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

    def read(self) -> dict:
        snap = {
            "percent": None,
            "remaining_wh": None,
            "full_charge_wh": None,
            "online": False,
            "charging": False,
            "discharge_w": None,
            "charge_w": None,
        }
        try:
            battery = list(self.wmi_default.InstancesOf("Win32_Battery"))
            status = list(self.wmi_battery.InstancesOf("BatteryStatus"))
            full = list(self.wmi_battery.InstancesOf("BatteryFullChargedCapacity"))
            b = battery[0] if battery else None
            s = status[0] if status else None
            f = full[0] if full else None
            if b is not None:
                snap["percent"] = int(b.EstimatedChargeRemaining)
            if s is not None:
                snap["remaining_wh"] = float(s.RemainingCapacity) / 1000.0
                snap["online"] = bool(s.PowerOnline)
                snap["charging"] = bool(s.Charging)
                snap["discharge_w"] = valid_battery_watts(s.DischargeRate)
                snap["charge_w"] = valid_battery_watts(s.ChargeRate)
            if f is not None:
                snap["full_charge_wh"] = float(f.FullChargedCapacity) / 1000.0
        except Exception as exc:
            print(f"Battery read failed: {exc}")
        return snap


def valid_battery_watts(value) -> float | None:
    try:
        watts = float(value) / 1000.0
    except Exception:
        return None
    if 0 < watts <= MAX_REASONABLE_BATTERY_WATTS:
        return watts
    return None


def set_sleep_guard(enabled: bool) -> None:
    flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED if enabled else ES_CONTINUOUS
    kernel32.SetThreadExecutionState(flags)


def turn_display_off() -> None:
    user32.SendMessageW(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF)


def average(values: list[float]) -> float | None:
    clean = [value for value in values if value is not None]
    if not clean:
        return None
    return statistics.fmean(clean)


def sample_phase(
    phase: str,
    seconds: int,
    interval: float,
    meter: PdhPowerMeter,
    battery: BatteryReader,
    writer: csv.writer,
) -> list[float]:
    values = []
    end_at = time.monotonic() + seconds
    while time.monotonic() < end_at:
        snap = battery.read()
        system_watts = meter.read_watts()
        battery_watts = snap["charge_w"] if snap["online"] else snap["discharge_w"]
        values.append(system_watts if system_watts is not None else battery_watts)
        writer.writerow(
            [
                time.strftime("%Y-%m-%d %H:%M:%S"),
                phase,
                "" if system_watts is None else f"{system_watts:.2f}",
                "" if battery_watts is None else f"{battery_watts:.2f}",
                snap["percent"],
                snap["remaining_wh"],
                snap["online"],
                snap["charging"],
            ]
        )
        time.sleep(interval)
    return values


def format_watts(value: float | None) -> str:
    return "--" if value is None else f"{value:.1f} W"


def main() -> int:
    parser = argparse.ArgumentParser(description="Measure Surface power draw with the display on and off.")
    parser.add_argument("--baseline-seconds", type=int, default=60, help="Screen-on baseline duration.")
    parser.add_argument("--screen-off-seconds", type=int, default=600, help="Screen-off measurement duration.")
    parser.add_argument("--interval", type=float, default=2.0, help="Sample interval in seconds.")
    parser.add_argument("--no-turn-off-display", action="store_true", help="Do not turn off the display.")
    parser.add_argument("--allow-sleep", action="store_true", help="Do not keep the system awake during measurement.")
    args = parser.parse_args()

    OUT_DIR.mkdir(exist_ok=True)
    output = OUT_DIR / f"screen_off_power_{time.strftime('%Y%m%d-%H%M%S')}.csv"
    meter = PdhPowerMeter()
    battery = BatteryReader()

    first = battery.read()
    if first["online"]:
        print("Warning: AC power is connected. For battery drain, unplug before measuring.")
    if not meter.ready:
        print("Warning: Surface power meter is unavailable; falling back to battery charge/discharge rate.")

    print(f"Writing samples to: {output}")
    print(f"Screen-on baseline: {args.baseline_seconds}s")
    print(f"Screen-off phase: {args.screen_off_seconds}s")
    print("Wake the display with keyboard/touchpad after the screen-off phase finishes.")

    if not args.allow_sleep:
        set_sleep_guard(True)

    try:
        with output.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "timestamp",
                    "phase",
                    "system_watts",
                    "battery_charge_or_discharge_watts",
                    "battery_percent",
                    "remaining_wh",
                    "online",
                    "charging",
                ]
            )

            baseline = sample_phase("screen_on", args.baseline_seconds, args.interval, meter, battery, writer)
            print("Turning display off now.")
            time.sleep(1)
            if not args.no_turn_off_display:
                turn_display_off()
            screen_off = sample_phase("screen_off", args.screen_off_seconds, args.interval, meter, battery, writer)
    finally:
        if not args.allow_sleep:
            set_sleep_guard(False)
        meter.close()

    baseline_avg = average(baseline)
    off_avg = average(screen_off)
    after = battery.read()
    full_wh = after["full_charge_wh"] or first["full_charge_wh"]

    print("")
    print("Result")
    print(f"Screen on average : {format_watts(baseline_avg)}")
    print(f"Screen off average: {format_watts(off_avg)}")
    if baseline_avg is not None and off_avg is not None:
        saved = baseline_avg - off_avg
        print(f"Display-off saving: {saved:.1f} W")
    if full_wh and off_avg and off_avg > 0:
        print(f"Estimated full-battery screen-off runtime: {full_wh / off_avg:.1f} h")
    print(f"CSV: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
