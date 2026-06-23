# Surface Battery Widget

A lightweight Windows 11 battery-runtime overlay developed on a Surface Pro 8.
The backend reads live system power draw and remaining battery energy, then
estimates the time left on battery.

The current V10 UI is a functional reference implementation, not a finished
visual design. The next task is to redesign the frontend while preserving the
working sensor, refresh, startup, positioning, and context-menu behavior.

## Current behavior

- Reads system power from the Windows PDH `Power Meter` counter.
- Reads remaining battery capacity and AC state through WMI.
- Recalculates the ETA once per second on a background worker.
- Uses a native Win32 per-pixel-alpha layered window.
- Supports per-monitor DPI scaling, always-on-top display, drag locking,
  right-edge positioning, a right-click menu, and per-user startup.
- Prevents duplicate instances with a named mutex.

## Requirements

- Windows 11
- Python 3.10 or newer
- A machine exposing `Power Meter` PDH data and WMI battery data

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Run:

```powershell
pythonw SurfaceBatteryWidgetV10.py
```

Or double-click `Start_SurfaceBatteryWidgetV10.cmd`.

## Files

- `SurfaceBatteryWidgetV10.py`: current backend and native reference UI.
- `Start_SurfaceBatteryWidgetV10.cmd`: portable Windows launcher.
- `GEMINI_FRONTEND_BRIEF.md`: frontend redesign brief and invariants.

Runtime logs, battery reports, local screenshots, caches, and personal context
are intentionally excluded from version control.
