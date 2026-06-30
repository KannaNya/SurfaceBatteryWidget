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
- Re-locks itself for several seconds after resume, display changes, DPI
  changes, and taskbar/work-area changes so it stays attached to the taskbar.
- Keeps a local power diary with total power and top process activity for
  battery-life diagnosis.
- The right-click menu can open the power summary, open the diary folder, and
  inspect local diagnostics.
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

## Power Diary

The widget writes local diagnostics under `power-diary/`:

- `summary.txt`: readable power and top-process summary with suggestions.
- `power_samples_v2.csv`: total power, battery, and Energy Saver samples.
- `process_activity.csv`: top process CPU activity samples.

These files are ignored by Git because they contain local usage history.

Run `python AnalyzeEnergySaver.py` to compare average power while Energy Saver
is on and off. It needs samples from both states to be meaningful.

## Files

- `SurfaceBatteryWidgetV10.py`: current backend and native reference UI.
- `Start_SurfaceBatteryWidgetV10.cmd`: portable Windows launcher.
- `AnalyzeEnergySaver.py`: summarizes Energy Saver on/off power samples.
- `GEMINI_FRONTEND_BRIEF.md`: frontend redesign brief and invariants.

Runtime logs, battery reports, local screenshots, caches, and personal context
are intentionally excluded from version control.
