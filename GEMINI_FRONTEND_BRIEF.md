# Frontend Redesign Brief

## Goal

Replace the V10 presentation layer with a polished Windows 11-quality battery
runtime widget. The primary information is estimated remaining time. Power draw
in watts is secondary. Do not repeat the battery percentage already shown by
Windows.

## Preserve these working behaviors

- `PdhPowerMeter` power reading and its unit conversion.
- `BatteryReader` WMI data retrieval.
- One-second background refresh and UI-thread message handoff.
- ETA calculation and AC/unavailable states.
- Single-instance mutex.
- Per-monitor DPI awareness.
- Always-on-top tool window behavior.
- Right-click menu, drag lock, right-edge relock, and exit actions.
- HKCU startup registration and cleanup of earlier startup entries.
- Positioning relative to the monitor work area, not a guessed taskbar size.

## Visual constraints from user testing

- Keep a compact, near-square rounded rectangle; avoid a long horizontal bar.
- Do not use a pill inside another pill.
- Do not use decorative green status bars.
- Do not show a `Remaining` label, battery percentage, or watt-hours.
- Keep the ETA on one line when practical, for example `47 min`.
- Number and unit on the same line must use the same font size, weight, and
  baseline.
- Watts should be clearly readable but visually subordinate to ETA.
- Avoid fuzzy fonts, hard chroma-key transparency, and jagged rounded edges.
- Verify at 200% Windows display scaling and on both light and dark backgrounds.

## Suggested separation

The current file mixes backend, rendering, window behavior, and startup logic.
For the redesign, isolate sensor/ETA state from the renderer so frontend work
does not destabilize refresh or context-menu handling. A small immutable view
model such as `{eta_text, watts_text, power_online, updated_at}` is sufficient.

## Acceptance checks

1. The visible ETA changes without restarting the process.
2. Right-click opens the menu reliably; touch long-press may use the same path.
3. Only one widget process remains after repeated launches.
4. The window remains sharp at 100%, 150%, and 200% scaling.
5. Transparent corners have smooth alpha with no black or colored fringe.
6. The layout does not shift when ETA changes between minutes, hours, AC, and
   unavailable states.
7. Startup launches the same current version exactly once.
