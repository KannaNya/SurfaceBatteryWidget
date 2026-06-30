import csv
import statistics
from collections import defaultdict
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DIARY_DIR = BASE_DIR / "power-diary"
SAMPLES = DIARY_DIR / "power_samples_v2.csv"


def as_float(value: str) -> float | None:
    try:
        return float(value)
    except Exception:
        return None


def main() -> int:
    if not SAMPLES.exists():
        print(f"No v2 power samples yet: {SAMPLES}")
        print("Keep the widget running for at least a minute, then try again.")
        return 1

    groups = defaultdict(list)
    expected_groups = defaultdict(list)
    rows = []
    with SAMPLES.open("r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            watts = as_float(row.get("system_watts", ""))
            if watts is None:
                continue
            status = row.get("energy_saver_status") or "Unknown"
            expected = row.get("energy_saver_auto_expected") or "Unknown"
            groups[status].append(watts)
            expected_groups[expected].append(watts)
            rows.append(row)

    print("Energy saver power analysis")
    print(f"Samples: {sum(len(values) for values in groups.values())}")
    print("")
    print("By actual EnergySaverStatus:")
    for status, values in sorted(groups.items()):
        print(f"- {status:<12} count={len(values):4d} avg={statistics.fmean(values):5.1f} W min={min(values):5.1f} W max={max(values):5.1f} W")

    print("")
    print("By auto-threshold expectation:")
    for expected, values in sorted(expected_groups.items()):
        print(f"- expected={expected:<5} count={len(values):4d} avg={statistics.fmean(values):5.1f} W")

    if len(groups) < 2:
        latest = rows[-1] if rows else {}
        print("")
        print("Not enough On/Off contrast yet.")
        print(f"Latest status: {latest.get('energy_saver_status', 'Unknown')}")
        print(f"Latest battery: {latest.get('battery_percent', '--')}%")
        print(f"Threshold: {latest.get('energy_saver_threshold', '--')}%")
        print("To prove savings, collect samples both above and below the threshold, or manually toggle Energy Saver.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
