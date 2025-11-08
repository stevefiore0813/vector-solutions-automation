#!/usr/bin/env python3
import json, os, sys

ROSTER = os.path.expanduser("~/projects/vector-solutions/02_Scripts/Tests/roster.json")
OUT    = os.path.expanduser("~/projects/vector-solutions/02_Scripts/Tests/participants.dry.json")
TARGET_UNITS = {"R1", "R26", "L26"}

with open(ROSTER, "r", encoding="utf-8") as f:
    roster = json.load(f)

by_unit = roster.get("by_unit", {})
picked  = []
for u in TARGET_UNITS:
    for name in by_unit.get(u, []):
        if name not in picked:
            picked.append(name)

payload = {
    "units": sorted(TARGET_UNITS),
    "participants": picked,
    "source": "MiniCAD",
    "generated_at": roster.get("generated_at"),
}

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(payload, f, indent=2, ensure_ascii=False)

print(f"[dry] Units: {', '.join(sorted(TARGET_UNITS))}")
print(f"[dry] Participants ({len(picked)}):")
for n in picked:
    print(f"  - {n}")
print(f"[dry] Wrote {OUT}")
