#!/usr/bin/env python3
import json, os, sys

ROSTER = os.path.expanduser(sys.argv[1] if len(sys.argv) > 1 else "~/projects/vector-solutions/roster.json")
UNITS  = set(sys.argv[2].split(",")) if len(sys.argv) > 2 else {"R1","R26","L26"}
OUT    = os.path.expanduser(sys.argv[3] if len(sys.argv) > 3 else "~/projects/vector-solutions/participants.dry.json")

with open(ROSTER, "r", encoding="utf-8") as f:
    roster = json.load(f)

unit_map = roster.get("by_unit", roster.get("units", {}))
picked = []
seen = set()
for u in UNITS:
    for n in unit_map.get(u, []):
        if n not in seen:
            picked.append(n)
            seen.add(n)

payload = {"units": sorted(UNITS), "participants": picked, "count": len(picked)}
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(payload, f, indent=2, ensure_ascii=False)

print(f"[dump] Units: {', '.join(sorted(UNITS))}")
for n in picked:
    print("  -", n)
print(f"[dump] Wrote {OUT}")
