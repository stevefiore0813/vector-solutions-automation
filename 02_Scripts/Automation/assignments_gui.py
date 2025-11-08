#!/usr/bin/env python3
"""
assignments_gui.py
Pull MiniCAD unit/staff data and emit a roster.json your training-bot can use.

Usage (CLI):
    python3 assignments_gui.py --output ~/projects/vector-solutions/02_Scripts/Rosters/roster.json --print

Optional GUI:
    python3 assignments_gui.py --gui

Env vars (preferred over hardcoded creds):
    MINICAD_USER=ccfrcad
    MINICAD_PASS=clayfireISGREAT2025
"""

import json
import os
import re
import sys
import argparse
import tempfile

from pathlib import Path
from datetime import datetime, timezone
from typing import Dict, List, Any

import requests
from requests.adapters import HTTPAdapter, Retry

MINICAD_URL = os.environ.get("MINICAD_URL", "").strip()
MINICAD_USER = os.environ.get("MINICAD_USER", "").strip()
MINICAD_PASS = os.environ.get("MINICAD_PASS", "").strip()

# Sensible default output where you've been dumping everything else
DEFAULT_OUTPUT = os.path.expanduser(
    "~/projects/vector-solutions/02_Scripts/Rosters/roster.json"
)

# Ranks/titles to strip from names so matching succeeds
RANK_PREFIXES = [
    r"Acting\s+", r"Interim\s+",
    r"(Chief|Battalion Chief|BC|Division Chief)\s+",
    r"(Captain|Capt)\.?\s+",
    r"(Lieutenant|Lt)\.?\s+",
    r"(Sergeant|Sgt)\.?\s+",
    r"(Engineer|Eng)\.?\s+",
    r"(Firefighter|FF)\.?\s+",
    r"(Paramedic|Medic)\.?\s+",
    r"(Officer)\s+",
]

RANK_EXTRACT = re.compile(rf"^({'|'.join(p[:-3] if p.endswith(r'\s+') else p for p in RANK_PREFIXES)})", re.IGNORECASE)
RANK_SPLIT = re.compile(r'^\s*(?:' + '|'.join(RANK_PREFIXES) + r')', re.IGNORECASE)
WRAPPER_KEYS = ("d", "results", "value", "Units", "units", "Data", "data")

def get_credentials() -> Dict[str, str]:
    user = os.getenv("MINICAD_USER", "").strip() or "ccfrcad"
    pwd = os.getenv("MINICAD_PASS", "").strip() or "clayfireISGREAT2025"
    return {"user": user, "pass": pwd}


def requests_session() -> requests.Session:
    sess = requests.Session()
    retries = Retry(
        total=5,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retries)
    sess.mount("http://", adapter)
    sess.mount("https://", adapter)
    return sess

def build_roster_structure(units: List[Dict[str, Any]]) -> Dict[str, Any]:
    out_units, flat, by_unit = [], [], {}

    for u in units:
        # Accept a variety of key spells because APIs are allergic to consistency
        def g(*names, default=""):
            for n in names:
                if n in u:
                    return u.get(n)
            return default

        unit_name = str(g("UnitName", "unit", "Unit", "Name", default="")).strip()
        if not unit_name:
            # skip entries with no unit id
            continue

        staff_field = g("Staff", "staff", "Personnel", "personnel", "Members", "members", default=[])
        staff = normalize_staff_list(staff_field)
        names_only = [m["name"] for m in staff if m.get("name")]

        by_unit[unit_name] = names_only
        flat.extend(n for n in names_only if n)

        out_units.append({
            "unit": unit_name,
            "unit_type": str(g("UnitType", "Type", "type", default="")).strip(),
            "status": str(g("UnitStatus", "Status", "status", default="")).strip(),
            "home_station": str(g("HomeStation", "Home", "Station", "station", default="")).strip(),
            "district": str(g("District", "district", default="")).strip(),
            "prime_officer": str(g("PrimeOfficer", "prime_officer", "Officer", default="")).strip(),
            "location": str(g("Location", "location", default="")).strip(),
            "staff": staff
        })

    # Deduplicate flat list while preserving order
    seen = set()
    flat_unique = []
    for n in flat:
        if n not in seen:
            flat_unique.append(n)
            seen.add(n)

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source": "MiniCAD",
        "units": out_units,
        "flat": flat_unique,
        "by_unit": by_unit
    }

def fetch_minicad() -> List[Dict[str, Any]]:
    creds = get_credentials()
    sess = requests_session()
    resp = sess.get(
        MINICAD_URL,
        auth=(creds["user"], creds["pass"]),
        timeout=20
    )

    # If auth failed but server returned 200 with HTML (some do that), catch it.
    content_type = resp.headers.get("Content-Type", "")
    body_text = None

    try:
        data = resp.json()
    except Exception:
        # Not JSON; dump to a temp file so you can eyeball it.
        body_text = resp.text[:5000]  # keep it sane
        dump = Path(tempfile.gettempdir()) / "minicad_dump.txt"
        dump.write_text(
            f"status={resp.status_code}\ncontent_type={content_type}\n\n{body_text}",
            encoding="utf-8",
        )
        if resp.status_code == 401:
            raise RuntimeError("MiniCAD 401 Unauthorized. Check username/password or IP restrictions.")
        if "text/html" in content_type.lower():
            raise RuntimeError(f"MiniCAD returned HTML (likely a login page). Dumped to: {dump}")
        raise RuntimeError(f"MiniCAD returned non-JSON. Dumped to: {dump}")

    try:
        units = _unwrap_units(data)
    except Exception as e:
        # Save the raw JSON so we can adapt without guesswork
        dump = Path(tempfile.gettempdir()) / "minicad_raw.json"
        import json as _json
        dump.write_text(_json.dumps(data, indent=2), encoding="utf-8")
        raise RuntimeError(f"Unexpected MiniCAD payload shape; saved raw JSON to: {dump}") from e

    # Basic sanity: list of dicts with UnitName-ish keys
    if not all(isinstance(u, dict) for u in units):
        raise RuntimeError("Units list isn't a list of objects. API changed?")
    return units


def strip_rank(name: str) -> str:
    # Remove common rank/title prefixes
    no_rank = RANK_SPLIT.sub("", name).strip()
    # Remove accidental double-spaces and trailing commas
    no_rank = re.sub(r"\s{2,}", " ", no_rank)
    no_rank = re.sub(r",\s*,+", ", ", no_rank)
    return no_rank


def split_role(name: str) -> (str, str): # pyright: ignore[reportInvalidTypeForm]
    # Try to capture role if present
    m = RANK_EXTRACT.match(name.strip())
    role = ""
    if m:
        role = m.group(1).strip()
    return strip_rank(name), role


def normalize_staff_list(staff_field: Any) -> List[Dict[str, str]]:
    """
    staff_field is expected to be a list of strings like:
        ["Officer Smith, Matthew", "Engineer Eng, Christopher"]
    Returns [{'name': 'Smith, Matthew', 'role': 'Officer'}, ...]
    """
    out = []
    if not staff_field:
        return out
    if not isinstance(staff_field, list):
        # Some feeds return a single string with delimiters. Be forgiving.
        if isinstance(staff_field, str):
            staff_field = [p.strip() for p in re.split(r"[;|]", staff_field) if p.strip()]
        else:
            return out

    for raw in staff_field:
        if not isinstance(raw, str):
            continue
        cleaned, role = split_role(raw)
        # If the name came in as "First Last", flip to "Last, First" for matching
        if "," not in cleaned:
            parts = cleaned.split()
            if len(parts) >= 2:
                cleaned = f"{parts[-1]}, {' '.join(parts[:-1])}"
        out.append({"name": cleaned, "role": role})
    return out


def _unwrap_units(payload):
    """
    Try very hard to find the actual list of unit dicts inside whatever
    nonsense shape the server returned.
    """
    # case 1: it's already a list
    if isinstance(payload, list):
        return payload

    # case 2: dict with a known wrapper key
    if isinstance(payload, dict):
        for k in WRAPPER_KEYS:
            if k in payload and isinstance(payload[k], list):
                return payload[k]

        # case 3: dict where exactly one value is a list
        list_values = [v for v in payload.values() if isinstance(v, list)]
        if len(list_values) == 1:
            return list_values[0]

        # case 4: dict where list is inside a nested dict under known keys
        for k in WRAPPER_KEYS:
            v = payload.get(k)
            if isinstance(v, dict):
                for vv in v.values():
                    if isinstance(vv, list):
                        return vv

    # case 5: string that might be JSON again
    if isinstance(payload, str):
        try:
            import json as _json
            return _unwrap_units(_json.loads(payload))
        except Exception:
            pass

    raise RuntimeError("Unexpected MiniCAD payload shape; expected list of units")



def write_json(path: str, payload: Dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def run_cli(output_path: str, do_print: bool) -> int:
    units = fetch_minicad()
    roster = build_roster_structure(units)
    write_json(output_path, roster)
    if do_print:
        print(json.dumps(roster, indent=2, ensure_ascii=False))
    else:
        print(f"[assignments] Wrote roster: {output_path}   "
              f"(units={len(roster['units'])}, people={len(roster['flat'])})")
    return 0


def run_gui(output_path: str) -> int:
    # Minimal Tk GUI so you can click a button like a civilized primate
    try:
        import tkinter as tk
        from tkinter import ttk, messagebox
    except Exception:
        print("tkinter not available; run without --gui", file=sys.stderr)
        return 1

    def do_fetch():
        btn.config(state="disabled")
        status_var.set("Fetching MiniCAD…")
        root.update_idletasks()
        try:
            units = fetch_minicad()
            roster = build_roster_structure(units)
            write_json(output_path, roster)
            tree.delete(*tree.get_children())
            for u in roster["units"]:
                unit = u["unit"]
                names = ", ".join([m["name"] for m in u["staff"]]) or "(no staff)"
                tree.insert("", "end", values=(unit, u["unit_type"], u["status"], names))
            status_var.set(f"Saved {output_path} | {len(roster['flat'])} personnel")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            status_var.set("Error")
        finally:
            btn.config(state="normal")

    root = tk.Tk()
    root.title("Assignments (MiniCAD → roster.json)")

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    btn = ttk.Button(frm, text="Fetch & Save Roster", command=do_fetch)
    btn.pack(anchor="w")

    cols = ("Unit", "Type", "Status", "Staff")
    tree = ttk.Treeview(frm, columns=cols, show="headings", height=16)
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, anchor="w", width=140 if c != "Staff" else 420)
    tree.pack(fill="both", expand=True, pady=(8, 4))

    status_var = tk.StringVar(value=f"Output: {output_path}")
    status = ttk.Label(frm, textvariable=status_var)
    status.pack(anchor="w")

    root.mainloop()
    return 0


def parse_args(argv: List[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Fetch MiniCAD assignments and build roster.json")
    p.add_argument("--output", default=DEFAULT_OUTPUT, help=f"Path to write roster.json (default: {DEFAULT_OUTPUT})")
    p.add_argument("--print", dest="do_print", action="store_true", help="Print JSON to stdout as well")
    p.add_argument("--gui", action="store_true", help="Launch minimal GUI")
    return p.parse_args(argv)


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    if args.gui:
        return run_gui(args.output)
    return run_cli(args.output, args.do_print)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
