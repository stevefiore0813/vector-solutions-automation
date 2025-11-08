#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
training-bot.py
Stable Playwright flow for Vector Solutions training entry with:
- Live roster fetch from MiniCAD (basic auth)
- Scenario source from CSV (preferred) or DOCX fallback
- Robust login and dashboard navigation
- "Save and Add Users" frame handling with retries
- Post-submit verification and structured logging

Dependencies:
  pip install playwright python-docx==1.1.0 pandas python-dotenv
  playwright install

Run example:
  python3 /home/training-bot/projects/vector-solutions/02_Scripts/Automation/training-bot.py   --vs-user "$VS_USER"   --vs-pass "$VS_PASS"   --scenario-csv "/home/training-bot/projects/vector-solutions/03_Training-Documentation/trainings.csv"   --roster-cmd "python3 /home/training-bot/projects/vector-solutions/02_Scripts/Automation/assignments_gui.py --print"   --units "R1,R26,L26"   --artifact-dir "/home/training-bot/projects/vector-solutions/04_Outputs/Logs/_artifacts"
"""

import os, sys, time, json, csv, re, traceback
import shlex, subprocess
import random
from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from contextlib import contextmanager

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from playwright.sync_api import ElementHandle
import pandas as pd
from dotenv import load_dotenv

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

# -----------------------------
# Config
# -----------------------------

VS_LOGIN_URL = "https://app.targetsolutions.com/ts/login"   # safe default; real SSO/login may redirect
VECTOR_DASHBOARD_READY_TEXT = "Dashboard"
FORM_OPEN_SELECTOR = (
    # brittle CSS from earlier; keep as fallback but prefer text-first approach
    "#single-spa-application\\:\\@target-solutions\\/home > section > section > vwc-tiling-grid "
    "> vwc-tiling-grid-tile.bulletin-board.is-schedule > vwc-card > div.bulletinBoard-content > div > p:nth-child(13) > a:nth-child(1) > img"
)

# Keys in your CSV
CSV_HEADERS = ["Location", "Checkbox Label", "Description", "Duration", "Instructor"]

# Fallback/defaults
DEFAULT_DURATION = "2 hours"
DEFAULT_INSTRUCTOR = "Lt. Fiore"
FORM_PROBES = [
    # the button we ultimately need
    {"role": ("button", re.compile(r"Save\s*&?\s*Add\s*Users", re.I))},
    # common fields on the Company Training form
    {"css": "textarea[name*='description' i], textarea[id*='description' i]"},
    {"css": "input[name*='instructor' i], input[id*='instructor' i], input[placeholder*='Instructor' i]"},
    {"css": "input[name*='location' i], input[id*='location' i], textarea[name*='location' i]"},
    # any label text that only shows on the form
    {"text": re.compile(r"\bCompany\s*Training\b", re.I)},
]

# XPath for the Submit button on the training form since it wants to be stubborn
SUBMIT_XPATH = "/html/body/div[1]/div/div/div[2]/div[2]/div/div[1]/div[1]/form/div[3]/input[1]"

# Where to write logs/artifacts
# Artifacts directory (override with env var if you want)
ARTIFACT_DIR = os.environ.get(
    "TB_ARTIFACT_DIR",
    "/home/training-bot/projects/vector-solutions/04_Outputs/Logs/_artifacts"
)
os.makedirs(ARTIFACT_DIR, exist_ok=True)

def _abs(p: str) -> str:
    import os
    return os.path.abspath(p)

def _safe_write(path: str, text: str):
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        log(f"[dump] wrote: {_abs(path)}")
        return path
    except Exception as e:
        # fallback to /tmp so it never silently fails
        alt = f"/tmp/{os.path.basename(path)}"
        with open(alt, "w", encoding="utf-8") as f:
            f.write(text)
        log(f"[dump:fallback:/tmp] {_abs(alt)} because: {e}")
        return alt

# -----------------------------
# Utilities
# -----------------------------

def log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def _log(level, msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Always convert to string first so it can’t be a function/obj anymore
    text = str(msg)
    try:
        text = text.replace("\n", " ")
    except Exception:
        # If for some cursed reason this still fails, just fall back to the raw repr
        text = repr(msg)
    print(f"[{ts}] [{level}] {text}", flush=True)

def info(msg):  _log("info", msg)

def warn(msg):  _log("warn", msg)

def fatal(msg): _log("fatal", msg)

def sanitize(s: str) -> str:
    return (s or "").strip()

def ensure_last_first(name: str) -> str:
    # Accept "First Last" or "Last, First" and normalize to "Last, First"
    name = sanitize(name)
    if not name:
        return name
    if "," in name:
        parts = [p.strip() for p in name.split(",", 1)]
        return f"{parts[0]}, {parts[1]}"
    parts = name.split()
    if len(parts) >= 2:
        first = parts[0]
        last = parts[-1]
        return f"{last}, {first}"
    return name

def write_json(path: str, data: Any):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def read_json_from_path_or_text(s: str) -> dict:
    # If s is a path to a file, load it; otherwise treat as JSON text.
    p = s.strip()
    if os.path.exists(p) and os.path.isfile(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return json.loads(p)

def fetch_roster_via_command(cmd: str, timeout: int = 60) -> Dict[str, Any]:
    """
    Runs an external program that either prints the full roster JSON to stdout
    or prints a file path that contains that JSON. Returns parsed JSON.
    """
    log(f"Running roster command: {cmd}")
    proc = subprocess.run(
        shlex.split(cmd),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout,
        text=True
    )
    if proc.returncode != 0:
        raise RuntimeError(f"Roster command failed [{proc.returncode}]: {proc.stderr.strip()[:300]}")
    out = proc.stdout.strip()
    if not out:
        raise RuntimeError("Roster command produced no output")
    try:
        data = read_json_from_path_or_text(out)
    except Exception as e:
        raise RuntimeError(f"Could not parse roster output as JSON or JSON file path: {e}")
    out_path = os.path.join(ARTIFACT_DIR, f"roster_cmd_{int(time.time())}.json")
    write_json(out_path, data)
    log(f"Roster (command) captured to {out_path}")
    return data

def normalize_unit(u: Optional[str]) -> str:
    return (u or "").strip().upper()

def _item_unit_value(item: dict) -> str:
    # Try several common keys
    for key in ("Unit", "UnitID", "UnitName", "UnitNumber", "Apparatus", "Resource", "CallSign"):
        if key in item:
            return normalize_unit(str(item[key]))
    return ""

def filter_roster_by_units(roster: Any, include_units: List[str]) -> Any:
    if not include_units:
        return roster
    wanted = {normalize_unit(u) for u in include_units}

    def matches(item: Any) -> bool:
        if not isinstance(item, dict):
            return False
        v = _item_unit_value(item)
        return v and (v in wanted or any(v.startswith(w) for w in wanted))

    def filter_any(obj: Any) -> Any:
        if isinstance(obj, list):
            out = []
            for x in obj:
                if isinstance(x, dict) and (matches(x) or any(isinstance(v, list) and any(matches(i) for i in v if isinstance(i, dict)) for v in x.values())):
                    out.append(x)
            return out
        if isinstance(obj, dict):
            out = {}
            for k, v in obj.items():
                if isinstance(v, list):
                    vv = [x for x in v if matches(x)]
                    if vv:
                        out[k] = vv
            return out if out else obj
        return obj

    filtered = filter_any(roster)
    # If filter obliterated everything, fall back to unfiltered so we still get names
    if (isinstance(filtered, list) and not filtered) or (isinstance(filtered, dict) and filtered == {}):
        log("[warn] Unit filter removed all entries; falling back to unfiltered roster")
        return roster
    return filtered

def fill_date_time_now(page):
    # Date input (try a few common patterns). Adjust if yours differs.
    # Try HTML5 date input first
    today = datetime.now().strftime("%Y-%m-%d")
    filled_date = False

    candidates = [
        'input[name*="endDate" i]',      # name contains endDate (case-insensitive)
        'input[name*="dateComplete" i]',
        'input[type="date"]'
    ]
    for sel in candidates:
        el = page.locator(sel).first
        if el.count() > 0 and el.is_enabled():
            try:
                el.fill(today)
                filled_date = True
                break
            except Exception:
                pass

    if not filled_date:
        # Some legacy widgets aren’t fillable; try clicking and typing
        for sel in candidates:
            el = page.locator(sel).first
            if el.count() > 0 and el.is_enabled():
                try:
                    el.click()
                    el.press("Control+A")
                    el.type(today)
                    filled_date = True
                    break
                except Exception:
                    pass

    # Time select. Use select_option, not fill.
    now = datetime.now()
    # Round to nearest 15 minutes to match your options
    minute = (now.minute // 15) * 15
    hh = now.hour % 12 or 12
    ampm = "AM" if now.hour < 12 else "PM"
    label = f"{hh}:{minute:02d} {ampm}"   # e.g. "9:15 AM"

    # Your snippet showed: <select name="nodeUserVal6_endTime"><option value="8:15 AM">8:15 AM</option>
    time_select = page.locator('select[name*="endTime" i]').first
    if time_select.count() > 0:
        try:
            time_select.select_option(label)     # value equals label
        except Exception:
            # Fall back to selecting by label if value mismatch
            time_select.select_option(label={"label": label})
    else:
        # Some forms use startTime/endTime pairs; try a broader net
        any_time = page.locator('select:has(option[value*="AM"], option[value*="PM"])').first
        if any_time.count() > 0:
            try:
                any_time.select_option(label)
            except Exception:
                any_time.select_option(label={"label": label})

def click_save_and_add_users(page):
    # Prefer role or value; these are stable
    try:
        page.get_by_role("button", name="Save and Add Users").click(timeout=4000)
    except Exception:
        page.locator('input[type="button"][value="Save and Add Users"]').first.click(timeout=4000)

def click_submit_as_complete(page):
    # Same idea
    try:
        page.get_by_role("button", name="Submit as Complete").click(timeout=4000)
    except Exception:
        page.locator('input[type="button"][value="Submit as Complete"]').first.click(timeout=4000)

def frame_by_url_contains(page, needle):
    for f in page.frames:
        if needle in (f.url or ""):
            return f
    return None

def within_users_frame(page):
    # Adjust the substring to something unique in that users page URL
    f = frame_by_url_contains(page, "/AddUsers")
    return f or page.main_frame

def debug_snap(page, tag):
    """Screenshot + URL log so we can see what the page actually is."""
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = ARTIFACT_DIR / f"login_{tag}_{ts}.png"
        page.screenshot(path=str(path), full_page=True)
        print(f"[debug] login stage={tag} url={page.url} screenshot={path}", flush=True)
    except Exception as e:
        print(f"[debug] login stage={tag} (screenshot failed: {e}) url={page.url}", flush=True)

# -------- FILL IN FIELDS --------

def fill_core_fields(page, scenario):
    """
    Fills: Location, Duration (hours), Date Complete, Instructor.
    Uses robust text-near-field targeting so it works even if the
    question text is in a <div> instead of a <label>.
    """
    log("Filling Location / Duration / Date / Instructor (core fields)")

     # 1) Location
    location = (scenario.get("Location") or "").strip() or "Station 26"
    with swallow("fill location"):
        fill_location_field(page, location)

    # 2) Duration hours  ("How long was the training")
    duration_hours = (scenario.get("Duration") or "").strip()
    # Accept "2 hours", "2", "1.5", etc. Pull the first number; default 2.
    m = re.search(r"(\d+(?:\.\d+)?)", duration_hours)
    hours = m.group(1) if m else "2"
    with swallow("fill duration hours"):
        dur_input = page.get_by_label(re.compile(r"How long was the training", re.I))
        if not dur_input.count():
            dur_input = page.locator(
                "xpath=//*[contains(translate(normalize-space(.), "
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
                "'how long was the training')]/following::input[1]"
            )
        if not dur_input.count():
            # extra fallback: any input with 'hours' in aria-label/name
            dur_input = page.locator(
                "input[name*='hour' i], input[id*='hour' i], input[aria-label*='hour' i]"
            )
        if dur_input.count():
            dur_input.first.fill(hours)

    # 3) Date Complete + Time
    now = datetime.now()
    mmddyyyy = now.strftime("%m/%d/%Y")
    hour_12 = now.strftime("%I").lstrip("0") or "12"
    minute = int(now.strftime("%M"))
    quarter = ("00", "15", "30", "45")[(minute + 7) // 15 % 4]
    ampm = now.strftime("%p")
    time_candidates = [
        f"{hour_12}:{quarter} {ampm}",
        f"{hour_12}:{quarter}{ampm.lower()}",
        f"{hour_12}:{quarter}",
    ]

    with swallow("fill date complete (text input)"):
        date_input = page.locator(
            "xpath=//*[contains(translate(normalize-space(.), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
            "'date complete')]/following::input[1]"
        )
        if date_input.count():
            date_input.first.fill(mmddyyyy)

    with swallow("select time (dropdown next to date)"):
        time_select = page.locator(
            "xpath=//*[contains(translate(normalize-space(.), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
            "'date complete')]/following::select[1]"
        )
        if time_select.count():
            opts = time_select.first.locator("option")
            chosen = False
            # try to select a matching label
            for cand in time_candidates:
                try:
                    time_select.first.select_option(label=cand)
                    chosen = True
                    break
                except Exception:
                    pass
            if not chosen and opts.count():
                # pick first non-empty option as fallback
                for i in range(opts.count()):
                    val = (opts.nth(i).get_attribute("value") or "").strip()
                    lab = (opts.nth(i).inner_text() or "").strip()
                    if val or lab:
                        time_select.first.select_option(index=i)
                        break

    # 4) Instructor  ("Who led the training")
    instructor = (scenario.get("Instructor") or "").strip() or "Lt. Fiore"
    with swallow("fill instructor"):
        instr_input = page.get_by_label(re.compile(r"Who led the training", re.I))
        if not instr_input.count():
            instr_input = page.locator(
                "xpath=//*[contains(translate(normalize-space(.), "
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
                "'who led the training')]/following::input[1]"
            )
        if not instr_input.count():
            # Fallback: any input that smells like instructor
            instr_input = page.locator(
                "input[name*='instructor' i], input[id*='instructor' i], "
                "input[placeholder*='Instructor' i]"
            )
        if instr_input.count():
            instr_input.first.fill(instructor)

    time.sleep(0.3)  # tiny pause so the portal can re-validate before submit

def fill_basic_fields(page, *, location, duration_hours, instructor_name):
    # Location
    loc = page.locator('input[name*="location" i], textarea[name*="location" i]').first
    if loc.count():
        loc.fill(location)

    # Duration: if it’s a text box in minutes; if it’s a select, adjust to select_option
    dur = page.locator('input[name*="duration" i], input[name*="minutes" i]').first
    if dur.count():
        dur.fill(str(duration_hours * 60)) # minutes
    else:
        dur_sel = page.locator('select[name*="duration" i]').first
        if dur_sel.count():
            # Try value first, then label
            try:
                dur_sel.select_option(str(duration_hours * 60))
            except Exception:
                dur_sel.select_option(label=str(duration_hours * 60))

    # Instructor
    inst = page.locator('input[name*="instructor" i], textarea[name*="instructor" i]').first
    if inst.count():
        inst.fill(instructor_name)

def fill_location_field(page, location: str):
    """
    Fill the Location textbox:
      <input type="text" name="nodeUserVal2" id="nodeUserVal2" ...>
    We search across all frames just in case the form lives inside one.
    """
    location = (location or "").strip() or "Station 26"
    log(f"Trying to fill Location with: {location!r}")

    # Try main page first
    candidates = [
        "input#nodeUserVal2",
        "input[name='nodeUserVal2']",
        "input[a_text*='Where did the training take place?']",
        (
            "xpath=//input[@type='text' and @a_type='textbox' "
            "and contains(@a_text, 'Where did the training take place')]"
        ),
    ]

    def _try_on(target_page):
        for sel in candidates:
            loc = target_page.locator(sel)
            if loc.count():
                log(f"[location] Found Location field via selector: {sel}")
                loc.first.fill(location)
                return True
        return False

    # 1) Try main frame
    if _try_on(page):
        return

    # 2) Try all frames if not in main
    for f in page.frames:
        try:
            if _try_on(f):
                log(f"[location] Filled Location inside frame URL={f.url}")
                return
        except Exception:
            continue

    # 3) If still not found, dump a hint so we’re not blind
    log("[warn] Location field (nodeUserVal2) not found in any frame; dumping DOM for debugging")
    html_path = os.path.join(
        ARTIFACT_DIR, f"missing_location_{int(time.time())}.html"
    )
    _safe_write(html_path, page.content())
    shot_path = os.path.join(
        ARTIFACT_DIR, f"missing_location_{int(time.time())}.png"
    )
    with swallow("screenshot missing location"):
        page.screenshot(path=shot_path, full_page=True)
        log(f"[dump] screenshot (missing location): {_abs(shot_path)}")

# --- ROSTER PARSING & FILTERING ---

def _norm_unit(u: str | None) -> str:
    return (u or "").strip().upper()

def extract_personnel_with_units(roster_json: Any) -> list[dict]:
    """
    Your schema: an array of unit objects like:
      { "unit": "R26", "staff": [ {"name": "Pruett, William"}, ... ] }
    Returns: [{ "name": "Last, First", "unit": "R26" }, ...]
    """
    out: list[dict] = []

    def maybe_add(n: str | None, u: str | None):
        n = (n or "").strip()
        u = _norm_unit(u)
        if n and u:
            out.append({"name": n, "unit": u})

    def walk(o: Any):
        if isinstance(o, list):
            for x in o: walk(x)
        elif isinstance(o, dict):
            unit = o.get("unit") or o.get("Unit")
            # staff list (your example)
            if isinstance(o.get("staff"), list):
                for p in o["staff"]:
                    if isinstance(p, dict):
                        maybe_add(p.get("name") or p.get("Name"), unit)
            # legacy fallbacks (doesn't hurt to keep)
            for k in ("Personnel", "UnitPersonnel", "Crew", "Members", "Staff"):
                if isinstance(o.get(k), list):
                    for p in o[k]:
                        if isinstance(p, dict):
                            maybe_add(p.get("name") or p.get("Name") or p.get("FullName") or p.get("LastFirst"), unit)
                        elif isinstance(p, str):
                            maybe_add(p, unit)
            for v in o.values():
                if isinstance(v, (dict, list)):
                    walk(v)

    walk(roster_json)

    # de-dupe by (name, unit)
    seen = set()
    uniq: list[dict] = []
    for r in out:
        key = (r["name"], r["unit"])
        if key not in seen:
            seen.add(key)
            uniq.append(r)

    # debug: what units did we see?
    from collections import Counter
    c = Counter([r["unit"] for r in uniq])
    log("Roster units seen: " + ", ".join(f"{u}:{c[u]}" for u in sorted(c)))
    return uniq

def filter_personnel_by_units(personnel: list[dict], include_units: list[str] | None) -> list[dict]:
    if not include_units:
        return personnel
    wanted = {_norm_unit(u) for u in include_units}
    keep = [p for p in personnel if p["unit"] in wanted]
    # STRICT: if nothing matched, do NOT fall back to unfiltered
    if not keep:
        raise RuntimeError(f"No roster entries match units {sorted(wanted)}. Check roster.json 'unit' values.")
    from collections import Counter
    c = Counter([p["unit"] for p in keep])
    log("Unit filter -> " + ", ".join(f"{u}:{c[u]}" for u in sorted(c)))
    return keep

def _extract_unit_from_item(item: dict) -> str:
    # Your real JSON uses 'unit'
    for key in ("unit", "Unit", "UnitID", "UnitName", "UnitNumber", "Apparatus", "Resource", "CallSign"):
        if key in item and item[key]:
            return _norm_unit(str(item[key]))
    # Try sniffing unit from name text if someone jammed it in there
    for key in ("name", "Name", "FullName", "LastFirst", "full_name", "last_first"):
        v = item.get(key)
        if isinstance(v, str):
            m = re.search(r"\b([REL]\d{1,2})\b", v.upper())  # R26, R1, L26
            if m:
                return _norm_unit(m.group(1))
    return ""

def fetch_minicad_roster(
    url: Optional[str],
    user: Optional[str],
    password: Optional[str],
    roster_cmd: Optional[str] = None,
    include_units: Optional[List[str]] = None,
    timeout: int = 60
) -> Dict[str, Any]:
    if roster_cmd:
        data = fetch_roster_via_command(roster_cmd, timeout=timeout)
    else:
        if not (url and user and password):
            raise RuntimeError("Provide either --roster-cmd or MiniCAD --minicad-url/--minicad-user/--minicad-pass")
        data = _fetch_minicad_roster_basic(url, user, password, timeout=timeout)
    # Optional unit filtering
    if include_units:
        data = filter_roster_by_units(data, include_units)
    out_path = os.path.join(ARTIFACT_DIR, f"minicad_roster_filtered_{int(time.time())}.json")
    write_json(out_path, data)
    log(f"Filtered roster written to {out_path}")
    return data

def _fetch_minicad_roster_basic(url: str, user: str, password: str, timeout: int = 20) -> Dict[str, Any]:
    import base64, urllib.request
    req = urllib.request.Request(url)
    token = base64.b64encode(f"{user}:{password}".encode()).decode()
    req.add_header("Authorization", f"Basic {token}")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read().decode("utf-8")
    try:
        parsed = json.loads(data)
    except Exception:
        parsed = json.loads(data.strip().rstrip(";"))
    out_path = os.path.join(ARTIFACT_DIR, f"minicad_roster_{int(time.time())}.json")
    write_json(out_path, parsed)
    log(f"Fetched roster (basic) to {out_path}")
    return parsed

def extract_names_from_roster(minicad_json: Any, include_units: Optional[List[str]] = None) -> List[str]:
    if include_units:
        minicad_json = filter_roster_by_units(minicad_json, include_units)

    names: List[str] = []

    def maybe_add(n: Optional[str]):
        n = sanitize(n)
        if n:
            names.append(ensure_last_first(n))

    def walk(obj: Any):
        if isinstance(obj, dict):
            # If a dict clearly looks like a person record
            person_keys = {"name", "fullname", "lastfirst", "last_name", "first_name"}
            for k, v in obj.items():
                lk = k.lower()
                if lk in person_keys and isinstance(v, str):
                    maybe_add(v)
                elif lk in {"personnel", "crew", "members", "unitpersonnel", "staff"} and isinstance(v, list):
                    for p in v:
                        walk(p)
                elif isinstance(v, (dict, list)):
                    walk(v)
        elif isinstance(obj, list):
            for x in obj:
                walk(x)

    walk(minicad_json)

    # Dedup
    seen, unique = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            unique.append(n)
    log(f"Extracted {len(unique)} unique personnel from roster")
    return unique

# --- ADD USERS (LAST NAME SEARCH) ---

def _last_name(s: str) -> str:
    """
    Extract a clean last name from strings like:
      - "Pruett Jr, William"
      - "Pruett, William Jr."
      - "William J Pruett Jr"
    Strips common suffixes (Jr, Sr, II, III, IV) off the last token.
    """
    s = (s or "").strip()
    if not s:
        return s

    # "Last, First ..." vs "First M Last Jr"
    if "," in s:
        last_part = s.split(",", 1)[0].strip()
    else:
        parts = s.split()
        last_part = parts[-1].strip() if parts else s

    SUFFIXES = {"JR", "SR", "II", "III", "IV", "JR.", "SR."}

    tokens = last_part.replace(".", " ").split()
    if tokens and tokens[-1].upper() in SUFFIXES and len(tokens) > 1:
        tokens = tokens[:-1]

    cleaned = " ".join(tokens).strip()
    return cleaned or last_part

def _split_first_last(raw: str) -> tuple[str, str]:
    """
    Best-effort split like:
      'WILLIAMS, JOHN A'  -> ('John', 'Williams')
      'John A Williams Jr' -> ('John', 'Williams')
    Uses the same suffix stripping logic as _last_name.
    """
    raw = (raw or "").strip()
    if not raw:
        return "", ""

    # Normalize spacing
    s = " ".join(raw.replace(",", " , ").split())

    # Handle "LAST, FIRST ..." format
    if "," in s:
        last_part, rest = s.split(",", 1)
        last = _last_name(last_part.strip())
        rest = rest.strip()
        if not rest:
            return "", last
        first = rest.split()[0]  # ignore middle / extra
        return first, last

    # Otherwise "First M Last Jr" style
    parts = s.split()
    if len(parts) == 1:
        return parts[0], _last_name(parts[0])
    first = parts[0]
    last = _last_name(" ".join(parts[1:]))
    return first, last

def _first_name_from_raw(raw: str) -> str:
    """
    Best-effort grab of FIRST name from roster string.
    Examples:
      "WILLIAMS, JOHN A" -> "JOHN"
      "John A Williams"  -> "John"
    """
    raw = (raw or "").strip()
    if not raw:
        return ""
    if "," in raw:
        # "Last, First ..."
        parts = raw.split(",", 1)[1].strip().split()
        return parts[0] if parts else ""
    parts = raw.split()
    if len(parts) >= 2:
        return parts[0]
    return ""

def choose_users_and_continue(page, names: list[str]):
    """
    On the 'Save and Add Users' screen:
      - For each roster name, search by multiple variants:
          * "Last, First"
          * "First Last"
          * raw string
          * last name only (suffix-stripped)
      - Click a matching entry in the available/middle list
      - If multiple matches (e.g. several 'Williams'), prefer the one whose
        text also contains the FIRST name from the roster.
      - Then click Continue to return to the form
    """
    log(f"'Add Users' page: targeting {len(names)} personnel (multi-name search, click-to-add)")

    import re

    # Search box
    search = page.locator("input[placeholder*='Search' i], input[type='search']")
    if not search.count():
        search = page.locator("label:has-text('Search')").locator("xpath=following::input[1]")

    # Try to detect the "available" (middle) list and "selected" list, but keep it best-effort
    available = page.locator(
        "[aria-label*='available' i], [aria-label*='middle' i], "
        "[role='listbox'], .listbox, .available-list, .middle-list"
    ).first
    selected = page.locator(
        "[aria-label*='selected' i], .selected-list, [data-selected-list]"
    ).first

    added = 0
    misses: list[str] = []

    for raw in names:
        raw = (raw or "").strip()
        if not raw:
            continue

        first_name = _first_name_from_raw(raw)
        first_upper = first_name.upper() if first_name else ""

        # Build variants: raw, "Last, First", "First Last", last name only
        variants = set()
        variants.add(raw)

        # split “Last, First” vs “First Last”
        if "," in raw:
            last, first = [p.strip() for p in raw.split(",", 1)]
            if first and last:
                variants.add(f"{last}, {first}")
                variants.add(f"{first} {last}")
        else:
            parts = raw.split()
            if len(parts) >= 2:
                first, last = parts[0], parts[-1]
                variants.add(f"{last}, {first}")
                variants.add(f"{first} {last}")

        # Always throw in last-name-only version
        last_only = _last_name(raw)
        if last_only:
            variants.add(last_only)

        success = False
        used_term = None

        for variant in variants:
            variant = variant.strip()
            if not variant:
                continue

            # For each variant, we try both the full variant and its last name
            trial_terms = []
            trial_terms.append(variant)
            lo = _last_name(variant)
            if lo and lo not in trial_terms:
                trial_terms.append(lo)

            for term in trial_terms:
                term = term.strip()
                if not term:
                    continue

                # Type into search box if present
                if search.count():
                    with swallow("search input fill"):
                        search.first.fill("")
                        search.first.type(term, delay=20)
                        wait_network_quiet(page, 700)

                pattern = re.compile(re.escape(term), re.I)

                # Prefer matches inside the available/middle pane if we have one
                if available.count():
                    cand = available.get_by_text(pattern, exact=False)
                else:
                    cand = page.get_by_text(pattern, exact=False)

                if not cand.count():
                    continue  # try next term / variant

                # >>> NEW: if we got multiple candidates and we know FIRST name,
                # prefer the one whose text contains that first name.
                target = None
                count = cand.count()
                if count > 1 and first_upper:
                    for i in range(count):
                        try:
                            txt = (cand.nth(i).inner_text() or "").strip()
                        except Exception:
                            continue
                        if first_upper in txt.upper():
                            target = cand.nth(i)
                            break

                # Fallback to old behavior if we didn't narrow it down
                if target is None:
                    target = cand.first

                # We found a candidate; click to add
                with swallow("scroll name into view"):
                    target.scroll_into_view_if_needed(timeout=4000)

                with swallow("click name to add"):
                    target.click(timeout=6000)
                    time.sleep(0.25)

                # If we can see a selected pane, verify it appeared there
                moved = False
                if selected.count():
                    sel_cand = selected.get_by_text(pattern, exact=False)
                    moved = sel_cand.count() > 0
                else:
                    moved = True  # no selected pane to confirm; assume success

                # Fallback: double-click if single-click didn’t move it
                if not moved:
                    with swallow("double-click user"):
                        target.dblclick(timeout=4000)
                        time.sleep(0.25)
                    if selected.count():
                        sel_cand = selected.get_by_text(pattern, exact=False)
                        moved = sel_cand.count() > 0

                if moved:
                    added += 1
                    success = True
                    used_term = term
                    log(f"[users] Added '{raw}' using search term '{term}'")
                    break  # break trial_terms

            if success:
                break  # break variants

        if not success:
            misses.append(raw)
            log(f"[warn] Could not add user '{raw}' (no match after all variants)")

    log(f"[users] Added {added}/{len(names)} personnel on Add Users page")
    if misses:
        log("[users] Still missing: " + "; ".join(misses[:12]) + (" ..." if len(misses) > 12 else ""))

    # Always return to the form
    _click_continue_from_add_users(page)

def add_user_by_last_first(page, last_first):
    f = within_users_frame(page)

    # Search box
    search = f.locator('input[placeholder*="Search" i], input[name*="search" i]').first
    search.fill(last_first)
    search.press("Enter")

    # Wait for results to render
    f.wait_for_timeout(400)  # tiny debounce

    # “Middle list” heuristic: a list between two sidebars; grab the center column list by structure.
    # Fallback: click the exact text.
    try:
        f.get_by_role("option", name=last_first).first.click(timeout=3000)
    except Exception:
        f.locator(f'text="{last_first}"').first.click(timeout=3000)

    # Add/continue if there’s an explicit button
    try:
        f.get_by_role("button", name="Continue").click(timeout=3000)
    except Exception:
        f.locator('input[type="button"][value="Continue"]').first.click(timeout=3000)

def _click_continue_from_add_users(page):
    log("Clicking Continue to return to form")
    cont = page.get_by_role("button", name=re.compile(r"Continue|Back to Form|Return", re.I))
    if not cont.count():
        cont = page.locator("button:has-text('Continue'), a:has-text('Continue')")
    cont.first.click(timeout=12000)
    wait_network_quiet(page, 8000)
    log("Returned to training form")

@contextmanager
def swallow(title: str):
    try:
        yield
    except Exception:
        log(f"[warn] {title}: {traceback.format_exc().strip()[:800]}")

# -----------------------------
# DATA SOURCES
# -----------------------------

def load_scenarios_from_csv(csv_path: str) -> List[Dict[str, str]]:
    df = pd.read_csv(csv_path)
    # Normalize expected headers
    rename_map = {c: c.strip() for c in df.columns}
    df.rename(columns=rename_map, inplace=True)

    missing = [h for h in CSV_HEADERS if h not in df.columns]
    if missing:
        raise ValueError(f"CSV missing required headers: {missing}")

    # Fill defaults if empty
    if "Duration" in df.columns:
        df["Duration"] = df["Duration"].fillna(DEFAULT_DURATION)
    if "Instructor" in df.columns:
        df["Instructor"] = df["Instructor"].fillna(DEFAULT_INSTRUCTOR)

    scenarios = df[CSV_HEADERS].fillna("").to_dict(orient="records")

    # 🔀 Randomize order so each run rotates scenarios
    random.shuffle(scenarios)

    log(f"Loaded {len(scenarios)} scenarios from CSV (shuffled order)")
    return scenarios

def load_scenarios_from_docx(docx_path: str) -> List[Dict[str, str]]:
    if DocxDocument is None:
        raise RuntimeError("python-docx not installed; cannot parse DOCX")

    doc = DocxDocument(docx_path)
    scenarios: List[Dict[str, str]] = []
    for p in doc.paragraphs:
        line = p.text.strip()
        if not line:
            continue
        # Expect pipe-delimited: Location | Checkbox Label | Description | Duration | Instructor
        parts = [x.strip() for x in line.split("|")]
        if len(parts) >= 5:
            scenarios.append({
                "Location": parts[0],
                "Checkbox Label": parts[1],
                "Description": parts[2],
                "Duration": parts[3] or DEFAULT_DURATION,
                "Instructor": parts[4] or DEFAULT_INSTRUCTOR,
            })
    log(f"Loaded {len(scenarios)} scenarios from DOCX")
    return scenarios

# -----------------------------
# PLAYWRIGHT FLOWS
# -----------------------------

def wait_network_quiet(page, timeout=8000):
    with swallow("wait networkidle"):
        page.wait_for_load_state("networkidle", timeout=timeout)

def perform_vector_login(page, login_url, username, password):
    print("[info] Navigating to Vector login...", flush=True)

    page.goto(login_url, wait_until="domcontentloaded")
    debug_snap(page, "loaded-login")

    # If we somehow landed on about:blank, reload once
    if page.url == "about:blank":
        print("[warn] Landed on about:blank; reloading login_url", flush=True)
        page.goto(login_url, wait_until="domcontentloaded")
        debug_snap(page, "reloaded-login")

    # Try a few common username/password selector patterns
    user_input = page.locator(
        'input[name="username"], input[name="j_username"], input#username, input[formcontrolname="username"]'
    ).first
    pass_input = page.locator(
        'input[name="password"], input[name="j_password"], input#password, input[formcontrolname="password"]'
    ).first

    if not user_input.count() or not pass_input.count():
        debug_snap(page, "login-no-fields")
        raise RuntimeError("Could not find username/password fields on login page")

    user_input.fill(username)
    pass_input.fill(password)

    # Find some sort of login button
    try:
        page.get_by_role("button", name=re.compile("log.?in", re.I)).first.click(timeout=5000)
    except Exception:
        # Fallback: look for input[type=submit] with 'Log' in value
        btn = page.locator(
            'button[type="submit"], input[type="submit"], input[type="button"]'
        ).filter(has_text=re.compile("log.?in", re.I))
        if btn.count():
            btn.first.click()
        else:
            # last resort: click the first submit/input button
            page.locator('button[type="submit"], input[type="submit"]').first.click()

    page.wait_for_load_state("domcontentloaded", timeout=15000)
    debug_snap(page, "after-login-submit")

    # Simple sanity check: not still on login page
    if any(token in (page.url or "").lower() for token in ["login", "signin"]):
        print("[warn] Still appear to be on a login page after submit.", flush=True)

def do_login(page, vs_user: str, vs_pass: str):
    # Handle generic TargetSolutions login. Your environment may SSO; if so, adapt selectors.
    log("Attempting Vector login")
    page.goto(VS_LOGIN_URL, wait_until="domcontentloaded", timeout=45000)
    wait_network_quiet(page, 20000)

    # Try several common selector sets
    candidates = [
        # Classic fields
        ("input[name='username']", "input[name='password']", "button[type='submit']"),
        # Alternate custom elements
        ("input#username", "input#password", "button#loginButton"),
        # Possible TS web components
        ("input[data-test='username']", "input[data-test='password']", "button[data-test='login']"),
    ]

    for u_sel, p_sel, s_sel in candidates:
        with swallow(f"login set {u_sel},{p_sel},{s_sel}"):
            user_box = page.locator(u_sel)
            pass_box = page.locator(p_sel)
            submit = page.locator(s_sel)
            if user_box.count() and pass_box.count():
                user_box.fill(vs_user, timeout=15000)
                pass_box.fill(vs_pass, timeout=15000)
                if submit.count():
                    submit.first.click()
                else:
                    pass_box.press("Enter")
                # Wait for navigation or dashboard token
                try:
                    page.wait_for_url(re.compile("dashboard|home", re.I), timeout=45000)
                except PWTimeout:
                    wait_network_quiet(page, 15000)
                    log("Login attempt finished, checking dashboard...")
                break

def _visible_anchor_dump(page):
    # log every visible anchor containing 'Company'
    try:
        hrefs = page.evaluate("""
            () => Array.from(document.querySelectorAll('a'))
                .filter(a => a && a.offsetParent !== null)
                .map(a => ({text:(a.innerText||'').trim(), href:a.href||a.getAttribute('href')||''}))
                .filter(x => /company/i.test(x.text))
        """)
        for h in hrefs[:30]:
            log(f"[anchor] {h.get('text')} -> {h.get('href')}")
    except Exception:
        pass

def _click_nearest_clickable(node_locator):
    # climb up to something actually clickable
    candidates = node_locator.locator("closest=a,closest=button,closest=[role='button'],closest=[role='link'],closest=vwc-card,closest=vwc-tiling-grid-tile,closest=div")
    for i in range(min(candidates.count(), 1)):
        try:
            candidates.nth(i).click(timeout=4000)
            return True
        except Exception:
            pass
    # last resort: force click original node
    try:
        node_locator.first.click(timeout=4000, force=True)
        return True
    except Exception:
        return False

def _try_click(locator, label):
    if locator.count():
        locator.first.scroll_into_view_if_needed(timeout=6000)
        locator.first.click(timeout=8000)
        wait_network_quiet(page=locator.page, timeout=10000)
        log(f"Clicked via {label}")
        return True
    return False

def goto_dashboard_open_form(page, timeout_sec: int = 30):
    """
    From the dashboard, find and click *some* entry point that clearly opens
    the company training form. Tries text on links/buttons first, then the
    brittle FORM_OPEN_SELECTOR, and loops for up to timeout_sec.
    """
    log("Navigating to training form from dashboard")
    deadline = time.time() + timeout_sec
    last_err = None

    # Regex for any reasonable text on that tile/button
    patt = re.compile(
        r"(company\s*training|training\s*form|training\s*record|enter\s*drill|new\s*record)",
        re.I,
    )

    while time.time() < deadline:
        wait_network_quiet(page, 5000)

        # 1) Try all visible links/buttons/role=button elements by text
        try:
            candidates = page.locator("a, button, [role='button']")
            count = min(candidates.count(), 200)
            for i in range(count):
                el = candidates.nth(i)
                try:
                    txt = (el.inner_text() or "").strip()
                except Exception as e:
                    last_err = e
                    continue

                if not txt:
                    continue

                if patt.search(txt):
                    log(f"Clicking training entry control with text: '{txt}'")
                    el.scroll_into_view_if_needed(timeout=4000)
                    el.click(timeout=8000)
                    wait_network_quiet(page, 10000)
                    return
        except Exception as e:
            last_err = e

        # 2) Fallback: your old CSS tile selector, if it exists
        try:
            loc = page.locator(FORM_OPEN_SELECTOR)
            if loc.count():
                log("Using FORM_OPEN_SELECTOR fallback to open training form")
                loc.first.scroll_into_view_if_needed(timeout=4000)
                loc.first.click(timeout=8000)
                wait_network_quiet(page, 10000)
                return
        except Exception as e:
            last_err = e

        time.sleep(0.5)

    # If we get here, nothing matched: dump HTML + screenshot
    html_path = os.path.join(
        ARTIFACT_DIR, f"missing_training_entry_{int(time.time())}.html"
    )
    _safe_write(html_path, page.content())
    shot_path = os.path.join(
        ARTIFACT_DIR, f"missing_training_entry_{int(time.time())}.png"
    )
    with swallow("screenshot missing training entry"):
        page.screenshot(path=shot_path, full_page=True)
        log(f"[dump] screenshot: {_abs(shot_path)}")

    raise RuntimeError(
        f"Could not locate training entry control after {timeout_sec}s. "
        f"Last error: {last_err}. Saved DOM: {html_path}"
    )

def go_to_company_training_from_dashboard(page):
    # Try the explicit CSS you shared first. If it’s brittle, fall back to link text.
    try:
        page.locator(
            'vwc-tiling-grid-tile.bulletin-board.is-schedule vwc-card div.bulletinBoard-content p a img'
        ).first.click(timeout=4000)
    except Exception:
        # Fallback: clickable text somewhere around "Company Training"
        page.get_by_role("link", name=lambda n: "training" in n.lower()).first.click(timeout=6000)

    # Now wait for a stable element on the training form page
    wait_for_training_form_ready(page)

def wait_for_training_form_ready(page):
    # Pick something unique on the training form (tweak if your form IDs differ)
    # Example: the big header or any required field that always exists
    # Avoid networkidle; the app probably keeps polling. Wait for a form control we actually use.
    page.locator('text=Company Training').first.wait_for(timeout=12000)

def click_company_training_tile(page):
    """
    Click the Company Training tile reliably.
    1) Try the exact CSS you provided (anchor parent of the img).
    2) Try a CSS that targets the paragraph/link by text.
    3) Try component-aware selectors.
    4) Last resort: JS querySelector on the exact path.
    Saves HTML/screenshot if it still fails.
    """
    log("Hunting the Company Training tile...")

    # 1) Your exact JS path, converted to a Playwright locator
    exact_css = (
        "#single-spa-application\\:\\@target-solutions\\/home > section > section > "
        "vwc-tiling-grid > vwc-tiling-grid-tile.bulletin-board.is-schedule > vwc-card > "
        "div.bulletinBoard-content > div > p:nth-child(13) > a > img"
    )

    try:
        img = page.locator(exact_css)
        if img.count():
            img.first.scroll_into_view_if_needed(timeout=5000)
            link = img.first.locator("xpath=ancestor::a[1]")
            if link.count():
                link.first.click(timeout=8000)
                wait_network_quiet(page, 8000)
                log("Clicked Company Training via exact CSS path (anchor parent).")
                return
            # If no anchor, force-click the image
            img.first.click(timeout=8000, force=True)
            wait_network_quiet(page, 8000)
            log("Force-clicked Company Training image via exact CSS.")
            return
    except Exception:
        pass

    # 2) Slightly less brittle: find the <p> with the text then its <a>
    try:
        p = page.locator("p:has-text('Company Training')")
        if p.count():
            a = p.first.locator("a")
            if a.count():
                a.first.scroll_into_view_if_needed(timeout=5000)
                a.first.click(timeout=8000)
                wait_network_quiet(page, 8000)
                log("Clicked Company Training via paragraph text + link.")
                return
    except Exception:
        pass

    # 3) Component-aware: click the tile/card that contains the text
    try:
        tile = page.locator("vwc-tiling-grid-tile.bulletin-board.is-schedule:has-text('Company Training')")
        if tile.count():
            tile.first.click(timeout=8000)
            wait_network_quiet(page, 8000)
            log("Clicked Company Training via vwc-tiling-grid-tile.")
            return
        card = page.locator("vwc-card:has-text('Company Training')")
        if card.count():
            card.first.click(timeout=8000)
            wait_network_quiet(page, 8000)
            log("Clicked Company Training via vwc-card.")
            return
    except Exception:
        pass

    # 4) Nuclear option: JS querySelector and click exactly what you pasted
    try:
        found = page.evaluate("""() => {
            const sel = "#single-spa-application\\\\:\\\\@target-solutions\\\\/home > section > section > vwc-tiling-grid > vwc-tiling-grid-tile.bulletin-board.is-schedule > vwc-card > div.bulletinBoard-content > div > p:nth-child(13) > a > img";
            const img = document.querySelector(sel);
            if (!img) return false;
            const a = img.closest('a') || img;
            a.scrollIntoView({block: 'center', inline: 'center'});
            a.click();
            return true;
        }""")
        if found:
            wait_network_quiet(page, 8000)
            log("Clicked Company Training via JS querySelector.")
            return
    except Exception:
        pass

    # Dump evidence so we can target exactly what’s there
    html_path = os.path.join(ARTIFACT_DIR, f"missing_company_training_{int(time.time())}.html")
    html_path = _safe_write(html_path, page.content())
    shot_path = os.path.join(ARTIFACT_DIR, f"missing_company_training_{int(time.time())}.png")
    with swallow("screenshot missing company training"):
        page.screenshot(path=shot_path, full_page=True)
        log(f"[dump] screenshot: {_abs(shot_path)}")

    raise RuntimeError(f"Could not click Company Training. Saved DOM: {html_path}")

def fill_training_form(page, scenario: Dict[str, str]):
    """
    Expect fields such as:
      - Location: maybe an input or select
      - Checkbox Label: one or more checkboxes matching given label text
      - Description: textarea
      - Duration: select with '8:00 AM' style or freeform duration field
      - Instructor: input
    """
    log(f"Filling training form for: {scenario}")
    loc = sanitize(scenario.get("Location"))
    label = sanitize(scenario.get("Checkbox Label"))
    desc = sanitize(scenario.get("Description"))
    dur = sanitize(scenario.get("Duration") or DEFAULT_DURATION)
    datetime = sanitize(scenario.get("Date/Time"))
    inst = sanitize(scenario.get("Instructor") or DEFAULT_INSTRUCTOR)

    # Try common patterns
    # Location
    with swallow("fill Location"):
        candidates = [
            ("input[name*='location' i]", None),
            ("input[id*='location' i]", None),
            ("textarea[name*='nodeUserVal2' i]", None),
            ("textarea[id*='nodeUserVal2' i]", None),
        ]
        for sel, _ in candidates:
            if page.locator(sel).count():
                page.locator(sel).first.fill(loc)
                break

    # Checkbox Label(s): may be multiple comma-separated
    def _norm(s: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()

    CHECKBOX_SYNONYMS = {
        "ves ventilation": ["ves", "vent enter search"],
        "fire streams and nozzles": ["fire streams", "nozzles"],
        "fire behaviour": ["fire behavior"],  # spelling drift
    }
    with swallow("click Checkbox Label"):
        wanted_raw = sanitize(label)
        labels = [l.strip() for l in re.split(r"[;,/]+", wanted_raw) if l.strip()] or [wanted_raw]
        for lbl in labels:
            targets = [_norm(lbl)]
        # add synonyms
            for canon, alts in CHECKBOX_SYNONYMS.items():
                if _norm(lbl) == canon or _norm(lbl) in [_norm(x) for x in alts]:
                    targets.extend([canon] + alts)

        # try role-based match first
            chk = page.get_by_role("checkbox", name=re.compile(re.escape(lbl), re.I))
            if chk.count():
                chk.first.check()
                continue

        # scan all labels and match by normalized text
            candidate_labels = page.locator("label")
            hit = False
            for i in range(min(candidate_labels.count(), 600)):
                lab = candidate_labels.nth(i)
                try:
                    t = _norm(lab.inner_text())
                except Exception:
                    t = ""
                if any(tt in t for tt in targets):
                    with swallow("click label"):
                        lab.click()
                        hit = True
                        break
            if not hit:
                log(f"[warn] Checkbox label not found after normalization: {lbl} (norm={_norm(lbl)})")

    # Description
    with swallow("fill Description"):
        desc_field = None
        for sel in ["textarea[name*='description' i]", "textarea[id*='description' i]", "textarea"]:
            if page.locator(sel).count():
                desc_field = page.locator(sel).first
                break
        if desc_field:
            desc_field.fill(desc)

    # Duration: try select time menus or freeform duration field
    with swallow("set Duration"):
        # Freeform if present
        dur_inputs = page.locator("input[name*='duration' i], input[id*='duration' i]")
        if dur_inputs.count():
            dur_inputs.first.fill(dur)
        else:
            # Some forms have completion time selects; we won't fight time math here. Best effort:
            # If there's a select with entries like "2 hours", pick by label text.
            selects = page.locator("select")
            chosen = False
            for i in range(min(selects.count(), 10)):
                sel = selects.nth(i)
                try:
                    options = [o.text_content().strip() for o in sel.locator("option").all()]
                except Exception:
                    options = []
                for opt in options:
                    if re.search(r"\b2\s*hours?\b", opt, re.I) and re.search(r"\b2\s*hours?\b", dur, re.I):
                        sel.select_option(label=opt)
                        chosen = True
                        break
                if chosen:
                    break

    # Date/Time
    with swallow("set Date/Time"):
        dt_inputs = page.locator("input[name*='date' i], input[id*='date' i], input[name*='time' i], input[id*='time' i]")
        if dt_inputs.count():
            dt_inputs.first.fill(datetime)

    # Instructor
    with swallow("fill Instructor"):
        instr_candidates = [
            "input[name*='instructor' i]",
            "input[id*='instructor' i]",
            "input[placeholder*='Instructor' i]",
        ]
        for sel in instr_candidates:
            if page.locator(sel).count():
                page.locator(sel).first.fill(inst)
                break

def wait_for_training_form(page, timeout_ms: int = 30000) -> None:
    deadline = time.time() + (timeout_ms / 1000.0)
    last_err = None
    while time.time() < deadline:
        for probe in FORM_PROBES:
            try:
                if "role" in probe:
                    role, name = probe["role"]
                    if page.get_by_role(role, name=name).first.count():
                        return
                elif "css" in probe:
                    if page.locator(probe["css"]).first.count():
                        return
                elif "text" in probe:
                    if page.get_by_text(probe["text"]).first.count():
                        return
            except Exception as e:
                last_err = e
        # small wait and keep probing
        time.sleep(0.35)
    raise RuntimeError(f"Training form did not present expected controls within {timeout_ms} ms. Last error: {last_err}")

def click_save_and_add_users(page) -> None:
    log("Clicking 'Save and Add Users'")
    patt = re.compile(r"save\s*&?\s*and\s*add\s*users", re.I)

    # 1) role=button
    btns = page.get_by_role("button")
    for i in range(min(btns.count(), 50)):
        b = btns.nth(i)
        with swallow("check button text"):
            txt = (b.inner_text() or "").strip()
            if patt.search(txt):
                b.click()
                wait_network_quiet(page, 20000)
                return

    # 2) generic clickable nodes
    nodes = page.locator("button, a, [role='button'], .btn, .button")
    for i in range(min(nodes.count(), 120)):
        n = nodes.nth(i)
        with swallow("check node text"):
            txt = (n.inner_text() or "").strip()
            if patt.search(txt):
                n.click()
                wait_network_quiet(page, 20000)
                return

    # 3) last resort: evaluate all text nodes once and click by JS
    with swallow("querySelectorAll text scan"):
        handle = page.evaluate_handle("""
            () => {
              const patt = /save\\s*&?\\s*and\\s*add\\s*users/i;
              const candidates = Array.from(document.querySelectorAll('button, a, [role="button"], .btn, .button, *'));
              for (const el of candidates) {
                const t = (el.innerText || '').trim();
                if (patt.test(t)) return el;
              }
              return null;
            }
        """)
        if handle:
            page.evaluate("(el)=>el.click()", handle)
            wait_network_quiet(page, 20000)
            return

    # dump page for debugging
    html_path = os.path.join(ARTIFACT_DIR, f"missing_save_add_users_{int(time.time())}.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(page.content())
    shot_path = os.path.join(ARTIFACT_DIR, f"missing_save_add_users_{int(time.time())}.png")
    with swallow("screenshot on missing Save and Add Users"):
        page.screenshot(path=shot_path, full_page=True)
    raise RuntimeError(f"Could not find 'Save and Add Users' control. Saved DOM: {html_path}, screenshot: {shot_path}")

def _name_variants(last_first: str) -> List[str]:
    # "Pruett, William" -> ["Pruett, William", "William Pruett", "William J Pruett"] if middle initial exists in source
    s = last_first.strip()
    if "," in s:
        last, first = [x.strip() for x in s.split(",", 1)]
    else:
        parts = s.split()
        if len(parts) >= 2:
            first, last = parts[0], parts[-1]
        else:
            return [s]
    variants = [f"{last}, {first}", f"{first} {last}"]
    # include middle initial variant if present like "Roy J."
    mid = re.search(r",\s*([A-Za-z]+)\s+([A-Za-z]\.)", s)
    if mid:
        first2, mi = mid.group(1), mid.group(2).rstrip(".")
        variants.append(f"{first2} {mi} {last}")
    return list(dict.fromkeys(variants))  # de-dupe, keep order

def _xpath_literal(s: str) -> str:
    # Build a valid XPath string literal even if s has both ' and "
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    parts = s.split("'")
    return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"

def _find_submit_anywhere(page):
    # 1) Your exact element by id (fast path)
    btn = page.locator("#submitBtn")
    if btn.count():
        return btn.first, page

    # 2) Your full XPath on the current page
    xp = "xpath=/html/body/div[1]/div/div/div[2]/div[2]/div/div[1]/div[1]/form/div[3]/input[1]"
    btn = page.locator(xp)
    if btn.count():
        return btn.first, page

    # 3) Look inside frames for either id or your XPath
    for f in page.frames:
        try:
            if f == page.main_frame:
                continue
        except Exception:
            pass
        b = f.locator("#submitBtn")
        if b.count():
            return b.first, f.page
        b = f.locator(xp)
        if b.count():
            return b.first, f.page

    # 4) Attribute match as last resort
    b = page.locator("input[type='button'][name='complete'][value='Submit as Complete']")
    if b.count():
        return b.first, page
    for f in page.frames:
        b = f.locator("input[type='button'][name='complete'][value='Submit as Complete']")
        if b.count():
            return b.first, f.page

    return None, None

def _find_submit_handle(page) -> ElementHandle | None:
    # A) direct ID
    h = page.locator("#submitBtn")
    if h.count(): return h.first.element_handle()
    # B) attributes on this page
    h = page.locator("input[type='button'][name='complete'][value='Submit as Complete']")
    if h.count(): return h.first.element_handle()
    # C) your full XPath on this page
    h = page.locator(f"xpath={SUBMIT_XPATH}")
    if h.count(): return h.first.element_handle()

    # D) frames: try the same three inside each frame
    for f in page.frames:
        if f == page.main_frame: 
            continue
        hh = f.locator("#submitBtn")
        if hh.count(): return hh.first.element_handle()
        hh = f.locator("input[type='button'][name='complete'][value='Submit as Complete']")
        if hh.count(): return hh.first.element_handle()
        hh = f.locator(f"xpath={SUBMIT_XPATH}")
        if hh.count(): return hh.first.element_handle()

    # E) shadow DOM deep scan: find an element whose value/text matches
    handle = page.evaluate_handle("""
        () => {
          const wants = (el) => {
            const v = (el.value || el.textContent || '').trim();
            return /submit\\s*(as)?\\s*complete/i.test(v);
          };
          function* walk(node) {
            if (!node) return;
            yield node;
            if (node.shadowRoot) {
              for (const c of node.shadowRoot.children) yield* walk(c);
            }
            for (const c of node.children || []) yield* walk(c);
          }
          for (const n of walk(document.documentElement)) {
            if (n.tagName && (n.tagName === 'INPUT' || n.tagName === 'BUTTON')) {
              if (wants(n)) return n;
            }
          }
          return null;
        }
    """)
    el = handle.as_element()
    return el

def submit_training(page) -> str:
    log("Submitting training record (id/xpath/attr search)")
    # try id
    btn = page.locator("#submitBtn")
    if not btn.count():
        # full XPath you provided
        btn = page.locator("xpath=/html/body/div[1]/div/div/div[2]/div[2]/div/div[1]/div[1]/form/div[3]/input[1]")
    if not btn.count():
        # attributes
        btn = page.locator("input[type='button'][name='complete'][value='Submit as Complete']")
    if not btn.count():
        # quick dump to artifact so we can key it exactly if this still fails
        _dump_submit_candidates(page)
        raise RuntimeError("Submit button not found with #submitBtn or provided XPath.")

    btn.first.scroll_into_view_if_needed(timeout=4000)
    btn.first.click(timeout=12000)
    wait_network_quiet(page, 12000)
    return "Submitted"

def _dump_submit_candidates(page):
    try:
        out = page.evaluate("""
            () => Array.from(document.querySelectorAll('input,button'))
              .filter(n => /submit|complete/i.test((n.value||n.textContent||'').trim()))
              .map(n => ({
                tag:n.tagName.toLowerCase(),
                id:n.id||'',
                name:n.name||'',
                value:n.value||'',
                class:n.className||'',
                outer:(n.outerHTML||'').slice(0,260)
              }))
        """)
        p = os.path.join(ARTIFACT_DIR, f"submit_candidates_{int(time.Date.now())}.json")
    except Exception:
        # fallback path if Date.now exploded
        p = os.path.join(ARTIFACT_DIR, f"submit_candidates_{int(time.time())}.json")
        out = page.evaluate("""
            () => Array.from(document.querySelectorAll('input,button'))
              .filter(n => /submit|complete/i.test((n.value||n.textContent||'').trim()))
              .map(n => ({
                tag:n.tagName.toLowerCase(),
                id:n.id||'',
                name:n.name||'',
                value:n.value||'',
                class:n.className||'',
                outer:(n.outerHTML||'').slice(0,260)
              }))
        """)
    with open(p, "w", encoding="utf-8") as f:
        import json; f.write(json.dumps(out, indent=2))
    log(f"[dump] wrote submit candidates: {p}")
    
# -----------------------------
# MAIN FLOW
# -----------------------------

def run_flow(
    vs_user: str,
    vs_pass: str,
    scenario_csv: Optional[str],
    scenario_docx: Optional[str],
    minicad_url: Optional[str],
    minicad_user: Optional[str],
    minicad_pass: Optional[str],
    headed: bool = False,
    max_scenarios: int = 1,
    roster_cmd: Optional[str] = None,
    include_units: Optional[List[str]] = None,
):
    
    # Load scenarios first
    scenarios: List[Dict[str, str]] = []
    if scenario_csv and os.path.exists(scenario_csv):
        scenarios = load_scenarios_from_csv(scenario_csv)
    elif scenario_docx and os.path.exists(scenario_docx):
        scenarios = load_scenarios_from_docx(scenario_docx)
    else:
        raise FileNotFoundError("No valid scenario source provided. Supply --scenario-csv or --scenario-docx.")
    if max_scenarios and max_scenarios > 0:
        scenarios = scenarios[:max_scenarios]

    # Fetch roster (command or URL) and extract personnel
    roster_raw = fetch_minicad_roster(
        url=minicad_url,
        user=minicad_user,
        password=minicad_pass,
        roster_cmd=roster_cmd,
        include_units=include_units,
    )
    raw_people = extract_personnel_with_units(roster_raw)     # pulls {'name','unit'} from your roster JSON (unit + staff[].name)
    people     = filter_personnel_by_units(raw_people, include_units)  # include_units comes from --units
    personnel_names = [p["name"] for p in people]
    log(f"Will attempt to add {len(personnel_names)} personnel for units={include_units}")


    # Playwright lifecycle (bulletproof)
    browser = None
    context = None
    page = None
    tracing_started = False
    results: List[Dict[str, Any]] = []

    pw_ctx = sync_playwright().start()
    try:
        browser = pw_ctx.chromium.launch(headless=not headed)
        context = browser.new_context()
        with swallow("tracing start"):
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
            tracing_started = True

        page = context.new_page()

       # ---- LOGIN ----
        log("Starting Vector login flow...")
        do_login(page, vs_user, vs_pass)
        wait_network_quiet(page, 10000)
        log("Login flow complete.") 

        # ---- OPEN TRAINING FORM ----
        goto_dashboard_open_form(page)          # single entry point
        wait_for_training_form_ready(page)
        log("Training form is ready.")      

        # Process scenarios
        for idx, scenario in enumerate(scenarios, 0):
            log(f"=== Scenario {idx}/{len(scenarios)} ===")
            fill_training_form(page, scenario)
            log(f"[debug] Before fill_core_fields, current URL: {page.url}")
            with swallow("debug screenshot before core fields"):
                dbg_path = os.path.join(ARTIFACT_DIR, f"before_core_fields_{int(time.time())}.png")
                page.screenshot(path=dbg_path, full_page=True)
                log(f"[debug] Screenshot before core fields: {_abs(dbg_path)}")
            fill_core_fields(page, scenario)
            log("Core fields filled, attempting submit...")
            click_save_and_add_users(page)
            choose_users_and_continue(page, personnel_names)
            with swallow("screenshot after choose users"):   # even if 0, it will press Continue
                shot_path = os.path.join(ARTIFACT_DIR, f"choose_users_{int(time.time())}_{idx}.png")
                page.screenshot(path=shot_path, full_page=True)
                log(f"Saved screenshot: {shot_path}")   
            conf = submit_training(page)
            with swallow("screenshot after submit"):
                shot_path = os.path.join(ARTIFACT_DIR, f"submit_{int(time.time())}_{idx}.png")
                page.screenshot(path=shot_path, full_page=True)
                log(f"Saved screenshot: {shot_path}")
            results.append({
                "scenario": scenario,
                "confirmation": conf,
                "url": page.url,
                "timestamp": int(time.time()),
            })
           
        # Persist results once
        out_path = os.path.join(ARTIFACT_DIR, f"run_results_{int(time.time())}.json")
        write_json(out_path, results)
        log(f"Wrote run results to {out_path}")

    finally:
        with swallow("tracing stop"):
            if context and tracing_started:
                trace_path = os.path.join(ARTIFACT_DIR, f"trace_{int(time.time())}.zip")
                context.tracing.stop(path=trace_path)
                log(f"Saved Playwright trace: {trace_path}")
        with swallow("context close"):
            if context:
                context.close()
        with swallow("browser close"):
            if browser:
                browser.close()
        with swallow("pw stop"):
            try:
                pw_ctx.stop()
            except Exception:
                pass

def parse_args():
    import argparse
    ap = argparse.ArgumentParser(description="Vector Solutions training-bot")
    ap.add_argument("--vs-user", required=False, default=os.getenv("VS_USER"))
    ap.add_argument("--vs-pass", required=False, default=os.getenv("VS_PASS"))
    ap.add_argument("--scenario-csv", required=False, default="/home/training-bot/projects/vector-solutions/trainings.csv")
    ap.add_argument("--scenario-docx", required=False)
    ap.add_argument("--minicad-url", required=False)
    ap.add_argument("--minicad-user", required=False)
    ap.add_argument("--minicad-pass", required=False)
    ap.add_argument("--headed", action="store_true", help="Run with a visible browser")
    ap.add_argument("--max-scenarios", type=int, default=1, help="Limit how many scenarios to submit this run")
    ap.add_argument("--roster-cmd", help="Command to run that outputs full roster JSON to stdout or a file path", required=False)
    ap.add_argument("--units", help="Comma-separated list of unit IDs to include (e.g., R1,E22,L26)", required=False)
    ap.add_argument("--artifact-dir", help="Where to write dumps, traces, screenshots", required=False)
    return ap.parse_args()

if __name__ == "__main__":
    load_dotenv()
    args = parse_args()

    if args.artifact_dir:
        os.makedirs(args.artifact_dir, exist_ok=True)
        ARTIFACT_DIR = args.artifact_dir  # override
        log(f"[config] ARTIFACT_DIR = {os.path.abspath(ARTIFACT_DIR)}")

    try:
        units_list = None
        if args.units:
            units_list = [u.strip() for u in args.units.split(",") if u.strip()]
        run_flow(
            vs_user=args.vs_user,
            vs_pass=args.vs_pass,
            scenario_csv=args.scenario_csv,
            scenario_docx=args.scenario_docx,
            minicad_url=args.minicad_url,
            minicad_user=args.minicad_user,
            minicad_pass=args.minicad_pass,
            headed=args.headed,
            max_scenarios=args.max_scenarios,
            roster_cmd=args.roster_cmd,
            include_units=units_list,
        )    
        
    except Exception as e:
        log(f"[fatal] {e}")
        sys.exit(1)