#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vector Solutions Training Bot — cleaned
- Safe form navigation with dashboard fallback
- Single, robust Add Users flow (auto-detects iframe or in-page)
- DEFAULT_UNITS honored for GUI and non-GUI flows
- No duplicate function names
"""

import os, sys, csv, json, time, argparse, hashlib, subprocess, requests, re
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Optional, List, Dict
from playwright.sync_api import sync_playwright
from playwright.sync_api import Page, TimeoutError as PWTimeout

# ========= Config / Paths =========
PROJECT_ROOT = Path("/home/training-bot/projects/vector-solutions")
DEFAULT_ROSTER = PROJECT_ROOT / "02_Scripts" / "Rosters" / "roster.json"
DEFAULT_CONFIG = PROJECT_ROOT / "05_Dev_Env" / "Dependencies" / "config.yaml"
GUI_PATH      = PROJECT_ROOT / "02_Scripts" / "Automation" / "assignments_gui.py"
SELECTED_PATH = PROJECT_ROOT / "02_Scripts" / "Runtime" / "selected_personnel.json"
CSV_DEFAULT   = "/home/training-bot/projects/vector-solutions/03_Training-Documentation/trainings.csv"
STATE_FILE    = "/home/training-bot/projects/vector-solutions/04_Outputs/Logs/.training_state.json"
AUTH_STATE    = "/home/training-bot/projects/vector-solutions/04_Outputs/Logs/.auth_state.json"
ARTIFACT_DIR  = "/home/training-bot/projects/vector-solutions/04_Outputs/Logs/"
SUBMIT_LOG    = "/home/training-bot/projects/vector-solutions/04_Outputs/Reports/submissions.jsonl"

# Optional: set via env
LOGIN_URL    = os.environ.get("VS_LOGIN_URL", "").strip()
FORM_URL     = os.environ.get("VS_FORM_URL", "").strip()
MINICAD_URL  = os.environ.get("MINICAD_URL", "").strip()
MINICAD_USER = os.environ.get("MINICAD_USER", "").strip()
MINICAD_PASS = os.environ.get("MINICAD_PASS", "").strip()
DEFAULT_UNITS_ENV = os.environ.get("DEFAULT_UNITS", "").strip()

# ========= Selectors / Constants =========
DEFAULT_TIMEOUT = 60000
RANK_WORDS = r"(Acting|Interim|Chief|Battalion Chief|Division Chief|BC|Captain|Capt|Lieutenant|Lt|Engineer|Eng|Firefighter|FF|Paramedic|Medic|Officer)"
SUFFIXES   = r"(Jr|Sr|II|III|IV|V)"

SEL_LOCATION    = 'input#nodeUserVal2'
SEL_DESCRIPTION = 'textarea#nodeUserVal4'
SEL_DURATION    = 'input#nodeUserVal5'
SEL_DATE        = 'input#nodeUserVal6'
SEL_TIME        = 'select[name="nodeUserVal6_endTime"]'
SEL_INSTRUCTOR  = 'input#nodeUserVal7'
SUBMIT_BUTTON_TEXT = "Submit"
SUCCESS_CUES       = ["success", "submitted", "saved", "completed"]

REQUIRED_HEADERS = [
    "Location", "Checkbox Label", "Description", "Duration", "Date", "Time", "Instructor"
]

# ========= Data Model =========
@dataclass
class TrainingRow:
    location: str
    checkbox_label: str
    description: str
    duration: str
    date_str: str
    time_str: str
    instructor: str
    raw: Dict[str, str]
    _hash: Optional[str] = None

# ========= State / CSV =========
def _hash_row(raw: Dict[str, str]) -> str:
    import hashlib
    m = hashlib.sha256()
    m.update(("\x1f".join(raw.get(h, "") for h in REQUIRED_HEADERS)).encode())
    return m.hexdigest()

def load_state():
    try:
        with open(STATE_FILE) as f: return json.load(f)
    except FileNotFoundError: return {"used": []}

def save_state(st):
    Path(STATE_FILE).parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w") as f: json.dump(st, f, indent=2)

def read_csv(csv_path: str) -> List[TrainingRow]:
    if not os.path.exists(csv_path):
        sys.exit(f"CSV not found: {csv_path}")
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        headers = [h.strip() for h in (reader.fieldnames or [])]
        missing = [h for h in REQUIRED_HEADERS if h not in headers]
        if missing:
            sys.exit(f"CSV missing columns: {missing}\nGot: {headers}")
        rows = []
        for r in reader:
            rows.append(TrainingRow(
                location=r["Location"].strip(),
                checkbox_label=r["Checkbox Label"].strip(),
                description=r["Description"].strip(),
                duration=r["Duration"].strip(),
                date_str=r["Date"].strip(),
                time_str=r["Time"].strip(),
                instructor=r["Instructor"].strip(),
                raw={k: (v or "").strip() for k, v in r.items()}
            ))
        if not rows: sys.exit("CSV parsed but contains zero data rows.")
        return rows

def pick_row(rows, mode="next") -> TrainingRow:
    st = load_state()
    used = set(st.get("used", []))
    candidates = [(r, _hash_row(r.raw)) for r in rows if _hash_row(r.raw) not in used]
    if not candidates:
        st["used"] = []; save_state(st)
        candidates = [(r, _hash_row(r.raw)) for r in rows]
    if mode == "random":
        import random; r, h = random.choice(candidates)
    elif mode == "bydate":
        idx = int(date.today().strftime("%j")) % len(candidates); r, h = candidates[idx]
    else:
        r, h = candidates[0]
    r._hash = h
    return r

# ========= Config / Roster =========
def _read_yaml(path: Path) -> dict:
    try:
        import yaml
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return yaml.safe_load(f) or {}
    except Exception:
        pass
    return {}

def load_config() -> dict:
    cfg = _read_yaml(DEFAULT_CONFIG)
    if DEFAULT_UNITS_ENV:
        cfg["default_units"] = [u.strip() for u in DEFAULT_UNITS_ENV.split(",") if u.strip()]
    cfg.setdefault("default_units", ["R1","R26"])
    cfg.setdefault("roster_path", str(DEFAULT_ROSTER))
    return cfg

def load_roster(roster_path: str) -> Dict:
    p = Path(roster_path)
    if not p.exists(): return {"date": "", "units": {}}
    with open(p, "r", encoding="utf-8") as f: return json.load(f)

def list_units(roster: Dict) -> List[str]:
    unit_map = roster.get("by_unit", roster.get("units", {}))
    if isinstance(unit_map, dict): return sorted(unit_map.keys())
    if isinstance(roster.get("units"), list):
        names = [str(u.get("unit") or u.get("UnitName") or "").strip() for u in roster["units"]]
        return sorted({n for n in names if n})
    return []

def gather_participants_from_units(roster: Dict, units: List[str]) -> List[str]:
    unit_map = roster.get("by_unit", roster.get("units", {}))
    out = []
    for u in units:
        out.extend(unit_map.get(u, []))
    return sorted(set(out))

def choose_units_interactively(roster: Dict, defaults: List[str]) -> List[str]:
    units = list_units(roster)
    if not units:
        print("[roster] No units found."); return []
    if defaults:
        # auto-pick defaults if they exist
        picks = [u for u in defaults if u in units]
        if picks:
            print(f"[roster] Using default units: {', '.join(picks)}")
            return sorted(set(picks))
    # fallback: prompt
    for i, u in enumerate(units):
        cnt = len(roster.get("by_unit", roster.get("units", {})).get(u, [])) if isinstance(roster.get("units"), dict) else 0
        print(f"[{i}] {u} ({cnt})")
    raw = input("\nSelect units by index or name (comma-separated): ").strip()
    if not raw: return []
    sel = []
    for x in [p.strip() for p in raw.split(",") if p.strip()]:
        if x.isdigit(): 
            i = int(x); 
            if 0 <= i < len(units): sel.append(units[i])
        elif x in units: sel.append(x)
    return sorted(set(sel))

# ========= Name helpers =========
def _strip_ranks_suffixes(n: str) -> str:
    n = re.sub(rf"^\s*{RANK_WORDS}\.?\s+", "", n, flags=re.I)
    n = re.sub(r"\s{2,}", " ", n).strip(" ,")
    n = re.sub(rf"\b{SUFFIXES}\b\.?", "", n, flags=re.I).strip(" ,")
    return re.sub(r"\s{2,}", " ", n).strip(" ,")

def normalize_to_last_first(name: str) -> str:
    n = _strip_ranks_suffixes(str(name or ""))
    if not n: return n
    if "," in n:
        last, firstrest = [p.strip() for p in n.split(",", 1)]
        first = (firstrest.split() or [""])[0]
        return f"{last}, {first}".strip(", ")
    parts = n.split()
    if len(parts) >= 2:
        first, last = parts[0], parts[-1]
        return f"{last}, {first}"
    return n

def build_match_patterns_last_first(last_first: str):
    if not last_first or "," not in last_first:
        return [re.compile(re.escape(last_first), re.I)]
    last, first = [p.strip() for p in last_first.split(",", 1)]
    tight = re.compile(rf"\b{re.escape(last)}\s*,\s*{re.escape(first)}\b", re.I)
    loose = re.compile(rf"\b{re.escape(last)}\s*,\s*{re.escape(first)}(?:\s+\w\.?)?\b", re.I)
    return [tight, loose]

# ========= Browser helpers =========
def capture_login_interactively(page: Page):
    print("[login] Complete login in the browser, then press ENTER.")
    input()

def ensure_dirs():
    Path(ARTIFACT_DIR).mkdir(parents=True, exist_ok=True)
    Path(Path(AUTH_STATE).parent).mkdir(parents=True, exist_ok=True)

def navigate_to(page: Page, url: str, label: str):
    if not url:
        raise RuntimeError(f"[nav] Missing URL for {label}")
    page.goto(url, wait_until="domcontentloaded", timeout=60000)

def ensure_cookies_saved(context): 
    try:
        os.makedirs(os.path.dirname(AUTH_STATE), exist_ok=True)
        context.storage_state(path=AUTH_STATE)
    except Exception:
        pass

def detect_form_ready(page: Page, timeout=30000):
    page.locator(SEL_LOCATION).wait_for(state="visible", timeout=timeout)
    print("[form] Detected form ready.", flush=True)    

def snapshot(page: Page, tag: str) -> str:
    fname = f"{int(time.time())}_{tag}.png"
    path = os.path.join(ARTIFACT_DIR, fname)
    try: page.screenshot(path=path, full_page=True)
    except Exception: pass
    return path

# ========= Live roster (MINICAD) =========
def fetch_minicad_live(url: str, user: str, pwd: str, timeout: int = 10) -> dict[str, list[str]]:
    if not url:  raise RuntimeError("MINICAD URL missing.")
    if not user or not pwd: raise RuntimeError("MINICAD username/password missing.")
    def _norm(n: str) -> str:
        n = (n or "").strip()
        if not n: return ""
        if "," in n:
            last, rest = [p.strip() for p in n.split(",", 1)]
            first = (rest.split() or [""])[0]
            return f"{last}, {first}".strip(", ")
        parts = n.split()
        if len(parts) >= 2:
            first, last = parts[0], parts[-1]
            return f"{last}, {first}"
        return n
    resp = requests.get(url, auth=(user, pwd), timeout=timeout)
    if resp.status_code == 401:
        raise RuntimeError("MINICAD auth failed (401).")
    resp.raise_for_status()
    try: data = resp.json()
    except Exception:
        data = json.loads(resp.text)
    # find list of unit dicts
    units_list = None
    if isinstance(data, list): units_list = data
    elif isinstance(data, dict):
        for k in ("Results","results","Data","data","Items","items","value","Value","rows","Rows","Units","units"):
            v = data.get(k)
            if isinstance(v, list) and v and isinstance(v[0], dict):
                units_list = v; break
        if units_list is None:
            for v in data.values():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    units_list = v; break
    if not isinstance(units_list, list):
        raise RuntimeError("Unexpected MINICAD format: cannot locate units list.")
    unit_map: dict[str, list[str]] = {}
    for obj in units_list:
        if not isinstance(obj, dict): continue
        unit = str(obj.get("UnitName") or obj.get("unit") or obj.get("Unit") or obj.get("Name") or "").strip()
        if not unit: continue
        people = obj.get("Staff") or obj.get("Personnel") or obj.get("staff") or obj.get("personnel") or []
        if isinstance(people, str): people = [p.strip() for p in people.replace(";", ",").split(",") if p.strip()]
        if not isinstance(people, list): continue
        norm = sorted({ _norm(p) for p in people if isinstance(p, str) and p.strip() })
        if norm: unit_map[unit] = norm
    if not unit_map:
        raise RuntimeError("Parsed zero staffed units from MINICAD.")
    print(f"[live] Parsed {len(unit_map)} units from MINICAD.")
    return unit_map

# ========= Participants acquisition =========
def run_assignments_gui_and_get_people(args) -> list[str]:
    if not GUI_PATH.exists():
        print(f"[gui] {GUI_PATH} not found; skipping GUI pre-step.")
        return []
    cmd = [sys.executable, str(GUI_PATH), "--output", str(SELECTED_PATH)]
    if getattr(args, "gui_roster", ""):       cmd += ["--roster", args.gui_roster]
    if getattr(args, "gui_roster_url", ""):   cmd += ["--roster-url", args.gui_roster_url]
    if getattr(args, "gui_auth_header", ""):  cmd += ["--auth-header", args.gui_auth_header]

    # Honor DEFAULT_UNITS if --gui-use is not provided
    gui_use = getattr(args, "gui_use", "") or DEFAULT_UNITS_ENV
    if gui_use: cmd += ["--use", gui_use]

    print(f"[gui] Launching: {' '.join(cmd)}")
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"[gui] assignments_gui exited non-zero: {e}; continuing without it.")
        return []
    try:
        with open(SELECTED_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        ppl = data.get("personnel") or data.get("participants") or []
        out = sorted({str(p).strip() for p in ppl if str(p).strip()})
        print(f"[gui] Loaded {len(out)} participant(s) from {SELECTED_PATH}")
        return out
    except Exception as e:
        print(f"[gui] Failed reading {SELECTED_PATH}: {e}")
        return []

def pre_run_collect_participants(args) -> List[str]:
    want_live = args.live or (args.minicad_user and args.minicad_pass and (args.minicad_url or MINICAD_URL))
    if want_live:
        try:
            murl = args.minicad_url or MINICAD_URL
            musr = args.minicad_user or MINICAD_USER
            mpwd = args.minicad_pass or MINICAD_PASS
            live_map = fetch_minicad_live(murl, musr, mpwd)
            units = [u.strip() for u in (args.units or DEFAULT_UNITS_ENV).split(",") if u.strip()] if (args.auto or args.units or DEFAULT_UNITS_ENV) else []
            if not units:
                # interactive as last resort
                units = sorted(live_map.keys())
            ppl = sorted({p for u in units for p in live_map.get(u, [])})
            print(f"[participants] From LIVE {units}: {len(ppl)} found")
            return ppl
        except Exception as e:
            print(f"[live][warn] {e}")
    if not args.allow_cached:
        print("[participants] Cached sources disabled (no --allow-cached). Returning empty list.")
        return []
    # GUI cache
    try:
        spath = SELECTED_PATH
        if spath.exists():
            with open(spath, "r", encoding="utf-8") as f: sel = json.load(f)
            ppl = sel.get("personnel", []) or sel.get("participants", [])
            if ppl:
                out = sorted(set(str(x).strip() for x in ppl if str(x).strip()))
                print(f"[participants] Using cached {spath} ({len(out)})")
                return out
    except Exception as e:
        print(f"[warn] Failed reading selected_personnel.json: {e}")
    # roster fallback
    cfg = load_config()
    roster_path = args.roster or cfg["roster_path"]
    roster = load_roster(roster_path)
    defaults = [u.strip() for u in (DEFAULT_UNITS_ENV or "").split(",") if u.strip()]
    units = [u.strip() for u in (args.units or "").split(",") if u.strip()]
    if not units:
        units = choose_units_interactively(roster, defaults if args.auto or defaults else [])
    ppl = gather_participants_from_units(roster, units)
    print(f"[participants] From roster {units}: {len(ppl)} found")
    return ppl

# ========= Navigation to form =========
def safe_go_to_form(page: Page, form_url: Optional[str]):
    if form_url:
        navigate_to(page, form_url, "form_url")
        return
    # Fall back to dashboard tile click
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    sel = "#single-spa-application\\:\\@target-solutions\\/home > section > section > vwc-tiling-grid > vwc-tiling-grid-tile.bulletin-board.is-schedule > vwc-card > div.bulletinBoard-content > div > p a"
    try:
        page.locator(sel).first.click(timeout=10000)
    except Exception:
        page.get_by_role("link", name=lambda n: "training" in n.lower()).first.click(timeout=15000)
    page.wait_for_load_state("networkidle", timeout=30000)
    assert ("training" in page.url.lower()) or ("record" in page.url.lower()), "Did not reach training form"

# ========= Add Users (supports iframe or in-page) =========
def go_to_participants(page: Page):
    for loc in [
        page.get_by_role("button", name=re.compile(r"^Save\s+and\s+Add\s+Users$", re.I)),
        page.get_by_role("button", name=re.compile(r"Add\s+Users", re.I)),
        page.locator("text=Save and Add Users"),
    ]:
        try:
            loc.scroll_into_view_if_needed(); loc.click(timeout=8000); break
        except Exception: pass
    page.wait_for_timeout(200)

def _detect_add_users_context(page: Page):
    # Return a tuple (context_page_or_frame, is_iframe)
    try:
        iframe = page.frame_locator("iframe[name^='TB_iframeContent'], iframe[src*='addUsers'], iframe#TB_iframeContent").first
        page.locator("iframe[name^='TB_iframeContent'], iframe[src*='addUsers'], iframe#TB_iframeContent").first.wait_for(state="visible", timeout=8000)
        return iframe, True
    except Exception:
        return page, False

def _wait_candidates(ctx, timeout=8000):
    try:
        lb = ctx.get_by_role("listbox", name=re.compile("(available|all|users)", re.I))
        lb.wait_for(state="visible", timeout=timeout)
        return lb, lambda pat: lb.get_by_role("option", name=pat)
    except Exception:
        pass
    # generic panel
    box = ctx.locator("body") if hasattr(ctx, "locator") else ctx
    return box, lambda pat: ctx.locator("text=" + pat.pattern).first

def _click_add_button(ctx):
    for b in [
        ctx.get_by_role("button", name=re.compile(r"^\s*Add\s*$", re.I)),
        ctx.get_by_role("button", name=re.compile(r"^\s*Add\s+Selected\s*$", re.I)),
        ctx.locator("button:has-text('Add')"),
        ctx.locator("button[title*='Add']"),
    ]:
        try:
            b.scroll_into_view_if_needed(); b.click(timeout=3000); return True
        except Exception: pass
    return False

def _verify_moved(ctx, last_first: str):
    try: right = ctx.get_by_role("listbox", name=re.compile("(selected|chosen)", re.I))
    except Exception: right = ctx.locator("div[aria-label*='Selected'], div.selected-grid, div.list-right")
    try: right.wait_for(state="visible", timeout=4000)
    except Exception: return False
    last, first = [p.strip() for p in last_first.split(",", 1)] if "," in last_first else (last_first, "")
    pat = re.compile(rf"\b{re.escape(last)}\s*,\s*{re.escape(first)}(?:\s+\w\.?)?\b", re.I) if first else re.compile(re.escape(last_first), re.I)
    try:
        node = right.get_by_text(pat).first if hasattr(right, "get_by_text") else right.locator("text=" + pat.pattern).first
        return node.count() > 0 if hasattr(node, "count") else node.is_visible()
    except Exception:
        return False

def _search_and_add(ctx, fullname: str, timeout=DEFAULT_TIMEOUT) -> bool:
    target = normalize_to_last_first(fullname)
    pats = build_match_patterns_last_first(target)
    # search box if present
    search = None
    try:
        s = ctx.get_by_placeholder(re.compile("Search", re.I))
        if s.is_visible(): search = s
    except Exception: pass
    if search:
        try:
            last = target.split(",", 1)[0]
            search.fill(""); search.fill(last)
            try: search.press("Enter")
            except Exception: pass
        except Exception: pass
    container, inside = _wait_candidates(ctx, timeout=8000)
    for pat in pats:
        try:
            item = inside(pat)
            try:
                if hasattr(item, "count") and item.count() == 0: continue
            except Exception:
                pass
            item.scroll_into_view_if_needed()
            try:
                item.dblclick(timeout=2500)
            except Exception:
                item.click(timeout=2500); _click_add_button(ctx)
            if _verify_moved(ctx, target): return True
            _click_add_button(ctx)
            if _verify_moved(ctx, target): return True
        except Exception:
            continue
    return False

def pick_participants(page: Page, participants: List[str]):
    ctx, is_iframe = _detect_add_users_context(page)
    added, missing = [], []
    for name in participants:
        ok = _search_and_add(ctx, name)
        (added if ok else missing).append(name)
        if not ok: print(f"[warn] Could not add: {name}")
    # continue button
    try:
        cont = ctx.get_by_role("button", name=re.compile("continue", re.I)).first
    except Exception:
        cont = ctx.locator("input[type='button'][value*='Continue']").first
    # observe iframe detach if needed
    if is_iframe:
        outer = page.locator("iframe[name^='TB_iframeContent'], iframe[src*='addUsers'], iframe#TB_iframeContent").first
        with page.expect_event("framedetached"):
            cont.click()
        outer.wait_for(state="detached", timeout=DEFAULT_TIMEOUT)
    else:
        cont.click()
    page.wait_for_load_state("domcontentloaded")
    print(f"[participants] added: {len(added)} | missing: {len(missing)}")
    if missing: print("[participants][missing]", ", ".join(missing))

# ========= Form fill / submit =========
def check_box(page, text):
    try:
        page.get_by_label(text, exact=False).check(); print(f"[check] {text}"); return
    except Exception: pass
    for sel in [
        f'xpath=//*[contains(normalize-space(.), "{text}")]/following::input[@type="checkbox"][1]',
        f'label:has-text("{text}") >> input[type="checkbox"]'
    ]:
        node = page.locator(sel).first
        if node.count():
            node.scroll_into_view_if_needed(); page.wait_for_timeout(100)
            node.check(); print(f"[check] {text} ({sel})"); return
    raise RuntimeError(f"Checkbox not found: {text}")

def submit_training(page):
    try:
        page.get_by_role("button", name=re.compile("^Submit$", re.I)).click(timeout=12000)
    except Exception:
        page.get_by_role("button", name=re.compile("^Save$", re.I)).click(timeout=12000)
        page.get_by_role("button", name=re.compile("^Submit$", re.I)).click(timeout=12000)

def normalize_duration(val: str) -> str:
    s = val.lower().strip()
    if not s: raise ValueError("Duration is empty")
    if s.isdigit(): hours = int(s)/60.0
    else:
        for token in ["hours","hour","hrs","hr","h"]: s = s.replace(token,"")
        hours = float(s.strip())
    hours = max(0.25, min(2.0, hours))
    return f"{hours:.2f}".rstrip("0").rstrip(".")

def normalize_date(s: str) -> str:
    s = s.strip()
    if not s: return datetime.today().strftime("%m/%d/%Y")
    fmts = ["%m/%d/%Y","%m/%d/%y","%Y-%m-%d","%m-%d-%Y","%m-%d-%y","%Y/%m/%d"]
    for fmt in fmts:
        try: return datetime.strptime(s, fmt).strftime("%m/%d/%Y")
        except: pass
    month, day, year = s.replace("-", "/").split("/")
    return datetime(int(year), int(month), int(day)).strftime("%m/%d/%Y")

def parse_time(s: str) -> datetime:
    s = s.strip().lower()
    if not s: return datetime(2000,1,1,12,0)
    s = s.replace("am"," am").replace("pm"," pm")
    for fmt in ["%I:%M %p","%I %p","%H:%M","%H"]:
        try: return datetime.strptime(s.upper(), fmt)
        except: pass
    parts = s.replace(".",":").split()
    hh, mm = (parts[0].split(":") + ["0"])[:2]
    hh = int(hh); mm = int(mm)
    ampm = parts[1].upper() if len(parts)>1 else ""
    if ampm == "PM" and hh < 12: hh += 12
    if ampm == "AM" and hh == 12: hh = 0
    return datetime(2000,1,1,hh%24,mm%60)

def quarter(dt: datetime) -> datetime:
    q = ((dt.minute + 7)//15)*15
    return dt.replace(minute=q%60, hour=(dt.hour + q//60)%24)

def normalize_time_option(s: str) -> str:
    dt = quarter(parse_time(s))
    return dt.strftime("%I:%M %p").lstrip("0")

def fill_form(page: Page, r: TrainingRow, participants: List[str]):
    hours_norm = normalize_duration(r.duration)
    date_norm  = normalize_date(r.date_str)
    time_opt   = normalize_time_option(r.time_str)

    pre_path = snapshot(page, "pre_submit")

    page.locator(SEL_LOCATION).scroll_into_view_if_needed(); page.wait_for_timeout(50)
    page.locator(SEL_LOCATION).fill(r.location)

    check_box(page, r.checkbox_label)

    page.locator(SEL_DESCRIPTION).scroll_into_view_if_needed(); page.wait_for_timeout(50)
    page.locator(SEL_DESCRIPTION).fill(r.description)

    page.locator(SEL_DURATION).scroll_into_view_if_needed(); page.wait_for_timeout(50)
    page.locator(SEL_DURATION).fill(hours_norm)

    page.locator(SEL_DATE).scroll_into_view_if_needed(); page.wait_for_timeout(50)
    page.locator(SEL_DATE).fill(date_norm)

    t = page.locator(SEL_TIME).first
    if t.count() == 0: raise RuntimeError(f"Time selector not found: {SEL_TIME}")
    try: t.select_option(value=time_opt)
    except Exception:
        try: t.select_option(label=time_opt)
        except Exception:
            opts = [o.strip() for o in t.locator("option").all_text_contents() if o.strip()]
            def mins(s): dt = parse_time(s); return dt.hour*60 + dt.minute
            target = mins(time_opt); best = min(opts, key=lambda o: abs(mins(o)-target))
            t.select_option(label=best); time_opt = best

    page.locator(SEL_INSTRUCTOR).scroll_into_view_if_needed(); page.wait_for_timeout(50)
    page.locator(SEL_INSTRUCTOR).fill(r.instructor)

    if participants:
        go_to_participants(page)
        pick_participants(page, participants)
    else:
        print("[participants] None provided; skipping Add Users flow.")

    submit_training(page)

    for _ in range(30):
        if any(x in page.content().lower() for x in SUCCESS_CUES): break
        time.sleep(0.25)

    post_path = snapshot(page, "post_submit")
    payload = {
        "ts": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "selected_hash": r._hash,
        "submitted": {
            "Location": r.location,
            "Checkbox Label": r.checkbox_label,
            "Description": r.description,
            "Duration_hours": hours_norm,
            "Date": date_norm,
            "Time_option": time_opt,
            "Instructor": r.instructor,
        },
        "artifacts": {"pre_screenshot": pre_path, "post_screenshot": post_path},
    }
    Path(Path(SUBMIT_LOG).parent).mkdir(parents=True, exist_ok=True)
    with open(SUBMIT_LOG, "a", encoding="utf-8") as f: f.write(json.dumps(payload) + "\n")
    print(f"[log] Submission written to {SUBMIT_LOG}")
    print(f"[snap] {pre_path}"); print(f"[snap] {post_path}")

def browser_ctx(pw, headed: bool):
    print("[browser] Launching bundled Chromium")
    browser = pw.chromium.launch(headless=not headed)
    context = browser.new_context()
    return browser, context

def finish(context, browser):
    try: context.close()
    except Exception: pass
    try: browser.close()
    except Exception: pass

# ========= Main flow =========
def run_browser_flow(args, r, login_url, form_url):
    with sync_playwright() as pw:
        print("[debug] entering playwright", flush=True)
        b, c = browser_ctx(pw, args.headed)
        p = c.new_page()
        try:
            if login_url:
                navigate_to(p, login_url, "login_url")
                if args.capture_login: capture_login_interactively(p)
                try: p.wait_for_load_state("networkidle", timeout=15000)
                except Exception: pass

            # Go to form (URL or dashboard fallback)
            safe_go_to_form(p, form_url)
            detect_form_ready(p)
            ensure_cookies_saved(c)

            # Collect participants
            participants = []
            if args.gui: participants = run_assignments_gui_and_get_people(args)
            if not participants: participants = pre_run_collect_participants(args)

            # Fill & submit
            fill_form(p, r, participants)

        except Exception as e:
            snap = os.path.join(ARTIFACT_DIR, f"fail_{int(time.time())}.png")
            try: p.screenshot(path=snap, full_page=True)
            except Exception: pass
            print(f"[ERROR] {e}\nScreenshot: {snap}", flush=True)
            raise
        finally:
            finish(c, b)

def main():
    ap = argparse.ArgumentParser(description="Vector Solutions Training Bot")
    ap.add_argument("--csv", default=CSV_DEFAULT)
    ap.add_argument("--mode", choices=["next","random","bydate"], default="next")
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--auto", action="store_true", help="Use DEFAULT_UNITS/env defaults without prompting")
    ap.add_argument("--units", type=str, default="", help="Explicit unit list, comma-separated (overrides defaults)")
    ap.add_argument("--roster", type=str, default="", help=f"Path to roster.json (default: {DEFAULT_ROSTER})")
    ap.add_argument("--gui", action="store_true", help="Run assignments_gui.py first")
    ap.add_argument("--gui-roster", type=str, default="")
    ap.add_argument("--gui-roster-url", type=str, default="")
    ap.add_argument("--gui-auth-header", type=str, default="")
    ap.add_argument("--gui-use", type=str, default="")
    ap.add_argument("--login-url", type=str, default="")
    ap.add_argument("--form-url",  type=str, default="")
    ap.add_argument("--capture-login", action="store_true")
    ap.add_argument("--user", type=str, default=os.environ.get("VS_USER",""))
    ap.add_argument("--password", type=str, default=os.environ.get("VS_PASS",""))
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--minicad-url", type=str, default=os.environ.get("MINICAD_URL",""))
    ap.add_argument("--minicad-user", type=str, default=os.environ.get("MINICAD_USER",""))
    ap.add_argument("--minicad-pass", type=str, default=os.environ.get("MINICAD_PASS",""))
    ap.add_argument("--allow-cached", action="store_true")

    args = ap.parse_args()

    ensure_dirs()
    rows = read_csv(args.csv)
    r = pick_row(rows, args.mode)
    print(f"[select] {r.location} | {r.checkbox_label} | {r.date_str} {r.time_str}")

    if args.dry_run:
        print("[dry-run] Not touching the site.")
        return

    login_url = args.login_url or LOGIN_URL
    form_url  = args.form_url  or FORM_URL

    print(f"[debug] will run browser flow headed={args.headed} live={args.live}", flush=True)
    run_browser_flow(args, r, login_url, form_url)
    print("[debug] browser flow returned", flush=True)

if __name__ == "__main__":
    main()
