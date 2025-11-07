from __future__ import annotations
import re, sys, unicodedata
from pathlib import Path
from typing import Dict, List
from docx import Document  # pip install python-docx

# Exact labels from the Vector form (checkbox group)
SITE_TRAINING_TYPES = {
    "Building Construction",
    "Fire Behavior",
    "Fire Detection, Alarm Systems, Suppression Systems",
    "Fire Extinguishers",
    "Fire Hose Evolutions",
    "Fire Streams and Nozzles",
    "Forcible Entry",
    "Ground Ladders",
    "Salvage and Overhaul",
    "Search and Rescue",
    "Ventilation",
    "VES - Ventilation, Enter, Search",
    "Water Supply (ex. hydrant operations, tender operations, dry hydrants)",
    "Firefighting Tactics and Strategies",
    "Extrication",
}

# Map common titles in your doc to the form labels
TOPIC_TO_LABELS: Dict[str, List[str]] = {
    "Search & Rescue": ["Search and Rescue"],
    "SCBA Maintenance & Skills": ["Firefighting Tactics and Strategies"],
    "Ladder Drills": ["Ground Ladders"],
    "Ventilation Tactics": ["Ventilation"],
    "Hose Handling & Stream Application": ["Fire Hose Evolutions", "Fire Streams and Nozzles"],
    "Fire Behavior": ["Fire Behavior"],
    "Rapid Intervention Team (RIT)": ["Firefighting Tactics and Strategies"],
    "Forcible Entry": ["Forcible Entry"],
    "Water Supply Operations": ["Water Supply (ex. hydrant operations, tender operations, dry hydrants)"],
    "Building Construction": ["Building Construction"],
    "Wildland Fire Ops": [],  # no exact checkbox; leave empty
    "Pump Operations": ["Water Supply (ex. hydrant operations, tender operations, dry hydrants)"],
    "EVOC Policy Review": [],  # leave empty
    "Incident Command System": ["Firefighting Tactics and Strategies"],  # ICS not listed; best-fit bucket
    "Fireground Communications": ["Firefighting Tactics and Strategies"],
    "Hydrant Operations": ["Water Supply (ex. hydrant operations, tender operations, dry hydrants)"],
    "Thermal Imaging Camera": ["Fire Behavior"],  # closest fit
    "Salvage & Overhaul": ["Salvage and Overhaul"],
    "First Responder Refresher": [],  # EMS; leave empty
    "Hazmat Awareness": [],  # not present; leave empty
    "Extrication": ["Extrication"],  # if it appears
}

PIPE_ROW = re.compile(r"^\s*([^|]+)\|\s*([0-9]+)\s*\|\s*(.+)$")

def sanitize_filename(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = re.sub(r"[^\w\s.-]", "", s)
    s = re.sub(r"\s+", "_", s.strip())
    return s.lower()

def minutes_to_hours(m: str) -> float:
    try:
        return round(int(m) / 60.0, 2)
    except Exception:
        return 1.0

def normalize_types(topic: str) -> List[str]:
    labels = TOPIC_TO_LABELS.get(topic.strip(), [])
    # keep only exact site labels
    return [l for l in labels if l in SITE_TRAINING_TYPES]

def parse_rows(docx_path: Path):
    doc = Document(str(docx_path))
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        m = PIPE_ROW.match(t)
        if m:
            title, minutes, desc = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
            yield title, minutes, desc

def write_module(out_dir: Path, title: str, hours: float, desc: str, labels: List[str]):
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = f"{sanitize_filename(title)}.md"
    path = out_dir / fname
    yml = []
    yml.append("---")
    yml.append(f'title: "{title}"')
    if labels:
        yml.append("training_types:")
        for l in labels:
            yml.append(f'  - {l}')
    yml.append(f"duration_hours: {hours}")
    yml.append("---\n")
    body = f"# Module: {title}\n\n{desc}\n"
    path.write_text("\n".join(yml) + body, encoding="utf-8")
    return path

def main():
    if len(sys.argv) < 3:
        print("Usage: convert_docx_to_modules.py <trainings.docx> <output_dir>", file=sys.stderr)
        return 2
    docx_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])

    seen = set()
    count_in, count_out = 0, 0
    for title, minutes, desc in parse_rows(docx_path):
        count_in += 1
        key = title.strip()
        if key in seen:
            continue  # dedupe repeated rows
        seen.add(key)
        hours = minutes_to_hours(minutes)
        labels = normalize_types(title)
        write_module(out_dir, title, hours, desc, labels)
        count_out += 1

    print(f"Parsed rows: {count_in} | Unique modules written: {count_out} | Output: {out_dir}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
