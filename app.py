import re
import io
import math
import json
from dataclasses import dataclass, asdict, field
from datetime import datetime
from difflib import get_close_matches
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="TEC-004 Running Hours Analyzer",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# STYLES
# =========================================================

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1500px;
}
.metric-card {
    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 16px 18px;
    box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
}
.metric-label {
    font-size: 12px;
    color: #64748b;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .04em;
}
.metric-value {
    font-size: 28px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1.1;
    margin-top: 4px;
}
.metric-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 6px;
}
.section-title {
    font-size: 24px;
    font-weight: 800;
    color: #0f172a;
    margin-top: 10px;
    margin-bottom: 8px;
}
.soft-note {
    color: #475569;
    font-size: 14px;
}
.badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
}
.badge-ok { background: #dcfce7; color: #166534; }
.badge-hp { background: #fef3c7; color: #92400e; }
.badge-od { background: #fee2e2; color: #991b1b; }
.badge-nd { background: #e2e8f0; color: #334155; }
.badge-val { background: #dbeafe; color: #1d4ed8; }
.badge-review { background: #fef3c7; color: #92400e; }
.small-muted {
    color: #64748b;
    font-size: 12px;
}
hr {
    margin-top: 1rem !important;
    margin-bottom: 1rem !important;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# CONFIG / CANONICAL MAPS
# =========================================================

SECTION_ALIASES = {
    "ME": {
        "MAIN ENGINE", "M/E", "ME", "MAIN ENGINE COMPONENTS"
    },
    "AUX": {
        "AUXILIARY ENGINES", "AUX ENGINES", "A/E", "AUX", "GENERATOR ENGINES"
    },
    "OE": {
        "OTHER EQUIPMENT", "TURBOCHARGER", "BOILER EQUIPMENT", "COMPRESSORS",
        "COOLING", "COOLERS / WATER SYSTEMS", "AIR CONDITIONING", "COOLING SYSTEMS"
    },
    "DG": {
        "D/G EQUIPMENT", "DG EQUIPMENT", "DIESEL GENERATOR", "DIESEL GENERATORS",
        "D/G", "DG"
    },
}

COMPONENT_ALIASES = {
    "CYLINDER COVER": {"CYLINDER COVER"},
    "PISTON ASSEMBLY": {"PISTON ASSEMBLY"},
    "STUFFING BOX": {"STUFFING BOX"},
    "PISTON CROWN": {"PISTON CROWN"},
    "CYLINDER LINER": {"CYLINDER LINER", "CYLINDER LINERS"},
    "EXHAUST VALVE": {"EXHAUST VALVE", "EXHAUST VALVES"},
    "STARTING VALVE": {"STARTING VALVE", "STARTING VALVES"},
    "SAFETY VALVE": {"SAFETY VALVE", "SAFETY VALVES"},
    "FUEL VALVES": {"FUEL VALVES", "FUEL VALVE", "FUEL VALVES (1)", "FUEL VALVES 1"},
    "FUEL PUMP": {"FUEL PUMP", "FUEL PUMPS"},
    "PLUNGER AND BARREL": {"PLUNGER AND BARREL", "PLUNGER & BARREL"},
    "FUEL PUMP SUCTION VALVE": {"FUEL PUMP SUCTION VALVE", "SUCTION VALVE"},
    "FUEL PUMP PUNCTURE VALVE": {"FUEL PUMP PUNCTURE VALVE", "PUNCTURE VALVE"},
    "CROSSHEAD BEARINGS": {"CROSSHEAD BEARINGS", "CROSSHEAD BEARING"},
    "BOTTOM END BEARINGS": {"BOTTOM END BEARINGS", "BOTTOM END BEARING"},
    "MAIN BEARING": {"MAIN BEARING", "MAIN BEARINGS"},
    "CYLINDER HEAD": {"CYLINDER HEAD", "CYLINDER HEADS"},
    "PISTON": {"PISTON", "PISTONS"},
    "CONNECTING ROD": {"CONNECTING ROD", "CONNECTING RODS"},
    "CRANK PIN BEARING": {"CRANK PIN BEARING", "CRANKPIN BEARING"},
    "ADJUST VALVE HEAD CLEARANCE": {"ADJUST VALVE HEAD CLEARANCE"},
    "GENERAL O/H": {"GENERAL O/H", "GENERAL OVERHAUL"},
    "BALANCING OF ROTOR SHAFT": {"BALANCING OF ROTOR SHAFT"},
    "AIR COOLER CLEANING": {"AIR COOLER CLEANING"},
    "AIR COND. COMPRESSOR NO.1": {"AIR COND. COMPRESSOR NO.1", "AIR COND COMPRESSOR NO.1", "AIR COND COMPRESSOR NO 1"},
    "AIR COND. COMPRESSOR NO.2": {"AIR COND. COMPRESSOR NO.2", "AIR COND COMPRESSOR NO.2", "AIR COND COMPRESSOR NO 2"},
    "REFRIGERATION COMPRESSOR NO.1": {"REFRIGERATION COMPRESSOR NO.1"},
    "REFRIGERATION COMPRESSOR NO.2": {"REFRIGERATION COMPRESSOR NO.2"},
    "STARTING MAIN AIR COMPRESSOR NO.1": {"STARTING MAIN AIR COMPRESSOR NO.1"},
    "STARTING MAIN AIR COMPRESSOR NO.2": {"STARTING MAIN AIR COMPRESSOR NO.2"},
    "SERVICE AIR COMPRESSOR": {"SERVICE AIR COMPRESSOR"},
    "EMERGENCY AIR COMPRESSOR NO.": {"EMERGENCY AIR COMPRESSOR NO.", "EMERGENCY AIR COMPRESSOR"},
    "FURNACE INSPECTION": {"FURNACE INSPECTION"},
    "BURNER ATOMIZER": {"BURNER ATOMIZER"},
    "FORCED DRAFT FAN": {"FORCED DRAFT FAN"},
    "FEED PUMPS NO.1": {"FEED PUMPS NO.1", "FEED PUMP NO.1"},
    "FEED PUMPS NO.2": {"FEED PUMPS NO.2", "FEED PUMP NO.2"},
    "WASHING THE TUBES": {"WASHING THE TUBES"},
    "O/H CIRC. PUMP NO.1": {"O/H CIRC. PUMP NO.1"},
    "O/H CIRC. PUMP NO.2": {"O/H CIRC. PUMP NO.2"},
    "M/E L.O.": {"M/E L.O."},
    "AIR. COND. COOLER CLEANING": {"AIR. COND. COOLER CLEANING"},
    "ATMOSPHERIC CONDENSER": {"ATMOSPHERIC CONDENSER"},
}

OE_SECTION_BY_COMPONENT = {
    "GENERAL O/H": "Turbocharger",
    "BALANCING OF ROTOR SHAFT": "Turbocharger",
    "AIR COOLER CLEANING": "Turbocharger",
    "M/E L.O.": "Coolers / Water Systems",
    "WASHING THE TUBES": "Coolers / Water Systems",
    "O/H CIRC. PUMP NO.1": "Coolers / Water Systems",
    "O/H CIRC. PUMP NO.2": "Coolers / Water Systems",
    "AIR COND. COMPRESSOR NO.1": "Compressors",
    "AIR COND. COMPRESSOR NO.2": "Compressors",
    "REFRIGERATION COMPRESSOR NO.1": "Compressors",
    "REFRIGERATION COMPRESSOR NO.2": "Compressors",
    "STARTING MAIN AIR COMPRESSOR NO.1": "Compressors",
    "STARTING MAIN AIR COMPRESSOR NO.2": "Compressors",
    "SERVICE AIR COMPRESSOR": "Compressors",
    "EMERGENCY AIR COMPRESSOR NO.": "Compressors",
    "FURNACE INSPECTION": "Boiler Equipment",
    "BURNER ATOMIZER": "Boiler Equipment",
    "FORCED DRAFT FAN": "Boiler Equipment",
    "FEED PUMPS NO.1": "Boiler Equipment",
    "FEED PUMPS NO.2": "Boiler Equipment",
    "AIR. COND. COOLER CLEANING": "Air Conditioning",
    "ATMOSPHERIC CONDENSER": "Cooling Systems",
}

EXPECTED_SECTION_ORDER = ["ME", "AUX", "OE", "DG"]


# =========================================================
# DATA MODEL
# =========================================================

@dataclass
class ParsedRow:
    source_text: str = ""
    raw_section: Optional[str] = None
    section: Optional[str] = None
    component_raw: Optional[str] = None
    component: Optional[str] = None
    engine_raw: Optional[str] = None
    engine: Optional[str] = None
    unit_raw: Optional[str] = None
    unit: Optional[str] = None
    periodicity_raw: Optional[str] = None
    periodicity: Optional[int] = None
    last_oh_raw: Optional[str] = None
    last_oh_iso: Optional[str] = None
    hrs_since_raw: Optional[str] = None
    hrs_since: Optional[int] = None
    pct_used: Optional[float] = None
    status: Optional[str] = None
    source_method: str = "text"
    confidence: float = 0.0
    validation_status: str = "review_needed"
    issues: List[str] = field(default_factory=list)
    row_key: Optional[str] = None
    display_order: int = 0


# =========================================================
# LOW LEVEL HELPERS
# =========================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ").strip()
    s = s.replace("DEC.22", "DEC 22").replace("OCT.", "OCT").replace("JAN.", "JAN")
    s = s.replace("NO 1", "NO.1").replace("NO 2", "NO.2")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def normalize_token(value: Any) -> str:
    s = clean_text(value).upper()
    s = s.replace("’", "'")
    s = re.sub(r"[(),]", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def parse_int(value: Any) -> Optional[int]:
    s = clean_text(value)
    if not s or s in {"—", "-", "N/A"}:
        return None
    s = s.replace(",", "")
    m = re.search(r"\d+", s)
    if not m:
        return None
    try:
        return int(m.group())
    except:
        return None

def parse_float_pct(value: Any) -> Optional[float]:
    s = clean_text(value)
    if not s or s in {"—", "-", "N/A"}:
        return None
    s = s.replace("%", "").strip()
    try:
        return float(s)
    except:
        return None

def normalize_date(value: Any) -> Optional[str]:
    s = clean_text(value).upper()
    if not s or s in {"—", "-", "N/A"}:
        return None

    replacements = {
        "SEPT": "SEP",
        "MARCH": "MAR",
        "JUNE": "JUN",
        "JULY": "JUL",
    }
    for old, new in replacements.items():
        s = re.sub(rf"\b{old}\b", new, s)

    s = s.replace(".", " ")
    s = re.sub(r"\s+", " ", s).strip()

    fmts = [
        "%d %b %y", "%d %b %Y",
        "%d/%m/%y", "%d/%m/%Y",
        "%d %m %y", "%d %m %Y",
    ]

    for fmt in fmts:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except:
            pass

    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)
    if m:
        d, mth, y = m.groups()
        d = int(d)
        mth = int(mth)
        y = int(y)
        if y < 100:
            y += 2000
        try:
            dt = datetime(y, mth, d)
            return dt.strftime("%Y-%m-%d")
        except:
            return None

    return None

def format_display_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return "—"
    try:
        return datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d %b %Y").upper()
    except:
        return iso_str

def canon_section(raw: str) -> Optional[str]:
    token = normalize_token(raw)
    alias_map = {}
    for canon, aliases in SECTION_ALIASES.items():
        for alias in aliases:
            alias_map[normalize_token(alias)] = canon
    if token in alias_map:
        return alias_map[token]
    return None

def canon_component(raw: str) -> Tuple[Optional[str], bool]:
    token = normalize_token(raw)
    alias_lookup = {}
    for canon, aliases in COMPONENT_ALIASES.items():
        for alias in aliases:
            alias_lookup[normalize_token(alias)] = canon

    if token in alias_lookup:
        return alias_lookup[token], True

    close = get_close_matches(token, list(alias_lookup.keys()), n=1, cutoff=0.92)
    if close:
        return alias_lookup[close[0]], False

    return None, False

def infer_oe_subsection(component: Optional[str]) -> Optional[str]:
    if not component:
        return None
    return OE_SECTION_BY_COMPONENT.get(component)

def normalize_engine(section: Optional[str], raw_engine: str, raw_unit: str) -> Tuple[Optional[str], Optional[str]]:
    e = clean_text(raw_engine).upper()
    u = clean_text(raw_unit)

    if section == "ME":
        if not u and re.search(r"CYL\s*\d+", e):
            u = re.sub(r"\s+", " ", e.title()).replace("Cyl", "Cyl")
        return "ME", (u if u else None)

    if section == "AUX":
        if e:
            e = e.replace(" ", "-")
        return (e if e else None), (u if u else None)

    if section == "DG":
        if e:
            e = e.replace("/", "").replace(" ", "-")
        return (e if e else None), (u if u else None)

    return (e if e else None), (u if u else None)

def compute_pct_used(periodicity: Optional[int], hrs_since: Optional[int]) -> Optional[float]:
    if not periodicity or periodicity <= 0 or hrs_since is None:
        return None
    return round((hrs_since / periodicity) * 100.0, 1)

def compute_status(periodicity: Optional[int], hrs_since: Optional[int]) -> str:
    if not periodicity or hrs_since is None:
        return "NO DATA"
    pct = (hrs_since / periodicity) * 100.0
    if pct >= 100:
        return "OVERDUE"
    if pct >= 80:
        return "HIGH PRIORITY"
    return "OK"

def safe_str(value: Any) -> str:
    return "" if value is None else str(value)

def row_key_from_parts(section, engine, unit, component, periodicity, last_oh_iso):
    return "|".join([
        safe_str(section),
        safe_str(engine),
        safe_str(unit),
        safe_str(component),
        safe_str(periodicity),
        safe_str(last_oh_iso),
    ])


# =========================================================
# RAW TEXT PARSER
# =========================================================

def extract_document_header(text: str) -> Dict[str, Any]:
    vessel = None
    report_date = None
    me_total_hrs = None
    me_this_month = None

    vessel_match = re.search(r"\b(?:Vessel)\s*\n?([A-Z][A-Z0-9 .\-]+)", text, re.IGNORECASE)
    if vessel_match:
        vessel = clean_text(vessel_match.group(1))

    date_match = re.search(r"\b(\d{1,2}\s+[A-Z]+\s+\d{4})\s*\n?\s*Report Date", text, re.IGNORECASE)
    if date_match:
        report_date = clean_text(date_match.group(1))

    total_match = re.search(r"\b([\d,]+)\s*\n?\s*M/E Total Hrs", text, re.IGNORECASE)
    if total_match:
        me_total_hrs = parse_int(total_match.group(1))

    month_match = re.search(r"\b([\d,]+)\s*\n?\s*M/E This Month", text, re.IGNORECASE)
    if month_match:
        me_this_month = parse_int(month_match.group(1))

    if not vessel:
        m = re.search(r"\bALEXIS\b", text, re.IGNORECASE)
        if m:
            vessel = "ALEXIS"

    return {
        "vessel": vessel,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
    }

def split_sections(lines: List[str]) -> Dict[str, List[str]]:
    sections = {"ME": [], "AUX": [], "OE": [], "DG": []}
    current = None

    for line in lines:
        n = normalize_token(line)

        if n == "MAIN ENGINE":
            current = "ME"
            continue
        elif n == "AUXILIARY ENGINES":
            current = "AUX"
            continue
        elif n == "OTHER EQUIPMENT":
            current = "OE"
            continue
        elif n == "D/G EQUIPMENT":
            current = "DG"
            continue

        if current:
            sections[current].append(line)

    return sections

def group_me_or_aux_rows(lines: List[str], section: str) -> List[Dict[str, Any]]:
    rows = []
    i = 0

    statuses = {"OK", "OVERDUE", "HIGH PRIORITY", "NO DATA"}

    while i < len(lines):
        line = clean_text(lines[i])
        if line in statuses:
            status_line = line
            if i + 7 < len(lines):
                component = clean_text(lines[i + 1])
                engine = clean_text(lines[i + 2])
                unit = clean_text(lines[i + 3])
                periodicity = clean_text(lines[i + 4])
                last_oh = clean_text(lines[i + 5])
                hrs_since = clean_text(lines[i + 6])
                pct = clean_text(lines[i + 7])

                looks_valid = (
                    component and
                    ((section == "ME" and engine == "ME") or (section == "AUX" and engine.startswith("AUX")))
                )

                if looks_valid:
                    rows.append({
                        "section": "MAIN ENGINE" if section == "ME" else "AUXILIARY ENGINES",
                        "component": component,
                        "engine": engine,
                        "unit": unit,
                        "periodicity": periodicity,
                        "last_oh": last_oh,
                        "hrs_since": hrs_since,
                        "pct_used_raw": pct,
                        "status_raw": status_line,
                        "source_method": "text_block",
                        "source_text": " | ".join(lines[i:i+8]),
                    })
                    i += 8
                    continue
        i += 1

    return rows

def group_oe_rows(lines: List[str]) -> List[Dict[str, Any]]:
    rows = []
    i = 0

    while i < len(lines):
        component = clean_text(lines[i])

        if i + 3 < len(lines):
            maybe_section = clean_text(lines[i + 1])
            periodicity = clean_text(lines[i + 2])
            last_date = clean_text(lines[i + 3])
            run_hrs = clean_text(lines[i + 4]) if i + 4 < len(lines) else ""

            recognized_section = maybe_section in {
                "Turbocharger", "Boiler Equipment", "Compressors",
                "Cooling", "Coolers / Water Systems", "Air Conditioning",
                "Cooling Systems"
            }

            if recognized_section:
                rows.append({
                    "section": "OTHER EQUIPMENT",
                    "component": component,
                    "engine": maybe_section,
                    "unit": "",
                    "periodicity": periodicity,
                    "last_oh": last_date,
                    "hrs_since": run_hrs,
                    "pct_used_raw": "",
                    "status_raw": "",
                    "source_method": "text_block",
                    "source_text": " | ".join(lines[i:i+5]),
                })
                i += 5
                continue

        i += 1

    return rows

def parse_text_to_raw_rows(text: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    lines = [clean_text(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    header = extract_document_header(text)
    sections = split_sections(lines)

    raw_rows = []
    raw_rows.extend(group_me_or_aux_rows(sections["ME"], "ME"))
    raw_rows.extend(group_me_or_aux_rows(sections["AUX"], "AUX"))
    raw_rows.extend(group_oe_rows(sections["OE"]))

    return raw_rows, header


# =========================================================
# NORMALIZATION / VALIDATION
# =========================================================

def normalize_row(raw: Dict[str, Any], display_order: int) -> ParsedRow:
    r = ParsedRow(
        source_text=clean_text(raw.get("source_text")),
        raw_section=clean_text(raw.get("section")),
        component_raw=clean_text(raw.get("component")),
        engine_raw=clean_text(raw.get("engine")),
        unit_raw=clean_text(raw.get("unit")),
        periodicity_raw=clean_text(raw.get("periodicity")),
        last_oh_raw=clean_text(raw.get("last_oh")),
        hrs_since_raw=clean_text(raw.get("hrs_since")),
        source_method=clean_text(raw.get("source_method")) or "unknown",
        display_order=display_order,
    )

    r.section = canon_section(r.raw_section) if r.raw_section else None
    if not r.section:
        r.issues.append("unknown_section")

    r.component, exact = canon_component(r.component_raw)
    if not r.component:
        r.issues.append("unmapped_component")
    elif not exact:
        r.issues.append("fuzzy_component_match")

    r.engine, r.unit = normalize_engine(r.section, r.engine_raw, r.unit_raw)
    r.periodicity = parse_int(r.periodicity_raw)
    r.last_oh_iso = normalize_date(r.last_oh_raw)
    r.hrs_since = parse_int(r.hrs_since_raw)

    if r.section == "OE" and not r.engine:
        inferred = infer_oe_subsection(r.component)
        if inferred:
            r.engine = inferred

    if r.last_oh_raw and not r.last_oh_iso and r.last_oh_raw not in {"—", "-"}:
        r.issues.append("invalid_date")

    r.pct_used = compute_pct_used(r.periodicity, r.hrs_since)
    r.status = compute_status(r.periodicity, r.hrs_since)

    status_raw = clean_text(raw.get("status_raw")).upper()
    if status_raw and status_raw != r.status:
        r.issues.append(f"status_mismatch:{status_raw}->{r.status}")

    if r.section in {"ME", "AUX", "DG"} and not r.engine:
        r.issues.append("missing_engine")
    if r.section in {"ME", "AUX", "DG"} and not r.unit:
        r.issues.append("missing_unit")

    if r.periodicity is None and r.hrs_since is not None:
        r.issues.append("hrs_since_without_periodicity")

    confidence = 1.0
    confidence -= 0.25 if "unknown_section" in r.issues else 0
    confidence -= 0.20 if "unmapped_component" in r.issues else 0
    confidence -= 0.10 if "fuzzy_component_match" in r.issues else 0
    confidence -= 0.15 if "invalid_date" in r.issues else 0
    confidence -= 0.10 if "missing_engine" in r.issues else 0
    confidence -= 0.10 if "missing_unit" in r.issues else 0
    confidence -= 0.08 if any(x.startswith("status_mismatch:") for x in r.issues) else 0
    confidence -= 0.08 if "hrs_since_without_periodicity" in r.issues else 0
    confidence -= 0.05 if r.source_method == "text_block" else 0

    r.confidence = max(0.0, round(confidence, 2))
    critical = {"unknown_section", "unmapped_component"} & set(r.issues)
    r.validation_status = "validated" if r.confidence >= 0.85 and not critical else "review_needed"
    r.row_key = row_key_from_parts(r.section, r.engine, r.unit, r.component, r.periodicity, r.last_oh_iso)

    return r

def dedupe_rows(rows: List[ParsedRow]) -> Tuple[List[ParsedRow], List[Dict[str, Any]]]:
    best_by_key: Dict[str, ParsedRow] = {}
    conflicts: List[Dict[str, Any]] = []

    for row in rows:
        key = row.row_key or f"__row__{row.display_order}"
        if key not in best_by_key:
            best_by_key[key] = row
            continue

        existing = best_by_key[key]
        same_payload = (
            existing.periodicity == row.periodicity and
            existing.last_oh_iso == row.last_oh_iso and
            existing.hrs_since == row.hrs_since and
            existing.status == row.status
        )

        if same_payload:
            if row.confidence > existing.confidence:
                best_by_key[key] = row
        else:
            conflicts.append({
                "row_key": key,
                "reason": "duplicate_key_conflict",
                "existing_component": existing.component,
                "incoming_component": row.component,
                "existing_engine": existing.engine,
                "incoming_engine": row.engine,
                "existing_last_oh": existing.last_oh_iso,
                "incoming_last_oh": row.last_oh_iso,
                "existing_hrs_since": existing.hrs_since,
                "incoming_hrs_since": row.hrs_since,
            })
            if row.confidence > existing.confidence:
                best_by_key[key] = row

    out = sorted(best_by_key.values(), key=lambda x: x.display_order)
    return out, conflicts

def audit_expected_sections(raw_text: str, rows: List[ParsedRow]) -> List[str]:
    warnings = []
    text = normalize_token(raw_text)

    found_in_source = set()
    for canon, aliases in SECTION_ALIASES.items():
        for a in aliases:
            if normalize_token(a) in text:
                found_in_source.add(canon)
                break

    found_in_output = {r.section for r in rows if r.section}

    for sec in found_in_source:
        if sec not in found_in_output:
            warnings.append(f"section_present_but_not_extracted:{sec}")

    if "DG" in found_in_source and not any(r.section == "DG" for r in rows):
        warnings.append("dg_section_missing")

    return warnings

def harden_parser_output(raw_rows: List[Dict[str, Any]], raw_document_text: str) -> Dict[str, Any]:
    normalized = [normalize_row(r, i+1) for i, r in enumerate(raw_rows)]
    deduped, conflicts = dedupe_rows(normalized)
    section_warnings = audit_expected_sections(raw_document_text, deduped)

    validated_rows = [r for r in deduped if r.validation_status == "validated"]
    review_rows = [r for r in deduped if r.validation_status != "validated"]

    summary = {
        "total_raw_rows": len(raw_rows),
        "total_normalized_rows": len(normalized),
        "total_deduped_rows": len(deduped),
        "validated_rows": len(validated_rows),
        "review_needed_rows": len(review_rows),
        "conflict_count": len(conflicts),
        "section_warnings": section_warnings,
        "counts_by_section": {sec: sum(1 for r in deduped if r.section == sec) for sec in EXPECTED_SECTION_ORDER},
        "overdue_count": sum(1 for r in deduped if r.status == "OVERDUE"),
        "high_priority_count": sum(1 for r in deduped if r.status == "HIGH PRIORITY"),
    }

    return {
        "summary": summary,
        "validated_rows": [asdict(r) for r in validated_rows],
        "review_rows": [asdict(r) for r in review_rows],
        "conflicts": conflicts,
        "all_rows": [asdict(r) for r in deduped],
    }


# =========================================================
# DATAFRAME / RENDER HELPERS
# =========================================================

def build_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=[
            "section", "component", "engine", "unit", "periodicity", "last_oh_iso",
            "hrs_since", "pct_used", "status", "confidence", "issues"
        ])
    df = pd.DataFrame(rows)
    for col in ["periodicity", "hrs_since", "pct_used", "confidence"]:
        if col in df.columns:
            df[col] = df[col]
    return df

def status_badge(status: str) -> str:
    status = clean_text(status).upper()
    if status == "OK":
        return '<span class="badge badge-ok">OK</span>'
    if status == "HIGH PRIORITY":
        return '<span class="badge badge-hp">HIGH PRIORITY</span>'
    if status == "OVERDUE":
        return '<span class="badge badge-od">OVERDUE</span>'
    return '<span class="badge badge-nd">NO DATA</span>'

def val_badge(status: str) -> str:
    if status == "validated":
        return '<span class="badge badge-val">VALIDATED</span>'
    return '<span class="badge badge-review">REVIEW NEEDED</span>'

def render_metrics(header: Dict[str, Any], summary: Dict[str, Any]):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Vessel</div>
            <div class="metric-value">{header.get("vessel") or "—"}</div>
            <div class="metric-sub">Parsed document header</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Report Date</div>
            <div class="metric-value" style="font-size:22px;">{header.get("report_date") or "—"}</div>
            <div class="metric-sub">Source date found in document</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        me_total = f'{header.get("me_total_hrs"):,}' if header.get("me_total_hrs") else "—"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">M/E Total Hrs</div>
            <div class="metric-value">{me_total}</div>
            <div class="metric-sub">Main engine total hours</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        me_month = f'{header.get("me_this_month"):,}' if header.get("me_this_month") else "—"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">M/E This Month</div>
            <div class="metric-value">{me_month}</div>
            <div class="metric-sub">Current month running hours</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    c5, c6, c7, c8 = st.columns(4)

    with c5:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Validated Rows</div>
            <div class="metric-value">{summary["validated_rows"]}</div>
            <div class="metric-sub">Rows safe for dashboard display</div>
        </div>
        """, unsafe_allow_html=True)

    with c6:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Review Needed</div>
            <div class="metric-value">{summary["review_needed_rows"]}</div>
            <div class="metric-sub">Rows needing audit attention</div>
        </div>
        """, unsafe_allow_html=True)

    with c7:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Overdue</div>
            <div class="metric-value">{summary["overdue_count"]}</div>
            <div class="metric-sub">Validated + deduped rows only</div>
        </div>
        """, unsafe_allow_html=True)

    with c8:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">High Priority</div>
            <div class="metric-value">{summary["high_priority_count"]}</div>
            <div class="metric-sub">Validated + deduped rows only</div>
        </div>
        """, unsafe_allow_html=True)

def apply_filters(df: pd.DataFrame, section_name: str) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    st.markdown(f'<div class="section-title">{section_name}</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns([1.2, 1.2, 1.2, 1.2])

    with col1:
        component_values = ["All"] + sorted([x for x in df["component"].dropna().unique().tolist() if x])
        component_filter = st.selectbox(f"{section_name} Component", component_values, key=f"{section_name}_component")

    with col2:
        engine_values = ["All"] + sorted([x for x in df["engine"].dropna().unique().tolist() if x])
        engine_filter = st.selectbox(f"{section_name} Engine", engine_values, key=f"{section_name}_engine")

    with col3:
        status_values = ["All", "OK", "HIGH PRIORITY", "OVERDUE", "NO DATA"]
        status_filter = st.selectbox(f"{section_name} Status", status_values, key=f"{section_name}_status")

    with col4:
        sort_mode = st.selectbox(
            f"{section_name} Sort",
            ["Document Order", "Priority → % Used", "Component → Unit"],
            key=f"{section_name}_sort"
        )

    out = df.copy()

    if component_filter != "All":
        out = out[out["component"] == component_filter]

    if engine_filter != "All":
        out = out[out["engine"] == engine_filter]

    if status_filter != "All":
        out = out[out["status"] == status_filter]

    if sort_mode == "Priority → % Used":
        status_rank = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}
        out["__status_rank"] = out["status"].map(status_rank).fillna(99)
        out["__pct_sort"] = out["pct_used"].fillna(-1)
        out = out.sort_values(["__status_rank", "__pct_sort"], ascending=[True, False]).drop(columns=["__status_rank", "__pct_sort"])
    elif sort_mode == "Component → Unit":
        out = out.sort_values(["component", "unit"], ascending=[True, True])
    else:
        if "display_order" in out.columns:
            out = out.sort_values("display_order", ascending=True)

    return out

def show_table(df: pd.DataFrame, columns: List[str]):
    if df.empty:
        st.info("No rows to display.")
        return

    show_df = df.copy()
    if "last_oh_iso" in show_df.columns:
        show_df["last_oh"] = show_df["last_oh_iso"].apply(format_display_date)
    if "pct_used" in show_df.columns:
        show_df["pct_used"] = show_df["pct_used"].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else "—")
    if "periodicity" in show_df.columns:
        show_df["periodicity"] = show_df["periodicity"].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else "—")
    if "hrs_since" in show_df.columns:
        show_df["hrs_since"] = show_df["hrs_since"].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else "—")
    if "confidence" in show_df.columns:
        show_df["confidence"] = show_df["confidence"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "—")
    if "issues" in show_df.columns:
        show_df["issues"] = show_df["issues"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)

    st.dataframe(
        show_df[columns],
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.title("TEC-004 Analyzer")
st.sidebar.markdown("Paste the extracted report text below or upload a `.txt` file.")
example_mode = st.sidebar.checkbox("Load demo sample", value=False)

uploaded_file = st.sidebar.file_uploader("Upload text file", type=["txt"])
manual_text = st.sidebar.text_area("Paste report text", height=300)

run_btn = st.sidebar.button("Analyze Report", type="primary", use_container_width=True)


# =========================================================
# DEMO DATA
# =========================================================

DEMO_TEXT = """
ALEXIS
Vessel
30 APRIL 2026
Report Date
71,984
M/E Total Hrs
516
M/E This Month

Main Engine
OK
CYLINDER COVER
ME
Cyl 1
16,000
12 JAN 26
1,308
8.2%
OVERDUE
CYLINDER COVER
ME
Cyl 2
16,000
16 DEC 22
17,560
109.7%
OK
PISTON ASSEMBLY
ME
Cyl 1
16,000
12 JAN 26
1,308
8.2%
OVERDUE
FUEL PUMP
ME
Cyl 1
16,000
22 DEC 22
17,560
109.7%
HIGH PRIORITY
FUEL PUMP SUCTION VALVE
ME
Cyl 6
8,000
02 DEC 24
7,858
98.2%

Auxiliary Engines
OK
Cylinder Head
AUX-1
Cyl 1
12,000
27/09/25
4,271
35.6%
HIGH PRIORITY
Piston
AUX-2
Cyl 3
10,000
14/11/24
9,707
97.1%
OVERDUE
Fuel Pumps
AUX-1
Cyl 2
5,000
11/1/25
7,250
145.0%
OK
Adjust Valve Head Clearance
AUX-3
Cyl 2
1,200
03/01/26
569
47.4%

Other Equipment
GENERAL O/H
Turbocharger
16,000
26 NOV 21
23,610
AIR COND. COMPRESSOR NO.1
Compressors
—
28 MAR 26
5,340
M/E L.O.
Coolers / Water Systems
—
30 DEC 23
—

D/G Equipment
0 records
No D/G equipment data found.
"""


# =========================================================
# MAIN EXECUTION
# =========================================================

st.title("⚙️ TEC-004 Running Hours Analyzer")
st.markdown('<div class="soft-note">Hardened parser shell with validation, dedupe, confidence scoring, and audit-first review.</div>', unsafe_allow_html=True)
st.write("")

input_text = ""
if example_mode:
    input_text = DEMO_TEXT
elif uploaded_file is not None:
    input_text = uploaded_file.read().decode("utf-8", errors="ignore")
elif manual_text.strip():
    input_text = manual_text

if run_btn:
    if not input_text.strip():
        st.error("Please upload a text file or paste report text before analyzing.")
        st.stop()

    raw_rows, header = parse_text_to_raw_rows(input_text)
    hardened = harden_parser_output(raw_rows, input_text)

    summary = hardened["summary"]
    validated_df = build_df(hardened["validated_rows"])
    review_df = build_df(hardened["review_rows"])
    all_df = build_df(hardened["all_rows"])
    conflicts_df = pd.DataFrame(hardened["conflicts"]) if hardened["conflicts"] else pd.DataFrame()

    render_metrics(header, summary)

    if summary["section_warnings"]:
        st.warning("Section warnings detected:")
        for warning in summary["section_warnings"]:
            st.write(f"- {warning}")

    st.write("")
    overview_tab, me_tab, aux_tab, oe_tab, dg_tab, audit_tab = st.tabs(
        ["Overview", "Main Engine", "Auxiliary Engines", "Other Equipment", "D/G", "Audit"]
    )

    with overview_tab:
        st.subheader("Validated Output")
        st.caption("Only validated rows are treated as trusted dashboard data.")
        ov = validated_df.copy()
        if not ov.empty:
            show_table(
                ov.sort_values("display_order"),
                ["section", "component", "engine", "unit", "periodicity", "last_oh_iso", "hrs_since", "pct_used", "status", "confidence"]
            )
        else:
            st.info("No validated rows found.")

    with me_tab:
        me_df = validated_df[validated_df["section"] == "ME"].copy() if not validated_df.empty else pd.DataFrame()
        me_df = apply_filters(me_df, "Main Engine")
        show_table(
            me_df,
            ["component", "engine", "unit", "periodicity", "last_oh_iso", "hrs_since", "pct_used", "status", "confidence"]
        )

    with aux_tab:
        aux_df = validated_df[validated_df["section"] == "AUX"].copy() if not validated_df.empty else pd.DataFrame()
        aux_df = apply_filters(aux_df, "Auxiliary Engines")
        show_table(
            aux_df,
            ["component", "engine", "unit", "periodicity", "last_oh_iso", "hrs_since", "pct_used", "status", "confidence"]
        )

    with oe_tab:
        oe_df = validated_df[validated_df["section"] == "OE"].copy() if not validated_df.empty else pd.DataFrame()
        oe_df = apply_filters(oe_df, "Other Equipment")
        show_table(
            oe_df,
            ["component", "engine", "periodicity", "last_oh_iso", "hrs_since", "status", "confidence"]
        )

    with dg_tab:
        dg_df = validated_df[validated_df["section"] == "DG"].copy() if not validated_df.empty else pd.DataFrame()
        if dg_df.empty:
            st.info("No validated D/G rows.")
        else:
            dg_df = apply_filters(dg_df, "D/G Equipment")
            show_table(
                dg_df,
                ["component", "engine", "unit", "periodicity", "last_oh_iso", "hrs_since", "pct_used", "status", "confidence"]
            )

    with audit_tab:
        st.subheader("Audit / Review Needed")
        st.caption("Anything uncertain is surfaced here instead of silently contaminating the trusted dashboard.")

        a1, a2, a3 = st.columns(3)
        a1.metric("Review Needed", summary["review_needed_rows"])
        a2.metric("Conflicts", summary["conflict_count"])
        a3.metric("Raw Rows", summary["total_raw_rows"])

        st.markdown("### Review Needed Rows")
        if review_df.empty:
            st.success("No review-needed rows.")
        else:
            show_table(
                review_df.sort_values("display_order"),
                [
                    "raw_section", "section", "component_raw", "component", "engine_raw", "engine", "unit_raw", "unit",
                    "periodicity_raw", "periodicity", "last_oh_raw", "last_oh_iso", "hrs_since_raw", "hrs_since",
                    "status", "confidence", "issues"
                ]
            )

        st.markdown("### Conflicts")
        if conflicts_df.empty:
            st.success("No dedupe conflicts detected.")
        else:
            st.dataframe(conflicts_df, use_container_width=True, hide_index=True)

        st.markdown("### All Rows Export Preview")
        if all_df.empty:
            st.info("No rows parsed.")
        else:
            export_df = all_df.copy()
            if "issues" in export_df.columns:
                export_df["issues"] = export_df["issues"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
            st.dataframe(export_df, use_container_width=True, hide_index=True)

            csv_bytes = export_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download CSV",
                data=csv_bytes,
                file_name="tec004_hardened_output.csv",
                mime="text/csv",
                use_container_width=False,
            )

else:
    st.info("Upload a `.txt` extraction or paste the report text, then click **Analyze Report**.")
    st.markdown("""
### What this version fixes
- Strict section normalization for **ME**, **AUX**, **OE**, and **DG**
- Canonical component matching to reduce label drift
- Date normalization to ISO internally
- Dedupe with conflict detection
- Confidence scoring and `review_needed` quarantine
- Audit-first UI so bad rows never silently become trusted rows
""")
