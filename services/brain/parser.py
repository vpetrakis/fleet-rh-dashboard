"""
Fleet Command Nexus — parser.py  v5.0
======================================
Key improvements over previous versions
────────────────────────────────────────
• Limits are read DYNAMICALLY from the document (cell +1 after each
  component name), so every vessel automatically gets its own limits
  without any hardcoding.  The database stores fallback limits only.

• Cell-offset extraction is EXACT (verified against a real fleet report):
      ME  section → "2" row at comp+14, values at comp+15…comp+14+me_cyls
      DG  section → "2" row at comp+18, values at comp+19…comp+18+dg_count
      AUX section → "2" row at comp+24, values at comp+25…comp+24+(dg×cyls)

• Section boundaries are auto-detected from sentinel cells so ME / DG / AUX
  searches never overlap and grab from the wrong section.

• Multi-value cells ("16801 95") take the first plausible running-hours
  number — fixing the "merged cell" artefact.

• European thousands-separator normalisation: "16.000" → 16000.

• Bracket-stripped values: "[2141]" → 2141.

• N/A / blank cells always return "-" (never crash the frontend).
"""

from __future__ import annotations

import logging
import re
import subprocess
import zipfile
import io
from dataclasses import dataclass, field
from pathlib import Path

log = logging.getLogger("fleet_nexus.parser")


# ══════════════════════════════════════════════════════════════════════════════
# ❶  FLEET REGISTRY
#
#    One entry per vessel.  Limits stored here are FALLBACKS only — the
#    parser will always try to read the real limit from the document first.
#
#    Adding a new vessel:
#      1. Copy any existing block.
#      2. Set me_cyls / dg_count / aux_cyls to match the vessel's engine.
#      3. Set aliases to every name variant the file/header might contain.
#      4. The component name lists rarely need changing — all 15 vessels use
#         the same document template; only limits and cylinder counts differ.
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class VesselConfig:
    me_cyls:    int
    dg_count:   int
    aux_cyls:   int
    # Keys = component names to search for (prefix-matched against doc cells).
    # Values = fallback limits used when the document says "Based on Observation" etc.
    me_limits:  dict[str, int]
    dg_limits:  dict[str, int]
    aux_limits: dict[str, int]
    aliases: list[str] = field(default_factory=list)


# All vessels use this template — component names are identical across the
# fleet.  Only limits, cylinder counts, and vessel name change per vessel.
_ME_COMPONENTS = {
    "CYLINDER COVER":           16_000,
    "PISTON ASSEMBLY":          16_000,
    "STUFFING BOX":             16_000,
    "PISTON CROWN":             32_000,
    "CYLINDER LINER":           32_000,
    "EXHAUST VALVE":            16_000,
    "STARTING VALVE":           12_000,
    "SAFETY VALVE":             12_000,
    "FUEL VALVES":               8_000,
    "FUEL PUMP":                16_000,
    "PLUNGER AND BARREL":       32_000,
    "FUEL PUMP SUCTION VALVE":   8_000,
    "FUEL PUMP PUNCTURE VALVE":  8_000,
    "CROSSHEAD BEARINGS":       32_000,
    "BOTTOM END BEARINGS":      32_000,
    "MAIN BEARINGS":            32_000,
}

_DG_COMPONENTS = {
    "Turbocharger":             12_000,
    "Air Cooler":                5_000,
    "L.O. Cooler":               8_000,
    "Cooling Water Pump":        5_000,
    "F.W. Cooler":               4_000,
    "L.O. Renewal":              1_500,
    "Alternator":                5_000,
    "Thrust Bearing":           12_000,
}

_AUX_COMPONENTS = {
    "Cylinder Head":            12_000,
    "Piston":                   10_000,
    "Connecting Rod":           10_000,
    "Cylinder Liners":          10_000,
    "Fuel Valves":               2_000,
    "Fuel Pumps":                5_000,
    "Crank Pin Bearing":        12_000,
}


FLEET_DATABASE: dict[str, VesselConfig] = {

    "MV ALEXIS": VesselConfig(
        me_cyls=5, dg_count=3, aux_cyls=6,
        me_limits=_ME_COMPONENTS,
        dg_limits=_DG_COMPONENTS,
        aux_limits=_AUX_COMPONENTS,
        aliases=["ALEXIS"],
    ),

    "MV OLYMPIA": VesselConfig(
        me_cyls=6, dg_count=4, aux_cyls=5,
        me_limits=_ME_COMPONENTS,   # limits will be overridden by document values
        dg_limits=_DG_COMPONENTS,
        aux_limits=_AUX_COMPONENTS,
        aliases=["OLYMPIA"],
    ),

    # ── Vessels 3-15: copy a block above, change me_cyls/dg_count/aux_cyls
    # and aliases.  Component name dicts can usually be left as the shared
    # _ME_COMPONENTS / _DG_COMPONENTS / _AUX_COMPONENTS defaults above.
    # ──────────────────────────────────────────────────────────────────────
}


# ══════════════════════════════════════════════════════════════════════════════
# ❷  CELL UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _clean(raw: str) -> str:
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x80-\x9f]", " ", raw)
    return re.sub(r"\s+", " ", s).strip()


def _parse_cells(raw_bytes: bytes) -> list[str]:
    text = raw_bytes.decode("latin-1", errors="ignore")
    return [_clean(c) for c in text.split("\x07")]


def _normalize(raw: str) -> str:
    """
    Convert a raw cell value to a plain integer string, or '-' if absent.
    Handles: blank, N/A, [brackets], European 16.000, commas, multi-value cells.
    """
    v = raw.strip("[]() \t")
    if not v or v.upper() in ("N/A", "N.A.", "NA", "-", "BASED ON OBSERVATION",
                               "OBSERVATION", "N.A", ""):
        return "-"

    def _try(tok: str) -> str | None:
        t = tok.strip("[](), \t")
        if not t:
            return None
        m = re.match(r"^(\d{1,3})\.(\d{3})$", t)      # European: 16.000
        if m:
            return m.group(1) + m.group(2)
        t2 = t.replace(",", "")
        try:
            n = int(t2)
            return str(n) if 0 <= n <= 999_999 else None
        except ValueError:
            return None

    # Fast path: single token
    r = _try(v)
    if r is not None:
        return r

    # Multi-value cell (e.g. "16801 95"): first 2+-digit token wins
    for tok in re.split(r"\s+", v):
        if len(tok.strip("[](),.")) < 2:
            continue
        r = _try(tok)
        if r is not None:
            return r
    return "-"


def _read_doc_limit(cells: list[str], comp_idx: int, fallback: int) -> int:
    """Read the overhaul limit from cell comp_idx+1 (always the limit cell)."""
    if comp_idx < 0 or comp_idx + 1 >= len(cells):
        return fallback
    v = _normalize(cells[comp_idx + 1])
    try:
        n = int(v)
        return n if n > 0 else fallback
    except (ValueError, TypeError):
        return fallback


# ══════════════════════════════════════════════════════════════════════════════
# ❸  SECTION BOUNDARY DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def _section_boundaries(cells: list[str]) -> tuple[int, int]:
    """
    Returns (aux_start, dg_start).
    • aux_start: index of the "AUX. ENGINE" header row
    • dg_start:  index just before the "D/G No1" header row
    """
    aux_start = len(cells)
    dg_start  = len(cells)
    for i, c in enumerate(cells):
        cu = c.upper()
        if aux_start == len(cells) and "AUX. ENGINE" in cu and len(c) < 60:
            aux_start = i
        if dg_start == len(cells) and cu in ("D/G NO1", "D/G NO2", "D/G NO3"):
            dg_start = max(0, i - 4)
            break
    return aux_start, dg_start


# ══════════════════════════════════════════════════════════════════════════════
# ❹  COMPONENT FINDER  (with limit-based disambiguation)
# ══════════════════════════════════════════════════════════════════════════════

def _find_component(
    cells: list[str],
    comp_name: str,
    db_limit: int,
    search_start: int,
    search_end: int,
    offset_row2: int,
) -> int:
    """
    Return the cell index of the best match for comp_name in the given range.
    When multiple cells start with the same prefix (e.g. "Turbocharger (2)" vs
    "Turbocharger (3)"), pick the one whose limit cell matches db_limit,
    or failing that, the one whose row-2 cell is literally "2".
    """
    upper = comp_name.upper()
    candidates: list[int] = []
    end = min(search_end, len(cells))

    for i in range(search_start, end):
        if cells[i].upper().startswith(upper):
            candidates.append(i)

    if not candidates:
        log.debug("  NOT FOUND: '%s'  (searched %d-%d)", comp_name, search_start, search_end)
        return -1
    if len(candidates) == 1:
        return candidates[0]

    # Disambiguation: prefer limit match
    limit_str = str(db_limit)
    lm = [idx for idx in candidates
          if idx + 1 < len(cells) and _normalize(cells[idx + 1]) == limit_str]
    if lm:
        return lm[0]

    # Fallback: prefer verified "2" at expected offset
    r2 = [idx for idx in candidates
          if idx + offset_row2 < len(cells) and cells[idx + offset_row2] == "2"]
    return r2[0] if r2 else candidates[0]


# ══════════════════════════════════════════════════════════════════════════════
# ❺  VALUE EXTRACTOR  (exact cell offsets, ±3 tolerance)
# ══════════════════════════════════════════════════════════════════════════════

def _extract(cells: list[str], comp_idx: int, offset_row2: int, n: int) -> list[str]:
    """
    Jump to cells[comp_idx + offset_row2] (the "2" marker), then read n values.
    Tolerates ±3 cell drift caused by unusual template layouts.
    """
    if comp_idx < 0:
        return ["-"] * n

    r2 = comp_idx + offset_row2
    if r2 < len(cells) and cells[r2] != "2":
        for d in range(-3, 4):
            probe = r2 + d
            if 0 <= probe < len(cells) and cells[probe] == "2":
                r2 = probe
                break
        else:
            log.warning("  '2' not found near cell %d for comp at %d — padding", r2, comp_idx)
            return ["-"] * n

    return [_normalize(cells[r2 + k]) if r2 + k < len(cells) else "-"
            for k in range(1, n + 1)]


# ══════════════════════════════════════════════════════════════════════════════
# ❻  VESSEL IDENTIFICATION
# ══════════════════════════════════════════════════════════════════════════════

def _best_text(raw_bytes: bytes, file_path: Path) -> str:
    """Return the cleanest possible plain text for vessel name searching."""

    # 1. antiword (installed on most Linux systems, gives perfect output)
    try:
        r = subprocess.run(
            ["antiword", str(file_path)],
            capture_output=True, text=True, timeout=15,
        )
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    # 2. .docx XML path
    try:
        parts: list[str] = []
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith(".xml"):
                    xml = zf.read(name).decode("utf-8", errors="ignore")
                    parts.append(re.sub(r"<[^>]+>", " ", xml))
        if parts:
            return re.sub(r"\s+", " ", " ".join(parts))
    except Exception:
        pass

    # 3. Binary latin-1 fallback
    body = raw_bytes[512:] if raw_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" else raw_bytes
    text = body.decode("latin-1", errors="ignore")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", " ", text)
    # Collapse character-spacing artefacts (S P A C E D → SPACED)
    text = re.sub(r"(?<=\w) (?=\w)", "", text)
    return re.sub(r"\s+", " ", text)


def identify_vessel(raw_bytes: bytes, file_path: Path) -> tuple[str, bool]:
    """4-tier identification: header name → body text → filename → fallback."""
    doc_text  = _best_text(raw_bytes, file_path)
    doc_upper = doc_text.upper()
    fname_upper = file_path.name.upper()

    def variants(vessel: str, cfg: VesselConfig) -> list[str]:
        base = re.sub(r"^(MV|M/V|MT|SS)\s*", "", vessel.upper()).strip()
        v = {vessel.upper(), base, base.replace(" ", "")}
        v.update(a.upper() for a in cfg.aliases)
        return sorted(v, key=len, reverse=True)

    # Tier 1: vessel name explicitly stated in document header
    m = re.search(
        r"VESSEL'?S?\s*NAME\s*[:\-]?\s*([A-Z0-9 ./\-]{2,40}?)(?:\t|DATE|$|\n|\r)",
        doc_upper,
    )
    if m:
        reported = m.group(1).strip()
        for vessel, cfg in FLEET_DATABASE.items():
            for v in variants(vessel, cfg):
                if v and (v in reported or reported in v):
                    log.info("Vessel (header) → %s", vessel)
                    return vessel, False

    # Tier 2 & 3: search document text then filename
    for source, haystack in [("text", doc_upper), ("filename", fname_upper)]:
        for vessel, cfg in FLEET_DATABASE.items():
            for v in variants(vessel, cfg):
                if v and v in haystack:
                    log.info("Vessel (%s) → %s", source, vessel)
                    return vessel, False

    fallback = next(iter(FLEET_DATABASE))
    log.warning("Vessel not identified — fallback '%s'", fallback)
    return fallback, True


# ══════════════════════════════════════════════════════════════════════════════
# ❼  INTEGRITY SCORING
# ══════════════════════════════════════════════════════════════════════════════

def _score(matrices: dict) -> int:
    total = missing = 0
    for row in matrices.get("main_engine", []):
        total  += len(row["hours"])
        missing += row["hours"].count("-")
    for row in matrices.get("dg_general", []):
        total  += len(row["hours"])
        missing += row["hours"].count("-")
    for row in matrices.get("aux_engine", []):
        for ch in row["dg_hours"]:
            total  += len(ch)
            missing += ch.count("-")
    return 0 if total == 0 else round(((total - missing) / total) * 100)


# ══════════════════════════════════════════════════════════════════════════════
# ❽  DIAGNOSTIC ENGINE  (local rule-based, zero external API cost)
# ══════════════════════════════════════════════════════════════════════════════

def _diagnostics(matrices: dict) -> list[dict]:
    """
    Scan every extracted value and return a list of alert dicts.
    Used by the dashboard's Diagnostics panel.
    """
    alerts: list[dict] = []

    def _check(section: str, component: str, limit: int, hours_list: list[str]) -> None:
        for idx, h in enumerate(hours_list):
            if h == "-":
                continue
            try:
                hrs = int(h)
            except ValueError:
                continue
            pct = (hrs / limit * 100) if limit else 0
            remaining = limit - hrs

            if hrs > limit:
                alerts.append({
                    "severity": "critical",
                    "section":   section,
                    "component": component,
                    "slot":      idx + 1,
                    "hours":     hrs,
                    "limit":     limit,
                    "pct":       round(pct, 1),
                    "message":   f"OVERDUE by {hrs - limit:,} hrs ({pct:.0f}% of limit)",
                })
            elif pct >= 90:
                alerts.append({
                    "severity": "warning",
                    "section":   section,
                    "component": component,
                    "slot":      idx + 1,
                    "hours":     hrs,
                    "limit":     limit,
                    "pct":       round(pct, 1),
                    "message":   f"Due in {remaining:,} hrs ({pct:.0f}% of limit)",
                })

    for row in matrices.get("main_engine", []):
        _check("Main Engine", row["component"], row["limit"], row.get("hours", []))

    for row in matrices.get("dg_general", []):
        _check("DG General", row["component"], row["limit"], row.get("hours", []))

    for row in matrices.get("aux_engine", []):
        for dg_idx, chunk in enumerate(row.get("dg_hours", [])):
            _check(f"Gen {dg_idx + 1}", row["component"], row["limit"], chunk)

    # Sort: critical first, then warning; within each severity sort by pct desc
    alerts.sort(key=lambda a: (0 if a["severity"] == "critical" else 1, -a["pct"]))
    return alerts


# ══════════════════════════════════════════════════════════════════════════════
# ❾  PUBLIC ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def extract_vessel_data(file_path: str) -> dict:
    """
    Parse a .doc / .docx engine running-hours report.
    Returns a fully structured dict ready for JSON serialisation.
    Never raises — all errors are caught and returned as status='error'.
    """
    path = Path(file_path)
    try:
        raw_bytes = path.read_bytes()
        cells     = _parse_cells(raw_bytes)
        log.info("Loaded %d cells from '%s'", len(cells), path.name)

        # Identify vessel
        vessel_name, is_fallback = identify_vessel(raw_bytes, path)
        cfg = FLEET_DATABASE[vessel_name]
        log.info("Vessel: %s  (ME×%d  DG×%d  Aux×%d)",
                 vessel_name, cfg.me_cyls, cfg.dg_count, cfg.aux_cyls)

        # Section boundaries
        aux_start, dg_start = _section_boundaries(cells)
        log.debug("Sections: ME=[0,%d)  AUX=[%d,%d)  DG=[%d,%d)",
                  aux_start, aux_start, dg_start, dg_start, len(cells))

        matrices: dict = {"main_engine": [], "dg_general": [], "aux_engine": []}

        # ── Main Engine ──────────────────────────────────────────────────────
        log.info("— Main Engine —")
        for comp, fb_limit in cfg.me_limits.items():
            idx   = _find_component(cells, comp, fb_limit, 0, aux_start, 14)
            limit = _read_doc_limit(cells, idx, fb_limit)
            vals  = _extract(cells, idx, 14, cfg.me_cyls)
            log.info("  %-32s limit=%-6d  %s", comp, limit, vals)
            matrices["main_engine"].append({
                "component": comp, "limit": limit, "hours": vals,
            })

        # ── DG General ───────────────────────────────────────────────────────
        log.info("— DG General —")
        for comp, fb_limit in cfg.dg_limits.items():
            idx   = _find_component(cells, comp, fb_limit, dg_start, len(cells), 18)
            limit = _read_doc_limit(cells, idx, fb_limit)
            vals  = _extract(cells, idx, 18, cfg.dg_count)
            log.info("  %-32s limit=%-6d  %s", comp, limit, vals)
            matrices["dg_general"].append({
                "component": comp, "limit": limit, "hours": vals,
            })

        # ── Aux Engine (cylinders) ───────────────────────────────────────────
        log.info("— Aux Engine —")
        total_aux = cfg.dg_count * cfg.aux_cyls
        for comp, fb_limit in cfg.aux_limits.items():
            idx   = _find_component(cells, comp, fb_limit, aux_start, dg_start, 24)
            limit = _read_doc_limit(cells, idx, fb_limit)
            flat  = _extract(cells, idx, 24, total_aux)
            chunks = [flat[g * cfg.aux_cyls: (g + 1) * cfg.aux_cyls]
                      for g in range(cfg.dg_count)]
            log.info("  %-32s limit=%-6d  DG1=%s", comp, limit, chunks[0])
            matrices["aux_engine"].append({
                "component": comp, "limit": limit, "dg_hours": chunks,
            })

        integrity = _score(matrices)
        diagnostics = _diagnostics(matrices)
        log.info("Integrity=%d%%  Alerts=%d", integrity, len(diagnostics))

        return {
            "status":      "success",
            "vessel":      vessel_name,
            "is_fallback": is_fallback,
            "config": {
                "me_cyls":  cfg.me_cyls,
                "dg_count": cfg.dg_count,
                "aux_cyls": cfg.aux_cyls,
            },
            "integrity":   integrity,
            "diagnostics": diagnostics,
            "matrices":    matrices,
        }

    except Exception as exc:
        log.exception("Parser crash on '%s': %s", path.name, exc)
        return {"status": "error", "detail": f"Parse error: {exc}"}