
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Streamlit App — Aircraft import validator & suggester
(CASE-SENSITIVE, hyphen-preserving, ICAO alias-aware, avoids 'N/A' code suggestions,
 original-type-first, base-family preference, canonical preference, preserves original import structure)

Features:
- Upload Import Excel (.xlsx).
- Master Excel:
    - Auto-resolve from OneDrive by scanning common fixed paths, then
      FALL BACK to an automatic recursive search across your OneDrive (configurable depth).
    - OR upload a Master file to override auto-resolve.
- Validates Manufacturer, Type, Engine, and their combination (strict, case-sensitive).
- Suggests corrections prioritizing:
    1) Exact combo
    2) ICAO/IATA code-derived exact combo (alias-aware, skips 'N/A' codes)
    3) Original-type-first within family
    4) Base family type (numeric-only canonical types, e.g., '737-800')
    5) Canonical family choice (closest to original, penalizing variant flags: W, ER, BCF, etc.)
    6) Global fuzzy fallback, excluding Types whose ICAO is 'N/A'
- Outputs:
    - Validation report (Summary + Validation) with diagnostic columns HIDDEN
    - Suggested Import workbook preserving original columns & order
      (only Manufacturer/Type/Engine and MasterID—if present—are updated)
    - Easy download buttons and optional OneDrive save

Requires:
- Python 3.8+
- pandas (openpyxl engine)
"""

import io
import os
import re
import difflib
import platform
from datetime import datetime
from typing import Optional, Tuple, Set, Dict, List, Iterable

import pandas as pd
import streamlit as st

# ---------------------- UI CONFIG --------------------------------------------
st.set_page_config(
    page_title="Aircraft Import Validator & Suggester",
    page_icon="🛫",
    layout="wide",
)

# ---------------------- VERBOSE LOGGING --------------------------------------
def get_logger(verbose: bool):
    logs = []

    def info(msg: str):
        if verbose:
            logs.append(f"[INFO] {msg}")

    def debug(msg: str):
        if verbose:
            logs.append(f"[DEBUG] {msg}")

    def flush():
        if logs:
            with st.expander("Logs (verbose)", expanded=False):
                st.write("\n".join(logs))

    return info, debug, flush

# --- Configuration -----------------------------------------------------------

# Dynamic OneDrive folder resolution under the current user's HOME
HOME = os.path.expanduser("~")

# macOS OneDrive mount root (typical)
MAC_ONEDRIVE_ROOTS = [
    os.path.join(HOME, "Library", "CloudStorage", "OneDrive-Comply365"),
    os.path.join(HOME, "Library", "CloudStorage", "OneDrive - Comply365"),
    os.path.join(HOME, "Library", "CloudStorage", "OneDrive"),
]

# Windows OneDrive root (typical)
WIN_ONEDRIVE_ROOTS = [
    os.path.join(HOME, "OneDrive - Comply365"),
    os.path.join(HOME, "OneDrive"),
]

PROJECT_FOLDER = os.path.join(
    "ASQS - Project and Account Management - Project and Account Management",
    "04. Tools"
)

# Historical/expected exact names
MASTER_FILENAMES = [
    "260101_Aicraft_Master.xlsx",
    "260101_Aircraft_Master.xlsx",
]

# Regex patterns (case-insensitive) to match more liberally during auto-search
MASTER_REGEX_PATTERNS = [
    r"(?i)\b(?:\d{6}_)?aicraft_master\.xlsx$",   # tolerate date prefix + typo 'Aicraft'
    r"(?i)\b(?:\d{6}_)?aircraft_master\.xlsx$",  # tolerate date prefix
    r"(?i)\b.*aircraft.*master.*\.xlsx$",        # broader fallback
]

MANUFACTURER_COL_CANDIDATES = ["manufacturer", "aircraft manufacturer", "mfr", "oem", "maker"]
TYPE_COL_CANDIDATES         = ["type", "aircraft type", "model", "family", "series", "aircraft model"]
ENGINE_COL_CANDIDATES       = ["engine", "engine type", "engine model", "powerplant", "motor"]
ID_COL_CANDIDATES           = ["id", "aircraft id", "aircraft_id", "uid", "key", "master id", "aircraft code"]

# Optional ICAO/IATA type designator column candidates (case-insensitive)
TYPE_ICAO_COL_CANDIDATES    = [
    "icao", "icao type", "icao code", "icao designator", "icao type designator",
    "type designator", "icao aircraft type", "icao model"
]

# Defaults (can be overridden via sidebar)
DEFAULT_SUGGESTION_COUNT  = 3
DEFAULT_SUGGESTION_CUTOFF = 0.60  # minimum similarity to keep a combo suggestion
DEFAULT_MASTER_SEARCH_DEPTH = 6   # depth for recursive OneDrive Master search

# --- Helpers ----------------------------------------------------------------

def sanitize_path(p: str) -> str:
    """Strip surrounding quotes, expand ~ and env vars, normalize."""
    if not p:
        return p
    p = p.strip()
    if (p.startswith("'") and p.endswith("'")) or (p.startswith('"') and p.endswith('"')):
        p = p[1:-1]
    p = os.path.expanduser(os.path.expandvars(p))
    return os.path.normpath(p)


def normalize_preserve_case(value) -> str:
    """
    Normalize while preserving case and hyphens:
    - Trim whitespace
    - Unify EN/EM dash to '-' and tighten spaces around hyphens ("737 - 700" -> "737-700")
    - Collapse remaining spaces
    - DO NOT lowercase
    """
    if pd.isna(value):
        return ""
    s = str(value).strip()
    s = s.replace("–", "-").replace("—", "-")  # unify dashes
    s = re.sub(r"\s*-\s*", "-", s)            # tighten hyphens
    s = re.sub(r"\s+", " ", s)                # collapse spaces
    return s


def pick_column(df: pd.DataFrame, candidates: Iterable[str]) -> str:
    """Pick best-matching column name (case-insensitive on headers)."""
    df_cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in df_cols_lower:
            return df_cols_lower[cand.lower()]
    # fuzzy fallback
    all_lower = list(df_cols_lower.keys())
    for cand in candidates:
        m = difflib.get_close_matches(cand.lower(), all_lower, n=1, cutoff=0.8)
        if m:
            return df_cols_lower[m[0]]
    raise ValueError(f"Could not auto-detect any of {list(candidates)}. Available: {list(df.columns)}")


def try_pick_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    try:
        return pick_column(df, candidates)
    except ValueError:
        return None


# ---------------- ICAO/IATA alias & family helpers ---------------------------

def normalize_code_alias(code: str) -> Tuple[str, Set[str]]:
    """
    Return a canonical key and a set of alias keys for an aircraft type code.
    - Canonical key removes spaces/hyphens (case preserved).
    - If pattern is letters? + digits (2-3) + letters? (e.g., 'B763', 'A321', 'B77W'),
      also include a numeric-only alias (digits + optional trailing letters).
    """
    if not code:
        return "", set()

    raw = str(code).strip()
    key = re.sub(r"[\s\-]+", "", raw)  # remove spaces/hyphens for mapping keys

    aliases = {key}
    m = re.match(r"^([A-Za-z]{0,2})(\d{2,3})([A-Za-z]{0,2})$", key)
    if m:
        num = m.group(2)
        tail = m.group(3)
        if tail:
            aliases.add(num + tail)  # e.g., '77W'
        else:
            aliases.add(num)         # e.g., '763'
    return key, aliases


def family_key(t: str) -> str:
    """
    Derive a family key from a Type string:
    - normalize preserving case/hyphen
    - strip all letters; keep digits and hyphens
    Examples:
      'B737-800'   -> '737-800'
      'B737-800W'  -> '737-800'
      '737-800BCF' -> '737-800'
      'A320-214'   -> '320-214'
    """
    tn = normalize_preserve_case(t)
    return re.sub(r"[A-Za-z]", "", tn)


def is_variant_type(t: str) -> bool:
    """
    Returns True if type string includes variant letters (prefix or suffix)
    beyond the canonical number-hyphen-number pattern:
    e.g., 'B737-800', 'B737-800W', '737-800BCF', '737-800ER', etc.
    """
    x = t.replace("-", "")
    # leading letters like 'B737800' are non-canonical
    if re.match(r"^[A-Za-z]+\d+", x):
        return True
    # trailing letters like '737800W' or '737800BCF'
    if re.search(r"[A-Za-z]+$", x) and not x.isdigit():
        return True
    return False


def man_eq(cand: str, given: str) -> bool:
    """Case-sensitive first, fallback to case-insensitive for candidate filtering."""
    return (cand == given) or (cand.upper() == given.upper())


# --- Master path resolution (fixed candidates) --------------------------------

def master_path_candidates() -> List[str]:
    roots = MAC_ONEDRIVE_ROOTS if platform.system() == "Darwin" else WIN_ONEDRIVE_ROOTS
    candidates: List[str] = []
    for root in roots:
        base = os.path.join(root, PROJECT_FOLDER)
        for fname in MASTER_FILENAMES:
            candidates.append(os.path.join(base, fname))
    return candidates


# --- OneDrive output resolution & save helpers -------------------------------

def resolve_onedrive_output_dir(custom_subdir: Optional[str] = None) -> Optional[str]:
    """
    Resolve a writable OneDrive output directory:
    - Picks the first existing OneDrive root (macOS or Windows).
    - Uses the configured PROJECT_FOLDER by default.
    - If custom_subdir is provided, it is sanitized and appended (or used as absolute if under OneDrive).
    Returns a normalized absolute path, or None if no root exists.
    """
    roots = MAC_ONEDRIVE_ROOTS if platform.system() == "Darwin" else WIN_ONEDRIVE_ROOTS
    existing_roots = [r for r in roots if os.path.exists(r)]
    if not existing_roots:
        return None

    base_root = existing_roots[0]
    if custom_subdir:
        sub = sanitize_path(custom_subdir)
        if os.path.isabs(sub) and sub.startswith(base_root):
            out_dir = sub
        else:
            out_dir = os.path.join(base_root, sub)
    else:
        out_dir = os.path.join(base_root, PROJECT_FOLDER)

    os.makedirs(out_dir, exist_ok=True)
    return os.path.normpath(out_dir)


def save_bytes_to_path(buf: io.BytesIO, target_path: str) -> None:
    """Persist a BytesIO to target_path."""
    buf.seek(0)
    with open(target_path, "wb") as f:
        f.write(buf.read())


def original_file_stem(filename: Optional[str]) -> str:
    """
    Return the original filename stem (without extension).
    Falls back to 'import' if missing, empty, or extensionless.
    """
    try:
        if not filename:
            return "import"
        base = os.path.basename(str(filename).strip().strip("'\""))
        stem, _ = os.path.splitext(base)
        return stem if stem else "import"
    except Exception:
        return "import"


# --- OneDrive Master auto-discovery ------------------------------------------

def list_existing_onedrive_roots() -> List[str]:
    roots = MAC_ONEDRIVE_ROOTS if platform.system() == "Darwin" else WIN_ONEDRIVE_ROOTS
    return [r for r in roots if os.path.exists(r)]


def depth_limited_walk(root: str, max_depth: int):
    """
    Yield (dirpath, dirnames, filenames) like os.walk but stop descending after max_depth.
    Depth is measured as levels below the root.
    """
    root = os.path.normpath(root)
    root_sep_count = root.count(os.sep)
    for dirpath, dirnames, filenames in os.walk(root):
        depth = os.path.normpath(dirpath).count(os.sep) - root_sep_count
        if depth >= max_depth:
            # prevent os.walk from descending further
            dirnames[:] = []
        yield dirpath, dirnames, filenames


@st.cache_data(show_spinner=False, ttl=300)
def find_master_in_onedrive_cached(roots: Tuple[str, ...],
                                   patterns: Tuple[str, ...],
                                   max_depth: int,
                                   stop_after_first: bool) -> List[str]:
    """Cached helper for performance — returns list of matching file paths."""
    compiled = [re.compile(p) for p in patterns]
    matches: List[str] = []
