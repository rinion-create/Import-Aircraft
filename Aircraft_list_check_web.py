
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
    for root in roots:
        if not os.path.exists(root):
            continue
        for dirpath, _, filenames in depth_limited_walk(root, max_depth=max_depth):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                # quick extension check before regex
                if not fn.lower().endswith(".xlsx"):
                    continue
                if any(rx.search(fn) for rx in compiled):
                    matches.append(full)
                    if stop_after_first:
                        return matches
    return matches


def find_master_candidates_via_auto_search(info, debug,
                                           max_depth: int = DEFAULT_MASTER_SEARCH_DEPTH,
                                           stop_after_first: bool = True) -> List[str]:
    roots = tuple(list_existing_onedrive_roots())
    if not roots:
        info("No OneDrive roots available for auto-search.")
        return []
    info(f"Auto-searching Master under OneDrive roots: {roots} (max_depth={max_depth}, stop_after_first={stop_after_first})")
    matches = find_master_in_onedrive_cached(roots, tuple(MASTER_REGEX_PATTERNS), max_depth, stop_after_first)
    info(f"Auto-search found {len(matches)} candidate(s).")
    for i, p in enumerate(matches[:5], 1):
        debug(f"[AutoSearch] Candidate {i}: {p}")
    return matches


# --- Master loading & lookup building ---------------------------------------

def load_master_from_path(master_path: str, info, debug) -> Dict:
    """Load master file and build lookups from a filesystem path."""
    info(f"Master path: {master_path}")
    try:
        master_df = pd.read_excel(master_path, engine="openpyxl")
        info(f"Master rows loaded: {len(master_df)}")
    except Exception as e:
        raise RuntimeError(f"Failed to read master Excel file: {e}")

    return build_master_lookups(master_df, master_path, info, debug)


def load_master_from_buffer(master_buffer, info, debug) -> Dict:
    """Load master file and build lookups from an uploaded file buffer."""
    try:
        master_df = pd.read_excel(master_buffer, engine="openpyxl")
        info(f"Master rows loaded (uploaded): {len(master_df)}")
    except Exception as e:
        raise RuntimeError(f"Failed to read uploaded master Excel file: {e}")

    return build_master_lookups(master_df, "<uploaded master>", info, debug)


def build_master_lookups(master_df: pd.DataFrame, master_label: str, info, debug) -> Dict:
    man_col       = pick_column(master_df, MANUFACTURER_COL_CANDIDATES)
    type_col      = pick_column(master_df, TYPE_COL_CANDIDATES)
    eng_col       = pick_column(master_df, ENGINE_COL_CANDIDATES)
    id_col        = pick_column(master_df, ID_COL_CANDIDATES)
    type_icao_col = try_pick_column(master_df, TYPE_ICAO_COL_CANDIDATES)  # optional
    debug(f"Master columns: man={man_col}, type={type_col}, engine={eng_col}, id={id_col}, icao={type_icao_col}")

    master_df["__man__"]    = master_df[man_col].apply(normalize_preserve_case)
    master_df["__type__"]   = master_df[type_col].apply(normalize_preserve_case)
    master_df["__engine__"] = master_df[eng_col].apply(normalize_preserve_case)
    master_df["__icao__"]   = master_df[type_icao_col].apply(normalize_preserve_case) if type_icao_col else ""

    # Helper to detect a valid ICAO (non-empty and not 'N/A')
    def is_valid_icao(v: str) -> bool:
        return bool(v) and str(v).strip().upper() != "N/A"

    # Combo -> ID (case-sensitive)
    lookup: Dict[Tuple[str, str, str], str] = {}
    for _, r in master_df.iterrows():
        key = (r["__man__"], r["__type__"], r["__engine__"])
        mid = r[id_col]
        if pd.notna(mid) and key not in lookup:
            lookup[key] = mid

    man_set  = sorted(set(master_df["__man__"]))
    type_set = sorted(set(master_df["__type__"]))
    eng_set  = sorted(set(master_df["__engine__"]))
    info(f"Unique counts — Manufacturers: {len(man_set)}, Types: {len(type_set)}, Engines: {len(eng_set)}")
    info(f"Valid (Manufacturer,Type,Engine) combinations: {len(lookup)}")

    # Type -> ICAO map & valid-code type set
    type_to_icao: Dict[str, str] = {}
    valid_code_types: Set[str] = set()
    if type_icao_col:
        for _, r in master_df.iterrows():
            t = r["__type__"]
            code = r["__icao__"]
            if t not in type_to_icao and code != "":
                type_to_icao[t] = code
            if is_valid_icao(code):
                valid_code_types.add(t)

    # ICAO/IATA -> Type(s) mapping (alias-aware), skip 'N/A'
    icao_to_types: Dict[str, Set[str]] = {}
    if type_icao_col:
        for _, r in master_df.iterrows():
            icao_raw = r["__icao__"]
            tval     = r["__type__"]
            if not is_valid_icao(icao_raw):
                continue
            key, aliases = normalize_code_alias(icao_raw)
            for a in ({key} | aliases):
                icao_to_types.setdefault(a, set()).add(tval)
        debug(f"ICAO map keys loaded: {len(icao_to_types)}")

    # Precompute combo lists
    combo_list_all = list(lookup.keys())
    # Exclude any combo whose Type has ICAO 'N/A' (or empty)
    if type_icao_col:
        combo_list_valid_code = [c for c in combo_list_all if c[1] in valid_code_types]
    else:
        combo_list_valid_code = combo_list_all[:]  # no ICAO column -> cannot filter

    debug(f"Combos (all): {len(combo_list_all)}, combos (valid code only): {len(combo_list_valid_code)}")

    return {
        "df": master_df,
        "man_col": man_col, "type_col": type_col, "eng_col": eng_col, "id_col": id_col,
        "type_icao_col": type_icao_col,
        "lookup": lookup,
        "man_set": man_set, "type_set": type_set, "eng_set": eng_set,
        "type_to_icao": type_to_icao,
        "valid_code_types": valid_code_types,
        "icao_to_types": icao_to_types,
        "combo_list_all": combo_list_all,
        "combo_list_valid_code": combo_list_valid_code,
        "combo_ids": lookup,
        "master_path": master_label,
    }


# --- Scoring helpers ---------------------------------------------------------

def score_combos_against(man: str, typ: str, eng: str, combos: list, combo_ids: dict, top_n: int, cutoff: float):
    """
    Score given master combos vs the triple using difflib similarity
    on "Manufacturer Type Engine" (case-sensitive, hyphen-preserving).
    Returns top matches with ratio >= cutoff: [(ratio, (m,t,e), master_id), ...]
    """
    target = f"{man} {typ} {eng}"
    scored = []
    for combo in combos:
        m, t, e = combo
        cand = f"{m} {t} {e}"
        ratio = difflib.SequenceMatcher(None, target, cand).ratio()
        if ratio >= cutoff:
            scored.append((ratio, combo, combo_ids[combo]))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:top_n]


# --- Processing core ---------------------------------------------------------

def process_import(
    import_df: pd.DataFrame,
    master: Dict,
    suggestion_count: int,
    suggestion_cutoff: float,
    info, debug,
    orig_stem: str,  # original filename stem from uploaded import
):
    # Detect import columns (keeps original header & order)
    info("Detecting Manufacturer/Type/Engine columns in import...")
    try:
        imp_man_col  = pick_column(import_df, MANUFACTURER_COL_CANDIDATES)
        imp_type_col = pick_column(import_df, TYPE_COL_CANDIDATES)
        imp_eng_col  = pick_column(import_df, ENGINE_COL_CANDIDATES)
    except ValueError as e:
        raise RuntimeError(f"Error detecting columns in import file: {e}")
    imp_id_col = try_pick_column(import_df, ID_COL_CANDIDATES)
    debug(f"Import columns: man={imp_man_col}, type={imp_type_col}, engine={imp_eng_col}, id={imp_id_col}")

    # Normalize to internal helper columns
    df = import_df.copy()
    df["__man__"]    = df[imp_man_col].apply(normalize_preserve_case)
    df["__type__"]   = df[imp_type_col].apply(normalize_preserve_case)
    df["__engine__"] = df[imp_eng_col].apply(normalize_preserve_case)

    total          = len(df)
    match_count    = 0
    warn_count     = 0

    validation_rows = []
    suggested_import = import_df.copy()

    info(f"Validating {total} rows...")
    progress = st.progress(0)
    for idx, row in df.iterrows():
        man_raw  = import_df.at[idx, imp_man_col]
        type_raw = import_df.at[idx, imp_type_col]
        eng_raw  = import_df.at[idx, imp_eng_col]

        # Keep a normalized version of the ORIGINAL input type for exact comparisons in family
        type_raw_norm = normalize_preserve_case(type_raw)

        man = row["__man__"]
        typ = row["__type__"]
        eng = row["__engine__"]

        debug(f"[Row {idx}] Raw:   MAN='{man_raw}', TYPE='{type_raw}', ENG='{eng_raw}'")
        debug(f"[Row {idx}] Norm:  MAN='{man}', TYPE='{typ}', ENG='{eng}'")

        key = (man, typ, eng)
        combo_ok  = key in master["lookup"]
        master_id = master["lookup"].get(key, "")

        man_ok  = man in master["man_set"]
        type_ok = typ in master["type_set"]
        eng_ok  = eng in master["eng_set"]

        # Field-level suggestions (for report only)
        type_universe_for_suggestions = (
            [t for t in master["type_set"] if t in master["valid_code_types"]]
            if master["type_icao_col"] else master["type_set"]
        )
        man_suggestions  = [] if man_ok  else difflib.get_close_matches(man, master["man_set"],  n=suggestion_count, cutoff=suggestion_cutoff)
        type_suggestions = [] if type_ok else difflib.get_close_matches(typ, type_universe_for_suggestions, n=suggestion_count, cutoff=suggestion_cutoff)
        eng_suggestions  = [] if eng_ok  else difflib.get_close_matches(eng, master["eng_set"],  n=suggestion_count, cutoff=suggestion_cutoff)

        # ---------- ICAO/IATA-aware mapping ----------
        icao_type_suggestions = []
        icao_combo_hits = []
        code_key, code_aliases = normalize_code_alias(typ)
        alias_candidates = ({code_key} | code_aliases) if code_key else set()

        mapped_types = set()
        if master["type_icao_col"]:
            for a in alias_candidates:
                mapped_types |= master["icao_to_types"].get(a, set())
        icao_type_suggestions = sorted(mapped_types)  # already exclude 'N/A' coded types

        # Prefer exact combos unlocked via ICAO/IATA-derived types (manufacturer/engine fixed)
        if master["type_icao_col"]:
            for t_candidate in icao_type_suggestions:
                candidate_combo = (man, t_candidate, eng)
                if candidate_combo in master["lookup"]:
                    icao_combo_hits.append((1.0, candidate_combo, master["lookup"][candidate_combo]))

        # ---------- FAMILY DETECTION ----------
        fam_key = family_key(type_raw_norm)
        family_types = {t for t in master["type_set"] if family_key(t) == fam_key}
        if master["type_icao_col"]:
            # Restrict to types with valid ICAO codes unless override is enabled
            family_types = {t for t in family_types if t in master["valid_code_types"]}

        # ---------- Fuzzy scoring sets ----------
        # Global fuzzy across combos **with valid ICAO code only**
        global_combos = master["combo_list_valid_code"]
        global_scored = [] if combo_ok else score_combos_against(
            man, typ, eng, global_combos, master["combo_ids"], top_n=suggestion_count, cutoff=suggestion_cutoff
        )

        # Fuzzy restricted to family Types (and robust manufacturer match)
        filtered_combos = [
            c for c in master["combo_list_valid_code"]
            if man_eq(c[0], man) and (c[1] in family_types)
        ] if family_types else []
        filtered_scored = [] if combo_ok else score_combos_against(
            man, typ, eng, filtered_combos, master["combo_ids"], top_n=suggestion_count, cutoff=suggestion_cutoff
        )

        # ---------- Original-type-first preference (works with family) ----------
        original_type_hits = []
        if (not combo_ok) and (type_raw_norm in family_types):
            preferred_type_combos = [
                c for c in master["combo_list_valid_code"]
                if man_eq(c[0], man) and (c[1] == type_raw_norm)
            ]
            if preferred_type_combos:
                original_type_hits = score_combos_against(
                    man, typ, eng, preferred_type_combos, master["combo_ids"], top_n=suggestion_count, cutoff=suggestion_cutoff
                )

        # ---------- Base-type-first within family (if original type not present) ----------
        base_type_hits = []
        if (not combo_ok) and (not original_type_hits) and family_types:
            base_family_types = [t for t in family_types if not is_variant_type(t)]
            if base_family_types:
                base_type_combos = [
                    c for c in master["combo_list_valid_code"]
                    if man_eq(c[0], man) and (c[1] in base_family_types)
                ]
                if base_type_combos:
                    base_type_hits = score_combos_against(
                        man, typ, eng, base_type_combos, master["combo_ids"], top_n=suggestion_count, cutoff=suggestion_cutoff
                    )

        # ---------- Canonical family preference ----------
        canonical_ranked = []
        if (not combo_ok) and filtered_combos and not original_type_hits and not base_type_hits:
            for (m, t, e) in filtered_combos:
                base_score  = difflib.SequenceMatcher(None, f"{man} {typ} {eng}", f"{m} {t} {e}").ratio()
                shape_score = difflib.SequenceMatcher(None, str(type_raw), t).ratio()  # original input vs candidate type
                penalty     = 0.25 if is_variant_type(t) else 0.0  # stronger penalty for variants
                final_score = (base_score * 0.4) + (shape_score * 0.6) - penalty
                canonical_ranked.append((final_score, (m, t, e), master["lookup"][(m, t, e)]))
            canonical_ranked.sort(reverse=True, key=lambda x: x[0])

        # ---------- Choose suggestion ----------
        suggestion_source = ""
        sug_combo = None
        sug_mid   = ""

        if combo_ok:
            match_count += 1
            sug_combo = key
            sug_mid   = master_id
            suggestion_source = "exact"
        else:
            warn_count += 1
            if original_type_hits:
                ratio, sug_combo, sug_mid = original_type_hits[0]
                suggestion_source = f"original_type_preferred(ratio={ratio:.3f})"
            elif base_type_hits:
                ratio, sug_combo, sug_mid = base_type_hits[0]
                suggestion_source = f"base_family_type_preferred(ratio={ratio:.3f})"
            elif icao_combo_hits:
                _, sug_combo, sug_mid = icao_combo_hits[0]
                suggestion_source = "icao_exact_combo"
            elif canonical_ranked:
                ratio, sug_combo, sug_mid = canonical_ranked[0]
                suggestion_source = f"canonical_family_choice(score={ratio:.3f})"
            elif filtered_scored:
                ratio, sug_combo, sug_mid = filtered_scored[0]
                suggestion_source = f"family_restricted_combo(ratio={ratio:.3f})"
            elif global_scored:
                ratio, sug_combo, sug_mid = global_scored[0]
                suggestion_source = f"closest_combo_valid_code_only(ratio={ratio:.3f})"
            else:
                sug_combo = key
                sug_mid   = ""
                suggestion_source = "none"

        debug(f"[Row {idx}] → Suggestion source: {suggestion_source}")
        debug(f"[Row {idx}] → Suggested combo: {sug_combo} | MasterID={sug_mid if sug_mid else '—'}")

        # Write into Suggested Import (preserving original columns/order)
        sug_man, sug_typ, sug_eng = sug_combo
        suggested_import.at[idx, imp_man_col]  = sug_man
        suggested_import.at[idx, imp_type_col] = sug_typ
        suggested_import.at[idx, imp_eng_col]  = sug_eng
        if imp_id_col is not None:
            suggested_import.at[idx, imp_id_col] = sug_mid  # update ID ONLY if original has such a column

        # Collect validation row (diagnostics)
        validation_rows.append({
            "RowIndex": idx,
            "Manufacturer": man_raw,
            "Type": type_raw,
            "Engine": eng_raw,
            "Manufacturer_norm": man,
            "Type_norm": typ,
            "Engine_norm": eng,
            "FieldManufacturerExists": man_ok,
            "FieldTypeExists": type_ok,
            "FieldEngineExists": eng_ok,
            "ComboMatch": combo_ok,
            "MasterID": master_id,
            "SuggestionsManufacturer": "; ".join(man_suggestions),
            "SuggestionsType": "; ".join(type_suggestions),
            "SuggestionsTypeFromICAO": "; ".join(icao_type_suggestions),
            "SuggestionsEngine": "; ".join(eng_suggestions),
            # Lists of combos for transparency
            "SuggestionsCombo_CodeRestricted": " | ".join([f"{m} {t} {e}" for _, (m,t,e), _ in filtered_scored]),
            "SuggestionsCombo_Global": " | ".join([f"{m} {t} {e}" for _, (m,t,e), _ in global_scored]),
            "SuggestionsComboFromICAO": " | ".join([f"{m} {t} {e}" for _, (m,t,e), _ in icao_combo_hits]),
            "ChosenSuggestionSource": suggestion_source,
            "ChosenManufacturer": sug_man,
            "ChosenType": sug_typ,
            "ChosenEngine": sug_eng,
            "ChosenMasterID": sug_mid,
        })

        progress.progress(min((idx + 1) / max(total, 1), 1.0))

    # --- Build outputs in-memory --------------------------------------------

    # 1) Validation report (diagnostics)
    summary_df = pd.DataFrame([{
        "Master file used": master["master_path"],
        "Total rows validated": total,
        "Successful combination matches": match_count,
        "Warnings (no combination match)": warn_count,
        "Type ICAO column detected": bool(master["type_icao_col"]),
        "Generated at": datetime.now().isoformat(timespec="seconds"),
    }])
    report_df  = pd.DataFrame(validation_rows)

    # Create Excel with hidden columns
    validation_bytes = io.BytesIO()
    with pd.ExcelWriter(validation_bytes, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        report_df.to_excel(writer, sheet_name="Validation", index=False)
    validation_bytes.seek(0)

    # Hide selected columns using openpyxl
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter

        wb = load_workbook(validation_bytes)
        ws = wb["Validation"]
        header_row = [cell.value for cell in ws[1]]

        # Columns to hide (by header name)
        cols_to_hide = [
            "Manufacturer_norm",
            "Type_norm",
            "Engine_norm",
            "SuggestionsType",
            "SuggestionsTypeFromICAO",
            "SuggestionsEngine",
            "SuggestionsCombo_CodeRestricted",
            "SuggestionsComboFromICAO",
        ]

        # Hide the specified columns if they exist
        name_to_idx = {str(v): i+1 for i, v in enumerate(header_row) if v is not None}
        for name in cols_to_hide:
            if name in name_to_idx:
                col_letter = get_column_letter(name_to_idx[name])
                ws.column_dimensions[col_letter].hidden = True

        out_validation = io.BytesIO()
        wb.save(out_validation)
        out_validation.seek(0)
    except Exception:
        # Fallback: use the original (unhidden) version
        out_validation = validation_bytes

    # 2) Suggested Import (keeps exact original header order & structure)
    out_suggested = io.BytesIO()
    with pd.ExcelWriter(out_suggested, engine="openpyxl") as writer:
        suggested_import.to_excel(writer, sheet_name="SuggestedImport", index=False)
    out_suggested.seek(0)

    # Filenames derived from original uploaded file name (stem)
    sugg_name = f"{orig_stem}_import_ready.xlsx"
    report_name = f"{orig_stem}_validation_report.xlsx"

    metrics = {
        "total": total,
        "matches": match_count,
        "warnings": warn_count,
        "report_name": report_name,
        "sugg_name": sugg_name,
    }

    return out_validation, out_suggested, metrics, summary_df, report_df, suggested_import


# ---------------------- UI LAYOUT -------------------------------------------

st.title("🛫 SafetyManager365")
st.subheader("Aircraft Import Validator & Suggester")
st.caption("Ruben Inion v0.2 2026 — original-name outputs, OneDrive save, auto Master discovery.")

with st.sidebar:
    st.header("Settings")
    verbose = st.toggle("Verbose logging", value=True)
    suggestion_count = st.slider("Suggestion count", min_value=1, max_value=10, value=DEFAULT_SUGGESTION_COUNT, step=1)
    suggestion_cutoff = st.slider("Similarity cutoff", min_value=0.30, max_value=0.90, value=DEFAULT_SUGGESTION_CUTOFF, step=0.05)

    st.divider()
    st.subheader("Master source")
    master_upload = st.file_uploader("Upload Master Excel (.xlsx) to override auto-detection", type=["xlsx"])
    auto_search_master = st.checkbox("Auto-search Master across OneDrive (fallback)", value=True)
    max_search_depth = st.slider("Auto-search depth", min_value=2, max_value=12, value=DEFAULT_MASTER_SEARCH_DEPTH, step=1,
                                 help="Limits recursion when scanning your OneDrive to find the Master file.")

    st.divider()
    st.subheader("Output")
    save_to_onedrive = st.checkbox("Save to OneDrive", value=False, help="Save output files to your OneDrive")
    onedrive_subfolder = st.text_input(
        "OneDrive subfolder (optional)",
        value=PROJECT_FOLDER,
        help="Relative to your OneDrive root. Leave as default to use the project tools folder."
    )

info, debug, flush_logs = get_logger(verbose)

col1, col2 = st.columns(2)
with col1:
    import_upload = st.file_uploader("Upload Import Excel (.xlsx)", type=["xlsx"])

with col2:
    st.markdown("**Master source status**")
    master_source_placeholder = st.empty()

run_btn = st.button("▶️ Validate & Suggest", type="primary", disabled=(import_upload is None))

# Resolve master
master_dict: Optional[Dict] = None
master_status_lines: List[str] = []

if master_upload is not None:
    try:
        master_dict = load_master_from_buffer(master_upload, info, debug)
        master_status_lines.append("Using **uploaded Master**.")
    except Exception as e:
        st.error(f"Failed to load uploaded Master: {e}")

else:
    # First try fixed candidates
    candidates = master_path_candidates()
    existing = [p for p in candidates if os.path.exists(p)]
    if existing:
        try:
            master_dict = load_master_from_path(existing[0], info, debug)
            master_status_lines.append(f"Master auto-detected (fixed path): `{existing[0]}`")
        except Exception as e:
            st.error(f"Failed to read detected Master: {e}")
    else:
        master_status_lines.append("No fixed-path OneDrive Master found.")
        # Auto-search (fallback)
        if auto_search_master:
            matches = find_master_candidates_via_auto_search(info, debug, max_depth=max_search_depth, stop_after_first=True)
            if matches:
                try:
                    master_dict = load_master_from_path(matches[0], info, debug)
                    master_status_lines.append(f"Master found via **OneDrive auto-search**: `{matches[0]}`")
                except Exception as e:
                    st.error(f"Failed to read auto-searched Master: {e}")
            else:
                master_status_lines.append("Auto-search did not find a Master file. You can upload a Master file in the sidebar.")
        else:
            master_status_lines.append("Auto-search disabled. You can upload a Master file in the sidebar.")

master_source_placeholder.markdown("\n\n".join(master_status_lines))

# Run processing
if run_btn and import_upload is not None:
    try:
        import_df = pd.read_excel(import_upload, engine="openpyxl")
        info(f"Import rows loaded: {len(import_df)}")

        if master_dict is None:
            st.error("Master dataset is not available. Please upload a Master file or enable OneDrive auto-search.")
        else:
            # Derive stem from uploaded filename
            import_stem = original_file_stem(getattr(import_upload, "name", None))

            out_validation, out_suggested, metrics, summary_df, report_df, suggested_import = process_import(
                import_df=import_df,
                master=master_dict,
                suggestion_count=suggestion_count,
                suggestion_cutoff=suggestion_cutoff,
                info=info,
                debug=debug,
                orig_stem=import_stem,
            )

            # Optional OneDrive Save
            if save_to_onedrive:
                output_dir = resolve_onedrive_output_dir(onedrive_subfolder)
                if output_dir is None:
                    st.warning("OneDrive root not found. Files were not saved. Please check your OneDrive setup.")
                    info("OneDrive save skipped: no OneDrive root detected.")
                else:
                    sugg_path = os.path.join(output_dir, metrics["sugg_name"])
                    report_path = os.path.join(output_dir, metrics["report_name"])
                    try:
                        save_bytes_to_path(out_suggested, sugg_path)
                        save_bytes_to_path(out_validation, report_path)
                        # Rewind buffers for subsequent download buttons
                        out_suggested.seek(0)
                        out_validation.seek(0)
                        st.success(f"Saved to OneDrive:\n- {sugg_path}\n- {report_path}")
                        info(f"Saved Suggested Import to {sugg_path}")
                        info(f"Saved Validation Report to {report_path}")
                    except Exception as e:
                        st.error(f"Failed to save to OneDrive: {e}")
                        info(f"OneDrive save failed: {e}")

            # Metrics
            st.success("Processing complete.")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Rows validated", metrics["total"])
            m2.metric("Exact matches", metrics["matches"])
            m3.metric("Warnings", metrics["warnings"])
            m4.metric("ICAO column detected", "Yes" if bool(master_dict["type_icao_col"]) else "No")

            # Previews
            st.subheader("Summary")
            st.dataframe(summary_df, use_container_width=True)

            st.subheader("Validation preview")
            st.dataframe(report_df.head(50), use_container_width=True)

            st.subheader("Suggested Import preview")
            st.dataframe(suggested_import.head(50), use_container_width=True)

            # Downloads
            st.subheader("Downloads")
            cdl1, cdl2 = st.columns(2)
            with cdl1:
                st.download_button(
                    label=f"⬇️ Download Validation Report ({metrics['report_name']})",
                    data=out_validation,
                    file_name=metrics["report_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with cdl2:
                st.download_button(
                    label=f"⬇️ Download Suggested Import ({metrics['sugg_name']})",
                    data=out_suggested,
                    file_name=metrics["sugg_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except Exception as e:
        st.error(f"Processing failed: {e}")

    flush_logs()
