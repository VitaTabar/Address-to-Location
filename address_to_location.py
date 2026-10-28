# =============================================================================
# address_to_location.py
#
# Matches traffic camera address records (Excel) to road segments (Shapefile).
#
# ── ADDRESS STRUCTURE ────────────────────────────────────────────────────────
#
# Each camera address in "نام مکان" follows this pattern:
#   [seg_name] از [origin] تا [destination] مسیر [direction] در منطقه [district]
#
# Mapped to shapefile columns:
#   seg_name    → Name  (fallback: OldName)
#   origin      → Origin
#   destination → Destinatio    (origin↔dest swap is always tried)
#   direction   → NetDirecti    (taken from "جهت" column, NOT parsed from text)
#   district    → District      (used only for preference, never as a hard filter)
#
# ── MATCHING PIPELINE ────────────────────────────────────────────────────────
#
# STAGE 1A – DirectMatch (exact MixValue, same district)
#   Normalize both sides and check for exact string equality against MixValue,
#   including the district suffix.
#
# STAGE 1B – DirectMatch (exact MixValue, cross-district)
#   Same as 1A but with the district suffix stripped from both sides before
#   comparison. Handles streets that run along district boundaries and may be
#   stored under a different district number in the shapefile than the one
#   written in the camera address.
#   When multiple shapefile rows match (same street in multiple districts),
#   the row whose District column matches the camera's district is preferred.
#   Rows matched in 1A and 1B are both saved to DirectMatch.xlsx.
#
# STAGE 1C – OldName DirectMatch
#   Some camera addresses use the previous/old street name (stored in OldName).
#   The parsed seg_name is matched against OldName; if found, the current Name
#   is substituted and the result checked against MixValue (district-agnostic).
#   These go into DirectMatch.xlsx.
#
# STAGE 2 – ComponentMatch (FuzzyMatch)
#   For rows not matched in Stage 1:
#   a) CANDIDATE FILTER – keep rows where Name or OldName scores ≥ NAME_FILTER_THRESH.
#      District is NOT used as a filter — cross-district segments are candidates.
#   b) COMPONENT SCORING (weighted):
#        Name / OldName  → weight W_NAME
#        Origin          → weight W_ORIG   (noise stripped, see od_score)
#        Destinatio      → weight W_DEST   (noise stripped, origin↔dest swap tried)
#        NetDirecti      → weight W_DIR    (order-sensitive, see dir_score)
#   c) DIRECTION PENALTY – if direction score < DIR_PENALTY_THRESHOLD, multiply
#      total by DIR_PENALTY_FACTOR to discourage wrong-direction matches.
#   d) Accept best candidate if total ≥ COMPONENT_THRESHOLD.
#
# ── SCORING DESIGN DECISIONS ─────────────────────────────────────────────────
#
# od_score (Origin / Destination):
#   Strips noise words (ورودی خروجی شمال شمالی جنوب جنوبی به از) before comparing.
#   Uses token_sort_ratio as base score.
#   Boosts with partial_ratio ONLY when the shorter stripped string is a literal
#   substring of the longer — handles 'شیان'⊂'بوستان جنگلی شیان' without
#   inflating unrelated words that share character sequences.
#
# dir_score (Direction):
#   Uses fuzz.ratio() (order-sensitive). token_sort_ratio would make
#   'شمال به جنوب' == 'جنوب به شمال' = 100, hiding opposite directions.
#   fuzz.ratio(): same direction = 100, opposite ≈ 33.
#
# Direction penalty:
#   If dir_score < DIR_PENALTY_THRESHOLD, total × DIR_PENALTY_FACTOR.
#   Prevents a wrong-direction segment from winning over a correct one even
#   when other components score well by coincidence.
#
# ── REQUIREMENTS ─────────────────────────────────────────────────────────────
#   pip install geopandas pandas openpyxl rapidfuzz
#
# ── OUTPUT FILES ─────────────────────────────────────────────────────────────
#   DirectMatch.xlsx    – rows matched exactly (Stage 1A, 1B, or 1C)
#   FuzzyMatch.xlsx     – rows matched by component scoring (Stage 2)
#   Unmatched.xlsx      – rows below threshold (needs manual review)
#   all_matched.gpkg    – combined spatial layer for QGIS / ArcGIS
# =============================================================================

import os
import re
import geopandas as gpd
import pandas as pd
from rapidfuzz import fuzz

# ── USER SETTINGS ─────────────────────────────────────────────────────────────
SHP_PATH   = os.environ.get("SHP_PATH",   os.path.join("data", "street_segments.shp"))   # .shp, .dbf, .shx, .prj must be together
EXCEL_PATH = os.environ.get("EXCEL_PATH", os.path.join("data", "cameras.xlsx"))

OUT_DIR = os.environ.get("CAMERA_MATCH_OUTPUT", "output")
os.makedirs(OUT_DIR, exist_ok=True)

OUT_DIRECT    = os.path.join(OUT_DIR, "DirectMatch.xlsx")
OUT_FUZZY     = os.path.join(OUT_DIR, "FuzzyMatch.xlsx")
OUT_UNMATCHED = os.path.join(OUT_DIR, "Unmatched.xlsx")
OUT_GPKG      = os.path.join(OUT_DIR, "all_matched.gpkg")

COMPONENT_THRESHOLD    = 87
NAME_FILTER_THRESH     = 65
W_NAME, W_ORIG, W_DEST, W_DIR = 0.40, 0.25, 0.25, 0.10
DIR_PENALTY_THRESHOLD  = 50
DIR_PENALTY_FACTOR     = 0.80


# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────

_NOISE = re.compile(r'\b(ورودی|خروجی|شمال|شمالی|جنوب|جنوبی|به|از)\b')

def clean(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = s.replace("\u200c", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", s).strip()


def strip_noise(s):
    return re.sub(r"\s+", " ", _NOISE.sub("", clean(s))).strip()


def od_score(cam_val, shp_val):
    sa = strip_noise(cam_val)
    sb = strip_noise(shp_val)
    if not sa or not sb:
        return 0
    base = fuzz.token_sort_ratio(sa, sb)
    shorter, longer = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
    if shorter in longer:
        return max(base, fuzz.partial_ratio(sa, sb))
    return base


def name_similarity(cam_names, shp_name, shp_oldname):
    """Max fuzzy similarity between any candidate camera-side name and the
    shapefile's Name/OldName. cam_names is a list because intersection-style
    addresses ("A - B") give two equally plausible street names to try.

    Guards against a rapidfuzz quirk: fuzz.token_sort_ratio('', '') == 100
    and fuzz.partial_ratio('', '') == 100 (two empty strings are treated as
    a perfect match). Without this guard, any row whose seg_name failed to
    parse (empty string) would score a false 100 against any shapefile row
    whose OldName is also blank -- silently corrupting the "best candidate"
    selection for every unparsed row, even though nothing was actually
    compared.
    """
    best = 0
    for cn in cam_names:
        if not cn:
            continue
        for sn in (shp_name, shp_oldname):
            if not sn:
                continue
            best = max(best, fuzz.token_sort_ratio(cn, sn), fuzz.partial_ratio(cn, sn))
    return best


_DIR_CANON = {
    "شمال": "شمال", "شمالی": "شمال",
    "جنوب": "جنوب", "جنوبی": "جنوب",
    "شرق":  "شرق",  "شرقی":  "شرق",
    "غرب":  "غرب",  "غربی":  "غرب",
}
_DIR_TOKEN_RE = re.compile(r"شمالی|شمال|جنوبی|جنوب|شرقی|شرق|غربی|غرب")
_TWO_WAY_RE   = re.compile(r"دو\s*طرفه")


def extract_directions(s):
    """Pull only the cardinal-direction words out of a string, in the order
    they appear, collapsed to canonical form (شمالی -> شمال, etc.), with
    consecutive duplicates removed. Ignores everything else (lane counts,
    'دو طرفه', 'به', etc.)."""
    s = clean(s)
    tokens = [_DIR_CANON[t] for t in _DIR_TOKEN_RE.findall(s)]
    seq = []
    for t in tokens:
        if not seq or seq[-1] != t:
            seq.append(t)
    return seq


def dir_score(cam_dir, shp_dir):
    ca, cb = clean(cam_dir), clean(shp_dir)
    if not ca or not cb:
        return 0

    seq_a = extract_directions(ca)
    seq_b = extract_directions(cb)
    if not seq_a or not seq_b:
        # fallback: no recognizable direction words on one side, use old behaviour
        return fuzz.ratio(ca, cb)

    # Two-way ("دو طرفه") segments serve traffic in both directions, so order
    # doesn't matter as long as the same set of directions is present.
    if _TWO_WAY_RE.search(cb) or _TWO_WAY_RE.search(ca):
        if set(seq_a) == set(seq_b):
            return 100

    joined_a = "".join(seq_a)
    joined_b = "".join(seq_b)

    if seq_a == seq_b:
        return 100
    if seq_a == list(reversed(seq_b)):
        return 0          # explicitly opposite direction -> hard fail
    if joined_a in joined_b or joined_b in joined_a:
        return 90         # one is a sub-sequence of the other (e.g. missing a leg)

    return fuzz.ratio(joined_a, joined_b)


def norm_with_district(s):
    """Normalise spacing only — keep district suffix for Stage 1A."""
    s = clean(s)
    return re.sub(r"منطقه\s+(\d)", r"منطقه\1", s)


def norm_no_district(s):
    """Remove district suffix — for cross-district Stage 1B/1C matching."""
    s = clean(s)
    s = re.sub(r"منطقه\s+(\d)", r"منطقه\1", s)
    return re.sub(r"\s*در منطقه\d+\s*$", "", s).strip()


def cam_district(raw):
    """Extract the district number from a camera address string."""
    m = re.search(r"منطقه\s*(\d+)", clean(raw))
    return int(m.group(1)) if m else None


def parse_camera_address(raw):
    """Parse a camera address into {seg_name, origin, dest, format}.

    The canonical pattern is:
        [seg_name] از [origin] تا [dest] مسیر [direction] در منطقه [d]

    But real data contains several other shapes that the old single-regex
    parser silently failed on (leaving origin/dest empty even though the
    information was present in the text). We now try each shape in order
    of specificity and record which one matched in "format", so failures
    are traceable instead of silently producing empty fields.
    """
    s = clean(raw)
    result = {"seg_name": "", "origin": "", "dest": "", "format": "unrecognized"}

    # seg_name = whatever precedes " از " (present in most non-intersection formats)
    m_name = re.search(r"^(.+?)\s+از\s+", s)
    if m_name:
        result["seg_name"] = m_name.group(1).strip()

    # 1) Full canonical form: از X تا Y مسیر ...
    m = re.search(r"از\s+(.+?)\s+تا\s+(.+?)\s+مسیر", s)
    if m:
        result["origin"], result["dest"] = m.group(1).strip(), m.group(2).strip()
        result["format"] = "full"
        return result

    # 2) از X تا Y but no مسیر/direction clause at all (destination still real)
    m = re.search(r"از\s+(.+?)\s+تا\s+(.+?)(?:\s+در\s+منطقه|\s*$)", s)
    if m:
        result["origin"], result["dest"] = m.group(1).strip(), m.group(2).strip()
        result["format"] = "no_masir"
        return result

    # 3) از X مسیر ... but no تا at all — single reference point, no real
    #    destination exists in the source data (e.g. feeder/ramp cameras).
    m = re.search(r"از\s+(.+?)\s+مسیر", s)
    if m:
        result["origin"] = m.group(1).strip()
        result["format"] = "single_point"
        return result

    # 3b) از X, no تا and no مسیر either — still a real reference point
    m = re.search(r"از\s+(.+?)(?:\s+در\s+منطقه\s*\d+\s*$|\s*$)", s)
    if m and result["seg_name"]:
        result["origin"] = m.group(1).strip()
        result["format"] = "origin_only"
        return result

    # 4) "X تا Y" with no "از" at all
    m = re.search(r"^(.+?)\s+تا\s+(.+?)(?:\s+در\s+منطقه|\s*$)", s)
    if m and not result["seg_name"]:
        result["seg_name"] = m.group(1).strip()
        result["dest"] = m.group(2).strip()
        result["format"] = "ta_only"
        return result

    # 5) Intersection / traffic-light style records:
    #      "[<code> -] A - B [چراغ ... تقاطع|میدان ...] [در منطقه N]"
    #    These describe an intersection of two named streets/places, not a
    #    single segment with an origin+destination — parsed into two parts.
    core = re.sub(r"\s*در\s+منطقه\s*\d+\s*$", "", s)
    core = re.sub(r"\s*چراغ.*$", "", core)          # strip "چراغ هوشمند/چشمک‌زن ... در محل تقاطع/میدان"
    core = re.sub(r"^\d+\s*-\s*", "", core)          # strip a leading numeric code like "4098 - "
    parts = [p.strip(" ()") for p in re.split(r"\s*-\s*", core) if p.strip(" ()")]
    if len(parts) >= 2:
        result["origin"], result["dest"] = parts[0], parts[-1]
        result["format"] = "intersection"
        return result

    if result["seg_name"]:
        result["format"] = "name_only"
    return result


def score_candidate(cam, jht, shp_name, shp_oldname,
                    shp_orig, shp_dest, shp_dir):
    """Weighted component score between a parsed camera record and a
    candidate shapefile segment.

    Two robustness fixes vs. the original version:
      1. Components the camera record genuinely doesn't have (e.g. no
         destination for a single-point/ramp record, no dedicated segment
         name for an intersection record) are excluded from the weighted
         sum and the remaining weights are renormalised — instead of being
         silently scored as 0 and dragging the total down for no reason.
      2. Fuzzy comparisons are only run when BOTH sides are non-empty.
         rapidfuzz defines ratio('', '') == 100, which previously produced
         a false "perfect" name score whenever the camera name failed to
         parse (empty) and the candidate segment's OldName was also blank.
    """
    def safe_fuzz(fn, a, b):
        return fn(a, b) if a and b else 0

    # For intersection-style records there is no dedicated seg_name — let
    # the two parsed parts (origin/dest) compete for the "name" slot too,
    # since one of them may actually be the segment's own street name.
    name_probe_list = [cam["seg_name"]] if cam["seg_name"] else \
                       [t for t in (cam["origin"], cam["dest"]) if t]

    name_s = 0
    for t in name_probe_list:
        name_s = max(
            name_s,
            safe_fuzz(fuzz.token_sort_ratio, t, shp_name),
            safe_fuzz(fuzz.token_sort_ratio, t, shp_oldname),
            safe_fuzz(fuzz.partial_ratio, t, shp_name),   # catches e.g. سلیمانی inside شهید قاسم سلیمانی
            safe_fuzz(fuzz.partial_ratio, t, shp_oldname),
        )

    has_orig, has_dest = bool(cam["origin"]), bool(cam["dest"])
    orig_s = dest_s = None
    if has_orig and has_dest:
        o_norm = od_score(cam["origin"], shp_orig)
        d_norm = od_score(cam["dest"],   shp_dest)
        o_swap = od_score(cam["origin"], shp_dest)
        d_swap = od_score(cam["dest"],   shp_orig)
        if (o_norm + d_norm) >= (o_swap + d_swap):
            orig_s, dest_s = o_norm, d_norm
        else:
            orig_s, dest_s = o_swap, d_swap
    elif has_orig:
        # single-point record: the one known point could match either end
        orig_s = max(od_score(cam["origin"], shp_orig), od_score(cam["origin"], shp_dest))
    elif has_dest:
        dest_s = max(od_score(cam["dest"], shp_orig), od_score(cam["dest"], shp_dest))

    d_s = dir_score(jht, shp_dir)

    weighted, weight_sum = 0.0, 0.0
    if name_probe_list:
        weighted += W_NAME * name_s; weight_sum += W_NAME
    if orig_s is not None:
        weighted += W_ORIG * orig_s; weight_sum += W_ORIG
    if dest_s is not None:
        weighted += W_DEST * dest_s; weight_sum += W_DEST
    weighted += W_DIR * d_s; weight_sum += W_DIR

    base = weighted / weight_sum if weight_sum else 0
    total = base * (DIR_PENALTY_FACTOR if d_s < DIR_PENALTY_THRESHOLD else 1.0)
    return total, name_s, (orig_s or 0), (dest_s or 0), d_s


# ── LOAD DATA ─────────────────────────────────────────────────────────────────
print("Loading shapefile …")
os.environ["SHAPE_RESTORE_SHX"] = "YES"
gdf = gpd.read_file(SHP_PATH, encoding="utf-8")
print(f"  {len(gdf):,} road segments  |  CRS: {gdf.crs}")
print(f"  Districts in shapefile: {sorted(gdf['District'].dropna().unique().tolist())}")
print(f"  Columns: {list(gdf.columns)}")

print("Loading Excel …")
excel = pd.read_excel(EXCEL_PATH)
print(f"  {len(excel):,} camera records\n")

for col in ["نام مکان", "جهت"]:
    if col not in excel.columns:
        raise ValueError(f"Required column '{col}' not found in Excel file.")

# ── PRE-COMPUTE CLEANED SHAPEFILE COLUMNS ─────────────────────────────────────
gdf["_mix_with_dist"] = gdf["MixValue"].apply(norm_with_district)
gdf["_mix_no_dist"]   = gdf["MixValue"].apply(norm_no_district)
gdf["_name"]          = gdf["Name"].apply(clean)
gdf["_oldname"]       = gdf["OldName"].apply(clean)
gdf["_orig"]          = gdf["Origin"].apply(clean)
gdf["_dest"]          = gdf["Destinatio"].apply(clean)
gdf["_dir"]           = gdf["NetDirecti"].apply(clean)
gdf["_district"]      = gdf["District"].apply(lambda x: int(x) if pd.notna(x) else 0)

print("Pre-computing shapefile lookup columns ... done")

# Stage 1A: exact match with district
mix_with_dist_idx = {}
for idx, val in gdf["_mix_with_dist"].items():
    if val:
        mix_with_dist_idx[val] = idx

# Stage 1B: cross-district exact match — key → list of (idx, district)
mix_no_dist_idx = {}
for idx, row in gdf.iterrows():
    key = row["_mix_no_dist"]
    if key:
        mix_no_dist_idx.setdefault(key, []).append((idx, row["_district"]))

print(f"Stage 1A lookup built: {len(mix_with_dist_idx):,} entries")
print(f"Stage 1B lookup built: {len(mix_no_dist_idx):,} entries")

SEG_COLS = [
    "OBJECTID", "TTCCID", "OldTTCCID",
    "Name", "LatinName", "OldName", "Summerized",
    "Origin", "Destinatio", "NetDirecti",
    "MixValue",
    "Type", "Class", "District",
    "XMidUTM", "YMidUTM",
    "XStartUTM", "YStartUTM",
    "XEndUTM", "YEndUTM",
    "SHAPE_STLe",
]
_INTERNAL = ["_mix_with_dist", "_mix_no_dist", "_name", "_oldname",
             "_orig", "_dest", "_dir", "_district"]


def attach_segment(result_row, seg_idx):
    seg = gdf.loc[seg_idx]
    for c in SEG_COLS:
        result_row[f"seg_{c}"] = seg.get(c, None)
    result_row["seg_geometry_wkt"] = seg.geometry.wkt
    return result_row


def make_direct_result(result, seg_idx, cam, stage_label):
    result = attach_segment(result, seg_idx)
    result["match_method"]          = f"DirectMatch ({stage_label})"
    result["match_score"]           = 100
    result["match_score_name"]      = 100
    result["match_score_origin"]    = 100
    result["match_score_dest"]      = 100
    result["match_score_direction"] = 100
    result["parsed_seg_name"]       = cam.get("seg_name", "")
    result["parsed_origin"]         = cam.get("origin", "")
    result["parsed_dest"]           = cam.get("dest", "")
    return result


# ── MATCHING LOOP ─────────────────────────────────────────────────────────────
direct_rows    = []
fuzzy_rows     = []
unmatched_rows = []

import time
_total_rows   = len(excel)
_progress_every = 500
_t0 = time.time()

print(f"\nStarting matching loop ({_total_rows:,} rows) ...")

for _row_num, (_, cam_row) in enumerate(excel.iterrows(), start=1):

    # ── Progress report every N rows ─────────────────────────────────────────
    if _row_num % _progress_every == 0 or _row_num == _total_rows:
        elapsed   = time.time() - _t0
        rate      = _row_num / elapsed if elapsed > 0 else 0
        remaining = (_total_rows - _row_num) / rate if rate > 0 else 0
        print(
            f"  [{_row_num:>{len(str(_total_rows))}}/{_total_rows}] "
            f"Direct={len(direct_rows)}  Fuzzy={len(fuzzy_rows)}  Unmatched={len(unmatched_rows)}  "
            f"| {elapsed:.0f}s elapsed  ~{remaining:.0f}s remaining"
        )

    raw_addr  = str(cam_row.get("نام مکان", ""))
    jht       = str(cam_row.get("جهت", ""))
    result    = cam_row.to_dict()
    cam       = parse_camera_address(raw_addr)
    cam_dist  = cam_district(raw_addr)

    # ── Stage 1A: DirectMatch — exact with district ───────────────────────────
    key_with = norm_with_district(raw_addr)
    if key_with and key_with in mix_with_dist_idx:
        direct_rows.append(make_direct_result(result, mix_with_dist_idx[key_with], cam, "exact"))
        continue

    # ── Stage 1B: DirectMatch — cross-district exact ──────────────────────────
    key_no = norm_no_district(raw_addr)
    if key_no and key_no in mix_no_dist_idx:
        candidates = mix_no_dist_idx[key_no]
        # Prefer the row whose district matches the camera address district
        same_dist = [idx for idx, d in candidates if d == cam_dist]
        seg_idx = same_dist[0] if same_dist else candidates[0][0]
        direct_rows.append(make_direct_result(result, seg_idx, cam, "cross-district"))
        continue

    # ── Stage 1C: DirectMatch — via OldName ──────────────────────────────────
    matched_oldname = False
    if cam["seg_name"]:
        for idx, row in gdf[gdf["_oldname"] != ""].iterrows():
            if fuzz.token_sort_ratio(cam["seg_name"], row["_oldname"]) >= 90:
                reconstructed = norm_no_district(raw_addr).replace(
                    cam["seg_name"], row["_name"], 1
                )
                if reconstructed == row["_mix_no_dist"]:
                    direct_rows.append(make_direct_result(result, idx, cam, "OldName"))
                    matched_oldname = True
                    break
    if matched_oldname:
        continue

    # ── Stage 2: ComponentMatch ───────────────────────────────────────────────
    result["parsed_seg_name"] = cam["seg_name"]
    result["parsed_origin"]   = cam["origin"]
    result["parsed_dest"]     = cam["dest"]

    # Candidate filter: Name or OldName similarity — district is NOT a filter.
    # Uses the same probe list as score_candidate: for intersection-style
    # records with no dedicated seg_name, try origin/dest as candidate names
    # too, so a real match isn't filtered out just because seg_name is empty.
    name_probe_list = [cam["seg_name"]] if cam["seg_name"] else \
                       [t for t in (cam["origin"], cam["dest"]) if t]
    if name_probe_list:
        name_scores = gdf.apply(
            lambda r: name_similarity(name_probe_list, r["_name"], r["_oldname"]),
            axis=1
        )
        candidates = gdf[name_scores >= NAME_FILTER_THRESH]
    else:
        candidates = gdf

    best_total, best_idx = -1, None
    best_components = (0, 0, 0, 0)

    if not candidates.empty:
        for idx, row in candidates.iterrows():
            total, ns, os_, ds, dirs = score_candidate(
                cam, jht,
                row["_name"], row["_oldname"],
                row["_orig"], row["_dest"],
                row["_dir"]
            )
            if total > best_total:
                best_total      = total
                best_idx        = idx
                best_components = (ns, os_, ds, dirs)

    if best_idx is not None and best_total >= COMPONENT_THRESHOLD:
        result = attach_segment(result, best_idx)
        result["match_method"]          = "FuzzyMatch"
        result["match_score"]           = round(best_total, 2)
        result["match_score_name"]      = best_components[0]
        result["match_score_origin"]    = best_components[1]
        result["match_score_dest"]      = best_components[2]
        result["match_score_direction"] = best_components[3]
        fuzzy_rows.append(result)
        continue

    # ── Unmatched ─────────────────────────────────────────────────────────────
    # No confirmed match. Attach the best candidate found (same columns as
    # matched rows) so the analyst can review it. match_method stays "Unmatched".
    if best_idx is not None:
        result = attach_segment(result, best_idx)
    else:
        for c in SEG_COLS:
            result[f"seg_{c}"] = None
        result["seg_geometry_wkt"] = None
    result["match_method"]          = "Unmatched"
    result["match_score"]           = round(best_total, 2) if best_idx is not None else None
    result["match_score_name"]      = best_components[0]
    result["match_score_origin"]    = best_components[1]
    result["match_score_dest"]      = best_components[2]
    result["match_score_direction"] = best_components[3]
    unmatched_rows.append(result)

# ── SUMMARY ───────────────────────────────────────────────────────────────────
_total_elapsed = time.time() - _t0
print(f"\nMatching loop finished in {_total_elapsed:.1f}s ({_total_elapsed/60:.1f} min)")

print(f"{'─'*60}")
print(f"  Total records       : {len(excel)}")
print(f"  DirectMatch         : {len(direct_rows)}")
print(f"  FuzzyMatch          : {len(fuzzy_rows)}")
print(f"  Unmatched           : {len(unmatched_rows)}")
print(f"{'─'*60}\n")

# ── SAVE EXCEL OUTPUTS ────────────────────────────────────────────────────────
DROP = ["seg_geometry_wkt"] + _INTERNAL

def save_excel(rows, path):
    if rows:
        df = pd.DataFrame(rows).drop(columns=DROP, errors="ignore")
        df.to_excel(path, index=False)
        print(f"Saved  →  {path}  ({len(rows)} rows)")
    else:
        print(f"Skipped (empty)  →  {path}")

print("Saving Excel outputs ...")
try:
    save_excel(direct_rows,    OUT_DIRECT)
except Exception as e:
    print(f"ERROR saving DirectMatch.xlsx: {e}")
try:
    save_excel(fuzzy_rows,     OUT_FUZZY)
except Exception as e:
    print(f"ERROR saving FuzzyMatch.xlsx: {e}")
try:
    save_excel(unmatched_rows, OUT_UNMATCHED)
except Exception as e:
    print(f"ERROR saving Unmatched.xlsx: {e}")

# ── SAVE GEOPACKAGE ───────────────────────────────────────────────────────────
print("Saving GeoPackage ...")
try:
    from shapely import wkt as shapely_wkt

    all_matched = direct_rows + fuzzy_rows
    if all_matched:
        df_all = pd.DataFrame(all_matched).drop(columns=_INTERNAL, errors="ignore")
        df_all["geometry"] = df_all["seg_geometry_wkt"].apply(
            lambda w: shapely_wkt.loads(w) if pd.notna(w) else None
        )
        gdf_out = gpd.GeoDataFrame(
            df_all.drop(columns=["seg_geometry_wkt"], errors="ignore"),
            geometry="geometry",
            crs=gdf.crs
        )
        gdf_out.to_file(OUT_GPKG, driver="GPKG", layer="camera_segments")
        print(f"Saved  →  {OUT_GPKG}  (CRS: {gdf.crs})")
except Exception as e:
    print(f"ERROR saving GeoPackage: {e}")

print("\nDone.")
