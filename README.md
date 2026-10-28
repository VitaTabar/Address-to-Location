# Address to Location

Matches traffic camera address records (Excel, Persian) to road segments in a shapefile, using exact matching first and weighted fuzzy matching as a fallback. Outputs Excel reports and a GeoPackage for QGIS/ArcGIS.

## How it works

Camera addresses follow this pattern:

```
[segment name] از [origin] تا [destination] مسیر [direction] در منطقه [district]
```

| Stage | Method | Description |
|---|---|---|
| 1A | Direct | Exact match on `MixValue`, same district |
| 1B | Direct | Exact match with the district suffix removed (cross-district); prefers the matching district when several rows match |
| 1C | Direct | Match via `OldName`, substituting the current `Name` |
| 2 | Fuzzy | Candidates filtered by name similarity (`NAME_FILTER_THRESH`), then weighted scoring of name, origin, destination and direction |

Scoring notes:

- Origin/destination: noise words are stripped, and origin↔destination swaps are tried.
- Direction: order-sensitive, so opposite directions score low; a wrong direction is penalized (`DIR_PENALTY_FACTOR`).
- Components missing from a camera record (e.g. no destination) are excluded and the remaining weights renormalized.
- Several address shapes are parsed (full, no `مسیر`, single point, intersection, etc.), and the detected shape is recorded.

## Requirements

- Python 3.9+
- Input files:
  - **Excel** with columns `نام مکان` (address) and `جهت` (direction)
  - **Shapefile** (with `.prj`) containing at least `Name`, `OldName`, `Origin`, `Destinatio`, `NetDirecti`, `MixValue`, `District`. Other segment attributes listed in `SEG_COLS` are copied when present.

## Installation

```bash
git clone https://github.com/VitaTabar/address-to-location.git
cd address-to-location
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

Place your inputs in `data/` (default names below) or set environment variables:

| Variable | Default | Description |
|---|---|---|
| `SHP_PATH` | `data/street_segments.shp` | Road segments shapefile |
| `EXCEL_PATH` | `data/cameras.xlsx` | Camera records |
| `CAMERA_MATCH_OUTPUT` | `output` | Output folder |

```bash
python address_to_location.py
```

Thresholds and weights (`COMPONENT_THRESHOLD`, `NAME_FILTER_THRESH`, `W_*`, `DIR_PENALTY_*`) are at the top of the script.

## Output

| File | Content |
|---|---|
| `DirectMatch.xlsx` | Exact matches (Stages 1A–1C) |
| `FuzzyMatch.xlsx` | Matches above `COMPONENT_THRESHOLD` |
| `Unmatched.xlsx` | Below threshold, with the best candidate for manual review |
| `all_matched.gpkg` | Direct + fuzzy matches as a spatial layer |

## Known limitations

- Stage 1C and Stage 2 scan the full shapefile for each unmatched row, so large datasets can be slow.
- The shapefile and Excel schemas are specific to the Tehran TTCC data; other datasets require adjusting column names.

## License

[MIT](LICENSE)
