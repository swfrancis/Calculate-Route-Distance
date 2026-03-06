"""
Route Distance Calculator using OSRM (free, no API key needed).
Run this locally via Claude Code or your terminal.

Prerequisites:
    pip install pandas openpyxl geopy requests

Usage:
    python calculate_route_distances.py

Input files (must be in same directory):
    - multi-drops.xlsx          (original spreadsheet)
    - address_pairs.csv         (generated address pairs)

Output:
    - multi-drops-with-distances.xlsx  (updated spreadsheet with column J filled)
"""

import time
import json
import hashlib
import pandas as pd
import requests
from pathlib import Path
from geopy.geocoders import Nominatim
from openpyxl import load_workbook

# --- Config ---
OSRM_URL = "https://router.project-osrm.org/route/v1/driving"
GEOCODE_CACHE_FILE = "geocode_cache.json"
DISTANCE_CACHE_FILE = "distance_cache.json"
INPUT_XLSX = "multi-drops.xlsx"
PAIRS_CSV = "address_pairs.csv"
OUTPUT_XLSX = "multi-drops-with-distances.xlsx"

# --- Cache helpers ---
def load_cache(path):
    if Path(path).exists():
        with open(path) as f:
            return json.load(f)
    return {}

def save_cache(data, path):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)

# --- Geocoding ---
def geocode_address(address, geolocator, cache):
    cache_key = address.strip().upper()
    if cache_key in cache:
        return cache[cache_key]

    # Clean address for geocoding: remove LOT prefixes, add Australia
    clean = address
    if clean.upper().startswith("LOT "):
        # Extract the part in parentheses if present, e.g. "LOT 732 (42) PARKLANDS ROAD" -> "42 PARKLANDS ROAD"
        import re
        m = re.match(r'LOT\s+\d+\s*\((\d+)\)\s*(.*)', clean, re.IGNORECASE)
        if m:
            clean = f"{m.group(1)} {m.group(2)}"
        else:
            # Just strip "LOT <number>" prefix
            clean = re.sub(r'^LOT\s+\d+\s*', '', clean, flags=re.IGNORECASE)

    query = f"{clean}, Australia"
    try:
        location = geolocator.geocode(query, timeout=10)
        if location:
            result = (location.latitude, location.longitude)
            cache[cache_key] = result
            return result

        # Fallback: try just city + state + postcode
        parts = address.split(",")
        if len(parts) >= 2:
            fallback = parts[-1].strip() + ", Australia"
            location = geolocator.geocode(fallback, timeout=10)
            if location:
                result = (location.latitude, location.longitude)
                cache[cache_key] = result
                return result
    except Exception as e:
        print(f"  Geocode error for '{address}': {e}")

    print(f"  WARNING: Could not geocode '{address}'")
    cache[cache_key] = None
    return None

# --- OSRM driving distance ---
def get_driving_distance_km(coord1, coord2, dist_cache):
    """Get driving distance in km between two (lat, lon) coordinates using OSRM."""
    cache_key = f"{coord1[0]:.6f},{coord1[1]:.6f}|{coord2[0]:.6f},{coord2[1]:.6f}"
    if cache_key in dist_cache:
        return dist_cache[cache_key]

    # OSRM expects lon,lat order
    url = f"{OSRM_URL}/{coord1[1]},{coord1[0]};{coord2[1]},{coord2[0]}"
    params = {"overview": "false", "geometries": "polyline"}

    try:
        resp = requests.get(url, params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") == "Ok" and data.get("routes"):
            dist_km = data["routes"][0]["distance"] / 1000.0
            dist_cache[cache_key] = round(dist_km, 1)
            return dist_cache[cache_key]
        else:
            print(f"  OSRM no route: {data.get('code')}")
    except Exception as e:
        print(f"  OSRM error: {e}")

    dist_cache[cache_key] = None
    return None

# --- Main ---
def main():
    print("=" * 60)
    print("Route Distance Calculator (OSRM)")
    print("=" * 60)

    # Load data
    pairs_df = pd.read_csv(PAIRS_CSV)
    print(f"\nLoaded {len(pairs_df)} address pairs across {pairs_df['load_number'].nunique()} loads")

    # Collect all unique addresses
    all_addresses = set(pairs_df['from_address'].unique()) | set(pairs_df['to_address'].unique())
    print(f"Unique addresses to geocode: {len(all_addresses)}")

    # Step 1: Geocode all addresses
    print("\n--- Step 1: Geocoding addresses ---")
    geo_cache = load_cache(GEOCODE_CACHE_FILE)
    geolocator = Nominatim(user_agent="route_distance_calculator_v1")
    
    cached = sum(1 for a in all_addresses if a.strip().upper() in geo_cache)
    print(f"Already cached: {cached}/{len(all_addresses)}")

    for i, addr in enumerate(sorted(all_addresses)):
        if addr.strip().upper() in geo_cache:
            continue
        print(f"  [{i+1}/{len(all_addresses)}] Geocoding: {addr[:60]}...")
        geocode_address(addr, geolocator, geo_cache)
        time.sleep(1.1)  # Nominatim rate limit: 1 req/sec
        if (i + 1) % 20 == 0:
            save_cache(geo_cache, GEOCODE_CACHE_FILE)

    save_cache(geo_cache, GEOCODE_CACHE_FILE)
    
    failed_geo = [a for a in all_addresses if geo_cache.get(a.strip().upper()) is None]
    if failed_geo:
        print(f"\nWARNING: {len(failed_geo)} addresses could not be geocoded:")
        for a in failed_geo[:10]:
            print(f"  - {a}")

    # Step 2: Get driving distances for each pair
    print("\n--- Step 2: Fetching driving distances via OSRM ---")
    dist_cache = load_cache(DISTANCE_CACHE_FILE)
    
    results = []
    for idx, row in pairs_df.iterrows():
        coord1 = geo_cache.get(row['from_address'].strip().upper())
        coord2 = geo_cache.get(row['to_address'].strip().upper())

        if coord1 is None or coord2 is None:
            print(f"  Load {row['load_number']} stop {row['from_stop']}->{row['to_stop']}: SKIP (geocode failed)")
            results.append(None)
            continue

        dist = get_driving_distance_km(coord1, coord2, dist_cache)
        if dist is not None:
            print(f"  Load {row['load_number']} stop {row['from_stop']}->{row['to_stop']}: {dist} km")
        else:
            print(f"  Load {row['load_number']} stop {row['from_stop']}->{row['to_stop']}: FAILED")
        results.append(dist)
        time.sleep(0.2)  # Be polite to OSRM

        if (idx + 1) % 50 == 0:
            save_cache(dist_cache, DISTANCE_CACHE_FILE)

    save_cache(dist_cache, DISTANCE_CACHE_FILE)
    pairs_df['leg_distance_km'] = results

    # Step 3: Write back to spreadsheet
    print("\n--- Step 3: Writing results to spreadsheet ---")
    import shutil
    shutil.copy(INPUT_XLSX, OUTPUT_XLSX)
    wb = load_workbook(OUTPUT_XLSX)
    ws = wb.active

    # Build a lookup: (load_number, from_stop, to_stop) -> leg_distance
    leg_lookup = {}
    for _, row in pairs_df.iterrows():
        key = (int(row['load_number']), str(row['from_stop']), str(row['to_stop']))
        leg_lookup[key] = row['leg_distance_km']

    # Walk through rows, calculate cumulative route distance per load
    current_load = None
    cumulative = 0
    prev_stop = None
    col_j = 10  # Column J

    for row_idx in range(2, ws.max_row + 1):
        load_num = ws.cell(row=row_idx, column=1).value
        stop_order = ws.cell(row=row_idx, column=2).value

        if load_num is None:
            continue

        load_num_int = int(load_num)
        stop_str = str(stop_order)

        if stop_str == 'DEPOT':
            current_load = load_num_int
            cumulative = 0
            prev_stop = 'DEPOT'
            ws.cell(row=row_idx, column=col_j).value = 0
        else:
            key = (current_load, prev_stop, stop_str)
            leg_dist = leg_lookup.get(key)
            if leg_dist is not None:
                cumulative += leg_dist
                ws.cell(row=row_idx, column=col_j).value = round(cumulative, 1)
            else:
                ws.cell(row=row_idx, column=col_j).value = "N/A"
            prev_stop = stop_str

    wb.save(OUTPUT_XLSX)
    print(f"\nDone! Output saved to: {OUTPUT_XLSX}")

    # Summary
    success = sum(1 for r in results if r is not None)
    print(f"\nSummary: {success}/{len(results)} legs calculated successfully")
    if failed_geo:
        print(f"         {len(failed_geo)} addresses failed geocoding")

if __name__ == "__main__":
    main()
