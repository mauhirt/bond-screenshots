#!/usr/bin/env python3
"""
Spatial join of NBI bridge data and CWNS facility data to municipality boundaries.

Uses Census TIGER/Line 2023 Place shapefiles to define city boundaries,
then assigns NBI bridges and CWNS wastewater facilities to cities based on
point-in-polygon spatial joins.

Input:
  - nbi_data/nbi_all_years.parquet  (NBI bridge inventory)
  - cwns_data/cwns_national_csv.zip (CWNS facility data)
  - geodata/tiger_place/*.zip       (Census TIGER/Line 2023 Place shapefiles)
  - green bonds excel.xlsx          (bond issuer names → city list)

Output:
  - geodata/us_places_2023.gpkg           (combined national Census Places)
  - geodata/bond_city_boundaries.gpkg     (filtered to bond-issuer cities)
  - geodata/nbi_bridges_by_city_2025.csv  (bridges within each city)
  - geodata/cwns_facilities_by_city.csv   (CWNS facilities within each city)
  - geodata/cwns_city_summary.csv         (per-city CWNS summary)
"""

import glob
import io
import os
import zipfile

import geopandas as gpd
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEODATA_DIR = os.path.join(BASE_DIR, "geodata")
TIGER_DIR = os.path.join(GEODATA_DIR, "tiger_place")

# State FIPS mapping
STATE_FIPS = {
    'AL': '01', 'AK': '02', 'AZ': '04', 'AR': '05', 'CA': '06', 'CO': '08',
    'CT': '09', 'DE': '10', 'DC': '11', 'FL': '12', 'GA': '13', 'HI': '15',
    'ID': '16', 'IL': '17', 'IN': '18', 'IA': '19', 'KS': '20', 'KY': '21',
    'LA': '22', 'ME': '23', 'MD': '24', 'MA': '25', 'MI': '26', 'MN': '27',
    'MS': '28', 'MO': '29', 'MT': '30', 'NE': '31', 'NV': '32', 'NH': '33',
    'NJ': '34', 'NM': '35', 'NY': '36', 'NC': '37', 'ND': '38', 'OH': '39',
    'OK': '40', 'OR': '41', 'PA': '42', 'RI': '44', 'SC': '45', 'SD': '46',
    'TN': '47', 'TX': '48', 'UT': '49', 'VT': '50', 'VA': '51', 'WA': '53',
    'WV': '54', 'WI': '55', 'WY': '56',
}

# Bond-issuer cities mapped to Census Place names and state FIPS
# (bond_city_name, state_fips, census_place_name)
CITY_MATCHES = [
    ("Albuquerque", "35", "Albuquerque"),
    ("Arvada", "08", "Arvada"),
    ("Atlanta", "13", "Atlanta"),
    ("Bayonne", "34", "Bayonne"),
    ("Cranford", "34", "Cranford"),
    ("Danville", "21", "Danville"),
    ("Dardanelle", "05", "Dardanelle"),
    ("East Rockaway", "36", "East Rockaway"),
    ("Economy", "42", "Economy"),
    ("Elmwood Park", "34", "Elmwood Park"),
    ("Freeland", "26", "Freeland"),
    ("Girard", "20", "Girard"),
    ("Hartford", "09", "Hartford"),
    ("Honolulu", "15", "Urban Honolulu"),
    ("Hot Springs", "05", "Hot Springs"),
    ("Jersey City", "34", "Jersey City"),
    ("Joliet", "17", "Joliet"),
    ("Le Mars", "19", "Le Mars"),
    ("Los Angeles", "06", "Los Angeles"),
    ("Middleton", "55", "Middleton"),
    ("Milwaukee", "55", "Milwaukee"),
    ("Mission", "48", "Mission"),
    ("Portland", "23", "Portland"),
    ("Sacramento", "06", "Sacramento"),
    ("San Francisco", "06", "San Francisco"),
    ("Sheridan", "08", "Sheridan"),
    ("South Lake Tahoe", "06", "South Lake Tahoe"),
    ("St. Marys", "20", "St. Marys"),
    ("St. Paul", "27", "St. Paul"),
    ("Stockton", "06", "Stockton"),
    ("Strongsville", "39", "Strongsville"),
    ("Tacoma", "53", "Tacoma"),
    ("Vineland", "34", "Vineland"),
    ("Wamego", "20", "Wamego"),
    ("Washington", "11", "Washington"),
]


def load_or_build_places_gpkg():
    """Load combined national places GeoPackage, building it if needed."""
    gpkg_path = os.path.join(GEODATA_DIR, "us_places_2023.gpkg")
    if os.path.exists(gpkg_path):
        print(f"Loading cached {gpkg_path}...", flush=True)
        return gpd.read_file(gpkg_path)

    print("Combining TIGER place shapefiles...", flush=True)
    zips = sorted(glob.glob(os.path.join(TIGER_DIR, "*.zip")))
    frames = [gpd.read_file(zf) for zf in zips]
    places = gpd.GeoDataFrame(pd.concat(frames, ignore_index=True), geometry='geometry')
    places.to_file(gpkg_path, driver="GPKG")
    print(f"Saved {gpkg_path} ({len(places):,} places)")
    return places


def build_city_boundaries(places):
    """Filter Census Places to bond-issuer cities."""
    gpkg_path = os.path.join(GEODATA_DIR, "bond_city_boundaries.gpkg")
    if os.path.exists(gpkg_path):
        print(f"Loading cached {gpkg_path}...", flush=True)
        return gpd.read_file(gpkg_path)

    matched = []
    for bond_name, state_fips, census_name in CITY_MATCHES:
        row = places[(places['STATEFP'] == state_fips) & (places['NAME'] == census_name)]
        if len(row):
            r = row.iloc[0].copy()
            r['bond_city_name'] = bond_name
            matched.append(r)
        else:
            print(f"  WARNING: {bond_name} ({state_fips}) not found in Census Places")

    city_bounds = gpd.GeoDataFrame(matched, crs=places.crs)
    city_bounds.to_file(gpkg_path, driver="GPKG")
    print(f"Saved {gpkg_path} ({len(city_bounds)} cities)")
    return city_bounds


def nbi_dms_to_decimal(dms_str, is_lon=False):
    """Convert NBI DMS coordinate (DDMMSSSS) to decimal degrees."""
    try:
        val = float(dms_str)
        if val == 0:
            return None
        deg = int(val / 1000000)
        remainder = val - deg * 1000000
        minutes = int(remainder / 10000)
        seconds = (remainder - minutes * 10000) / 100
        decimal = deg + minutes / 60 + seconds / 3600
        return -decimal if is_lon else decimal
    except (ValueError, TypeError):
        return None


def spatial_join_nbi(cities):
    """Spatial join NBI 2025 bridges to city boundaries."""
    print("\n=== NBI Spatial Join ===", flush=True)
    nbi_cols = [
        'year', 'STATE_CODE_001', 'STRUCTURE_NUMBER_008', 'LAT_016', 'LONG_017',
        'COUNTY_CODE_003', 'PLACE_CODE_004', 'FEATURES_DESC_006A',
        'FACILITY_CARRIED_007', 'YEAR_BUILT_027', 'ADT_029',
        'DECK_COND_058', 'SUPERSTRUCTURE_COND_059', 'SUBSTRUCTURE_COND_060',
        'STRUCTURAL_EVAL_067', 'SUFFICIENCY_RATING',
    ]
    nbi = pd.read_parquet(os.path.join(BASE_DIR, "nbi_data", "nbi_all_years.parquet"), columns=nbi_cols)
    nbi_latest = nbi[nbi['year'] == '2025'].copy()
    print(f"NBI 2025: {len(nbi_latest):,} bridges", flush=True)

    nbi_latest['lat'] = nbi_latest['LAT_016'].apply(lambda x: nbi_dms_to_decimal(x, False))
    nbi_latest['lon'] = nbi_latest['LONG_017'].apply(lambda x: nbi_dms_to_decimal(x, True))
    valid = nbi_latest.dropna(subset=['lat', 'lon'])
    valid = valid[(valid['lat'] > 17) & (valid['lat'] < 72) & (valid['lon'] > -180) & (valid['lon'] < -60)]

    nbi_gdf = gpd.GeoDataFrame(valid, geometry=gpd.points_from_xy(valid['lon'], valid['lat']), crs="EPSG:4269")
    joined = gpd.sjoin(nbi_gdf, cities[['bond_city_name', 'GEOID', 'NAMELSAD', 'geometry']],
                       how='inner', predicate='within')
    print(f"Bridges matched: {len(joined):,}", flush=True)

    out = os.path.join(GEODATA_DIR, "nbi_bridges_by_city_2025.csv")
    joined.drop(columns=['geometry']).to_csv(out, index=False)
    print(f"Saved: {out}")
    return joined


def spatial_join_cwns(cities):
    """Spatial join CWNS facilities to city boundaries."""
    print("\n=== CWNS Spatial Join ===", flush=True)
    zf = zipfile.ZipFile(os.path.join(BASE_DIR, "cwns_data", "cwns_national_csv.zip"))
    phys_loc = pd.read_csv(io.BytesIO(zf.read("PHYSICAL_LOCATION.csv")))
    facilities = pd.read_csv(io.BytesIO(zf.read("FACILITIES.csv")))
    needs = pd.read_csv(io.BytesIO(zf.read("NEEDS_COST_BY_CATEGORY.csv")))

    points = phys_loc[phys_loc['LOCATION_TYPE'] == 'Point'].copy()
    valid = points[(points['LATITUDE'] > 17) & (points['LATITUDE'] < 72) &
                   (points['LONGITUDE'] > -180) & (points['LONGITUDE'] < -60)]
    print(f"CWNS point locations: {len(valid):,}", flush=True)

    cwns_gdf = gpd.GeoDataFrame(valid, geometry=gpd.points_from_xy(valid['LONGITUDE'], valid['LATITUDE']), crs="EPSG:4269")
    joined = gpd.sjoin(cwns_gdf, cities[['bond_city_name', 'GEOID', 'NAMELSAD', 'geometry']],
                       how='inner', predicate='within')
    print(f"Facilities matched: {len(joined):,}", flush=True)

    merged = joined.merge(
        facilities[['CWNS_ID', 'FACILITY_ID', 'FACILITY_NAME', 'INFRASTRUCTURE_TYPE', 'OWNER_TYPE', 'NO_NEEDS']],
        on=['CWNS_ID', 'FACILITY_ID'], how='left'
    )
    needs_agg = needs.groupby(['CWNS_ID', 'FACILITY_ID']).agg(
        total_needs_cost=('OFFICIAL_AMOUNT', 'sum'),
        needs_count=('NEEDS_CATEGORY', 'count'),
        needs_categories=('NEEDS_CATEGORY', lambda x: '; '.join(sorted(x.unique())))
    ).reset_index()
    merged = merged.merge(needs_agg, on=['CWNS_ID', 'FACILITY_ID'], how='left')

    out = os.path.join(GEODATA_DIR, "cwns_facilities_by_city.csv")
    merged.drop(columns=['geometry']).to_csv(out, index=False)
    print(f"Saved: {out}")

    summary = merged.groupby('bond_city_name').agg(
        facility_count=('FACILITY_ID', 'nunique'),
        total_needs_cost=('total_needs_cost', 'sum'),
        infra_types=('INFRASTRUCTURE_TYPE', lambda x: ', '.join(sorted(x.dropna().unique()))),
    ).sort_values('total_needs_cost', ascending=False)
    summary_path = os.path.join(GEODATA_DIR, "cwns_city_summary.csv")
    summary.to_csv(summary_path)
    print(f"Saved: {summary_path}")
    return merged


def main():
    os.makedirs(GEODATA_DIR, exist_ok=True)

    places = load_or_build_places_gpkg()
    print(f"Census Places: {len(places):,}", flush=True)

    cities = build_city_boundaries(places)
    cities = cities.to_crs("EPSG:4269")
    print(f"Bond-issuer cities: {len(cities)}", flush=True)

    nbi_joined = spatial_join_nbi(cities)
    cwns_joined = spatial_join_cwns(cities)

    print("\n=== Summary ===", flush=True)
    print(f"NBI bridges assigned to cities: {len(nbi_joined):,}")
    print(f"CWNS facilities assigned to cities: {len(cwns_joined):,}")
    print(f"Cities with NBI data: {nbi_joined['bond_city_name'].nunique()}")
    print(f"Cities with CWNS data: {cwns_joined['bond_city_name'].nunique()}")


if __name__ == "__main__":
    main()
