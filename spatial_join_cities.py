#!/usr/bin/env python3
"""
Spatial join of NBI bridge data and CWNS facility data to municipality boundaries.

Uses LargeCities_places_2023.gpkg.zip (578 large US cities) as city boundaries,
then assigns NBI bridges and CWNS wastewater facilities to cities based on
point-in-polygon spatial joins.

Input:
  - LargeCities.xlsx                      (city list: fips, geo_name, state_abb)
  - LargeCities_places_2023.gpkg.zip      (city boundary polygons)
  - nbi_data/nbi_all_years.parquet        (NBI bridge inventory)
  - cwns_data/cwns_national_csv.zip       (CWNS facility data)

Output:
  - geodata/nbi_bridges_by_city_2025.csv  (bridges within each city)
  - geodata/cwns_facilities_by_city.csv   (CWNS facilities within each city)
  - geodata/cwns_city_summary.csv         (per-city CWNS summary)
"""

import io
import os
import zipfile

import geopandas as gpd
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEODATA_DIR = os.path.join(BASE_DIR, "geodata")


def load_city_boundaries():
    """Load city boundaries from LargeCities GeoPackage.

    Extracts the GeoPackage from the ZIP if needed, cross-checks against
    LargeCities.xlsx, and assigns CRS (EPSG:4269) since the source file
    has no CRS metadata.
    """
    gpkg_zip = os.path.join(BASE_DIR, "LargeCities_places_2023.gpkg.zip")
    gpkg_path = os.path.join(GEODATA_DIR, "LargeCities_places_2023.gpkg")

    # Extract GeoPackage from ZIP if needed
    if not os.path.exists(gpkg_path):
        print(f"Extracting {gpkg_zip}...", flush=True)
        with zipfile.ZipFile(gpkg_zip) as zf:
            zf.extractall(GEODATA_DIR)

    # Load city boundaries
    cities = gpd.read_file(gpkg_path)
    if cities.crs is None:
        cities = cities.set_crs("EPSG:4269")
    else:
        cities = cities.to_crs("EPSG:4269")
    print(f"City boundaries loaded: {len(cities)} cities", flush=True)

    # Cross-check against Excel city list
    excel_path = os.path.join(BASE_DIR, "LargeCities.xlsx")
    if os.path.exists(excel_path):
        city_list = pd.read_excel(excel_path)
        gpkg_fips = set(cities['fips'].astype(str))
        excel_fips = set(city_list['fips'].astype(str))
        missing = excel_fips - gpkg_fips
        if missing:
            missing_names = city_list[city_list['fips'].astype(str).isin(missing)][['geo_name', 'state_abb', 'fips']]
            print(f"  WARNING: {len(missing)} cities in Excel but not in GeoPackage:")
            for _, row in missing_names.iterrows():
                print(f"    {row['geo_name']}, {row['state_abb']} (fips={row['fips']})")
        else:
            print(f"  All {len(excel_fips)} cities from Excel found in GeoPackage")

    return cities


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
    city_cols = ['fips', 'geo_name', 'state_abb', 'GEOID', 'NAMELSAD', 'geometry']
    joined = gpd.sjoin(nbi_gdf, cities[city_cols], how='inner', predicate='within')
    print(f"Bridges matched: {len(joined):,}", flush=True)
    print(f"Cities with bridges: {joined['geo_name'].nunique()}", flush=True)

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
    flow = pd.read_csv(io.BytesIO(zf.read("FLOW.csv")))

    points = phys_loc[phys_loc['LOCATION_TYPE'] == 'Point'].copy()
    valid = points[(points['LATITUDE'] > 17) & (points['LATITUDE'] < 72) &
                   (points['LONGITUDE'] > -180) & (points['LONGITUDE'] < -60)]
    print(f"CWNS point locations: {len(valid):,}", flush=True)

    cwns_gdf = gpd.GeoDataFrame(valid, geometry=gpd.points_from_xy(valid['LONGITUDE'], valid['LATITUDE']), crs="EPSG:4269")
    city_cols = ['fips', 'geo_name', 'state_abb', 'GEOID', 'NAMELSAD', 'geometry']
    joined = gpd.sjoin(cwns_gdf, cities[city_cols], how='inner', predicate='within')
    print(f"Facilities matched: {len(joined):,}", flush=True)

    # Merge facility metadata
    merged = joined.merge(
        facilities[['CWNS_ID', 'FACILITY_ID', 'FACILITY_NAME', 'INFRASTRUCTURE_TYPE', 'OWNER_TYPE', 'NO_NEEDS']],
        on=['CWNS_ID', 'FACILITY_ID'], how='left'
    )

    # Aggregate needs costs per facility
    needs_agg = needs.groupby(['CWNS_ID', 'FACILITY_ID']).agg(
        total_needs_cost=('OFFICIAL_AMOUNT', 'sum'),
        needs_count=('NEEDS_CATEGORY', 'count'),
        needs_categories=('NEEDS_CATEGORY', lambda x: '; '.join(sorted(x.unique())))
    ).reset_index()
    merged = merged.merge(needs_agg, on=['CWNS_ID', 'FACILITY_ID'], how='left')

    # Aggregate design flow per facility
    flow_agg = flow.groupby(['CWNS_ID', 'FACILITY_ID']).agg(
        total_design_flow=('CURRENT_DESIGN_FLOW', 'sum')
    ).reset_index()
    merged = merged.merge(flow_agg, on=['CWNS_ID', 'FACILITY_ID'], how='left')

    out = os.path.join(GEODATA_DIR, "cwns_facilities_by_city.csv")
    merged.drop(columns=['geometry']).to_csv(out, index=False)
    print(f"Saved: {out}")

    # Per-city summary
    summary = merged.groupby(['geo_name', 'state_abb', 'fips', 'GEOID']).agg(
        facility_count=('FACILITY_ID', 'nunique'),
        total_needs_cost=('total_needs_cost', 'sum'),
        total_design_flow=('total_design_flow', 'sum'),
        infra_types=('INFRASTRUCTURE_TYPE', lambda x: ', '.join(sorted(x.dropna().unique()))),
    ).sort_values('total_needs_cost', ascending=False)
    summary_path = os.path.join(GEODATA_DIR, "cwns_city_summary.csv")
    summary.to_csv(summary_path)
    print(f"Saved: {summary_path}")
    return merged


def main():
    os.makedirs(GEODATA_DIR, exist_ok=True)

    cities = load_city_boundaries()

    nbi_joined = spatial_join_nbi(cities)
    cwns_joined = spatial_join_cwns(cities)

    print("\n=== Summary ===", flush=True)
    print(f"NBI bridges assigned to cities: {len(nbi_joined):,}")
    print(f"CWNS facilities assigned to cities: {len(cwns_joined):,}")
    print(f"Cities with NBI data: {nbi_joined['geo_name'].nunique()}")
    print(f"Cities with CWNS data: {cwns_joined['geo_name'].nunique()}")


if __name__ == "__main__":
    main()
