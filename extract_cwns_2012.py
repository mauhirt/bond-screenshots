#!/usr/bin/env python3
"""
Extract 2012 CWNS facility data from the national HQ.mdb Access database and
spatially join to LargeCities boundaries.

Reads the HQ.mdb file using JayDeBeApi + UCanAccess JDBC driver (since mdbtools
is not available in this environment). Extracts three key tables:

  - SUMMARY_FACILITY: facility locations (CWNS_NUMBER, LATITUDE, LONGITUDE)
  - SUMMARY_NEEDS: investment needs (CWNS_NUMBER, TOTAL_OFFICIAL_NEEDS in Jan 2012$)
  - SUMMARY_FLOW: design flow (CWNS_NUMBER, PRES_TOTAL in MGD)

Lat/lon are strings like "32.8859 N" / "87.7404 W" that need parsing.

The spatial join matches facilities to LargeCities_places_2023.gpkg boundaries
(577 cities, EPSG:4269).

Prerequisites:
  - Java runtime (openjdk)
  - UCanAccess JDBC JARs in /tmp/ucanaccess/lib/:
      ucanaccess-5.0.1.jar, jackcess-4.0.5.jar, hsqldb-2.7.2.jar,
      commons-lang3-3.14.0.jar, commons-logging-1.3.0.jar

Input:
  - HQ.mdb.zip                               (21MB, extract to cwns_data/cwns_2012/)
  - LargeCities_places_2023.gpkg.zip          (city boundary polygons)

Output:
  - cwns_data/cwns_2012/facilities.csv        (27,016 rows)
  - cwns_data/cwns_2012/needs.csv             (13,668 rows)
  - cwns_data/cwns_2012/flow.csv              (15,359 rows)
  - cwns_data/cwns_2012/facilities_merged.csv (21,686 rows with valid coords)
  - geodata/cwns_2012_city_summary.csv        (501 cities, $128.1B total)
"""

import os
import zipfile

import geopandas as gpd
import jaydebeapi
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEODATA_DIR = os.path.join(BASE_DIR, "geodata")
CWNS_DIR = os.path.join(BASE_DIR, "cwns_data", "cwns_2012")


def parse_latlon(val):
    """Parse CWNS lat/lon string like '32.8859 N' or '149.3779 W' to float."""
    if pd.isna(val) or not isinstance(val, str):
        return None
    parts = val.strip().split()
    try:
        num = float(parts[0])
        if len(parts) > 1 and parts[1] in ("S", "W"):
            num = -num
        return num
    except (ValueError, IndexError):
        return None


def extract_from_mdb():
    """Extract facility, needs, and flow tables from HQ.mdb."""
    os.makedirs(CWNS_DIR, exist_ok=True)

    # Extract MDB from ZIP if needed
    mdb_path = os.path.join(CWNS_DIR, "HQ.mdb")
    if not os.path.exists(mdb_path):
        zip_path = os.path.join(BASE_DIR, "HQ.mdb.zip")
        print(f"Extracting {zip_path}...")
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(CWNS_DIR)

    # Connect via UCanAccess JDBC
    lib_dir = "/tmp/ucanaccess/lib"
    jars = [os.path.join(lib_dir, f) for f in os.listdir(lib_dir) if f.endswith(".jar")]
    url = f"jdbc:ucanaccess://{os.path.abspath(mdb_path)};memory=false"
    conn = jaydebeapi.connect("net.ucanaccess.jdbc.UcanaccessDriver", url, ["", ""], jars)

    # SUMMARY_FACILITY
    fac = pd.read_sql(
        "SELECT CWNS_NUMBER, FACILITY_NAME, CITY, STATE, COUNTY, ZIP_CODE, "
        "LATITUDE, LONGITUDE, CONGRESSIONAL_DISTRICT, OWNER_TYPE, FACILITY_TYPE "
        "FROM SUMMARY_FACILITY", conn
    )
    fac.to_csv(os.path.join(CWNS_DIR, "facilities.csv"), index=False)
    print(f"SUMMARY_FACILITY: {len(fac):,} rows")

    # SUMMARY_NEEDS
    needs = pd.read_sql(
        "SELECT CWNS_NUMBER, TOTAL_OFFICIAL_NEEDS, TOTAL_ELIGIBLE_NEEDS, "
        "I_OFFICIAL, II_OFFICIAL, IIIA_OFFICIAL, IIIB_OFFICIAL, IVA_OFFICIAL, IVB_OFFICIAL, V_OFFICIAL "
        "FROM SUMMARY_NEEDS", conn
    )
    needs.to_csv(os.path.join(CWNS_DIR, "needs.csv"), index=False)
    print(f"SUMMARY_NEEDS: {len(needs):,} rows")

    # SUMMARY_FLOW
    flow = pd.read_sql(
        "SELECT CWNS_NUMBER, PRES_TOTAL, PROJ_TOTAL FROM SUMMARY_FLOW", conn
    )
    flow.to_csv(os.path.join(CWNS_DIR, "flow.csv"), index=False)
    print(f"SUMMARY_FLOW: {len(flow):,} rows")

    conn.close()
    return fac, needs, flow


def merge_and_geocode(fac, needs, flow):
    """Merge tables, parse lat/lon, filter to valid coordinates."""
    merged = fac.merge(needs[["CWNS_NUMBER", "TOTAL_OFFICIAL_NEEDS"]], on="CWNS_NUMBER", how="left")
    merged = merged.merge(flow[["CWNS_NUMBER", "PRES_TOTAL"]], on="CWNS_NUMBER", how="left")

    merged["lat"] = merged["LATITUDE"].apply(parse_latlon)
    merged["lon"] = merged["LONGITUDE"].apply(parse_latlon)
    valid = merged.dropna(subset=["lat", "lon"])
    valid = valid[(valid["lat"] > 17) & (valid["lat"] < 72) & (valid["lon"] > -180) & (valid["lon"] < -60)]

    out = os.path.join(CWNS_DIR, "facilities_merged.csv")
    valid.to_csv(out, index=False)
    total = valid["TOTAL_OFFICIAL_NEEDS"].sum()
    print(f"Merged: {len(valid):,} facilities with valid coords, ${total/1e9:.1f}B total needs")
    print(f"Saved: {out}")
    return valid


def spatial_join_to_cities(facilities):
    """Spatial join 2012 CWNS facilities to LargeCities boundaries."""
    os.makedirs(GEODATA_DIR, exist_ok=True)

    # Load city boundaries
    gpkg_zip = os.path.join(BASE_DIR, "LargeCities_places_2023.gpkg.zip")
    gpkg_path = os.path.join(GEODATA_DIR, "LargeCities_places_2023.gpkg")
    if not os.path.exists(gpkg_path):
        with zipfile.ZipFile(gpkg_zip) as zf:
            zf.extractall(GEODATA_DIR)

    cities = gpd.read_file(gpkg_path)
    if cities.crs is None:
        cities = cities.set_crs("EPSG:4269")
    else:
        cities = cities.to_crs("EPSG:4269")

    # Create GeoDataFrame from facilities
    fac_gdf = gpd.GeoDataFrame(
        facilities,
        geometry=gpd.points_from_xy(facilities["lon"], facilities["lat"]),
        crs="EPSG:4269",
    )

    # Spatial join
    city_cols = ["fips", "geo_name", "state_abb", "geometry"]
    joined = gpd.sjoin(fac_gdf, cities[city_cols], how="inner", predicate="within")
    print(f"Facilities in cities: {len(joined):,} ({joined['geo_name'].nunique()} cities)")

    # Aggregate per city
    summary = joined.groupby(["fips", "geo_name", "state_abb"]).agg(
        cwns_facility_count_2012=("CWNS_NUMBER", "nunique"),
        total_needs_2012=("TOTAL_OFFICIAL_NEEDS", "sum"),
        design_flow_2012=("PRES_TOTAL", "sum"),
    ).reset_index().sort_values("total_needs_2012", ascending=False)

    out = os.path.join(GEODATA_DIR, "cwns_2012_city_summary.csv")
    summary.to_csv(out, index=False)
    total = summary["total_needs_2012"].sum()
    print(f"City summary: {len(summary)} cities, ${total/1e9:.1f}B total needs")
    print(f"Saved: {out}")
    return summary


def main():
    fac, needs, flow = extract_from_mdb()
    valid = merge_and_geocode(fac, needs, flow)
    summary = spatial_join_to_cities(valid)

    print(f"\nTop 5 cities by 2012 CWNS needs:")
    for _, r in summary.head(5).iterrows():
        print(f"  {r['geo_name']:25s} {r['state_abb']}  ${r['total_needs_2012']/1e9:>5.1f}B  ({r['cwns_facility_count_2012']:.0f} facilities)")


if __name__ == "__main__":
    main()
