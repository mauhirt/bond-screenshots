#!/usr/bin/env python3
"""
Build time-varying CWNS wastewater investment needs (2012-2025) by interpolating
between the 2012 and 2022 CWNS surveys, then merge with NBI bridge data into a
city × year panel with all dollar values deflated to constant 2017 dollars.

CWNS needs are allocated to LargeCities using a tiered approach:

  Tier 0 — Point-level spatial join: CWNS facilities with precise lat/lon are
           spatially joined to LargeCities boundaries.
  Tier 1 — City-centroid spatial join: CWNS facilities with city-level centroids
           (stormwater, decentralized WW) are spatially joined to boundaries.
  Tier 2 — County-level allocation: CWNS facilities with only county-level geography
           are allocated to LargeCities in that county, weighted by city land area
           share of the county's total LargeCity area (via Crosswalk.csv).
  Tier 3 — State-level allocation: CWNS facilities with only state-level geography
           are allocated to LargeCities in that state by area share.

2012 CWNS: Tier 0 from HQ.mdb spatial join + Tier 3 state allocation.
           550 cities, $141.6B total (Jan 2012 dollars).
2022 CWNS: Tiers 0-3 from EPA CSV download.
           578 cities, $293.1B total (Jan 2022 dollars).

Deflation:  BEA price index for state & local government gross investment in
            structures (FRED series Y650RG3A086NBEA), base year 2017=100.

Order of operations:
  1. Deflate 2012 city totals to 2017$ using deflator[2012] = 93.381
  2. Deflate 2022 city totals to 2017$ using deflator[2022] = 121.390
  3. Linearly interpolate between deflated values for 2013-2021
  4. Hold at 2022 deflated value for 2022-2025
  5. Cities in only one survey use that survey's deflated value for all years
  6. Deflate NBI bridge cost columns to 2017$ using each year's deflator

Input:
  - geodata/cwns_2012_city_summary.csv   (2012 per-city, tiered allocation)
  - geodata/cwns_city_summary_v2.csv     (2022 per-city, tiered allocation)
  - geodata/city_year_investment_needs.csv  (NBI bridge panel, v1)
  - Crosswalk.csv                        (city-to-county FIPS mapping)

Output:
  - geodata/cwns_interpolated_panel.csv       (standalone CWNS panel, 2012-2025)
  - geodata/city_year_investment_needs_v3.csv  (combined panel, 31 columns)
  - geodata/bea_deflator_Y650RG3A086NBEA.csv  (deflator series)
"""

import os

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEODATA_DIR = os.path.join(BASE_DIR, "geodata")

# BEA price index: state & local govt gross investment in structures (2017=100)
DEFLATOR = {
    2000: 64.132, 2001: 66.202, 2002: 67.657, 2003: 69.786, 2004: 72.855,
    2005: 76.026, 2006: 79.121, 2007: 82.365, 2008: 86.153, 2009: 86.122,
    2010: 88.275, 2011: 91.143, 2012: 93.381, 2013: 96.052, 2014: 98.151,
    2015: 98.056, 2016: 97.963, 2017: 100.000, 2018: 103.295, 2019: 104.433,
    2020: 106.338, 2021: 112.508, 2022: 121.390, 2023: 123.613, 2024: 126.565,
}
DEFLATOR[2025] = round(DEFLATOR[2024] + (DEFLATOR[2024] - DEFLATOR[2023]), 3)  # ~129.517


def build_cwns_interpolated():
    """Interpolate CWNS city-level needs between 2012 and 2022.

    Uses tiered allocation for both surveys:
    - 2012: 550 cities, $141.6B (Tier 0 spatial join + Tier 3 state allocation)
    - 2022: 578 cities, $293.1B (Tiers 0-3: point, city, county, state)

    Cities in both surveys (550) get linear interpolation.
    Cities only in 2022 (28) use constant 2022 value for all years.
    """

    # 2012 city summary (Jan 2012 dollars, tiered allocation)
    cwns_2012 = pd.read_csv(os.path.join(GEODATA_DIR, "cwns_2012_city_summary.csv"))
    cwns_2012["fips"] = cwns_2012["fips"].astype(str)

    # 2022 city summary (Jan 2022 dollars, tiered allocation)
    cwns_2022 = pd.read_csv(os.path.join(GEODATA_DIR, "cwns_city_summary_v2.csv"))
    cwns_2022["fips"] = cwns_2022["fips"].astype(str)

    # Outer join: keep cities from either survey
    both = cwns_2012[["fips", "geo_name", "state_abb",
                       "facility_count_2012", "total_needs_2012", "design_flow_2012"]].merge(
        cwns_2022[["fips", "geo_name", "state_abb",
                    "facility_count_2022", "total_needs_2022", "design_flow_2022"]],
        on="fips", how="outer", suffixes=("_12", "_22"),
    )
    both["geo_name"] = both["geo_name_22"].fillna(both["geo_name_12"])
    both["state_abb"] = both["state_abb_22"].fillna(both["state_abb_12"])
    both = both.drop(columns=["geo_name_12", "geo_name_22", "state_abb_12", "state_abb_22"])

    n_both = both["total_needs_2012"].notna() & both["total_needs_2022"].notna()
    n_only12 = both["total_needs_2012"].notna() & both["total_needs_2022"].isna()
    n_only22 = both["total_needs_2012"].isna() & both["total_needs_2022"].notna()
    print(f"CWNS cities: {n_both.sum()} in both, {n_only12.sum()} only-2012, {n_only22.sum()} only-2022")

    # Deflate both endpoints to constant 2017$
    both["needs_2012_real"] = both["total_needs_2012"] * (100.0 / DEFLATOR[2012])
    both["needs_2022_real"] = both["total_needs_2022"] * (100.0 / DEFLATOR[2022])

    # Build year-by-year panel
    rows = []
    for _, c in both.iterrows():
        n12 = c["needs_2012_real"] if pd.notna(c["needs_2012_real"]) else None
        n22 = c["needs_2022_real"] if pd.notna(c["needs_2022_real"]) else None
        fac_12 = c["facility_count_2012"] if pd.notna(c.get("facility_count_2012")) else None
        fac_22 = c["facility_count_2022"] if pd.notna(c.get("facility_count_2022")) else None
        flow_12 = c["design_flow_2012"] if pd.notna(c.get("design_flow_2012")) else None
        flow_22 = c["design_flow_2022"] if pd.notna(c.get("design_flow_2022")) else None

        for year in range(2012, 2026):
            if year == 2012:
                needs = n12
                src = "2012" if n12 is not None else None
                f, fl = fac_12, flow_12
            elif year == 2022:
                needs = n22
                src = "2022" if n22 is not None else None
                f, fl = fac_22, flow_22
            elif year < 2022 and n12 is not None and n22 is not None:
                t = (year - 2012) / 10.0
                needs = n12 + t * (n22 - n12)
                src = "interpolated"
                f = fac_12 + t * ((fac_22 or fac_12) - fac_12) if fac_12 is not None else fac_22
                fl = flow_12 + t * ((flow_22 or flow_12) - flow_12) if flow_12 is not None else flow_22
            elif year > 2022:
                needs = n22
                src = "2022" if n22 is not None else None
                f, fl = fac_22, flow_22
            elif n12 is not None:
                needs, src, f, fl = n12, "2012", fac_12, flow_12
            elif n22 is not None:
                needs, src, f, fl = n22, "2022", fac_22, flow_22
            else:
                needs, src, f, fl = None, None, None, None

            rows.append({
                "fips": c["fips"], "geo_name": c["geo_name"], "state_abb": c["state_abb"],
                "year": year, "cwns_needs_real": needs, "cwns_flow_interp": fl,
                "cwns_facilities_interp": f, "cwns_source": src,
            })

    panel = pd.DataFrame(rows)
    out = os.path.join(GEODATA_DIR, "cwns_interpolated_panel.csv")
    panel.to_csv(out, index=False)
    print(f"CWNS interpolated: {len(panel):,} rows ({panel['geo_name'].nunique()} cities, {panel['year'].nunique()} years)")
    print(f"Saved: {out}")
    return panel


def build_final_panel(cwns_panel):
    """Merge interpolated CWNS with NBI bridge panel, deflate all costs."""

    # Load bridge panel v1
    bridge = pd.read_csv(os.path.join(GEODATA_DIR, "city_year_investment_needs.csv"))
    old_cwns = ["cwns_facility_count", "cwns_total_needs_cost", "cwns_total_design_flow"]
    bridge = bridge.drop(columns=[c for c in old_cwns if c in bridge.columns])

    # Deflate bridge costs
    bridge["deflator"] = bridge["year"].map(DEFLATOR)
    for col in ["bridge_imp_cost_k", "roadway_imp_cost_k", "total_imp_cost_k"]:
        real_col = col.replace("_k", "_real_k")
        bridge[real_col] = np.where(bridge["deflator"].notna(), bridge[col] * (100.0 / bridge["deflator"]), np.nan)
    bridge["imp_cost_per_bridge_real_k"] = np.where(
        bridge["deflator"].notna(), bridge["imp_cost_per_bridge_k"] * (100.0 / bridge["deflator"]), np.nan
    )

    # Merge CWNS
    cwns_panel["fips"] = cwns_panel["fips"].astype(str)
    bridge["fips"] = bridge["fips"].astype(str)
    merged = bridge.merge(
        cwns_panel[["fips", "year", "cwns_needs_real", "cwns_flow_interp", "cwns_facilities_interp", "cwns_source"]],
        on=["fips", "year"], how="left",
    )

    # Column order
    id_cols = ["fips", "geo_name", "state_abb", "year"]
    count_cols = ["total_bridges", "deficient_bridges", "poor_bridges", "fair_bridges", "good_bridges"]
    cost_nom = ["bridge_imp_cost_k", "roadway_imp_cost_k", "total_imp_cost_k", "imp_cost_per_bridge_k"]
    cost_real = ["bridge_imp_cost_real_k", "roadway_imp_cost_real_k", "total_imp_cost_real_k", "imp_cost_per_bridge_real_k"]
    other = ["total_deck_area_sqm", "deficient_deck_area_sqm", "total_adt",
             "avg_min_condition", "avg_sufficiency", "avg_bridge_age", "scour_critical_count",
             "pct_deficient", "pct_poor"]
    cwns = ["cwns_needs_real", "cwns_flow_interp", "cwns_facilities_interp", "cwns_source"]
    meta = ["deflator"]
    col_order = [c for c in id_cols + count_cols + cost_nom + cost_real + other + cwns + meta if c in merged.columns]
    merged = merged[col_order].sort_values(["geo_name", "state_abb", "year"]).reset_index(drop=True)

    out = os.path.join(GEODATA_DIR, "city_year_investment_needs_v3.csv")
    merged.to_csv(out, index=False)
    print(f"\nFinal panel: {len(merged):,} rows, {merged['geo_name'].nunique()} cities, {len(merged.columns)} columns")
    print(f"Saved: {out}")
    return merged


def main():
    os.makedirs(GEODATA_DIR, exist_ok=True)

    # Save deflator
    defl_df = pd.DataFrame([{"year": k, "deflator": v} for k, v in sorted(DEFLATOR.items())])
    defl_df.to_csv(os.path.join(GEODATA_DIR, "bea_deflator_Y650RG3A086NBEA.csv"), index=False)

    cwns = build_cwns_interpolated()
    panel = build_final_panel(cwns)

    # Summary
    p22 = panel[panel["year"] == 2022].copy()
    p22["combined"] = p22["total_imp_cost_real_k"].fillna(0) * 1000 + p22["cwns_needs_real"].fillna(0)
    print(f"\nTop 5 cities by 2022 combined real needs:")
    for _, r in p22.nlargest(5, "combined").iterrows():
        print(f"  {r['geo_name']:25s} {r['state_abb']}  bridge: ${r['total_imp_cost_real_k']/1000:>8,.0f}M  cwns: ${r['cwns_needs_real']/1e9:>5.1f}B")


if __name__ == "__main__":
    main()
