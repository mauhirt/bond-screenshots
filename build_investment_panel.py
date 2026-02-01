#!/usr/bin/env python3
"""
Build a city × year panel dataset quantifying infrastructure investment needs.

Combines:
  - NBI bridge data (1992-2025): condition ratings, improvement costs, deck area,
    traffic, and structural deficiency metrics — aggregated per city per year.
  - CWNS wastewater data (2022 snapshot): facility counts, total needs costs,
    and design flow — merged as static cross-section columns.

Bridge-to-city mapping uses the 2025 NBI spatial join (by STRUCTURE_NUMBER)
to assign bridges to LargeCities boundaries, then traces each bridge backward
through all 34 years of NBI data.

Input:
  - geodata/nbi_bridges_by_city_2025.csv   (spatial join output)
  - geodata/cwns_city_summary.csv          (per-city CWNS summary)
  - nbi_data/nbi_all_years.parquet         (full NBI 1992-2025)

Output:
  - geodata/city_year_investment_needs.csv  (city × year panel, 25 columns)

Columns:
  fips, geo_name, state_abb, year          — identifiers
  total_bridges                            — bridge count in city that year
  deficient_bridges, poor/fair/good        — condition breakdown
  bridge_imp_cost_k                        — bridge improvement cost ($thousands)
  roadway_imp_cost_k                       — roadway improvement cost ($thousands)
  total_imp_cost_k                         — total improvement cost ($thousands)
  total_deck_area_sqm                      — total bridge deck area (m²)
  deficient_deck_area_sqm                  — deck area of deficient bridges (m²)
  total_adt                                — total average daily traffic
  avg_min_condition                        — mean of min(deck,super,sub) rating (0-9)
  avg_sufficiency                          — mean sufficiency rating (0-100, 1992-2018)
  avg_bridge_age                           — mean bridge age in years
  scour_critical_count                     — bridges with scour rating ≤ 3
  pct_deficient, pct_poor                  — percentage rates
  imp_cost_per_bridge_k                    — avg improvement cost per bridge ($k)
  cwns_facility_count                      — CWNS facilities in city (2022)
  cwns_total_needs_cost                    — CWNS total needs cost ($) (2022)
  cwns_total_design_flow                   — CWNS total design flow MGD (2022)
"""

import os

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEODATA_DIR = os.path.join(BASE_DIR, "geodata")


def build_bridge_city_lookup():
    """Build a (STATE_CODE, STRUCTURE_NUMBER) → city lookup from 2025 spatial join."""
    bridges = pd.read_csv(
        os.path.join(GEODATA_DIR, "nbi_bridges_by_city_2025.csv"),
        usecols=["STATE_CODE_001", "STRUCTURE_NUMBER_008", "fips", "geo_name", "state_abb"],
    )
    bridges["STATE_CODE_001"] = bridges["STATE_CODE_001"].astype(str).str.strip().str.zfill(2)
    bridges["STRUCTURE_NUMBER_008"] = bridges["STRUCTURE_NUMBER_008"].astype(str).str.strip()

    lookup = bridges.drop_duplicates(subset=["STATE_CODE_001", "STRUCTURE_NUMBER_008"])
    lookup = lookup[["STATE_CODE_001", "STRUCTURE_NUMBER_008", "fips", "geo_name", "state_abb"]]
    print(f"Bridge→city lookup: {len(lookup):,} unique bridges in {lookup['geo_name'].nunique()} cities")
    return lookup


def load_nbi_with_city(lookup):
    """Load all NBI years and join to city lookup."""
    nbi_cols = [
        "year", "STATE_CODE_001", "STRUCTURE_NUMBER_008",
        "YEAR_BUILT_027", "ADT_029",
        "STRUCTURE_LEN_MT_049", "DECK_WIDTH_MT_052",
        "DECK_COND_058", "SUPERSTRUCTURE_COND_059", "SUBSTRUCTURE_COND_060",
        "CULVERT_COND_062", "STRUCTURAL_EVAL_067",
        "BRIDGE_IMP_COST_094", "ROADWAY_IMP_COST_095", "TOTAL_IMP_COST_096",
        "SUFFICIENCY_RATING", "SCOUR_CRITICAL_113",
        "BRIDGE_CONDITION", "LOWEST_RATING", "DECK_AREA",
    ]
    print("Reading NBI parquet (all years)...", flush=True)
    nbi = pd.read_parquet(os.path.join(BASE_DIR, "nbi_data", "nbi_all_years.parquet"), columns=nbi_cols)
    nbi["STATE_CODE_001"] = nbi["STATE_CODE_001"].astype(str).str.strip().str.zfill(2)
    nbi["STRUCTURE_NUMBER_008"] = nbi["STRUCTURE_NUMBER_008"].astype(str).str.strip()

    city_nbi = nbi.merge(lookup, on=["STATE_CODE_001", "STRUCTURE_NUMBER_008"], how="inner")
    print(f"NBI records matched to cities: {len(city_nbi):,} ({city_nbi['geo_name'].nunique()} cities, {city_nbi['year'].nunique()} years)")
    return city_nbi


def compute_bridge_metrics(city_nbi):
    """Compute derived condition/cost columns."""
    for col in [
        "DECK_COND_058", "SUPERSTRUCTURE_COND_059", "SUBSTRUCTURE_COND_060",
        "CULVERT_COND_062", "STRUCTURAL_EVAL_067", "LOWEST_RATING",
        "BRIDGE_IMP_COST_094", "ROADWAY_IMP_COST_095", "TOTAL_IMP_COST_096",
        "ADT_029", "SUFFICIENCY_RATING", "SCOUR_CRITICAL_113",
        "STRUCTURE_LEN_MT_049", "DECK_WIDTH_MT_052", "DECK_AREA", "YEAR_BUILT_027",
    ]:
        city_nbi[col] = pd.to_numeric(city_nbi[col], errors="coerce")

    city_nbi["min_condition"] = city_nbi[
        ["DECK_COND_058", "SUPERSTRUCTURE_COND_059", "SUBSTRUCTURE_COND_060"]
    ].min(axis=1)

    city_nbi["is_deficient"] = (city_nbi["min_condition"] <= 4) | (city_nbi["CULVERT_COND_062"] <= 4)
    city_nbi["cond_poor"] = city_nbi["min_condition"] <= 4
    city_nbi["cond_fair"] = (city_nbi["min_condition"] >= 5) & (city_nbi["min_condition"] <= 6)
    city_nbi["cond_good"] = city_nbi["min_condition"] >= 7

    city_nbi["deck_area_calc"] = np.where(
        city_nbi["DECK_AREA"].notna(),
        city_nbi["DECK_AREA"],
        city_nbi["STRUCTURE_LEN_MT_049"] * city_nbi["DECK_WIDTH_MT_052"],
    )

    city_nbi["year_num"] = pd.to_numeric(city_nbi["year"], errors="coerce")
    city_nbi["bridge_age"] = city_nbi["year_num"] - city_nbi["YEAR_BUILT_027"]
    city_nbi["is_scour_critical"] = city_nbi["SCOUR_CRITICAL_113"] <= 3
    city_nbi["deficient_deck_area"] = np.where(city_nbi["is_deficient"], city_nbi["deck_area_calc"], 0)

    return city_nbi


def aggregate_panel(city_nbi):
    """Aggregate to city × year level."""
    print("Aggregating per city × year...", flush=True)
    panel = city_nbi.groupby(["fips", "geo_name", "state_abb", "year"]).agg(
        total_bridges=("STRUCTURE_NUMBER_008", "count"),
        deficient_bridges=("is_deficient", "sum"),
        poor_bridges=("cond_poor", "sum"),
        fair_bridges=("cond_fair", "sum"),
        good_bridges=("cond_good", "sum"),
        bridge_imp_cost_k=("BRIDGE_IMP_COST_094", "sum"),
        roadway_imp_cost_k=("ROADWAY_IMP_COST_095", "sum"),
        total_imp_cost_k=("TOTAL_IMP_COST_096", "sum"),
        total_deck_area_sqm=("deck_area_calc", "sum"),
        deficient_deck_area_sqm=("deficient_deck_area", "sum"),
        total_adt=("ADT_029", "sum"),
        avg_min_condition=("min_condition", "mean"),
        avg_sufficiency=("SUFFICIENCY_RATING", "mean"),
        avg_bridge_age=("bridge_age", "mean"),
        scour_critical_count=("is_scour_critical", "sum"),
    ).reset_index()

    for col in ["deficient_bridges", "poor_bridges", "fair_bridges", "good_bridges", "scour_critical_count"]:
        panel[col] = panel[col].astype(int)

    panel["pct_deficient"] = (panel["deficient_bridges"] / panel["total_bridges"] * 100).round(2)
    panel["pct_poor"] = (panel["poor_bridges"] / panel["total_bridges"] * 100).round(2)
    panel["imp_cost_per_bridge_k"] = (panel["total_imp_cost_k"] / panel["total_bridges"]).round(1)

    return panel.sort_values(["geo_name", "state_abb", "year"]).reset_index(drop=True)


def merge_cwns(panel):
    """Merge CWNS 2022 wastewater needs as static cross-section columns."""
    cwns_path = os.path.join(GEODATA_DIR, "cwns_city_summary.csv")
    if not os.path.exists(cwns_path):
        print("CWNS summary not found, skipping wastewater data")
        return panel

    cwns = pd.read_csv(cwns_path)
    cwns = cwns[["fips", "facility_count", "total_needs_cost", "total_design_flow"]].rename(columns={
        "facility_count": "cwns_facility_count",
        "total_needs_cost": "cwns_total_needs_cost",
        "total_design_flow": "cwns_total_design_flow",
    })
    cwns["fips"] = cwns["fips"].astype(str)
    panel["fips"] = panel["fips"].astype(str)

    merged = panel.merge(cwns, on="fips", how="left")
    n_cwns = merged["cwns_facility_count"].notna().groupby(merged["geo_name"]).first().sum()
    print(f"CWNS data merged for {n_cwns} cities")
    return merged


def main():
    os.makedirs(GEODATA_DIR, exist_ok=True)

    lookup = build_bridge_city_lookup()
    city_nbi = load_nbi_with_city(lookup)
    city_nbi = compute_bridge_metrics(city_nbi)
    panel = aggregate_panel(city_nbi)
    panel = merge_cwns(panel)

    out = os.path.join(GEODATA_DIR, "city_year_investment_needs.csv")
    panel.to_csv(out, index=False)

    print(f"\nPanel: {len(panel):,} rows, {panel['geo_name'].nunique()} cities, "
          f"years {panel['year'].min()}-{panel['year'].max()}, {len(panel.columns)} columns")
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
