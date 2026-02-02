"""
Build Clean Water SRF (State Revolving Fund) city-year panel.

Input:
  - CWAgreementReport.csv: EPA CWSRF agreement data (latin1 encoding)
  - Crosswalk.csv: city-to-county FIPS mapping for 578 LargeCities
  - geodata/LargeCities_places_2023.gpkg: city boundaries with ALAND

Matching tiers (each tier only matches agreements NOT matched in prior tiers):
  Tier 1 (strict): normalized borrower name == normalized city name, same state
  Tier 2 (fuzzy):  city name appears as word boundary in borrower name, same state,
                    with exclusion rules for county entities, different municipalities,
                    regional authorities, etc.
  Tier 3 (county): unmatched agreements allocated proportionally by city land-area
                    share within their county (via Crosswalk.csv)

Output:
  - geodata/srf_city_year.csv: panel with strict and inclusive (strict+fuzzy+county)
    SRF columns per city-year
"""

import pandas as pd
import numpy as np
import re
import geopandas as gpd

STATE_ABBR = {
    'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA',
    'Colorado':'CO','Connecticut':'CT','Delaware':'DE','Florida':'FL','Georgia':'GA',
    'Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN','Iowa':'IA',
    'Kansas':'KS','Kentucky':'KY','Louisiana':'LA','Maine':'ME','Maryland':'MD',
    'Massachusetts':'MA','Michigan':'MI','Minnesota':'MN','Mississippi':'MS',
    'Missouri':'MO','Montana':'MT','Nebraska':'NE','Nevada':'NV',
    'New Hampshire':'NH','New Jersey':'NJ','New Mexico':'NM','New York':'NY',
    'North Carolina':'NC','North Dakota':'ND','Ohio':'OH','Oklahoma':'OK',
    'Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC',
    'South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT',
    'Virginia':'VA','Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY',
    'District of Columbia':'DC','Puerto Rico':'PR',
}


def normalize_name(name):
    """Normalize city/borrower name for matching."""
    if pd.isna(name):
        return ''
    name = name.lower().strip()
    name = re.sub(r',?\s*(city|town|village|borough|township|municipality)\s+of\s*', '', name)
    name = re.sub(r'^(city|town|village|borough|township|municipality)\s+of\s+', '', name)
    name = re.sub(r'\s+(city|town|village|borough|township)', '', name)
    name = re.sub(r',\s*(city|town|village|borough|township)', '', name)
    name = re.sub(r'\s*\(.*\)', '', name)
    name = re.sub(r',\s*', '', name)
    return name.strip()


def parse_dollar(s):
    if pd.isna(s):
        return 0.0
    return float(str(s).replace('$', '').replace(',', ''))


def parse_counties(val):
    if pd.isna(val):
        return []
    return re.findall(r'0500000US(\d{5})', str(val))


def is_valid_fuzzy(borrower_name_raw, borrower_norm, city_norm, state_abb):
    """Check whether a fuzzy (substring) match is legitimate."""
    bn = borrower_norm
    bn_raw = borrower_name_raw.lower() if pd.notna(borrower_name_raw) else ''
    cn = city_norm

    # Rule 1: City name must appear as word boundary
    if not re.search(r'\b' + re.escape(cn) + r'\b', bn):
        return False

    # Rule 2: Exclude county entities (unless "city and county")
    if 'county' in bn_raw and 'city and county' not in bn_raw:
        return False

    # Rule 3: Exclude different municipalities (directional prefix + cityname)
    if re.search(r'\b(east|west|north|south|new)\s+' + re.escape(cn) + r'\b', bn):
        return False

    # Rule 4: Exclude "X Heights", "X Shores", etc. (different municipalities)
    if re.search(re.escape(cn) + r'\s+(heights|shores|springs|beach)\b', bn):
        return False

    # Rule 5: Exclude parish (Louisiana county equivalent)
    if 'parish' in bn_raw:
        return False

    # Rule 6: Exclude regional/conservation/irrigation/port/misc entities
    exclude_words = [
        'valley sc', 'valley water commission', 'hampton roads',
        'union sanitary', 'conservation commission', 'conservation dist',
        'conservancy', 'irrigation district',
        'port commission', 'port of catoosa',
        'waterkeeper', 'water office',
        'regional', 'redevelopment',
    ]
    for w in exclude_words:
        if w in bn_raw:
            return False

    # Rule 7: Exclude state-level entities (state name containing city name)
    state_names = ['new jersey', 'new york', 'new mexico', 'new hampshire']
    for sn in state_names:
        if sn in bn_raw and cn in sn:
            return False

    # Rule 8: Exclude "greater X" (multi-jurisdiction)
    if re.search(r'greater\s+' + re.escape(cn), bn):
        return False

    # Rule 9: Exclude "township"
    if 'township' in bn_raw:
        return False

    # Rule 10: Explicit deny list (using normalized names)
    deny_pairs = {
        ('dell rapids', 'rapid'),
        ('arkansas', 'kansas'),
        ('sioux center', 'sioux'),
        ('barcalo buffalo', 'buffalo'),
        ('east central', 'oklahoma'),
        ('central oklahoma master', 'oklahoma'),
    }
    for bad_b, bad_c in deny_pairs:
        if bad_b in bn and cn == bad_c:
            return False

    return True


def main():
    # ── Load inputs ─────────────────────────────────────────────
    srf_raw = pd.read_csv('CWAgreementReport.csv', encoding='latin1')
    cw = pd.read_csv('Crosswalk.csv')
    cities = gpd.read_file('geodata/LargeCities_places_2023.gpkg')

    # ── Prep SRF data ──────────────────────────────────────────
    srf = srf_raw.copy()
    srf['state_abb'] = srf['State'].map(STATE_ABBR)
    srf = srf[srf['state_abb'].notna()].copy()
    srf['amount'] = srf['Initial Agreement Amount'].apply(parse_dollar)
    srf['year'] = pd.to_datetime(srf['Initial Agreement Date'], errors='coerce').dt.year
    srf = srf[srf['year'].notna()].copy()
    srf['year'] = srf['year'].astype(int)
    srf['srf_idx'] = range(len(srf))
    srf['norm_borrower'] = srf['Borrower Name'].apply(normalize_name)

    # ── Prep city data ─────────────────────────────────────────
    cw['fips'] = cw['fips'].astype(str)
    cities['fips'] = cities['GEOID'].astype(str)
    city_info = cw[['fips', 'geo_name', 'state_abb']].copy()
    city_info['norm_city'] = city_info['geo_name'].apply(normalize_name)
    city_info = city_info.merge(cities[['fips', 'ALAND']], on='fips', how='left')

    # City → county mapping with area shares
    cw['county_list'] = cw['relevant_counties'].apply(parse_counties)
    city_county_rows = []
    for _, row in cw.iterrows():
        for cfips in row['county_list']:
            city_county_rows.append({'fips': str(row['fips']), 'county_fips': cfips})
    city_county = pd.DataFrame(city_county_rows)
    city_county = city_county.merge(city_info[['fips', 'ALAND', 'state_abb']], on='fips', how='left')
    county_totals = city_county.groupby('county_fips')['ALAND'].sum().reset_index()
    county_totals.columns = ['county_fips', 'county_total_ALAND']
    city_county = city_county.merge(county_totals, on='county_fips', how='left')
    city_county['area_share'] = city_county['ALAND'] / city_county['county_total_ALAND']

    # County name→FIPS mapping from crosswalk
    county_name_to_fips = {}
    for _, row in cw.iterrows():
        st = row['state_abb']
        for part in str(row.get('relevant_counties', '')).split(';'):
            m = re.match(r'(.+?)\s*\[0500000US(\d{5})\]', part.strip())
            if m:
                county_name_to_fips[(st, m.group(1).strip().lower())] = m.group(2)

    def srf_county_to_fips(row):
        if pd.isna(row['County']):
            return None
        return county_name_to_fips.get((row['state_abb'], row['County'].strip().lower()))

    srf['county_fips'] = srf.apply(srf_county_to_fips, axis=1)

    # ── Tier 1: Strict name match ──────────────────────────────
    matched_strict = srf.merge(
        city_info[['fips', 'norm_city', 'state_abb']],
        left_on=['norm_borrower', 'state_abb'],
        right_on=['norm_city', 'state_abb'],
        how='inner'
    )
    strict_idx = set(matched_strict['srf_idx'])

    # ── Tier 2: Fuzzy match (with exclusion rules) ─────────────
    srf_remaining = srf[~srf['srf_idx'].isin(strict_idx)].copy()
    fuzzy_matches = []
    for _, city_row in city_info.iterrows():
        cname = city_row['norm_city']
        st = city_row['state_abb']
        fips = city_row['fips']
        if len(cname) < 4:
            continue
        candidates = srf_remaining[
            (srf_remaining['state_abb'] == st) &
            (srf_remaining['norm_borrower'].str.contains(re.escape(cname), case=False, na=False))
        ]
        for _, sr in candidates.iterrows():
            if is_valid_fuzzy(sr['Borrower Name'], sr['norm_borrower'], cname, st):
                fuzzy_matches.append({
                    'srf_idx': sr['srf_idx'], 'fips': fips,
                    'amount': sr['amount'], 'year': sr['year'],
                })

    fuzzy_df = pd.DataFrame(fuzzy_matches) if fuzzy_matches else pd.DataFrame(
        columns=['srf_idx', 'fips', 'amount', 'year'])

    # Deduplicate: one agreement → one city (longest city name = most specific)
    if len(fuzzy_df) > 0:
        fuzzy_df = fuzzy_df.merge(
            city_info[['fips', 'norm_city']], on='fips', how='left')
        fuzzy_df['city_len'] = fuzzy_df['norm_city'].str.len()
        fuzzy_df = fuzzy_df.sort_values('city_len', ascending=False)
        fuzzy_df = fuzzy_df.drop_duplicates(subset='srf_idx', keep='first')
        fuzzy_df = fuzzy_df.drop(columns=['city_len', 'norm_city'])

    fuzzy_idx = set(fuzzy_df['srf_idx']) if len(fuzzy_df) > 0 else set()

    # ── Tier 3: County allocation ──────────────────────────────
    used_idx = strict_idx | fuzzy_idx
    srf_unmatched = srf[~srf['srf_idx'].isin(used_idx) & srf['county_fips'].notna()].copy()
    valid_counties = set(city_county['county_fips'])
    srf_county_alloc = srf_unmatched[srf_unmatched['county_fips'].isin(valid_counties)].copy()
    county_alloc = srf_county_alloc.merge(
        city_county[['county_fips', 'fips', 'area_share']],
        on='county_fips', how='inner'
    )
    county_alloc['alloc_amount'] = county_alloc['amount'] * county_alloc['area_share']

    # ── Build output panels ────────────────────────────────────
    # Strict panel
    strict_panel = matched_strict.groupby(['fips', 'year']).agg(
        srf_strict_count=('amount', 'count'),
        srf_strict_amount=('amount', 'sum')
    ).reset_index()

    # Inclusive panel (strict + fuzzy + county)
    direct_rows = []
    for _, r in matched_strict.iterrows():
        direct_rows.append({'fips': r['fips'], 'year': r['year'], 'amount': r['amount']})
    for _, r in fuzzy_df.iterrows():
        direct_rows.append({'fips': r['fips'], 'year': r['year'], 'amount': r['amount']})
    for _, r in county_alloc.iterrows():
        direct_rows.append({'fips': r['fips'], 'year': r['year'], 'amount': r['alloc_amount']})

    incl_df = pd.DataFrame(direct_rows)
    incl_panel = incl_df.groupby(['fips', 'year']).agg(
        srf_incl_count=('amount', 'count'),
        srf_incl_amount=('amount', 'sum')
    ).reset_index()

    # Merge strict and inclusive
    srf_panel = strict_panel.merge(incl_panel, on=['fips', 'year'], how='outer')
    srf_panel = srf_panel.sort_values(['fips', 'year']).reset_index(drop=True)

    # Deflate to 2017$ using BEA price deflator
    deflator = pd.read_csv(
        'geodata/city_year_investment_needs_v3.csv',
        usecols=['year', 'deflator']
    ).drop_duplicates()
    srf_panel = srf_panel.merge(deflator, on='year', how='left')
    srf_panel['srf_strict_amount_real'] = (
        srf_panel['srf_strict_amount'] * (100.0 / srf_panel['deflator']))
    srf_panel['srf_incl_amount_real'] = (
        srf_panel['srf_incl_amount'] * (100.0 / srf_panel['deflator']))
    srf_panel = srf_panel.drop(columns='deflator')

    srf_panel.to_csv('geodata/srf_city_year.csv', index=False)

    # Print summary
    total_srf = srf['amount'].sum()
    s_total = matched_strict['amount'].sum()
    f_total = fuzzy_df['amount'].sum() if len(fuzzy_df) > 0 else 0
    c_total = county_alloc['alloc_amount'].sum()
    print(f"Strict:  {len(matched_strict):>5} agreements → {matched_strict['fips'].nunique():>3} cities  ${s_total:>15,.0f}")
    print(f"Fuzzy:   {len(fuzzy_df):>5} agreements → {fuzzy_df['fips'].nunique():>3} cities  ${f_total:>15,.0f}")
    print(f"County:  {len(srf_county_alloc):>5} src agr → {county_alloc['fips'].nunique():>3} cities  ${c_total:>15,.0f}")
    print(f"Total:   ${s_total + f_total + c_total:>15,.0f} ({(s_total + f_total + c_total) / total_srf * 100:.1f}%)")
    print(f"Panel:   {len(srf_panel)} rows, {srf_panel['fips'].nunique()} cities")


if __name__ == '__main__':
    main()
