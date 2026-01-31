#!/usr/bin/env python3
"""
Download EPA Clean Watersheds Needs Survey (CWNS) 2022 data from:
  https://sdwis.epa.gov/ords/sfdw_pub/r/sfdw/cwns_pub/data-download

Downloads:
  1. National CSV ZIP  (all tables, nationwide)
  2. National Access Database ZIP
  3. Data Dictionary (XLSX)
  4. Per-state CSV ZIPs (56 states/territories)

The portal uses Oracle APEX with session-state protection, so we:
  - Establish a session by visiting the main data-download page
  - Use popup URLs (with valid checksums) to set P2_LOCATION_ID in session
  - Use AJAX calls to change P2_LOCATION_ID for per-state downloads
  - Fetch page 2 to trigger the actual file download
"""

import os
import re
import sys
import time
import html

import requests

BASE = "https://sdwis.epa.gov/ords/sfdw_pub"
OUTPUT_DIR = "cwns_data"
STATES_DIR = os.path.join(OUTPUT_DIR, "states")

STATES = [
    'AK', 'AL', 'AR', 'AS', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE',
    'FL', 'GA', 'GU', 'HI', 'IA', 'ID', 'IL', 'IN', 'KS', 'KY',
    'LA', 'MA', 'MD', 'ME', 'MI', 'MN', 'MO', 'MP', 'MS', 'MT',
    'NC', 'ND', 'NE', 'NH', 'NJ', 'NM', 'NV', 'NY', 'OH', 'OK',
    'OR', 'PA', 'PR', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VA',
    'VI', 'VT', 'WA', 'WI', 'WV', 'WY',
]


def get_session():
    """Establish an APEX session and extract popup URLs."""
    s = requests.Session()
    r = s.get(f"{BASE}/r/sfdw/cwns_pub/data-download", timeout=60)
    r.raise_for_status()

    session_id = re.search(r'APP_SESSION.*?["\'](\d+)["\']', r.text).group(1)
    decoded = html.unescape(r.text.replace('\\u002F', '/').replace('\\u0026', '&'))

    popup_urls = re.findall(
        r'/ords/sfdw_pub/r/sfdw/cwns_pub/download-popup\?[^\'\"]+', decoded
    )
    dd_match = re.search(r'download-state-zip\?p2_location_id=DD[^"\']+', decoded)

    urls = {}
    for u in popup_urls:
        if 'NA_CSV' in u:
            urls['national_csv'] = BASE.rsplit('/', 1)[0] + u
        elif 'NA_AC' in u:
            urls['national_access'] = BASE.rsplit('/', 1)[0] + u
        elif 'p3_type=State' in u:
            urls['state_popup'] = BASE.rsplit('/', 1)[0] + u
    if dd_match:
        urls['data_dict'] = f"{BASE}/r/sfdw/cwns_pub/" + dd_match.group(0)

    return s, session_id, urls


def download_via_popup(session, session_id, popup_url, output_path, label=""):
    """Visit a popup URL (sets session state), then fetch page 2 for download."""
    if os.path.exists(output_path) and os.path.getsize(output_path) > 100:
        print(f"  [{label}] Already exists, skipping.", flush=True)
        return True

    # Visit popup to set P2_LOCATION_ID in session state
    r = session.get(popup_url, timeout=60)
    if r.status_code != 200:
        print(f"  [{label}] Popup failed: {r.status_code}", flush=True)
        return False

    # Fetch page 2 to trigger download
    r2 = session.get(f"{BASE}/f?p=148:2:{session_id}:::::", timeout=300)
    cd = r2.headers.get('Content-Disposition', '')
    if r2.status_code == 200 and len(r2.content) > 100 and 'html' not in r2.headers.get('Content-Type', ''):
        with open(output_path, 'wb') as f:
            f.write(r2.content)
        fname = re.search(r'filename="([^"]+)"', cd)
        print(f"  [{label}] Saved {fname.group(1) if fname else output_path} ({len(r2.content):,} bytes)", flush=True)
        return True
    else:
        print(f"  [{label}] Download failed: status={r2.status_code}, type={r2.headers.get('Content-Type')}", flush=True)
        return False


def download_state(session, session_id, state_code, output_path):
    """Download a state ZIP by setting P2_LOCATION_ID via AJAX, then fetching page 2."""
    if os.path.exists(output_path) and os.path.getsize(output_path) > 100:
        print(f"  [{state_code}] Already exists, skipping.", flush=True)
        return True

    # Set P2_LOCATION_ID in session state via AJAX
    ajax_url = f"{BASE}/wwv_flow.ajax"
    data = {
        'p_flow_id': '148',
        'p_flow_step_id': '2',
        'p_instance': session_id,
        'p_arg_names': 'P2_LOCATION_ID',
        'p_arg_values': state_code,
    }
    r = session.post(ajax_url, data=data, timeout=30,
                     headers={'X-Requested-With': 'XMLHttpRequest'})
    if r.status_code != 200:
        print(f"  [{state_code}] AJAX failed: {r.status_code}", flush=True)
        return False

    # Fetch page 2 for download
    r2 = session.get(f"{BASE}/f?p=148:2:{session_id}:::::", timeout=300)
    cd = r2.headers.get('Content-Disposition', '')
    ct = r2.headers.get('Content-Type', '')
    if r2.status_code == 200 and len(r2.content) > 100 and 'html' not in ct:
        with open(output_path, 'wb') as f:
            f.write(r2.content)
        fname = re.search(r'filename="([^"]+)"', cd)
        print(f"  [{state_code}] Saved {fname.group(1) if fname else output_path} ({len(r2.content):,} bytes)", flush=True)
        return True
    else:
        print(f"  [{state_code}] Download failed: status={r2.status_code}, type={ct}", flush=True)
        return False


def main():
    os.makedirs(STATES_DIR, exist_ok=True)

    print("Establishing APEX session...", flush=True)
    session, session_id, urls = get_session()
    print(f"  Session ID: {session_id}", flush=True)

    # 1. Data Dictionary
    if 'data_dict' in urls:
        print("\nDownloading Data Dictionary...", flush=True)
        dd_path = os.path.join(OUTPUT_DIR, "cwns_data_dictionary.xlsx")
        if not (os.path.exists(dd_path) and os.path.getsize(dd_path) > 100):
            r = session.get(urls['data_dict'], timeout=60)
            if r.status_code == 200 and 'html' not in r.headers.get('Content-Type', ''):
                with open(dd_path, 'wb') as f:
                    f.write(r.content)
                print(f"  Saved data dictionary ({len(r.content):,} bytes)", flush=True)
        else:
            print("  Already exists, skipping.", flush=True)

    # 2. National CSV
    if 'national_csv' in urls:
        print("\nDownloading National CSV ZIP...", flush=True)
        download_via_popup(
            session, session_id, urls['national_csv'],
            os.path.join(OUTPUT_DIR, "cwns_national_csv.zip"),
            label="National CSV"
        )

    # 3. National Access Database
    if 'national_access' in urls:
        print("\nDownloading National Access Database ZIP...", flush=True)
        download_via_popup(
            session, session_id, urls['national_access'],
            os.path.join(OUTPUT_DIR, "cwns_national_access.zip"),
            label="National Access"
        )

    # 4. Per-state downloads
    print(f"\nDownloading {len(STATES)} state/territory ZIPs...", flush=True)

    # First, visit the state popup to establish proper session state context
    if 'state_popup' in urls:
        session.get(urls['state_popup'], timeout=60)

    failed = []
    for i, state in enumerate(STATES):
        out_path = os.path.join(STATES_DIR, f"cwns_{state}.zip")
        ok = download_state(session, session_id, state, out_path)
        if not ok:
            failed.append(state)
        # Small delay to be polite to the server
        if (i + 1) % 10 == 0:
            time.sleep(1)

    # Summary
    print(f"\n{'='*60}", flush=True)
    print(f"Download complete!", flush=True)
    total_files = sum(1 for f in os.listdir(OUTPUT_DIR) if os.path.isfile(os.path.join(OUTPUT_DIR, f)))
    state_files = sum(1 for f in os.listdir(STATES_DIR) if f.endswith('.zip'))
    total_size = sum(
        os.path.getsize(os.path.join(dp, f))
        for dp, _, files in os.walk(OUTPUT_DIR)
        for f in files
    )
    print(f"  National files: {total_files}", flush=True)
    print(f"  State ZIPs: {state_files}/{len(STATES)}", flush=True)
    print(f"  Total size: {total_size / 1e6:.1f} MB", flush=True)
    if failed:
        print(f"  Failed states: {failed}", flush=True)


if __name__ == "__main__":
    main()
