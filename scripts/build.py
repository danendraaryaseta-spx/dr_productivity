"""
Daily Rent Vendor Trip Analysis - GitHub Actions pipeline.

Replaces the live Apps Script version (which was timing out reading 200K+/240K+
row Google Sheets on every page load) with a scheduled pull: this script runs on
a cron via GitHub Actions, re-derives the whole dashboard from the 4 source
sheets, and writes a static docs/index.html for GitHub Pages. Page load is then
just a static file - no live Sheets reads at request time.

Mirrors the logic already proven out in:
  - Daily Rent Performance/scripts/tracker_raw/build_firstleg.py (first-leg detection)
  - Daily Rent Performance/scripts/productivity_onsite.py (onsite + productivity)
  - Daily Rent Performance/scripts/ordered_vs_onsite.py (order sheet comparison)
  - Daily Rent Performance/appscript-dashboard/Code.gs (the live version's port of all of the above)
"""
import os
import json
import math
from datetime import datetime, timedelta, timezone

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

SHEET_IDS = {
    'SOC_LM': '1nLurOQ1JJRRVcyA_-egi32J6nGj9darabA-al3OwG7E',
    'FM_SOC': '1wbM3PzJBWweJ0lOHljWvPmBAHqAZPO4fYYKJky32ON0',
    'ONSITE': '1yp7eVkhZftRjXCzEWye0hkCkXTLSlo-wlC1HaGf-A_Y',
    'ORDER': '1nMWBta_RA7jrNWfluMSS4J07VoPkpVKqXeUbIOmYr5E',
}
WINDOW_DAYS = 7
ORDER_HEADER_ROW = 12  # header sits at row 12 on that specific tab as of 7.7 campaign

TRACKER_COLS = ['trip_date_v2', 'trip_route', 'slot_number', 'trip_number', 'cost_type',
                'origin_station', 'dest_station', 'vehicle_type_name', 'total_loaded',
                'total_unloaded', 'trip_std', 'trip_atd', 'trip_sta', 'trip_ata',
                'trip_source', 'trip_status', 'dest_sta', 'dest_ata', 'agency_name']

# [29309] On Site Registration sheet - NOTE the header row's own text labels (row 3)
# are misaligned with the actual data by a few columns (left over from when the
# agency_name column was added). This mapping is based on verified data content,
# not the header text. No plate/registration ID here either - see build_onsite().
ONSITE_HEADER_ROW = 3
ONSITE_COLS = ['trip_date', 'original_soc_station', 'original_vehicle_type', 'agency_name', 'arrival_status', 'cost_type']


def get_client():
    key_json = os.environ['GCP_SERVICE_ACCOUNT_KEY']
    info = json.loads(key_json)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    return gspread.Client(auth=creds)


def to_date_str(v):
    if v is None or v == '':
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return pd.to_datetime(s).strftime('%Y-%m-%d')
    except Exception:
        return None


def read_tracker(gc, sheet_id, source_label):
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet('raw')
    values = ws.get('A2:S')
    rows = []
    for r in values:
        r = list(r) + [''] * (19 - len(r))
        if not r[3]:
            continue
        rows.append({
            'trip_number': r[3],
            'cost_type': r[4],
            'origin_station': r[5],
            'vehicle_type_name': r[7],
            'trip_atd': r[11],
            'agency_name': r[18],
            'source_sheet': source_label,
        })
    print(f'  {source_label}: {len(rows)} rows with a trip_number')
    return rows


def build_trips_first_leg(gc):
    all_rows = []
    for key, label in (('SOC_LM', 'SOC-LM'), ('FM_SOC', 'FM-SOC')):
        all_rows.extend(read_tracker(gc, SHEET_IDS[key], label))

    df = pd.DataFrame(all_rows)
    df['atd_dt'] = pd.to_datetime(df['trip_atd'], errors='coerce')
    departed = df[df['atd_dt'].notna()].copy()

    idx = departed.groupby('trip_number')['atd_dt'].idxmin()
    first_leg = departed.loc[idx].copy()

    anchor_date = first_leg['atd_dt'].max()
    if pd.isna(anchor_date):
        anchor_date = pd.Timestamp.now()
    window_dates = [(anchor_date.normalize() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(WINDOW_DAYS - 1, -1, -1)]

    first_leg['Date'] = first_leg['atd_dt'].dt.strftime('%Y-%m-%d')
    first_leg = first_leg[first_leg['Date'].isin(window_dates)]
    first_leg = first_leg[first_leg['origin_station'].astype(str).str.endswith(' DC')]

    trips = first_leg.rename(columns={
        'agency_name': 'Vendor', 'origin_station': 'Origin DC',
        'vehicle_type_name': 'Vehicle Type', 'cost_type': 'Cost Type',
    })[['Vendor', 'Origin DC', 'Vehicle Type', 'Cost Type', 'Date']]

    print(f'  First-leg reduction: {len(departed)} departed legs -> {len(first_leg)} distinct trips in window')
    return trips, window_dates


def build_onsite(gc, window_dates):
    min_d, max_d = window_dates[0], window_dates[-1]
    sh = gc.open_by_key(SHEET_IDS['ONSITE'])
    ws = sh.worksheet('raw')
    values = ws.get(f'A{ONSITE_HEADER_ROW + 1}:F')

    rows = []
    for i, r in enumerate(values):
        r = list(r) + [''] * (6 - len(r))
        date_str = to_date_str(r[0])
        if not date_str or date_str < min_d or date_str > max_d:
            continue
        if r[5] != 'By Day':
            continue
        status = r[4]
        if status in ('Expired', 'No Show'):
            continue
        rows.append({
            'Date': date_str,
            'Origin DC': r[1],
            'Vendor': r[3],
            'Vehicle Type': r[2] or 'Unknown',
            # No plate/registration ID in this sheet - each row is its own check-in
            # event with nothing to dedupe against, counted as one distinct unit.
            'row_key': i,
        })
    print(f'  Onsite: {len(rows)} qualifying "By Day" check-ins in window (no dedup key available)')
    return pd.DataFrame(rows)


def bucket_vehicle_type(vt):
    return 'WB' if vt == 'TRONTON (10WH)' else 'CDD L'


def build_order(gc):
    try:
        sh = gc.open_by_key(SHEET_IDS['ORDER'])
        ws = sh.worksheet('DR Campaign (LH) 7.7')
        values = ws.get(f'A{ORDER_HEADER_ROW}:AH')
        header = values[0]
        idx = {h: i for i, h in enumerate(header) if h}
        rows = []
        for r in values[1:]:
            dc = r[idx['Hub/DC Name']] if idx['Hub/DC Name'] < len(r) else ''
            if not dc:
                continue
            rows.append({
                'Origin DC': dc,
                'Order Type': r[idx['Type Unit']] if idx['Type Unit'] < len(r) else '',
                'Qty': pd.to_numeric(r[idx['Qty']] if idx['Qty'] < len(r) else 0, errors='coerce') or 0,
                'Start': to_date_str(r[idx['Contract Period Start Date']]) if idx['Contract Period Start Date'] < len(r) else None,
                'End': to_date_str(r[idx['Contract Period End Date']]) if idx['Contract Period End Date'] < len(r) else None,
            })
        print(f'  Order sheet: {len(rows)} booking rows')
        return pd.DataFrame(rows)
    except Exception as e:
        print(f'  Order sheet unavailable: {e}')
        return None


def build_dashboard_data(trips, window_dates):
    byday = trips[trips['Cost Type'] == 'By Day'].copy()
    by_vendor = byday.groupby('Vendor').size().reset_index(name='Trips').sort_values('Trips', ascending=False)
    by_dc = byday.groupby('Origin DC').size().reset_index(name='Trips').sort_values('Trips', ascending=False)
    detail = byday.groupby(['Vendor', 'Origin DC', 'Vehicle Type']).size().reset_index(name='Trips')
    by_dc_date = byday.groupby(['Origin DC', 'Date']).size().reset_index(name='Trips')
    by_vendor_date = byday.groupby(['Vendor', 'Date']).size().reset_index(name='Trips')
    by_date = byday.groupby('Date').size().reset_index(name='Trips').sort_values('Date')

    summary = {
        'total_trips': int(len(byday)),
        'total_vendors': int(byday['Vendor'].nunique()),
        'total_dcs': int(byday['Origin DC'].nunique()),
        'total_vehicle_types': int(byday['Vehicle Type'].nunique()),
        'date_min': by_date['Date'].min() if len(by_date) else '',
        'date_max': by_date['Date'].max() if len(by_date) else '',
    }
    return {
        'summary': summary,
        'by_vendor': by_vendor.to_dict('records'),
        'by_dc': by_dc.to_dict('records'),
        'detail': detail.to_dict('records'),
        'by_dc_date': by_dc_date.to_dict('records'),
        'by_vendor_date': by_vendor_date.to_dict('records'),
        'by_date': by_date.to_dict('records'),
    }


def clean_nan(records):
    for r in records:
        for k, v in r.items():
            if isinstance(v, float) and pd.isna(v):
                r[k] = None
    return records


def build_productivity(trips, onsite, window_dates):
    byday = trips[trips['Cost Type'] == 'By Day'].copy()
    trip_counts = byday.groupby(['Vendor', 'Origin DC', 'Vehicle Type', 'Date']).size().reset_index(name='LT_Trips')

    if len(onsite):
        onsite_counts = onsite.groupby(['Vendor', 'Origin DC', 'Vehicle Type', 'Date']).size().reset_index(name='Onsited')
    else:
        onsite_counts = pd.DataFrame(columns=['Vendor', 'Origin DC', 'Vehicle Type', 'Date', 'Onsited'])

    merged = trip_counts.merge(onsite_counts, on=['Vendor', 'Origin DC', 'Vehicle Type', 'Date'], how='outer').fillna(0)
    merged['Onsited'] = merged['Onsited'].astype(int)
    merged['LT_Trips'] = merged['LT_Trips'].astype(int)
    onsited_f = merged['Onsited'].astype(float).replace(0, float('nan'))
    merged['Productivity'] = (merged['LT_Trips'] / onsited_f).round(2)
    merged = merged.sort_values('LT_Trips', ascending=False)

    def agg(keys):
        g = merged.groupby(keys).agg(Onsited=('Onsited', 'sum'), LT_Trips=('LT_Trips', 'sum')).reset_index()
        g['Productivity'] = (g['LT_Trips'] / g['Onsited'].replace(0, float('nan'))).round(2)
        return g

    by_vendor = agg(['Vendor']).sort_values('LT_Trips', ascending=False)
    by_dc = agg(['Origin DC']).sort_values('LT_Trips', ascending=False)
    by_date = agg(['Date']).sort_values('Date')
    by_vendor_date = agg(['Vendor', 'Date'])

    total_onsited = int(merged['Onsited'].sum())
    total_trips = int(merged['LT_Trips'].sum())
    overall = round(total_trips / total_onsited, 2) if total_onsited else None

    return {
        'dates': window_dates,
        'total_onsited': total_onsited,
        'total_trips': total_trips,
        'overall_productivity': overall,
        'by_vendor': clean_nan(by_vendor.to_dict('records')),
        'by_dc': clean_nan(by_dc.to_dict('records')),
        'by_date': clean_nan(by_date.to_dict('records')),
        'by_vendor_date': clean_nan(by_vendor_date.to_dict('records')),
        'detail': clean_nan(merged.to_dict('records')),
    }


def build_ordered_vs_onsite(order, onsite, window_dates):
    if order is None:
        return {'rows': [], 'daily': {'dates': window_dates, 'rows': []}}

    order = order.copy()
    order['Start_dt'] = pd.to_datetime(order['Start'], errors='coerce')
    order['End_dt'] = pd.to_datetime(order['End'], errors='coerce')

    total_ordered = order.groupby(['Origin DC', 'Order Type'])['Qty'].sum().reset_index(name='Ordered')

    onsite = onsite.copy()
    if len(onsite):
        onsite['Bucket'] = onsite['Vehicle Type'].apply(bucket_vehicle_type)
        total_onsited = onsite.groupby(['Origin DC', 'Bucket']).size().reset_index(name='Onsited').rename(columns={'Bucket': 'Vehicle Type'})
    else:
        total_onsited = pd.DataFrame(columns=['Origin DC', 'Vehicle Type', 'Onsited'])

    total_ordered = total_ordered.rename(columns={'Order Type': 'Vehicle Type'})
    rows = total_ordered.merge(total_onsited, on=['Origin DC', 'Vehicle Type'], how='outer').fillna(0)
    rows['Ordered'] = rows['Ordered'].astype(int)
    rows['Onsited'] = rows['Onsited'].astype(int)

    daily_ordered_rows = []
    for dt in window_dates:
        d_ts = pd.Timestamp(dt)
        active = order[(order['Start_dt'] <= d_ts) & (order['End_dt'] >= d_ts)]
        g = active.groupby(['Origin DC', 'Order Type'])['Qty'].sum().reset_index(name='Ordered')
        g['Date'] = dt
        daily_ordered_rows.append(g)
    daily_ordered = pd.concat(daily_ordered_rows, ignore_index=True) if daily_ordered_rows else pd.DataFrame(columns=['Origin DC', 'Order Type', 'Ordered', 'Date'])
    daily_ordered = daily_ordered.rename(columns={'Order Type': 'Vehicle Type'})

    if len(onsite):
        daily_onsited = onsite.groupby(['Origin DC', 'Bucket', 'Date']).size().reset_index(name='Onsited').rename(columns={'Bucket': 'Vehicle Type'})
    else:
        daily_onsited = pd.DataFrame(columns=['Origin DC', 'Vehicle Type', 'Date', 'Onsited'])

    daily = daily_ordered.merge(daily_onsited, on=['Origin DC', 'Vehicle Type', 'Date'], how='outer').fillna(0)
    daily['Ordered'] = daily['Ordered'].astype(int)
    daily['Onsited'] = daily['Onsited'].astype(int)

    return {
        'rows': rows.to_dict('records'),
        'daily': {'dates': window_dates, 'rows': daily.to_dict('records')},
    }


def main():
    print('Connecting to Google Sheets...')
    gc = get_client()

    print('Fetching trip trackers (SOC-LM/FM-SOC)...')
    trips, window_dates = build_trips_first_leg(gc)
    print(f'  Window: {window_dates}')

    print('Fetching onsite registrations...')
    onsite = build_onsite(gc, window_dates)

    print('Fetching order sheet...')
    order = build_order(gc)

    print('Computing aggregates...')
    dashboard = build_dashboard_data(trips, window_dates)
    productivity = build_productivity(trips, onsite, window_dates)
    ordered_vs_onsite = build_ordered_vs_onsite(order, onsite, window_dates)

    raw = dict(dashboard)
    raw['productivity'] = productivity
    raw['ordered_vs_onsite'] = ordered_vs_onsite['rows']
    raw['ordered_vs_onsite_daily'] = ordered_vs_onsite['daily']
    raw['generated_at'] = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
    raw['window_days'] = WINDOW_DAYS
    raw['order_sheet_available'] = order is not None

    errors = []
    if order is None:
        errors.append({
            'source': 'DR order sheet',
            'message': 'Could not be read (likely a sharing/permission issue for the service account). '
                       'Ordered-vs-onsited figures are unavailable; onsited-only figures elsewhere are unaffected.',
        })
    raw['errors'] = errors

    template_path = os.path.join(os.path.dirname(__file__), 'template.html')
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    raw_json = json.dumps(raw, separators=(',', ':'), default=str)
    html = template.replace('__RAW_JSON__', raw_json)

    out_dir = os.path.join(os.path.dirname(__file__), '..', 'docs')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, 'index.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'Wrote {out_path} ({len(html):,} bytes)')
    print(f'Summary: {dashboard["summary"]}')
    print(f'Productivity: {productivity["total_trips"]} trips / {productivity["total_onsited"]} onsited = {productivity["overall_productivity"]}')


if __name__ == '__main__':
    main()
