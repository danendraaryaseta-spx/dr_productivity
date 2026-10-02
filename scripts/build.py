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
    'SOC_SOC': '11CidDsjDqCWxZ0u9AWXDm2gyNQeWi8vKd_lOInv6_lU',
}
ORDER_HEADER_ROW = 12  # header sits at row 12 on that specific tab as of 7.7 campaign

TRACKER_COLS = ['trip_date_v2', 'trip_route', 'slot_number', 'trip_number', 'cost_type',
                'origin_station', 'dest_station', 'vehicle_type_name', 'total_loaded',
                'total_unloaded', 'trip_std', 'trip_atd', 'trip_sta', 'trip_ata',
                'trip_source', 'trip_status', 'dest_sta', 'dest_ata', 'agency_name']

# [29309] On Site Registration sheet - a Nopol (plate number) column and a
# Region column were added (Region appended at the end, after cost_type). The
# header row's own text labels now match this layout (verified against sample
# data). Plate lets us dedupe duplicate check-in rows for the same vehicle -
# see build_onsite(). Region lets the dashboard filter DCs by region even
# though the trip trackers (SOC-LM/FM-SOC) have no region field of their own -
# see build_dc_regions().
ONSITE_HEADER_ROW = 3
ONSITE_COLS = ['trip_date', 'original_soc_station', 'nopol', 'original_vehicle_type', 'agency_name', 'arrival_status', 'cost_type', 'region']


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
    values = ws.get('A1:AZ')
    header = [str(h).strip().lower() for h in values[0]] if values else []
    # cost_type is located by header name: it's column E on SOC-LM/FM-SOC but column T on SOC-SOC.
    ct_idx = header.index('cost_type') if 'cost_type' in header else None
    width = max(19, (ct_idx or 0) + 1)
    rows = []
    for r in values[1:]:
        r = list(r) + [''] * (width - len(r))
        if not r[3]:
            continue
        rows.append({
            'trip_number': r[3],
            'trip_route': r[1],
            'cost_type': r[ct_idx] if ct_idx is not None else '',
            'origin_station': r[5],
            'dest_station': r[6],
            'vehicle_type_name': r[7],
            'total_loaded': r[8],
            'trip_atd': r[11],
            'trip_ata': r[13],
            'dest_ata': r[17],
            'agency_name': r[18],
            'source_sheet': source_label,
        })
    by_day = sum(1 for r in rows if r['cost_type'] == 'By Day')
    print(f'  {source_label}: {len(rows)} rows with a trip_number, {by_day} By Day (cost_type column: {"yes" if ct_idx is not None else "MISSING"})')
    return rows


def station_kind(name):
    s = str(name).strip()
    if s.endswith(' DC') or s.endswith(' RDC'):
        return 'SOC'
    if 'first mile' in s.lower():
        return 'FM'
    if s.lower().endswith('hub'):
        return 'LM'
    return 'Other'


def build_lt_stats(df):
    # Only finished LTs (every leg, departed or not, has a dest_ata) are judged, since an
    # in-progress LT would look empty or too short.
    # "Empty" only looks at legs that should carry parcels: FM hub -> SOC (pickup coming in)
    # and SOC -> LM hub (delivery going out). SOC -> FM and LM -> SOC normally run empty.
    # total_loaded is what was loaded at that leg's origin; blank counts as 0.
    #   SOC -> LM is empty when nothing was loaded at the SOC.
    #   FM -> SOC is empty when the truck arrives carrying nothing: 0 loaded at every hub since
    #   it last left a SOC. Parcels loaded AT a SOC are deliveries dropped at LM hubs, so they
    #   don't count - and a multi-stop pickup whose last FM hub had nothing isn't empty.
    kinds = {s: station_kind(s) for s in pd.unique(pd.concat([df['origin_station'], df['dest_station']]))}
    o_kind, d_kind = df['origin_station'].map(kinds), df['dest_station'].map(kinds)
    inbound = (o_kind == 'FM') & (d_kind == 'SOC')
    outbound = (o_kind == 'SOC') & (d_kind == 'LM')
    loaded = pd.to_numeric(df['total_loaded'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    s = df[['trip_number', 'atd_dt']].sort_values(['trip_number', 'atd_dt'])
    from_soc = (o_kind.loc[s.index] == 'SOC').values
    seg = pd.Series(from_soc, index=s.index).groupby(s['trip_number'].values).cumsum()
    hub_loaded = loaded.loc[s.index].where(~from_soc, 0)
    on_board = hub_loaded.groupby([s['trip_number'].values, seg.values]).cumsum().reindex(df.index)
    legs = pd.DataFrame({
        'trip_number': df['trip_number'],
        'ata': pd.to_datetime(df['trip_ata'], errors='coerce'),
        'dest_ata': pd.to_datetime(df['dest_ata'], errors='coerce'),
        'relevant': inbound | outbound,
        'empty_leg': (inbound & (on_board == 0)) | (outbound & (loaded == 0)),
    })
    lt = legs.groupby('trip_number').agg(
        n_legs=('trip_number', 'size'), n_arrived=('dest_ata', 'count'),
        n_relevant=('relevant', 'sum'), n_empty=('empty_leg', 'sum'),
        first_ata=('ata', 'min'), last_dest_ata=('dest_ata', 'max'),
    )
    hrs = (lt['last_dest_ata'] - lt['first_ata']).dt.total_seconds() / 3600
    finished = (lt['n_arrived'] == lt['n_legs']) & (hrs >= 0)
    stats = pd.DataFrame({
        'Finished': finished,
        'Empty': finished & (lt['n_relevant'] > 0) & (lt['n_empty'] == lt['n_relevant']),
        'Work_Hrs': hrs.where(finished),
    })
    rel = df.loc[inbound | outbound, ['trip_number', 'origin_station', 'dest_station']].copy()
    rel['Direction'] = inbound[inbound | outbound].map({True: 'FM → SOC', False: 'SOC → LM'})
    rel['SOC'] = rel['dest_station'].where(rel['Direction'] == 'FM → SOC', rel['origin_station'])
    rel['Hub'] = rel['origin_station'].where(rel['Direction'] == 'FM → SOC', rel['dest_station'])
    rel['Empty'] = legs.loc[rel.index, 'empty_leg']
    return stats, rel[['trip_number', 'Direction', 'SOC', 'Hub', 'Empty']]


def build_trips_first_leg(gc):
    all_rows = []
    for key, label in (('SOC_LM', 'SOC-LM'), ('FM_SOC', 'FM-SOC')):
        all_rows.extend(read_tracker(gc, SHEET_IDS[key], label))
    try:
        all_rows.extend(read_tracker(gc, SHEET_IDS['SOC_SOC'], 'SOC-SOC'))
        socsoc_ok = True
    except Exception as e:
        print(f'  SOC-SOC unavailable: {e}')
        socsoc_ok = False

    df = pd.DataFrame(all_rows)
    df['atd_dt'] = pd.to_datetime(df['trip_atd'], errors='coerce')
    departed = df[df['atd_dt'].notna()].copy()

    idx = departed.groupby('trip_number')['atd_dt'].idxmin()
    first_leg = departed.loc[idx].copy()

    first_leg['Date'] = first_leg['atd_dt'].dt.strftime('%Y-%m-%d')
    lt_stats, rel_legs = build_lt_stats(df)
    # Every departed LT, any cost type or starting station - the empty-leg section covers
    # FM pickup routes that start at a hub too.
    all_lts = first_leg.join(lt_stats, on='trip_number')
    first_leg = all_lts[all_lts['origin_station'].astype(str).str.endswith(' DC')]

    # No fixed rolling window - the full date range actually present in the
    # trackers is exposed to the client, which lets the user pick any sub-range
    # via the date-window control while still showing daily granularity.
    all_dates = sorted(first_leg['Date'].dropna().unique().tolist())

    trips = first_leg.rename(columns={
        'agency_name': 'Vendor', 'origin_station': 'Origin DC',
        'vehicle_type_name': 'Vehicle Type', 'cost_type': 'Cost Type',
        'trip_number': 'LT', 'trip_route': 'Route',
    })[['LT', 'Vendor', 'Origin DC', 'Vehicle Type', 'Cost Type', 'Date', 'Route', 'Finished', 'Empty', 'Work_Hrs']]

    print(f'  First-leg reduction: {len(departed)} departed legs -> {len(first_leg)} distinct trips across {len(all_dates)} days')
    return trips, all_dates, socsoc_ok, rel_legs, all_lts


def build_onsite(gc, window_dates):
    min_d, max_d = window_dates[0], window_dates[-1]
    sh = gc.open_by_key(SHEET_IDS['ONSITE'])
    ws = sh.worksheet('raw')
    values = ws.get(f'A{ONSITE_HEADER_ROW + 1}:H')

    rows = []
    for i, r in enumerate(values):
        r = list(r) + [''] * (8 - len(r))
        date_str = to_date_str(r[0])
        if not date_str or date_str < min_d or date_str > max_d:
            continue
        if r[6] != 'By Day':
            continue
        status = r[5]
        if status in ('Expired', 'No Show'):
            continue
        vendor = r[4]
        plate = str(r[2]).strip().upper()
        rows.append({
            'Date': date_str,
            'Origin DC': r[1],
            'Vendor': vendor,
            'Vehicle Type': r[3] or 'Unknown',
            'Region': r[7] or 'Unknown',
            # Dedupe key: distinct plate+vendor *per day*. A vehicle checked in
            # twice the same day for the same vendor/DC/type is one physical unit,
            # not two - but the same vehicle onsite on different days still counts
            # once per day (this feeds "vehicle-days" totals). Falls back to a
            # per-row key when Nopol is blank (legacy rows before this column
            # existed) so those still count individually.
            'unit_key': f'{date_str}|{vendor}|{plate}' if plate else f'ROW_{i}',
        })
    print(f'  Onsite: {len(rows)} qualifying "By Day" check-ins in window')
    return pd.DataFrame(rows)


def build_dc_regions(trips, onsite, order):
    # Region only exists on the onsite sheet, not the trip trackers or the order
    # sheet - build a DC -> Region map from onsite data and reuse it everywhere
    # Origin DC shows up: trips, onsite, AND the order sheet (Ordered-vs-Onsited
    # table). A DC seen only in trips/order (never onsite) has no known region
    # and is bucketed 'Unknown' rather than dropped or silently missing.
    region_map = {}
    if len(onsite):
        region_map = onsite.groupby('Origin DC')['Region'].agg(lambda s: s.mode().iat[0]).to_dict()
    all_dcs = set(trips['Origin DC'].dropna().unique())
    if len(onsite):
        all_dcs |= set(onsite['Origin DC'].dropna().unique())
    if order is not None and len(order):
        all_dcs |= set(order['Origin DC'].dropna().unique())
    return {dc: region_map.get(dc, 'Unknown') for dc in all_dcs}


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
    trip_counts = byday.groupby(['Vendor', 'Origin DC', 'Vehicle Type', 'Date']).agg(
        LT_Trips=('Date', 'size'), Finished_LT=('Finished', 'sum'),
        Empty_LT=('Empty', 'sum'), Work_Hrs=('Work_Hrs', 'sum'),
    ).reset_index()

    if len(onsite):
        onsite_counts = onsite.groupby(['Vendor', 'Origin DC', 'Vehicle Type', 'Date'])['unit_key'].nunique().reset_index(name='Onsited')
    else:
        onsite_counts = pd.DataFrame(columns=['Vendor', 'Origin DC', 'Vehicle Type', 'Date', 'Onsited'])

    merged = trip_counts.merge(onsite_counts, on=['Vendor', 'Origin DC', 'Vehicle Type', 'Date'], how='outer').fillna(0)
    merged['Onsited'] = merged['Onsited'].astype(int)
    for col in ('LT_Trips', 'Finished_LT', 'Empty_LT'):
        merged[col] = merged[col].astype(int)
    merged['Work_Hrs'] = merged['Work_Hrs'].astype(float).round(2)

    def add_ratios(g):
        g['Productivity'] = (g['LT_Trips'] / g['Onsited'].replace(0, float('nan'))).round(2)
        g['Avg_Work_Hrs'] = (g['Work_Hrs'] / g['Finished_LT'].replace(0, float('nan'))).round(1)
        return g

    merged = add_ratios(merged).sort_values('LT_Trips', ascending=False)

    def agg(keys):
        return add_ratios(merged.groupby(keys).agg(
            Onsited=('Onsited', 'sum'), LT_Trips=('LT_Trips', 'sum'), Finished_LT=('Finished_LT', 'sum'),
            Empty_LT=('Empty_LT', 'sum'), Work_Hrs=('Work_Hrs', 'sum'),
        ).reset_index())

    by_vendor = agg(['Vendor']).sort_values('LT_Trips', ascending=False)
    by_dc = agg(['Origin DC']).sort_values('LT_Trips', ascending=False)
    by_date = agg(['Date']).sort_values('Date')
    by_vendor_date = agg(['Vendor', 'Date'])

    total_onsited = int(merged['Onsited'].sum())
    total_trips = int(merged['LT_Trips'].sum())
    overall = round(total_trips / total_onsited, 2) if total_onsited else None
    total_finished = int(merged['Finished_LT'].sum())
    total_empty = int(merged['Empty_LT'].sum())
    avg_hrs = round(merged['Work_Hrs'].sum() / total_finished, 1) if total_finished else None
    print(f'  DR LTs: {total_finished} finished, {total_empty} empty (0 parcels), avg working hrs {avg_hrs}')

    return {
        'dates': window_dates,
        'total_onsited': total_onsited,
        'total_trips': total_trips,
        'overall_productivity': overall,
        'total_finished_lt': total_finished,
        'total_empty_lt': total_empty,
        'avg_work_hrs': avg_hrs,
        'by_vendor': clean_nan(by_vendor.to_dict('records')),
        'by_dc': clean_nan(by_dc.to_dict('records')),
        'by_date': clean_nan(by_date.to_dict('records')),
        'by_vendor_date': clean_nan(by_vendor_date.to_dict('records')),
        'detail': clean_nan(merged.to_dict('records')),
    }


def normalize_cost_type(ct):
    s = str(ct or '').strip()
    if not s:
        return 'Unknown'
    return 'In-House' if s.lower() == 'in-house' else s


def build_empty_lt(all_lts, rel_legs):
    # Leg-level view of finished LTs of every cost type: each FM -> SOC and SOC -> LM leg,
    # credited to its SOC, dated by its LT's first-leg date so it follows the Date Window.
    # Only SOC/direction pairs with at least one empty leg are shipped - the others never
    # appear in the section.
    fin = all_lts[all_lts['Finished'].fillna(False).astype(bool)][
        ['trip_number', 'Date', 'agency_name', 'vehicle_type_name', 'cost_type', 'trip_route', 'Work_Hrs']]
    legs = rel_legs.merge(fin, on='trip_number')
    legs['cost_type'] = legs['cost_type'].map(normalize_cost_type)
    legs['agency_name'] = legs['agency_name'].replace('', 'Unknown')
    hot = set(map(tuple, legs.loc[legs['Empty'], ['SOC', 'Direction']].drop_duplicates().values))
    shown = legs[[(s, d) in hot for s, d in zip(legs['SOC'], legs['Direction'])]]
    g = shown.groupby(['Date', 'cost_type', 'agency_name', 'SOC', 'Direction']).agg(
        Legs=('Empty', 'size'), Empty=('Empty', 'sum')).reset_index()

    # Repeated names (routes, stations, vendors...) are sent once in `strings` and referenced
    # by index - otherwise the few thousand empty legs alone add ~800 KB to the page.
    pool = {}
    ix = lambda s: pool.setdefault(s, len(pool))
    empty_legs = legs[legs['Empty']].sort_values(['Date', 'SOC'], ascending=[False, True])
    items = [[r['Date'], r['trip_number'], ix(r['cost_type']), ix(r['agency_name']), ix(r['vehicle_type_name']),
              ix(r['Direction']), ix(r['SOC']), ix(r['Hub']), ix(r['trip_route']),
              None if pd.isna(r['Work_Hrs']) else round(float(r['Work_Hrs']), 1)]
             for r in empty_legs.to_dict('records')]
    rows = [[d, ix(c), ix(v), ix(s), ix(di), int(n), int(e)] for d, c, v, s, di, n, e in g.itertuples(index=False)]
    print(f'  Empty-leg section: {len(legs)} FM->SOC / SOC->LM legs (all cost types), {len(items)} empty, '
          f'{len(hot)} SOC/direction pairs, {len(rows)} rows')
    return {
        'strings': list(pool),
        'rows': rows,    # Date, Cost Type*, Vendor*, SOC*, Direction*, Legs, Empty   (* = index into strings)
        'legs': items,   # Date, LT, Cost Type*, Vendor*, Vehicle Type*, Direction*, SOC*, Hub*, Route*, Work_Hrs
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
        total_onsited = onsite.groupby(['Origin DC', 'Bucket'])['unit_key'].nunique().reset_index(name='Onsited').rename(columns={'Bucket': 'Vehicle Type'})
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
        daily_onsited = onsite.groupby(['Origin DC', 'Bucket', 'Date'])['unit_key'].nunique().reset_index(name='Onsited').rename(columns={'Bucket': 'Vehicle Type'})
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

    print('Fetching trip trackers (SOC-LM/FM-SOC/SOC-SOC)...')
    trips, window_dates, socsoc_ok, rel_legs, all_lts = build_trips_first_leg(gc)
    print(f'  Date range: {window_dates[0]} to {window_dates[-1]} ({len(window_dates)} days)')

    print('Fetching onsite registrations...')
    onsite = build_onsite(gc, window_dates)

    print('Fetching order sheet...')
    order = build_order(gc)

    print('Computing aggregates...')
    dashboard = build_dashboard_data(trips, window_dates)
    productivity = build_productivity(trips, onsite, window_dates)
    ordered_vs_onsite = build_ordered_vs_onsite(order, onsite, window_dates)
    dc_regions = build_dc_regions(trips, onsite, order)

    raw = dict(dashboard)
    raw['productivity'] = productivity
    raw['empty_lt'] = build_empty_lt(all_lts, rel_legs)
    raw['ordered_vs_onsite'] = ordered_vs_onsite['rows']
    raw['ordered_vs_onsite_daily'] = ordered_vs_onsite['daily']
    raw['dc_regions'] = dc_regions
    raw['generated_at'] = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
    raw['order_sheet_available'] = order is not None

    errors = []
    if order is None:
        errors.append({
            'source': 'DR order sheet',
            'message': 'Could not be read (likely a sharing/permission issue for the service account). '
                       'Ordered-vs-onsited figures are unavailable; onsited-only figures elsewhere are unaffected.',
        })
    if not socsoc_ok:
        errors.append({
            'source': 'SOC-SOC trip tracker',
            'message': 'Could not be read (likely a sharing/permission issue for the service account). '
                       'SOC-SOC trips are missing, so DR productivity is understated where units ran SOC-SOC.',
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
