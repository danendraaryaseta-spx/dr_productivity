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
    'SOC_SOC': '11CidDsjDqCWxZ0u9AWXDm2gyNQeWi8vKd_lOInv6_lU',
}

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
            'total_unloaded': r[9],
            'trip_std': r[10],
            'trip_atd': r[11],
            'trip_ata': r[13],
            'trip_status': r[15],
            'dest_ata': r[17],
            'agency_name': r[18],
            'source_sheet': source_label,
        })
    by_day = sum(1 for r in rows if r['cost_type'] == 'By Day')
    print(f'  {source_label}: {len(rows)} rows with a trip_number, {by_day} By Day (cost_type column: {"yes" if ct_idx is not None else "MISSING"})')
    return rows


# ESB (Eco Super Bulky) SOCs run on units borrowed from regular DCs, so planners need to be able
# to take them out of the numbers. List from the planning team; two are named "EDC".
ESB_DCS = ['Banyumas 2 DC', 'Cakung 5 DC', 'Cirebon 2 DC', 'Jember 2 EDC', 'Madiun 3 DC',
           'Semarang 4 DC', 'Sidoarjo 2 DC', 'Tegal 3 EDC']


def is_dc(name):
    return str(name).strip().endswith((' DC', ' RDC', ' EDC'))


def station_kind(name):
    s = str(name).strip()
    if is_dc(s):
        return 'SOC'
    if 'first mile' in s.lower():
        return 'FM'
    if s.lower().endswith('hub'):
        return 'LM'
    return 'Other'


def to_num(col):
    return pd.to_numeric(col.astype(str).str.replace(',', ''), errors='coerce').fillna(0)


def build_lt_stats(df):
    # An LT can sit in both the SOC-LM and FM-SOC trackers, so legs are deduped first.
    # "Not Depart" legs never ran (no ATD/ATA) and are ignored throughout.
    legs = df.drop_duplicates(['trip_number', 'origin_station', 'dest_station', 'trip_std']).copy()
    legs['std_dt'] = pd.to_datetime(legs['trip_std'], errors='coerce')
    legs['ata'] = pd.to_datetime(legs['trip_ata'], errors='coerce')
    legs['dest_ata_dt'] = pd.to_datetime(legs['dest_ata'], errors='coerce')
    legs['departed'] = legs['trip_status'].astype(str).str.strip().str.lower() != 'not depart'
    legs = legs.sort_values(['trip_number', 'std_dt', 'atd_dt'])

    # Finished (for working hours / the DR Empty LT card): the last scheduled leg ran and every
    # departed leg has arrived - an LT still on the road would look too short.
    last_ran = legs.groupby('trip_number')['departed'].last()
    d = legs[legs['departed']].copy()
    g = d.groupby('trip_number')
    hrs = (g['dest_ata_dt'].max() - g['ata'].min()).dt.total_seconds() / 3600
    finished = last_ran.reindex(hrs.index, fill_value=False) & (g['dest_ata_dt'].count() == g.size()) & (hrs >= 0)

    # Empty only looks at legs that should carry parcels; SOC -> FM and LM -> SOC normally run
    # empty. total_loaded / total_unloaded belong to the leg's ORIGIN station; blank counts as 0.
    # A "run" is the stretch of legs from leaving a SOC until the next SOC.
    #   FM -> SOC is empty when the truck arrives carrying nothing: 0 loaded at every hub since it
    #   last left a SOC (parcels loaded AT a SOC are deliveries dropped at LM hubs).
    #   SOC -> LM is empty when 0 was loaded at the SOC and 0 was unloaded at every hub on the run.
    #   Unloaded alone isn't enough: hubs often leave it blank even on full delivery runs. Not
    #   judged until the run's last leg has arrived, since unloaded is filled in after arrival.
    kinds = {s: station_kind(s) for s in pd.unique(pd.concat([d['origin_station'], d['dest_station']]))}
    o_kind, d_kind = d['origin_station'].map(kinds), d['dest_station'].map(kinds)
    loaded, unloaded = to_num(d['total_loaded']), to_num(d['total_unloaded'])
    from_soc = o_kind == 'SOC'
    run = [d['trip_number'], from_soc.groupby(d['trip_number']).cumsum()]
    on_board = loaded.where(~from_soc, 0).groupby(run).cumsum()
    run_unloaded = unloaded.where(~from_soc, 0).groupby(run).transform('sum')
    run_done = d['dest_ata_dt'].notna().groupby(run).transform('last')
    inbound = (o_kind == 'FM') & (d_kind == 'SOC')
    outbound = (o_kind == 'SOC') & (d_kind == 'LM') & run_done
    empty_leg = (inbound & (on_board == 0)) | (outbound & (loaded == 0) & (run_unloaded == 0))

    judged = inbound | outbound
    n_judged = judged.groupby(d['trip_number']).sum()
    n_empty = empty_leg.groupby(d['trip_number']).sum()
    stats = pd.DataFrame({
        'Finished': finished,
        'Empty': finished & (n_judged > 0) & (n_empty == n_judged),
        'Work_Hrs': hrs.where(finished),
    })
    rel = d.loc[judged, ['trip_number', 'origin_station', 'dest_station', 'trip_std', 'trip_atd', 'dest_ata']].copy()
    rel['Direction'] = inbound[judged].map({True: 'FM → SOC', False: 'SOC → LM'})
    rel['SOC'] = rel['dest_station'].where(rel['Direction'] == 'FM → SOC', rel['origin_station'])
    rel['Hub'] = rel['origin_station'].where(rel['Direction'] == 'FM → SOC', rel['dest_station'])
    rel['Empty'] = empty_leg[judged]
    return stats, rel[['trip_number', 'Direction', 'SOC', 'Hub', 'Empty', 'trip_std', 'trip_atd', 'dest_ata']]


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
    first_leg = all_lts[all_lts['origin_station'].map(is_dc)]

    # No fixed rolling window - the full date range actually present in the
    # trackers is exposed to the client, which lets the user pick any sub-range
    # via the date-window control while still showing daily granularity.
    all_dates = sorted(first_leg['Date'].dropna().unique().tolist())
    # The trackers keep a rolling ~17 days, so the oldest day is cut off part-way: trucks are
    # onsite but most of its LTs are already gone, which reads as near-zero productivity.
    per_day = first_leg.groupby('Date').size()
    while len(all_dates) > 1 and per_day.get(all_dates[0], 0) < 0.5 * per_day.median():
        print(f'  Dropping {all_dates[0]}: only {per_day.get(all_dates[0], 0)} LTs (tracker window cut it off)')
        all_dates = all_dates[1:]

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


def build_dc_regions(trips, onsite, empty_socs):
    # Region only exists on the onsite sheet, not the trip trackers - build a
    # DC -> Region map from onsite data and reuse it everywhere a DC shows up
    # (trips, onsite, and the SOCs in the empty-leg section). A DC never seen
    # onsite has no known region and is bucketed 'Unknown' rather than dropped.
    region_map = {}
    if len(onsite):
        region_map = onsite.groupby('Origin DC')['Region'].agg(lambda s: s.mode().iat[0]).to_dict()
    all_dcs = set(trips['Origin DC'].dropna().unique()) | set(empty_socs)
    if len(onsite):
        all_dcs |= set(onsite['Origin DC'].dropna().unique())
    return {dc: region_map.get(dc, 'Unknown') for dc in all_dcs}


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


def build_empty_lt(all_lts, rel_legs, dc_regions):
    # Leg-level view of every cost type: each departed FM -> SOC and SOC -> LM leg (see
    # build_lt_stats), dated by its LT's first-leg date so it follows the Date Window.
    # Two roll-ups, each shipped only where there's at least one empty leg (the rest never
    # appear in the section): per SOC + direction, and per route (the LT's trip_route).
    lts = all_lts[['trip_number', 'Date', 'cost_type', 'trip_route', 'agency_name', 'vehicle_type_name']]
    legs = rel_legs.merge(lts, on='trip_number')
    legs['cost_type'] = legs['cost_type'].map(normalize_cost_type)
    legs['agency_name'] = legs['agency_name'].replace('', 'Unknown')

    def rollup(key):
        hot = set(map(tuple, legs.loc[legs['Empty'], [key, 'Direction']].drop_duplicates().values))
        shown = legs[[(k, d) in hot for k, d in zip(legs[key], legs['Direction'])]]
        return shown.groupby(['Date', 'cost_type', key, 'Direction']).agg(
            Legs=('Empty', 'size'), Empty=('Empty', 'sum')).reset_index()

    # Repeated names are sent once in `strings` and referenced by index to keep the page small.
    pool = {}
    ix = lambda s: pool.setdefault(s, len(pool))
    pack = lambda g: [[d, ix(c), ix(k), ix(di), int(n), int(e)] for d, c, k, di, n, e in g.itertuples(index=False)]
    socs, routes = pack(rollup('SOC')), pack(rollup('trip_route'))
    legs['Region'] = legs['SOC'].map(dc_regions).fillna('Unknown')
    legs['ESB'] = legs['SOC'].isin(ESB_DCS)
    totals = [[d, ix(c), ix(rg), ix(di), int(esb), int(n), int(e)] for d, c, rg, di, esb, n, e in
              legs.groupby(['Date', 'cost_type', 'Region', 'Direction', 'ESB']).agg(
                  Legs=('Empty', 'size'), Empty=('Empty', 'sum')).reset_index().itertuples(index=False)]
    # One row per empty leg, for the planners' CSV download.
    empty = legs[legs['Empty']].fillna('').sort_values(['Date', 'SOC', 'trip_std'], ascending=[False, True, True])
    detail = [[r['Date'], r['trip_number'], ix(r['cost_type']), ix(r['agency_name']), ix(r['vehicle_type_name']),
               ix(r['Direction']), ix(r['SOC']), ix(r['Hub']), ix(r['trip_route']),
               r['trip_std'], r['trip_atd'], r['dest_ata']] for r in empty.to_dict('records')]
    print(f'  Empty-leg section: {len(legs)} FM->SOC / SOC->LM legs (all cost types), {len(detail)} empty, '
          f'{len(socs)} SOC rows, {len(routes)} route rows')
    return {
        'strings': list(pool),
        'socs': socs,      # Date, Cost Type*, SOC*, Direction*, Legs, Empty   (* = index into strings)
        'routes': routes,  # Date, Cost Type*, Route*, Direction*, Legs, Empty
        'totals': totals,  # Date, Cost Type*, Region*, Direction*, SOC is ESB (0/1), Legs, Empty - every leg
        # Date, LT, Cost Type*, Vendor*, Vehicle Type*, Direction*, SOC*, Hub*, Route*, Leg STD, Leg ATD, Leg arrival
        'detail': detail,
    }


def main():
    print('Connecting to Google Sheets...')
    gc = get_client()

    print('Fetching trip trackers (SOC-LM/FM-SOC/SOC-SOC)...')
    trips, window_dates, socsoc_ok, rel_legs, all_lts = build_trips_first_leg(gc)
    print(f'  Date range: {window_dates[0]} to {window_dates[-1]} ({len(window_dates)} days)')

    print('Fetching onsite registrations...')
    onsite = build_onsite(gc, window_dates)

    print('Computing aggregates...')
    dashboard = build_dashboard_data(trips, window_dates)
    productivity = build_productivity(trips, onsite, window_dates)
    dc_regions = build_dc_regions(trips, onsite, rel_legs['SOC'].unique())
    empty_lt = build_empty_lt(all_lts, rel_legs, dc_regions)

    raw = dict(dashboard)
    raw['productivity'] = productivity
    raw['empty_lt'] = empty_lt
    raw['dc_regions'] = dc_regions
    raw['esb_dcs'] = ESB_DCS
    raw['generated_at'] = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')

    errors = []
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
