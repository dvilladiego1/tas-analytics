#!/usr/bin/env python3
"""
TAS Weekly Consolidated Deck Generator — W12 (updated with data through Mar 22)
Generates: ~/Downloads/TAS_Weekly_Consolidated_W12_2026.html
"""

import csv
from datetime import datetime, date, timedelta
from collections import defaultdict
import json
import os

# ─── CONFIG ──────────────────────────────────────────────────────────────────
CSV7 = os.path.expanduser("~/Downloads/Summary MKP _ TAS & BULK - Snapshot_Diario (7).csv")
CSV8 = os.path.expanduser("~/Downloads/Summary MKP _ TAS & BULK - Snapshot_Diario (8).csv")
PEOPLE_CSV = os.path.expanduser("~/Downloads/People Cost - Liquidity Marketplace - Sheet1.csv")
BP_CSV = os.path.expanduser("~/Downloads/Liquidity BP 2026 - Liquidity.csv")
OUTPUT = os.path.expanduser("~/Downloads/TAS_Weekly_Consolidated_W13_2026.html")

# Weekly periods (Mon-Sat)
WEEKS = {
    'w-5': (date(2026,2,16), date(2026,2,21)),
    'w-4': (date(2026,2,23), date(2026,2,28)),
    'w-3': (date(2026,3,2),  date(2026,3,7)),
    'w-2': (date(2026,3,9),  date(2026,3,14)),
    'w-1': (date(2026,3,16), date(2026,3,21)),
}
WEEK_LABELS = {
    'w-5': '16-21<br/>Feb',
    'w-4': '23-28<br/>Feb',
    'w-3': '2-7<br/>Mar',
    'w-2': '9-14<br/>Mar',
    'w-1': '16-21<br/>Mar',
}

# Monthly periods
MONTHS = ['Oct','Nov','Dic','Ene','Feb','Mar']
MONTH_RANGES = {
    'Oct': (date(2025,10,1), date(2025,10,31)),
    'Nov': (date(2025,11,1), date(2025,11,30)),
    'Dic': (date(2025,12,1), date(2025,12,31)),
    'Ene': (date(2026,1,1),  date(2026,1,31)),
    'Feb': (date(2026,2,1),  date(2026,2,28)),
    'Mar': (date(2026,3,1),  date(2026,3,31)),
}

# MTD day 20 for Feb vs Mar comparison
FEB_MTD_END = date(2026,2,20)
MAR_MTD_END = date(2026,3,21)  # Latest data: purchases through Sat Mar 21

# Business days in March 2026 = 25 (Mon-Sat, excluding Mar 16 holiday)
MAR_BIZ_DAYS = 25
# Elapsed business days through Mar 22 (count Mon-Sat from Mar 1-22 minus Mar 16 holiday)
# Mar 2,3,4,5,6,7 (6) + Mar 9,10,11,12,13,14 (6) + Mar 17,18,19,20,21 (5) = 17
MAR_ELAPSED_BIZ = 17

# Holidays
HOLIDAYS = {date(2026,3,16)}  # Benito Juárez

# BP Targets for March
BP_TARGETS = {
    'Quotes': 1909,
    'Made': 1013,
    'Approved': 822,
    'Purchased': 370,
}

# FTEs per month (White Label)
FTE_BY_MONTH = {'Oct': 35, 'Nov': 37, 'Dic': 46, 'Ene': 46, 'Feb': 59, 'Mar': 59}

# Group name normalization
GRUPO_MAP = {
    'GRUPO CONTINENTAL': 'CONTINENTAL',
    'GRUPO PREMIER': 'PREMIER',
    'GRUPO ANDRADE': 'ANDRADE',
    'GRUPO PLASENCIA': 'PLASENCIA',
    'GRUPO ISMO TLALNEPANTLA': 'ISMO',
    'GRUPO GP AUTO': 'GP AUTO',
    'GRUPO AUTOPOLIS': 'AUTOPOLIS',
    'GRUPO SONI': 'SONI',
    'GRUPO ISMO AGUASCALIENTES': 'ISMO',
    'GRUPO TOLLOCAN': 'TOLLOCAN',
    'GRUPO POTOSINA': 'POTOSINA',
    'GRUPO ISMO LEON': 'ISMO',
    'GRUPO TORRES CORZO': 'TORRES CORZO',
    'GRUPO WECARS': 'WECARS',
    'GRUPO ISMO': 'ISMO',
    'GRUPO MEGA': 'MEGA',
    'GRUPO MISOL': 'MISOL',
    'ALIADO COAPA': 'COAPA',
    'ALIADO LEON': 'LEON',
    'ALIADO TLAHUAC': 'TLAHUAC',
    'ALIADO CONTINENTAL': 'CONTINENTAL',
    'ALIADO TOLLOCAN': 'TOLLOCAN',
    'ALIADO MEGA': 'MEGA',
    'ALIADO TORRES CORZO': 'TORRES CORZO',
    'ALIADO ANDRADE': 'ANDRADE',
    'ALIADO ISMO': 'ISMO',
    'ALIADO SONI': 'SONI',
    'ALIADO POTOSINA': 'POTOSINA',
    'ALIADO PREMIER': 'PREMIER',
    'ALIADO WECARS': 'WECARS',
    'ALIADO MISOL': 'MISOL',
    'ALIADO AGUASCALIENTES': 'ISMO',
}

# Region mapping
REGION_MAP = {
    'PREMIER': 'GDL', 'PLASENCIA': 'GDL',
    'LEON': 'QRO', 'POTOSINA': 'QRO',
    'SONI': 'CDMX',
    'WECARS': 'MTY', 'MISOL': 'MTY', 'AUTOPOLIS': 'MTY',
}

# Groups to exclude (noise / non-real TAS groups)
EXCLUDE_GROUPS = {
    'ALBACAR', 'AUTOKLIC', 'B2B', 'MENA',
    'NO ES TAS | Oportunidad de Compra', 'Rechazo 7 días',
}

def parse_date(s):
    """Parse dd/mm/yyyy date string."""
    if not s or '/' not in s:
        return None
    parts = s.strip().split('/')
    if len(parts) != 3:
        return None
    try:
        return date(int(parts[2]), int(parts[1]), int(parts[0]))
    except (ValueError, IndexError):
        return None

def is_biz_day(d):
    """Mon-Sat excluding holidays."""
    return d.weekday() < 6 and d not in HOLIDAYS

def biz_days_in_range(start, end):
    """Count business days (Mon-Sat excl holidays) in range inclusive."""
    count = 0
    d = start
    while d <= end:
        if is_biz_day(d):
            count += 1
        d += timedelta(days=1)
    return count

def in_range(d, start, end):
    return start <= d <= end

# ─── LOAD DATA ───────────────────────────────────────────────────────────────
print("Loading CSV data...")

records = []

def load_csv(path, label):
    rows = []
    with open(path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        headers = next(reader)
        for row in reader:
            if len(row) < 55:
                continue
            grupo_raw = row[30].strip()
            celula_raw = row[31].strip()
            channel_raw = row[36].strip()

            if celula_raw == 'TEST-TEST':
                continue
            if channel_raw not in ('TAS', 'ALIADO'):
                continue

            # Overrides per reference categories file
            if 'POTOSINA' in grupo_raw:
                channel_raw = 'TAS'
            if celula_raw == 'ANDRADE AEROPUERTO' and channel_raw == 'ALIADO':
                channel_raw = 'TAS'
            if celula_raw == 'SONI PACHUCA' and channel_raw == 'ALIADO':
                channel_raw = 'TAS'

            fecha_creacion = parse_date(row[2])
            purchase_date = parse_date(row[46])

            # Funnel flags (columns 48=scheduled.1, 49=made.1, 50=approved, 54=purchased)
            scheduled = row[48].strip() if len(row) > 48 else ''
            made = row[49].strip() if len(row) > 49 else ''
            approved = row[50].strip() if len(row) > 50 else ''
            purchased = row[54].strip() if len(row) > 54 else ''

            is_quote = scheduled not in ('', '0', 'FALSE', 'false')
            is_made = made not in ('', '0', 'FALSE', 'false')
            is_approved = approved not in ('', '0', 'FALSE', 'false')
            is_purchased = purchased not in ('', '0', 'FALSE', 'false')

            # Normalize grupo
            grupo_norm = GRUPO_MAP.get(grupo_raw, grupo_raw.replace('GRUPO ', '').replace('ALIADO ', ''))

            rows.append({
                'grupo_raw': grupo_raw,
                'grupo': grupo_norm,
                'celula': celula_raw,
                'channel': channel_raw,
                'fecha_creacion': fecha_creacion,
                'purchase_date': purchase_date,
                'is_quote': is_quote,
                'is_made': is_made,
                'is_approved': is_approved,
                'is_purchased': is_purchased,
                'approved_date': parse_date(row[3]) if len(row) > 3 else None,  # Fecha Inspección as proxy
            })
    return rows

# Load CSV7 (complete historical) and CSV8 (latest with purchases through Mar 21)
csv7_rows = load_csv(CSV7, 'CSV7')
csv8_rows = load_csv(CSV8, 'CSV8')

# CSV8 has incomplete Jan data but UPDATED purchase flags through Mar 21.
# Strategy:
#   - Q/M/A records: CSV7 for all months + CSV8 for Mar 21-22 quotes
#   - Purchase records: CSV7 for Oct-Jan + CSV8 for Feb-Mar (latest purchase data)
print("Merging data...")

# Set 1: for Q/M/A (grouped by fecha_creacion)
qma_records = list(csv7_rows)
mar21 = date(2026, 3, 21)
mar22 = date(2026, 3, 22)
added_qma = 0
for r in csv8_rows:
    fc = r['fecha_creacion']
    if fc and fc in (mar21, mar22):
        qma_records.append(r)
        added_qma += 1
print(f"Q/M/A records: {len(qma_records)} (added {added_qma} from CSV8 for Mar 21-22)")

# Set 2: for purchases (grouped by purchase_date)
# Use CSV7 for Oct-Jan purchases, CSV8 for Feb-Mar purchases
purchase_records = []
for r in csv7_rows:
    if r['is_purchased'] and r['purchase_date']:
        if r['purchase_date'] < date(2026, 2, 1):  # Oct-Jan
            purchase_records.append(r)
for r in csv8_rows:
    if r['is_purchased'] and r['purchase_date']:
        if r['purchase_date'] >= date(2026, 2, 1):  # Feb-Mar from CSV8
            purchase_records.append(r)
print(f"Purchase records: {len(purchase_records)} (Feb-Mar from CSV8)")

# Combined for aggregate function — tag records with source
all_records = list(qma_records)  # All records for Q/M/A
# Note: purchase counting will use purchase_records separately

print(f"Total records: {len(all_records)}")

# ─── AGGREGATE DATA ─────────────────────────────────────────────────────────
def aggregate(records, channel_filter=None, grupo_filter=None, p_records=None):
    """Aggregate funnel metrics for given filters.
    records: used for Q/M/A (by fecha_creacion)
    p_records: used for purchases (by purchase_date) — defaults to records if None
    """
    filtered = [r for r in records if r['grupo'] not in EXCLUDE_GROUPS]
    if channel_filter:
        filtered = [r for r in filtered if r['channel'] == channel_filter]
    if grupo_filter:
        filtered = [r for r in filtered if r['grupo'] == grupo_filter]

    # Separate purchase records (use p_records if provided)
    if p_records is not None:
        p_filtered = [r for r in p_records if r['grupo'] not in EXCLUDE_GROUPS]
        if channel_filter:
            p_filtered = [r for r in p_filtered if r['channel'] == channel_filter]
        if grupo_filter:
            p_filtered = [r for r in p_filtered if r['grupo'] == grupo_filter]
    else:
        p_filtered = filtered

    result = {}

    # Weekly metrics (Q/M/A grouped by fecha_creacion, P grouped by purchase_date)
    for wk_name in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk_name]
        q = sum(1 for r in filtered if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
        m = sum(1 for r in filtered if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
        a = sum(1 for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
        p = sum(1 for r in p_filtered if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], ws, we))
        result[wk_name] = [q, m, a, p]

    # Monthly metrics
    for mo in MONTHS:
        ms, me = MONTH_RANGES[mo]
        q = sum(1 for r in filtered if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
        m = sum(1 for r in filtered if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
        a = sum(1 for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
        p = sum(1 for r in p_filtered if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], ms, me))
        result[mo] = [q, m, a, p]

    # MTD comparisons (Feb MTD day 20 vs Mar MTD latest)
    feb_mtd_start = date(2026,2,1)
    q_feb_mtd = sum(1 for r in filtered if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], feb_mtd_start, FEB_MTD_END))
    m_feb_mtd = sum(1 for r in filtered if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], feb_mtd_start, FEB_MTD_END))
    a_feb_mtd = sum(1 for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], feb_mtd_start, FEB_MTD_END))
    p_feb_mtd = sum(1 for r in p_filtered if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], feb_mtd_start, FEB_MTD_END))
    result['Feb MTD'] = [q_feb_mtd, m_feb_mtd, a_feb_mtd, p_feb_mtd]

    mar_mtd_start = date(2026,3,1)
    q_mar_mtd = sum(1 for r in filtered if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], mar_mtd_start, MAR_MTD_END))
    m_mar_mtd = sum(1 for r in filtered if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], mar_mtd_start, MAR_MTD_END))
    a_mar_mtd = sum(1 for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], mar_mtd_start, MAR_MTD_END))
    p_mar_mtd = sum(1 for r in p_filtered if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], mar_mtd_start, MAR_MTD_END))
    result['Mar MTD'] = [q_mar_mtd, m_mar_mtd, a_mar_mtd, p_mar_mtd]

    # A→P% cohort cerrada: of approvals ≥7 days old, % with purchase within 7d of approval
    # For weekly: approvals in that week whose approval date is ≥7 days ago
    # This is complex - approximate using approved rows where fecha_creacion is ≥7d before end of period
    for wk_name in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk_name]
        cutoff = MAR_MTD_END - timedelta(days=7)
        approved_in_week = [r for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we) and r['fecha_creacion'] <= cutoff]
        purchased_within_7d = sum(1 for r in approved_in_week if r['is_purchased'] and r['purchase_date'] and (r['purchase_date'] - r['fecha_creacion']).days <= 7)
        total_closed = len(approved_in_week)
        result[f'{wk_name}_cohort'] = (purchased_within_7d, total_closed)

    # Monthly cohort
    for mo in MONTHS:
        ms, me = MONTH_RANGES[mo]
        cutoff = MAR_MTD_END - timedelta(days=7)
        approved_in_mo = [r for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me) and r['fecha_creacion'] <= cutoff]
        purchased_within_7d = sum(1 for r in approved_in_mo if r['is_purchased'] and r['purchase_date'] and (r['purchase_date'] - r['fecha_creacion']).days <= 7)
        total_closed = len(approved_in_mo)
        result[f'{mo}_cohort'] = (purchased_within_7d, total_closed)

    # MTD cohort
    for mtd_key, (mtd_start, mtd_end) in [('Feb MTD', (feb_mtd_start, FEB_MTD_END)), ('Mar MTD', (mar_mtd_start, MAR_MTD_END))]:
        cutoff = MAR_MTD_END - timedelta(days=7)
        approved_in_mtd = [r for r in filtered if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], mtd_start, mtd_end) and r['fecha_creacion'] <= cutoff]
        purchased_within_7d = sum(1 for r in approved_in_mtd if r['is_purchased'] and r['purchase_date'] and (r['purchase_date'] - r['fecha_creacion']).days <= 7)
        total_closed = len(approved_in_mtd)
        result[f'{mtd_key}_cohort'] = (purchased_within_7d, total_closed)

    return result

# Discover all unique groups per channel (excluding noise)
groups_tas = sorted(set(r['grupo'] for r in all_records if r['channel'] == 'TAS' and r['grupo'] not in EXCLUDE_GROUPS))
groups_ali = sorted(set(r['grupo'] for r in all_records if r['channel'] == 'ALIADO' and r['grupo'] not in EXCLUDE_GROUPS))

print(f"TAS groups: {groups_tas}")
print(f"ALIADO groups: {groups_ali}")

# Compute aggregates (pass purchase_records for P counting)
print("Computing aggregates...")
agg_total = aggregate(all_records, p_records=purchase_records)
agg_tas_total = aggregate(all_records, channel_filter='TAS', p_records=purchase_records)
agg_ali_total = aggregate(all_records, channel_filter='ALIADO', p_records=purchase_records)

agg_tas = {}
for g in groups_tas:
    agg_tas[g] = aggregate(all_records, channel_filter='TAS', grupo_filter=g, p_records=purchase_records)

agg_ali = {}
for g in groups_ali:
    agg_ali[g] = aggregate(all_records, channel_filter='ALIADO', grupo_filter=g, p_records=purchase_records)

# Count active cells per period
def count_active_cells(records, channel, period_key, period_range):
    """Count cells with ≥1 quote in period."""
    cells = set()
    ms, me = period_range
    for r in records:
        if r['channel'] == channel and r['grupo'] not in EXCLUDE_GROUPS and r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me):
            cells.add(r['celula'])
    return len(cells)

# ─── HTML GENERATION ─────────────────────────────────────────────────────────
print("Generating HTML...")

def fmt(v):
    """Format integer with comma separator."""
    if v == 0:
        return '0'
    return f'{v:,}'

def fmtR(v):
    """Format rate as percentage."""
    if v == 0:
        return '0.0%'
    return f'{v:.1f}%'

def rate(num, den):
    if den == 0:
        return 0
    return num / den * 100

def wow_pct(w1, w2):
    """WoW change for volume metrics."""
    if w2 == 0:
        if w1 > 0:
            return 'NEW'
        return '—'
    return f'{(w1-w2)/w2*100:+.1f}%'

def wow_pp(w1, w2):
    """WoW change for rate metrics (percentage points)."""
    if w2 == 0 and w1 == 0:
        return '—'
    if w2 == 0 and w1 > 0:
        return 'NEW'
    diff = w1 - w2
    return f'{diff:+.1f}pp'

def mom_pct(cur, prev):
    """MoM change."""
    if prev == 0:
        if cur > 0:
            return '+0.0%'
        return '+0.0%'
    return f'{(cur-prev)/prev*100:+.1f}%'

def color_class(val_str):
    """Return CSS class for colored delta."""
    if 'NEW' in str(val_str):
        return 'color:#F59E0B;font-weight:700'
    if '—' in str(val_str):
        return 'color:#94A3B8;font-weight:700'
    try:
        if val_str.startswith('+') and float(val_str.replace('%','').replace('pp','')) > 0:
            return 'color:#22C55E;font-weight:700'
        elif val_str.startswith('-'):
            return 'color:#EF4444;font-weight:700'
        else:
            return 'color:#F59E0B;font-weight:700'
    except:
        return 'color:#94A3B8;font-weight:700'

def mom_color_class(val_str):
    if not val_str or val_str == '—':
        return 'neu'
    try:
        v = float(val_str.replace('%','').replace('+',''))
        if v > 0:
            return 'pos'
        elif v < 0:
            return 'neg'
        return 'neu'
    except:
        return 'neu'

def best_month(agg, metric_idx):
    """Find best month value for a metric."""
    best_val = 0
    best_label = '—'
    for mo in MONTHS:
        v = agg[mo][metric_idx]
        if v > best_val:
            best_val = v
            best_label = f'{mo} ({fmt(v)})'
    return best_label

def best_month_rate(agg, num_idx, den_idx):
    """Find best month rate."""
    best_val = 0
    best_label = '—'
    for mo in MONTHS:
        den = agg[mo][den_idx]
        if den > 0:
            v = agg[mo][num_idx] / den * 100
            if v > best_val:
                best_val = v
                best_label = f'{mo} ({v:.1f}%)'
    return best_label

def rr_7d(agg, metric_idx, is_purchase=False):
    """Run rate 7d = last 7 days value / 7 * biz days remaining + MTD."""
    # Last 7 calendar days from Mar 22
    last7_start = date(2026,3,16)
    last7_end = date(2026,3,22)
    # For simplicity, use w-1 data (Mar 16-21) as proxy for last 7 days
    w1_val = agg['w-1'][metric_idx]
    biz_in_w1 = biz_days_in_range(date(2026,3,16), date(2026,3,21))  # 5 days (holiday excluded)
    daily_rate = w1_val / biz_in_w1 if biz_in_w1 > 0 else 0
    projected = round(daily_rate * MAR_BIZ_DAYS)
    return projected

def rr_seas(agg, metric_idx):
    """Run rate seasonality = MTD + weighted avg of last 5 weeks × remaining days."""
    mtd = agg['Mar MTD'][metric_idx]
    # Weighted average of weekly daily rates
    week_rates = []
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk]
        bd = biz_days_in_range(ws, we)
        if bd > 0:
            week_rates.append(agg[wk][metric_idx] / bd)
    if not week_rates:
        return mtd
    # Weighted: more recent weeks get more weight
    weights = [1,2,3,4,5][:len(week_rates)]
    weighted_avg = sum(r*w for r,w in zip(week_rates, weights)) / sum(weights)
    remaining_days = MAR_BIZ_DAYS - MAR_ELAPSED_BIZ
    projected = round(mtd + weighted_avg * remaining_days)
    return projected

def generate_group_rows(agg, grupo_name, is_consolidated=False, channel='TAS'):
    """Generate 8 data rows for a group."""
    rows = []
    metrics = [
        ('Quotes', 0, 'dr'),
        ('Made', 1, 'dr'),
        ('Approved', 2, 'dr'),
        ('Purchased', 3, 'dr bl'),
        ('Q→M%', (1,0), 'dr'),
        ('M→A%', (2,1), 'dr'),
        ('A→P% (purchase)', (3,2), 'dr'),
        ('A→P% (cohort cerrada)', 'cohort', 'dr pu'),
        ('Q→P%', (3,0), 'dr bl qp'),
    ]

    for label, idx, css in metrics:
        row = f'<tr class="{css}"><td>{label}</td>'

        if isinstance(idx, int):
            # Volume metric
            vals = []
            # Weekly columns
            for wk in ['w-5','w-4','w-3','w-2','w-1']:
                v = agg[wk][idx]
                vals.append(v)
                row += f'<td>{fmt(v)}</td>'

            # WoW
            w1, w2 = agg['w-1'][idx], agg['w-2'][idx]
            wow = wow_pct(w1, w2)
            cc = color_class(wow)
            row += f'<td style="background:#FEF9E8;{cc}">{wow}</td>'

            # Monthly
            prev_mo_val = None
            for mo in MONTHS:
                v = agg[mo][idx]
                row += f'<td>{fmt(v)}</td>'

            # Feb MTD, Mar MTD
            feb_mtd = agg['Feb MTD'][idx]
            mar_mtd = agg['Mar MTD'][idx]
            row += f'<td>{fmt(feb_mtd)}</td>'
            row += f'<td>{fmt(mar_mtd)}</td>'

            # MoM
            mom = mom_pct(mar_mtd, feb_mtd)
            mc = mom_color_class(mom)
            row += f'<td class="mc {mc}" style="background:#FEF9E8">{mom}</td>'

            # Best
            row += f'<td>{best_month(agg, idx)}</td>'

            # RR 7d, RR Seas
            rr7 = rr_7d(agg, idx)
            rrs = rr_seas(agg, idx)
            row += f'<td>{fmt(rr7)}</td>'
            row += f'<td>{fmt(rrs)}</td>'

            # Target / %BP (only for consolidated)
            if is_consolidated and label in BP_TARGETS:
                target = BP_TARGETS[label]
                pct_bp = mar_mtd / target * 100 if target > 0 else 0
                bp_cls = 'pos' if pct_bp >= 100 else 'neg'
                row += f'<td style="background:#F3E8FF;font-weight:700">{fmt(target)}</td>'
                row += f'<td class="{bp_cls}" style="font-weight:700">{pct_bp:.1f}%</td>'
            else:
                row += '<td>—</td><td>—</td>'

        elif isinstance(idx, tuple):
            # Rate metric
            num_idx, den_idx = idx

            # Weekly
            for wk in ['w-5','w-4','w-3','w-2','w-1']:
                num, den = agg[wk][num_idx], agg[wk][den_idx]
                v = rate(num, den)
                row += f'<td>{fmtR(v)}</td>'

            # WoW
            w1_num, w1_den = agg['w-1'][num_idx], agg['w-1'][den_idx]
            w2_num, w2_den = agg['w-2'][num_idx], agg['w-2'][den_idx]
            w1r = rate(w1_num, w1_den)
            w2r = rate(w2_num, w2_den)
            wow = wow_pp(w1r, w2r)
            cc = color_class(wow)
            row += f'<td style="background:#FEF9E8;{cc}">{wow}</td>'

            # Monthly
            for mo in MONTHS:
                num, den = agg[mo][num_idx], agg[mo][den_idx]
                v = rate(num, den)
                row += f'<td>{fmtR(v)}</td>'

            # Feb MTD, Mar MTD
            feb_r = rate(agg['Feb MTD'][num_idx], agg['Feb MTD'][den_idx])
            mar_r = rate(agg['Mar MTD'][num_idx], agg['Mar MTD'][den_idx])
            row += f'<td>{fmtR(feb_r)}</td>'
            row += f'<td>{fmtR(mar_r)}</td>'

            # MoM (pp)
            mom = wow_pp(mar_r, feb_r)
            mc = mom_color_class(mom.replace('pp','%'))
            row += f'<td class="mc {mc}" style="background:#FEF9E8">{mom}</td>'

            # Best
            row += f'<td>{best_month_rate(agg, num_idx, den_idx)}</td>'

            # RR rates
            rr7_num = rr_7d(agg, num_idx)
            rr7_den = rr_7d(agg, den_idx)
            rr7_r = rate(rr7_num, rr7_den)
            rrs_num = rr_seas(agg, num_idx)
            rrs_den = rr_seas(agg, den_idx)
            rrs_r = rate(rrs_num, rrs_den)
            row += f'<td>{fmtR(rr7_r)}</td>'
            row += f'<td>{fmtR(rrs_r)}</td>'

            # Target columns
            row += '<td>—</td><td>—</td>'

        elif idx == 'cohort':
            # Cohort cerrada
            for wk in ['w-5','w-4','w-3','w-2','w-1']:
                p_wk, t_wk = agg.get(f'{wk}_cohort', (0,0))
                v = rate(p_wk, t_wk)
                row += f'<td>{fmtR(v)}</td>'

            # WoW
            p1, t1 = agg.get('w-1_cohort', (0,0))
            p2, t2 = agg.get('w-2_cohort', (0,0))
            r1 = rate(p1, t1)
            r2 = rate(p2, t2)
            wow = wow_pp(r1, r2)
            cc = color_class(wow)
            row += f'<td style="background:#FEF9E8;{cc}">{wow}</td>'

            # Monthly
            for mo in MONTHS:
                p_mo, t_mo = agg.get(f'{mo}_cohort', (0,0))
                v = rate(p_mo, t_mo)
                row += f'<td>{fmtR(v)}</td>'

            # MTD
            p_feb, t_feb = agg.get('Feb MTD_cohort', (0,0))
            p_mar, t_mar = agg.get('Mar MTD_cohort', (0,0))
            feb_r = rate(p_feb, t_feb)
            mar_r = rate(p_mar, t_mar)
            row += f'<td>{fmtR(feb_r)}</td>'
            row += f'<td>{fmtR(mar_r)}</td>'

            # MoM
            mom = wow_pp(mar_r, feb_r)
            mc = mom_color_class(mom.replace('pp','%'))
            row += f'<td class="mc {mc}" style="background:#FEF9E8">{mom}</td>'

            # Best/RR
            row += '<td>—</td><td>—</td><td>—</td>'

            # Target
            row += '<td>—</td><td>—</td>'

        row += '</tr>\n'
        rows.append(row)

    return ''.join(rows)

def generate_grupo_header(name, agg, is_aliado=False, is_consolidated=False):
    """Generate grupo separator row with MoM deltas."""
    # MoM: Feb MTD vs Mar MTD
    feb_q, feb_m, feb_a, feb_p = agg['Feb MTD']
    mar_q, mar_m, mar_a, mar_p = agg['Mar MTD']

    def delta_str(cur, prev, label):
        if prev == 0 and cur > 0:
            return f'{label} <span style="color:#F59E0B">NEW</span>'
        if prev == 0:
            return f'{label} <span style="color:#94A3B8">—</span>'
        pct = (cur - prev) / prev * 100
        color = '#22C55E' if pct > 0 else '#EF4444' if pct < 0 else '#F59E0B'
        return f'{label} <span style="color:{color}">{pct:+.0f}%</span>'

    q_delta = delta_str(mar_q, feb_q, 'Q')
    m_delta = delta_str(mar_m, feb_m, 'M')
    p_delta = delta_str(mar_p, feb_p, 'P')

    css = 'grp cons' if (is_aliado and is_consolidated) else 'grp'
    bg = ''
    if is_consolidated:
        bg = ' style="background:#004E98!important"' if not is_aliado else ' style="background:#3B7DDD!important"'

    return f'<tr class="{css}"><td colspan="21"{bg}>{name} (vs Feb MTD) → {q_delta} | {m_delta} | {p_delta}</td></tr>\n'

# Sort groups by March MTD purchases descending
def sort_key(g, aggs):
    return aggs[g]['Mar MTD'][3]

groups_tas_sorted = sorted(groups_tas, key=lambda g: sort_key(g, agg_tas), reverse=True)
groups_ali_sorted = sorted(groups_ali, key=lambda g: sort_key(g, agg_ali), reverse=True)

# ─── BUILD EXECUTIVE VIEW TABLE ─────────────────────────────────────────────
def build_exec_table():
    html = '<table>\n<thead>\n'

    # Header rows
    html += '<tr class="gh"><th></th>'
    html += '<th colspan="6">Últimas 5 Semanas</th>'
    html += '<th colspan="6">Mensual</th>'
    html += '<th colspan="2">MTD (día 21)</th>'
    html += '<th colspan="1">Comp.</th>'
    html += '<th colspan="1">Best</th>'
    html += '<th colspan="2">Runrates Mar</th>'
    html += '<th colspan="2" style="background:#7C3AED">Target / % BP</th>'
    html += '</tr>\n'

    html += '<tr class="sh"><th>Métrica</th>'
    html += '<th>w-5</th><th>w-4</th><th>w-3</th><th>w-2</th><th>w-1</th>'
    html += '<th style="background:#F59E0B;color:#fff;font-weight:700">WoW</th>'
    html += '<th>Oct</th><th>Nov</th><th>Dic</th><th>Ene</th><th>Feb</th><th>Mar*</th>'
    html += '<th>Feb<br/>MTD</th><th>Mar<br/>MTD</th>'
    html += '<th style="background:#F59E0B;color:#fff;font-weight:700">MoM</th>'
    html += '<th>Best</th>'
    html += '<th>RR<br/>7d</th><th>RR<br/>Seas</th>'
    html += '<th style="background:#7C3AED;color:#fff">Target<br/>BP</th>'
    html += '<th style="background:#7C3AED;color:#fff">%<br/>BP</th>'
    html += '</tr>\n'

    # Calendar date row
    html += '<tr style="border-bottom:2px solid var(--brd)"><th style="background:var(--hdr2)"></th>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        html += f'<th style="font-size:9px;color:#64748B;background:#E8ECF1;font-weight:400;padding:2px 4px">{WEEK_LABELS[wk]}</th>'
    html += '<th style="background:#FEF3C7;padding:2px 4px"></th>'
    for _ in range(6):
        html += '<th style="background:var(--hdr2)"></th>'
    for _ in range(4):
        html += '<th style="background:var(--hdr2)"></th>'
    for _ in range(2):
        html += '<th style="background:var(--hdr2)"></th>'
    html += '<th style="background:#7C3AED"></th><th style="background:#7C3AED"></th>'
    html += '</tr>\n'

    # P MoM% row in header
    html += '<tr style="border-bottom:1px solid var(--brd)"><th style="background:var(--hdr2);font-size:9px;font-weight:600;padding:2px 4px;color:#94A3B8">P MoM%</th>'
    for _ in range(5):
        html += '<th style="background:#E8ECF1;padding:2px 4px"></th>'
    html += '<th style="background:#FEF3C7;padding:2px 4px"></th>'

    # Monthly P MoM%
    prev_p = None
    for mo in MONTHS:
        p = agg_total[mo][3]
        if prev_p is not None and prev_p > 0:
            pct = (p - prev_p) / prev_p * 100
            color = '#22C55E' if pct > 0 else '#EF4444'
            html += f'<th style="background:var(--hdr2);font-size:9px;font-weight:600;padding:2px 4px;color:{color}">{pct:+.1f}%</th>'
        else:
            html += '<th style="background:var(--hdr2);font-size:9px;font-weight:600;padding:2px 4px;color:#94A3B8">—</th>'
        prev_p = p

    for _ in range(4):
        html += '<th style="background:var(--hdr2);padding:2px 4px"></th>'
    for _ in range(2):
        html += '<th style="background:var(--hdr2);padding:2px 4px"></th>'
    html += '<th style="background:#7C3AED;padding:2px 4px"></th><th style="background:#7C3AED;padding:2px 4px"></th>'
    html += '</tr>\n'

    html += '</thead>\n<tbody>\n'

    # TAS Consolidado (= grand total TAS + ALIADO, matching reference naming)
    html += generate_grupo_header('TAS Consolidado', agg_total, is_consolidated=True)
    html += generate_group_rows(agg_total, 'TAS Consolidado', is_consolidated=True)

    # Individual TAS groups
    for g in groups_tas_sorted:
        html += generate_grupo_header(g, agg_tas[g])
        html += generate_group_rows(agg_tas[g], g)

    # Aliado Consolidado
    html += generate_grupo_header('Aliado Consolidado', agg_ali_total, is_aliado=True, is_consolidated=True)
    html += generate_group_rows(agg_ali_total, 'Aliado Consolidado')

    # Individual Aliado groups
    for g in groups_ali_sorted:
        html += generate_grupo_header(g, agg_ali[g], is_aliado=True)
        html += generate_group_rows(agg_ali[g], g, channel='ALIADO')

    # Bottom rows: FTE, daily rates, cells
    html += '<tr class="sp"><td colspan="21">Indicadores Operativos</td></tr>\n'

    # Compras/FTE TAS
    html += '<tr class="bt"><td>Compras / FTE TAS</td>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        html += '<td>—</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        p = agg_tas_total[mo][3]
        fte = FTE_BY_MONTH.get(mo, 0)
        v = p / fte if fte > 0 else 0
        html += f'<td>{v:.1f}</td>'
    # MTD
    feb_p = agg_tas_total['Feb MTD'][3]
    mar_p = agg_tas_total['Mar MTD'][3]
    feb_fte = FTE_BY_MONTH.get('Feb', 0)
    mar_fte = FTE_BY_MONTH.get('Mar', 0)
    html += f'<td>{feb_p/feb_fte:.1f}</td>' if feb_fte > 0 else '<td>—</td>'
    html += f'<td>{mar_p/mar_fte:.1f}</td>' if mar_fte > 0 else '<td>—</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    # Quotes/día
    html += '<tr class="bt"><td>Quotes / día</td>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk]
        bd = biz_days_in_range(ws, we)
        v = agg_total[wk][0] / bd if bd > 0 else 0
        html += f'<td>{v:.0f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        ms, me = MONTH_RANGES[mo]
        bd = biz_days_in_range(ms, me)
        v = agg_total[mo][0] / bd if bd > 0 else 0
        html += f'<td>{v:.0f}</td>'
    html += f'<td>{agg_total["Feb MTD"][0]/biz_days_in_range(date(2026,2,1),FEB_MTD_END):.0f}</td>'
    html += f'<td>{agg_total["Mar MTD"][0]/MAR_ELAPSED_BIZ:.0f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    # Mades/día
    html += '<tr class="bt"><td>Mades / día</td>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk]
        bd = biz_days_in_range(ws, we)
        v = agg_total[wk][1] / bd if bd > 0 else 0
        html += f'<td>{v:.0f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        ms, me = MONTH_RANGES[mo]
        bd = biz_days_in_range(ms, me)
        v = agg_total[mo][1] / bd if bd > 0 else 0
        html += f'<td>{v:.0f}</td>'
    html += f'<td>{agg_total["Feb MTD"][1]/biz_days_in_range(date(2026,2,1),FEB_MTD_END):.0f}</td>'
    html += f'<td>{agg_total["Mar MTD"][1]/MAR_ELAPSED_BIZ:.0f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    # Purchases/día
    html += '<tr class="bt"><td>Purchases / día</td>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        ws, we = WEEKS[wk]
        bd = biz_days_in_range(ws, we)
        v = agg_total[wk][3] / bd if bd > 0 else 0
        html += f'<td>{v:.1f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        ms, me = MONTH_RANGES[mo]
        bd = biz_days_in_range(ms, me)
        v = agg_total[mo][3] / bd if bd > 0 else 0
        html += f'<td>{v:.1f}</td>'
    html += f'<td>{agg_total["Feb MTD"][3]/biz_days_in_range(date(2026,2,1),FEB_MTD_END):.1f}</td>'
    html += f'<td>{agg_total["Mar MTD"][3]/MAR_ELAPSED_BIZ:.1f}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    # Células TAS
    html += '<tr class="bt"><td>Células TAS</td>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        c = count_active_cells(all_records, 'TAS', wk, WEEKS[wk])
        html += f'<td>{c}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        c = count_active_cells(all_records, 'TAS', mo, MONTH_RANGES[mo])
        html += f'<td>{c}</td>'
    html += '<td>—</td><td>—</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    # FTEs TAS
    html += '<tr class="bt"><td>FTEs TAS</td>'
    for _ in range(5):
        html += '<td>—</td>'
    html += '<td style="background:#FEF9E8"></td>'
    for mo in MONTHS:
        html += f'<td>{FTE_BY_MONTH.get(mo, 0)}</td>'
    html += f'<td>{FTE_BY_MONTH.get("Feb", 0)}</td>'
    html += f'<td>{FTE_BY_MONTH.get("Mar", 0)}</td>'
    html += '<td style="background:#FEF9E8"></td>'
    html += '<td>—</td><td>—</td><td>—</td><td>—</td><td>—</td></tr>\n'

    html += '</tbody>\n</table>\n'
    return html

# ─── BUILD FUNNEL TABLES ────────────────────────────────────────────────────
def build_funnel_table(channel, groups_sorted, aggs, agg_consol, consol_name):
    html = '<table>\n<thead>\n'
    html += '<tr class="gh"><th></th>'
    html += '<th colspan="6">Últimas 5 Semanas</th>'
    html += '<th colspan="6">Mensual</th>'
    html += '<th colspan="2">MTD (día 21)</th>'
    html += '<th colspan="1">Comp.</th>'
    html += '<th colspan="1">Best</th>'
    html += '<th colspan="2">Runrates Mar</th>'
    html += '<th colspan="2" style="background:#7C3AED">Target / % BP</th>'
    html += '</tr>\n'
    html += '<tr class="sh"><th>Métrica</th>'
    html += '<th>w-5</th><th>w-4</th><th>w-3</th><th>w-2</th><th>w-1</th>'
    html += '<th style="background:#F59E0B;color:#fff;font-weight:700">WoW</th>'
    html += '<th>Oct</th><th>Nov</th><th>Dic</th><th>Ene</th><th>Feb</th><th>Mar*</th>'
    html += '<th>Feb<br/>MTD</th><th>Mar<br/>MTD</th>'
    html += '<th style="background:#F59E0B;color:#fff;font-weight:700">MoM</th>'
    html += '<th>Best</th>'
    html += '<th>RR<br/>7d</th><th>RR<br/>Seas</th>'
    html += '<th style="background:#7C3AED;color:#fff">Target<br/>BP</th>'
    html += '<th style="background:#7C3AED;color:#fff">%<br/>BP</th>'
    html += '</tr>\n'
    html += '<tr style="border-bottom:2px solid var(--brd)"><th style="background:var(--hdr2)"></th>'
    for wk in ['w-5','w-4','w-3','w-2','w-1']:
        html += f'<th style="font-size:9px;color:#64748B;background:#E8ECF1;font-weight:400;padding:2px 4px">{WEEK_LABELS[wk]}</th>'
    html += '<th style="background:#FEF3C7;padding:2px 4px"></th>'
    for _ in range(14):
        html += '<th style="background:var(--hdr2)"></th>'
    html += '</tr>\n'
    html += '</thead>\n<tbody>\n'

    is_ali = channel == 'ALIADO'
    html += generate_grupo_header(consol_name, agg_consol, is_aliado=is_ali, is_consolidated=True)
    html += generate_group_rows(agg_consol, consol_name)

    for g in groups_sorted:
        html += generate_grupo_header(g, aggs[g], is_aliado=is_ali)
        html += generate_group_rows(aggs[g], g, channel=channel)

    html += '</tbody>\n</table>\n'
    return html

# ─── BUILD SCORECARD DATA ───────────────────────────────────────────────────
def build_scorecard_html():
    """Build scorecard tab — Mensual + Semanal views with metric pills and channel filters."""

    # ── Monthly columns (newest first)
    SC_MONTHS = ['Mar 26', 'Feb 26', 'Ene 26', 'Dic 25', 'Nov 25', 'Oct 25',
                 'Sep 25', 'Ago 25', 'Jul 25', 'Jun 25', 'May 25']
    SC_MONTH_MAP = {
        'Mar 26': 'Mar', 'Feb 26': 'Feb', 'Ene 26': 'Ene',
        'Dic 25': 'Dic', 'Nov 25': 'Nov', 'Oct 25': 'Oct',
        'Sep 25': (date(2025,9,1), date(2025,9,30)),
        'Ago 25': (date(2025,8,1), date(2025,8,31)),
        'Jul 25': (date(2025,7,1), date(2025,7,31)),
        'Jun 25': (date(2025,6,1), date(2025,6,30)),
        'May 25': (date(2025,5,1), date(2025,5,31)),
    }

    # ── Weekly columns (S13*=current partial, S12..S8 = last 5 complete weeks)
    SC_WEEKS = ['S13', 'S12', 'S11', 'S10', 'S9', 'S8']
    SC_WEEK_RANGES = {
        'S13': (date(2026,3,23), date(2026,3,28)),  # current partial week
        'S12': (date(2026,3,16), date(2026,3,21)),   # = w-1
        'S11': (date(2026,3,9),  date(2026,3,14)),   # = w-2
        'S10': (date(2026,3,2),  date(2026,3,7)),    # = w-3
        'S9':  (date(2026,2,23), date(2026,2,28)),   # = w-4
        'S8':  (date(2026,2,16), date(2026,2,21)),   # = w-5
    }

    def get_group_data(grupo, channel):
        """Get monthly + weekly [Q, M, A, P] for a grupo."""
        aggs = agg_tas if channel == 'TAS' else agg_ali

        # Monthly data
        monthly = {}
        for sc_mo in SC_MONTHS:
            mapped = SC_MONTH_MAP[sc_mo]
            if isinstance(mapped, str):
                if grupo in aggs and mapped in aggs[grupo]:
                    monthly[sc_mo] = list(aggs[grupo][mapped])
                else:
                    monthly[sc_mo] = [0, 0, 0, 0]
            else:
                ms, me = mapped
                ch_recs = [r for r in all_records if r['grupo'] == grupo and r['channel'] == channel and r['grupo'] not in EXCLUDE_GROUPS]
                q = sum(1 for r in ch_recs if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
                m = sum(1 for r in ch_recs if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
                a = sum(1 for r in ch_recs if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ms, me))
                p_ch = [r for r in purchase_records if r['grupo'] == grupo and r['channel'] == channel and r['grupo'] not in EXCLUDE_GROUPS]
                p = sum(1 for r in p_ch if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], ms, me))
                monthly[sc_mo] = [q, m, a, p]

        # Weekly data — use existing agg week keys where possible
        week_key_map = {'S12': 'w-1', 'S11': 'w-2', 'S10': 'w-3', 'S9': 'w-4', 'S8': 'w-5'}
        weekly = {}
        for sw in SC_WEEKS:
            if sw in week_key_map and grupo in aggs and week_key_map[sw] in aggs[grupo]:
                weekly[sw] = list(aggs[grupo][week_key_map[sw]])
            else:
                ws, we = SC_WEEK_RANGES[sw]
                ch_recs = [r for r in all_records if r['grupo'] == grupo and r['channel'] == channel and r['grupo'] not in EXCLUDE_GROUPS]
                q = sum(1 for r in ch_recs if r['is_quote'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
                m = sum(1 for r in ch_recs if r['is_made'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
                a = sum(1 for r in ch_recs if r['is_approved'] and r['fecha_creacion'] and in_range(r['fecha_creacion'], ws, we))
                p_ch = [r for r in purchase_records if r['grupo'] == grupo and r['channel'] == channel and r['grupo'] not in EXCLUDE_GROUPS]
                p = sum(1 for r in p_ch if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], ws, we))
                weekly[sw] = [q, m, a, p]

        return monthly, weekly

    # Build data for all groups
    print("  Building scorecard data for all groups...")
    sc_data = []
    for g in groups_tas_sorted:
        mdata, wdata = get_group_data(g, 'TAS')
        sc_data.append({'n': g, 'ch': 'TAS', 'mo': mdata, 'wk': wdata})
    for g in groups_ali_sorted:
        mdata, wdata = get_group_data(g, 'ALIADO')
        sc_data.append({'n': g, 'ch': 'ALIADO', 'mo': mdata, 'wk': wdata})

    sc_json = json.dumps(sc_data)
    total_q = sum(d['mo']['Mar 26'][0] for d in sc_data)
    total_p = sum(d['mo']['Mar 26'][3] for d in sc_data)

    html = f'''
<div id="sc-root">
<div class="sc-top">
  <div>
    <div class="sc-title">Grupos</div>
    <div class="sc-subtitle" id="sc-sub">{len(sc_data)} grupos · {fmt(total_q)} quotes · {fmt(total_p)} compras · Mar 2026</div>
  </div>
  <div class="sc-filters">
    <button class="sc-ch-btn active" data-ch="all">General</button>
    <button class="sc-ch-btn" data-ch="TAS">TAS</button>
    <button class="sc-ch-btn" data-ch="ALIADO">Aliado</button>
  </div>
</div>
<div class="sc-bar">
  <div class="sc-view-toggle">
    <button class="sc-vbtn active" data-v="mo">Mensual</button>
    <button class="sc-vbtn" data-v="wk">Semanal</button>
  </div>
  <div class="sc-metric-pills">
    <button class="sc-mpill active" data-m="3">Purchased</button>
    <button class="sc-mpill" data-m="0">Quotes</button>
    <button class="sc-mpill" data-m="1">Made</button>
    <button class="sc-mpill" data-m="2">Approved</button>
    <button class="sc-mpill" data-m="qm">Q→M</button>
    <button class="sc-mpill" data-m="ma">M→A</button>
    <button class="sc-mpill" data-m="ap">A→P</button>
    <button class="sc-mpill" data-m="qp">Q→P</button>
  </div>
</div>
<div class="sc-table-wrap">
<table class="sc-tbl2">
<thead id="sc-thead"></thead>
<tbody id="sc-tbody2"></tbody>
</table>
</div>
</div>
<script>
var _scD={sc_json};
var _scMoCols={json.dumps(SC_MONTHS)};
var _scWkCols={json.dumps(SC_WEEKS)};
var _scElapsed={MAR_ELAPSED_BIZ};
var _scTotal={MAR_BIZ_DAYS};
</script>
'''
    return html

def build_scorecard_data():
    """Legacy compat."""
    return '[]', '[]'

# ─── LOAD 2DOS PAGOS DATA ────────────────────────────────────────────────────
def load_2dos_pagos():
    """Load 2dos Pagos CSV and compute SLA metrics."""
    pagos_csv = os.path.expanduser("~/Downloads/Summary MKP _ TAS & BULK - 2dos Pagos (1).csv")
    if not os.path.exists(pagos_csv):
        return None

    with open(pagos_csv, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    # Find header row
    header_idx = None
    for i, row in enumerate(all_rows):
        if row[0].strip() == 'Transaction Id':
            header_idx = i
            break
    if header_idx is None:
        return None

    data_rows = all_rows[header_idx+1:]
    result = {
        'total': 0, 'over15': 0, 'within15': 0,
        'tas': {'total': 0, 'over15': 0},
        'ali': {'total': 0, 'over15': 0},
        'by_grupo': defaultdict(lambda: {'total': 0, 'over15': 0}),
        'aging_sum': 0, 'aging_max': 0,
        'worst_agent': '', 'worst_grupo': '',
        'cohort': defaultdict(lambda: {'total': 0, 'over15': 0}),
    }

    for row in data_rows:
        if len(row) < 9 or not row[0].strip():
            continue
        try:
            aging = int(row[7].strip())
        except (ValueError, IndexError):
            continue

        channel = row[8].strip()
        agente = row[1].strip()
        grupo_raw = row[6].strip()
        grupo = grupo_raw.split('-')[0].strip() if '-' in grupo_raw else grupo_raw
        grupo_norm = GRUPO_MAP.get(grupo, grupo.replace('GRUPO ', ''))

        result['total'] += 1
        result['aging_sum'] += aging
        if aging > result['aging_max']:
            result['aging_max'] = aging
            result['worst_agent'] = agente
            result['worst_grupo'] = grupo_norm
        if aging > 15:
            result['over15'] += 1
        else:
            result['within15'] += 1

        ch_key = 'tas' if channel == 'TAS' else 'ali'
        result[ch_key]['total'] += 1
        if aging > 15:
            result[ch_key]['over15'] += 1

        result['by_grupo'][grupo_norm]['total'] += 1
        if aging > 15:
            result['by_grupo'][grupo_norm]['over15'] += 1

        # Cohort by Item Receipt week (Fecha de PO = col 4)
        po_date_str = row[4].strip() if len(row) > 4 else ''
        if '/' in po_date_str:
            parts = po_date_str.split('/')
            if len(parts) == 3:
                try:
                    d = date(int(parts[2]), int(parts[1]), int(parts[0]))
                    iso_week = d.isocalendar()[1]
                    wk_key = f'S{iso_week}'
                    result['cohort'][wk_key]['total'] += 1
                    if aging > 15:
                        result['cohort'][wk_key]['over15'] += 1
                except:
                    pass

    result['avg_aging'] = result['aging_sum'] / result['total'] if result['total'] > 0 else 0
    result['pct_over15'] = result['over15'] / result['total'] * 100 if result['total'] > 0 else 0
    return result

print("Loading 2dos Pagos data...")
pagos_data = load_2dos_pagos()
if pagos_data:
    print(f"  2dos Pagos: {pagos_data['total']} pendientes, {pagos_data['over15']} fuera SLA ({pagos_data['pct_over15']:.1f}%)")

# ─── BUILD SUMMARY SEMANAL ───────────────────────────────────────────────────
def build_summary_semanal():
    """Build Summary Semanal tab with KPI cards."""

    def _build_cohort_rows(pd):
        cohort = pd['cohort']
        weeks_sorted = sorted(cohort.keys(), key=lambda x: int(x[1:]), reverse=True)
        rows = ''
        for wk in weeks_sorted:
            d = cohort[wk]
            if d['total'] == 0:
                continue
            pct = d['over15'] / d['total'] * 100
            color = '#22C55E' if pct <= 15 else '#F59E0B' if pct <= 50 else '#EF4444'
            star = '*' if wk == 'S13' else ''
            rows += f'<tr><td>{wk}{star}</td><td>{d["total"]}</td><td style="color:{color};font-weight:700">{d["over15"]}</td><td style="color:{color};font-weight:700">{pct:.0f}%</td></tr>'
        return rows

    def _build_pagos_section():
        if not pagos_data:
            return '<div class="sm-wip"><div class="sm-wip-icon">⚠️</div><div class="sm-wip-text">Archivo de 2dos Pagos no encontrado</div></div>'

        pd = pagos_data
        pct = pd['pct_over15']
        sla_color = '#22C55E' if pct <= 15 else '#F59E0B' if pct <= 25 else '#EF4444'
        tas_pct = pd['tas']['over15'] / pd['tas']['total'] * 100 if pd['tas']['total'] > 0 else 0
        ali_pct = pd['ali']['over15'] / pd['ali']['total'] * 100 if pd['ali']['total'] > 0 else 0
        tas_color = '#22C55E' if tas_pct <= 15 else '#F59E0B' if tas_pct <= 25 else '#EF4444'
        ali_color = '#22C55E' if ali_pct <= 15 else '#F59E0B' if ali_pct <= 25 else '#EF4444'

        # Top offenders
        top_grupos = sorted(pd['by_grupo'].items(), key=lambda x: x[1]['over15'], reverse=True)[:6]
        grupo_rows = ''
        for g, d in top_grupos:
            if d['over15'] == 0:
                continue
            g_pct = d['over15'] / d['total'] * 100 if d['total'] > 0 else 0
            g_color = '#22C55E' if g_pct <= 15 else '#F59E0B' if g_pct <= 25 else '#EF4444'
            grupo_rows += f'<tr><td style="text-align:left;font-weight:600">{g}</td><td>{d["total"]}</td><td style="color:{g_color};font-weight:700">{d["over15"]}</td><td style="color:{g_color};font-weight:700">{g_pct:.0f}%</td></tr>'

        return f'''
  <div class="sm-kpi-row" style="grid-template-columns:repeat(5,1fr)">
    <div class="sm-kpi">
      <div class="sm-kpi-label">Fuera de SLA</div>
      <div class="sm-kpi-val" style="font-size:32px;color:{sla_color}">{pct:.1f}%</div>
      <div class="sm-kpi-delta">{pd["over15"]} de {pd["total"]} pendientes</div>
    </div>
    <div class="sm-kpi">
      <div class="sm-kpi-label">TAS</div>
      <div class="sm-kpi-val" style="color:{tas_color}">{tas_pct:.1f}%</div>
      <div class="sm-kpi-delta">{pd["tas"]["over15"]} de {pd["tas"]["total"]}</div>
    </div>
    <div class="sm-kpi">
      <div class="sm-kpi-label">Aliado</div>
      <div class="sm-kpi-val" style="color:{ali_color}">{ali_pct:.1f}%</div>
      <div class="sm-kpi-delta">{pd["ali"]["over15"]} de {pd["ali"]["total"]}</div>
    </div>
    <div class="sm-kpi">
      <div class="sm-kpi-label">Aging Promedio</div>
      <div class="sm-kpi-val">{pd["avg_aging"]:.0f}d</div>
      <div class="sm-kpi-sub">SLA = 15 días</div>
    </div>
    <div class="sm-kpi">
      <div class="sm-kpi-label">Aging Máximo</div>
      <div class="sm-kpi-val" style="color:#EF4444">{pd["aging_max"]}d</div>
      <div class="sm-kpi-sub" style="font-size:10px;color:#64748B">LM: {pd["worst_agent"]}</div>
      <div class="sm-kpi-sub">{pd["worst_grupo"]}</div>
    </div>
  </div>
  <div style="margin-top:12px;display:grid;grid-template-columns:1fr 1fr;gap:12px">
    <div>
      <div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:6px">Por Grupo</div>
      <table class="sm-pagos-tbl">
        <thead><tr><th style="text-align:left">Grupo</th><th>Pendientes</th><th>Fuera SLA</th><th>%</th></tr></thead>
        <tbody>{grupo_rows}</tbody>
      </table>
    </div>
    <div>
      <div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:6px">Cohort por Semana Item Receipt</div>
      <table class="sm-pagos-tbl">
        <thead><tr><th>Semana IR</th><th>Pendientes</th><th>Fuera SLA</th><th>%</th></tr></thead>
        <tbody>{_build_cohort_rows(pd)}</tbody>
      </table>
    </div>
  </div>'''

    # ── Section 1: Compras volumes
    w1_p = agg_total['w-1'][3]
    w2_p = agg_total['w-2'][3]
    mar_mtd_p = agg_total['Mar MTD'][3]
    feb_mtd_p = agg_total['Feb MTD'][3]
    wow_val = ((w1_p - w2_p) / w2_p * 100) if w2_p > 0 else 0
    mom_val = ((mar_mtd_p - feb_mtd_p) / feb_mtd_p * 100) if feb_mtd_p > 0 else 0
    rr7 = rr_7d(agg_total, 3)
    rrs = rr_seas(agg_total, 3)
    target_p = BP_TARGETS.get('Purchased', 370)
    pct_bp = mar_mtd_p / target_p * 100 if target_p > 0 else 0

    # ── Section 2: Funnel conversions
    def funnel_rates(agg):
        mar = agg['Mar MTD']
        feb = agg['Feb MTD']
        q, m, a, p = mar
        fq, fm, fa, fp = feb
        return {
            'Q': q, 'M': m, 'A': a, 'P': p,
            'qm': rate(m, q), 'ma': rate(a, m), 'ap': rate(p, a), 'qp': rate(p, q),
            'f_qm': rate(fm, fq), 'f_ma': rate(fa, fm), 'f_ap': rate(fp, fa), 'f_qp': rate(fp, fq),
        }

    total_f = funnel_rates(agg_total)
    tas_f = funnel_rates(agg_tas_total)
    ali_f = funnel_rates(agg_ali_total)

    def delta_html(cur, prev, suffix='%', is_pp=False):
        if prev == 0 and cur == 0:
            return '<span style="color:#94A3B8">—</span>'
        if is_pp:
            d = cur - prev
            color = '#22C55E' if d >= 0 else '#EF4444'
            return f'<span style="color:{color}">{d:+.1f}pp</span>'
        if prev == 0:
            return '<span style="color:#22C55E">NEW</span>'
        d = (cur - prev) / prev * 100
        color = '#22C55E' if d >= 0 else '#EF4444'
        return f'<span style="color:{color}">{d:+.1f}%</span>'

    def kpi_card(label, value, delta_str, sub='', big=False):
        sz = '32px' if big else '24px'
        return f'''<div class="sm-kpi">
          <div class="sm-kpi-label">{label}</div>
          <div class="sm-kpi-val" style="font-size:{sz}">{value}</div>
          <div class="sm-kpi-delta">{delta_str}</div>
          {f'<div class="sm-kpi-sub">{sub}</div>' if sub else ''}
        </div>'''

    def funnel_weekly_table(title, agg, color):
        """Build a weekly funnel table with 5 weeks + WoW."""
        wks = ['w-5','w-4','w-3','w-2','w-1']
        wk_labels = ['W8','W9','W10','W11','W12']
        metrics = [
            ('Quotes', 0), ('Made', 1), ('Approved', 2), ('Purchased', 3),
            ('Q→M%', (1,0)), ('M→A%', (2,1)), ('A→P%', (3,2)), ('Q→P%', (3,0)),
        ]
        rows_html = ''
        for label, idx in metrics:
            is_rate = isinstance(idx, tuple)
            is_pur = label == 'Purchased'
            cls = ' class="ft-bold"' if is_pur else ' class="ft-rate"' if is_rate else ''
            rows_html += f'<tr{cls}><td>{label}</td>'
            for wk in wks:
                d = agg[wk]
                if is_rate:
                    v = rate(d[idx[0]], d[idx[1]])
                    rows_html += f'<td>{fmtR(v)}</td>'
                else:
                    rows_html += f'<td>{fmt(d[idx])}</td>'
            # WoW
            w1, w2 = agg['w-1'], agg['w-2']
            if is_rate:
                v1 = rate(w1[idx[0]], w1[idx[1]])
                v2 = rate(w2[idx[0]], w2[idx[1]])
                diff = v1 - v2
                color_cls = 'color:#22C55E' if diff >= 0 else 'color:#EF4444'
                rows_html += f'<td style="{color_cls};font-weight:700">{diff:+.1f}pp</td>'
            else:
                v1, v2 = w1[idx], w2[idx]
                if v2 > 0:
                    pct = (v1 - v2) / v2 * 100
                    color_cls = 'color:#22C55E' if pct >= 0 else 'color:#EF4444'
                    rows_html += f'<td style="{color_cls};font-weight:700">{pct:+.1f}%</td>'
                elif v1 > 0:
                    rows_html += '<td style="color:#22C55E;font-weight:700">NEW</td>'
                else:
                    rows_html += '<td style="color:#94A3B8">—</td>'
            rows_html += '</tr>\n'

        return f'''<div class="sm-ftable" style="border-top:3px solid {color}">
          <div class="sm-ft-title">{title}</div>
          <table class="sm-ft">
            <thead><tr><th></th>{''.join(f"<th>{l}</th>" for l in wk_labels)}<th style="background:#FEF3C7;color:#92400E">WoW</th></tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>'''

    wow_color = '#22C55E' if wow_val >= 0 else '#EF4444'
    mom_color = '#22C55E' if mom_val >= 0 else '#EF4444'

    html = f'''
<div class="sm-container">

<!-- Section 1: Compras -->
<div class="sm-section">
  <div class="sm-section-title">Compras — Semana W12 (Mar 16–21)</div>
  <div class="sm-kpi-row">
    {kpi_card('Compras W12', fmt(w1_p), f'<span style="color:{wow_color}">WoW {wow_val:+.1f}%</span>', f'vs W11: {fmt(w2_p)}', big=True)}
    {kpi_card('Mar MTD', fmt(mar_mtd_p), f'<span style="color:{mom_color}">MoM {mom_val:+.1f}%</span>', f'vs Feb MTD: {fmt(feb_mtd_p)}')}
    {kpi_card('Run Rate 7d', fmt(rr7), f'{pct_bp:.0f}% del BP ({fmt(target_p)})', 'Proyección lineal')}
    {kpi_card('RR Seasonality', fmt(rrs), f'{rrs/target_p*100:.0f}% del BP' if target_p>0 else '', 'Prom ponderado 5 sem')}
  </div>
</div>

<!-- Section 2: Conversiones Funnel (weekly tables) -->
<div class="sm-section">
  <div class="sm-section-title">Conversiones Funnel — Últimas 5 Semanas</div>
  <div class="sm-funnel-row">
    {funnel_weekly_table('Total (TAS + Aliado)', agg_total, '#1E293B')}
    {funnel_weekly_table('TAS', agg_tas_total, '#1B2A4A')}
    {funnel_weekly_table('Aliado', agg_ali_total, '#3B7DDD')}
  </div>
</div>

<!-- Section 3: Sentinel Score (WIP) -->
<div class="sm-section">
  <div class="sm-section-title">Sentinel Score</div>
  <div class="sm-wip">
    <div class="sm-wip-icon">🚧</div>
    <div class="sm-wip-text">Work in Progress — Próximamente: score compuesto de salud operativa por célula</div>
  </div>
</div>

<!-- Section 4: Days to Hub / Days to Produced (WIP) -->
<div class="sm-section">
  <div class="sm-section-title">Days to Hub Destino / % Days to Produced</div>
  <div class="sm-wip">
    <div class="sm-wip-icon">🚧</div>
    <div class="sm-wip-text">Work in Progress — Próximamente: tiempos de traslado y producción post-compra</div>
  </div>
</div>

<!-- Section 5: % Autos en 2do pagos > 15 días -->
<div class="sm-section">
  <div class="sm-section-title">2do Pagos — SLA Item Receipt (&gt; 15 días = fuera de SLA)</div>
  {_build_pagos_section()}
</div>

</div>'''
    return html

# ─── BUILD CHART DATA ───────────────────────────────────────────────────────
def build_chart_data():
    """Compute values for all 7 charts."""
    # Chart 1: Monthly purchases TAS vs Aliado
    tas_monthly_p = [agg_tas_total[mo][3] for mo in MONTHS]
    ali_monthly_p = [agg_ali_total[mo][3] for mo in MONTHS]
    tas_rr = rr_7d(agg_tas_total, 3)
    ali_rr = rr_7d(agg_ali_total, 3)

    # Chart 2: Waterfall total
    feb_total = agg_total['Feb MTD'][3]
    mar_total = agg_total['Mar MTD'][3]
    tas_delta = agg_tas_total['Mar MTD'][3] - agg_tas_total['Feb MTD'][3]
    ali_delta = agg_ali_total['Mar MTD'][3] - agg_ali_total['Feb MTD'][3]

    # Chart 3: TAS waterfall by funnel
    tas_feb_funnel = agg_tas_total['Feb MTD']
    tas_mar_funnel = agg_tas_total['Mar MTD']
    tas_deltas = [tas_mar_funnel[i] - tas_feb_funnel[i] for i in range(4)]

    # Chart 4: Aliado waterfall by funnel
    ali_feb_funnel = agg_ali_total['Feb MTD']
    ali_mar_funnel = agg_ali_total['Mar MTD']
    ali_deltas = [ali_mar_funnel[i] - ali_feb_funnel[i] for i in range(4)]

    # Cells annotation
    feb_tas_cells = count_active_cells(all_records, 'TAS', 'Feb', MONTH_RANGES['Feb'])
    mar_tas_cells = count_active_cells(all_records, 'TAS', 'Mar', MONTH_RANGES['Mar'])
    feb_ali_cells = count_active_cells(all_records, 'ALIADO', 'Feb', MONTH_RANGES['Feb'])
    mar_ali_cells = count_active_cells(all_records, 'ALIADO', 'Mar', MONTH_RANGES['Mar'])

    feb_q_per_cell = round(tas_feb_funnel[0] / feb_tas_cells) if feb_tas_cells > 0 else 0
    mar_q_per_cell = round(tas_mar_funnel[0] / mar_tas_cells) if mar_tas_cells > 0 else 0
    q_per_cell_pct = round((mar_q_per_cell - feb_q_per_cell) / feb_q_per_cell * 100) if feb_q_per_cell > 0 else 0

    feb_q_per_ali = round(ali_feb_funnel[0] / feb_ali_cells) if feb_ali_cells > 0 else 0
    mar_q_per_ali = round(ali_mar_funnel[0] / mar_ali_cells) if mar_ali_cells > 0 else 0
    q_per_ali_pct = round((mar_q_per_ali - feb_q_per_ali) / feb_q_per_ali * 100) if feb_q_per_ali > 0 else 0

    # Chart 5 & 6: Conversion rates
    def conv_rates(agg, months):
        qm = [rate(agg[m][1], agg[m][0]) for m in months]
        ma = [rate(agg[m][2], agg[m][1]) for m in months]
        ap = [rate(agg[m][3], agg[m][2]) for m in months]
        qp = [rate(agg[m][3], agg[m][0]) for m in months]
        return qm, ma, ap, qp

    tas_qm, tas_ma, tas_ap, tas_qp = conv_rates(agg_tas_total, MONTHS)
    ali_qm, ali_ma, ali_ap, ali_qp = conv_rates(agg_ali_total, MONTHS)

    # Chart 7: Region
    region_purchases = defaultdict(lambda: {'TAS': 0, 'ALIADO': 0})
    for r in purchase_records:
        if r['is_purchased'] and r['purchase_date'] and in_range(r['purchase_date'], date(2026,3,1), MAR_MTD_END):
            grupo = r['grupo']
            channel = r['channel']
            region = REGION_MAP.get(grupo, 'Lerma')
            region_purchases[region][channel] += r['is_purchased']

    regions_order = ['Lerma', 'GDL', 'QRO', 'MTY', 'CDMX']
    region_tas = [region_purchases[rg]['TAS'] for rg in regions_order]
    region_ali = [region_purchases[rg]['ALIADO'] for rg in regions_order]

    return {
        'tas_p': tas_monthly_p, 'ali_p': ali_monthly_p, 'tas_rr': tas_rr, 'ali_rr': ali_rr,
        'feb_total': feb_total, 'mar_total': mar_total, 'tas_delta': tas_delta, 'ali_delta': ali_delta,
        'tas_feb_funnel': list(tas_feb_funnel), 'tas_mar_funnel': list(tas_mar_funnel),
        'tas_deltas': tas_deltas,
        'ali_feb_funnel': list(ali_feb_funnel), 'ali_mar_funnel': list(ali_mar_funnel),
        'ali_deltas': ali_deltas,
        'feb_tas_cells': feb_tas_cells, 'mar_tas_cells': mar_tas_cells,
        'feb_ali_cells': feb_ali_cells, 'mar_ali_cells': mar_ali_cells,
        'feb_q_per_cell': feb_q_per_cell, 'mar_q_per_cell': mar_q_per_cell, 'q_per_cell_pct': q_per_cell_pct,
        'feb_q_per_ali': feb_q_per_ali, 'mar_q_per_ali': mar_q_per_ali, 'q_per_ali_pct': q_per_ali_pct,
        'tas_qm': [round(x,1) for x in tas_qm], 'tas_ma': [round(x,1) for x in tas_ma],
        'tas_ap': [round(x,1) for x in tas_ap], 'tas_qp': [round(x,1) for x in tas_qp],
        'ali_qm': [round(x,1) for x in ali_qm], 'ali_ma': [round(x,1) for x in ali_ma],
        'ali_ap': [round(x,1) for x in ali_ap], 'ali_qp': [round(x,1) for x in ali_qp],
        'region_tas': region_tas, 'region_ali': region_ali,
    }

# ─── ASSEMBLE HTML ──────────────────────────────────────────────────────────
chart_data = build_chart_data()
scorecard_html = build_scorecard_html()
summary_html = build_summary_semanal()

exec_table = build_exec_table()
funnel_tas_table = build_funnel_table('TAS', groups_tas_sorted, agg_tas, agg_tas_total, 'TAS Consolidado')
funnel_ali_table = build_funnel_table('ALIADO', groups_ali_sorted, agg_ali, agg_ali_total, 'Aliado Consolidado')

# Feb total for chart annotation
feb_total_p = agg_total['Feb'][3]
oct_total_p = agg_total['Oct'][3]
growth_pct = round((chart_data['tas_rr'] + chart_data['ali_rr'] - feb_total_p) / feb_total_p * 100) if feb_total_p > 0 else 0

full_html = f'''<!DOCTYPE html>

<html lang="es">
<head>
<meta charset="utf-8"/>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" name="viewport"/>
<title>TAS Weekly Consolidated — W13</title>
<style>
:root{{
  --dark:#1A1A2E;--hdr:#2B478B;--hdr2:#3A569A;
  --g:#22C55E;--r:#EF4444;--a:#F59E0B;--p:#8B5CF6;
  --lg:#F8FAFC;--brd:#E2E8F0;--sep:#E8ECF1;
  --grupo-bg:#1B2A4A;--ali:#3B7DDD;
}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;
  background:#F1F5F9;color:#1E293B;
  min-height:100vh;min-height:100dvh;
  display:flex;flex-direction:column;
}}
.hdr{{
  background:var(--dark);padding:14px 16px 10px;
  position:sticky;top:0;z-index:200;flex-shrink:0;
}}
.hdr h1{{color:#fff;font-size:17px;font-weight:700;margin-bottom:3px}}
.hdr .sub{{color:#94A3B8;font-size:10px;line-height:1.4}}

/* Tab bar */
.tab-bar{{
  position:sticky;top:52px;z-index:190;
  background:#F1F5F9;
  padding:8px 12px 6px;
  display:flex;gap:6px;
  overflow-x:auto;
  -webkit-overflow-scrolling:touch;
  border-bottom:1px solid var(--brd);
  flex-shrink:0;
}}
.tab-btn{{
  flex-shrink:0;
  padding:7px 14px;
  border:2px solid var(--brd);
  border-radius:20px;
  background:#fff;
  color:#64748B;
  font-size:12px;font-weight:600;
  cursor:pointer;
  transition:all .15s;
  white-space:nowrap;
}}
.tab-btn:hover{{border-color:var(--hdr);color:var(--hdr)}}
.tab-btn.active{{
  background:var(--hdr);color:#fff;border-color:var(--hdr);
}}

/* Tab content */
.tab-content{{display:none;flex:1;flex-direction:column;overflow:hidden}}
.tab-content.active{{display:flex}}

/* Table wrapper */
.wrap{{flex:1;overflow:auto;-webkit-overflow-scrolling:touch}}

/* Table styles */
table{{border-collapse:collapse;width:max-content;min-width:100%;font-size:11px}}
th,td{{padding:5px 7px;text-align:center;white-space:nowrap;border-bottom:1px solid var(--brd);min-width:54px}}
th:first-child,td:first-child{{
  position:sticky;left:0;z-index:10;
  text-align:left;min-width:105px;max-width:125px;
  white-space:normal;font-weight:600;
  border-right:2px solid var(--brd);
}}
.gh th{{background:var(--hdr);color:#fff;font-size:10px;font-weight:700;padding:4px 5px;border-right:1px solid rgba(255,255,255,.15)}}
.gh th:first-child{{background:var(--hdr)}}
.sh th{{background:var(--hdr2);color:#fff;font-size:9px;font-weight:600;padding:4px 5px}}
.sh th:first-child{{background:var(--hdr2)}}
tr.grp td{{
  background:var(--grupo-bg)!important;color:#fff;font-weight:700;
  font-size:12px;text-align:left;padding:6px 8px;
  border-bottom:2px solid var(--hdr);
}}
tr.grp.cons td{{background:var(--ali)!important}}
tr.dr td{{background:#fff}}
tr.dr:nth-child(even) td{{background:var(--lg)}}
tr.dr td:first-child{{background:inherit}}
tr.bl td{{font-weight:700}}
tr.qp td{{color:var(--hdr)}}
tr.qp td:first-child{{color:var(--dark)}}
tr.qp td.mc{{color:inherit}}
tr.pu td:first-child{{color:var(--p)}}
.pos{{color:var(--g);font-weight:700}}
.neg{{color:var(--r);font-weight:700}}
.neu,.neutral{{color:var(--a);font-weight:700}}
tr.sp td{{background:var(--sep)!important;color:var(--hdr);font-weight:700;font-size:10px;text-align:left;padding:5px 8px}}
tr.bt td{{font-size:10px;background:#fff}}
tr.bt:nth-child(even) td{{background:var(--lg)}}
tr.bt td:first-child{{background:inherit}}
.ft{{padding:10px 12px;font-size:8px;color:#94A3B8;line-height:1.5;flex-shrink:0}}

/* Summary Semanal */
.sm-container{{padding:12px;display:flex;flex-direction:column;gap:14px;overflow-y:auto;flex:1}}
.sm-section{{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.sm-section-title{{font-size:13px;font-weight:700;color:#1E293B;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #E2E8F0}}
.sm-kpi-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}}
.sm-kpi{{background:#F8FAFC;border-radius:8px;padding:12px;text-align:center;border:1px solid #E2E8F0}}
.sm-kpi-label{{font-size:10px;font-weight:600;color:#94A3B8;text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px}}
.sm-kpi-val{{font-size:24px;font-weight:800;color:#1E293B;line-height:1.2}}
.sm-kpi-delta{{font-size:11px;font-weight:600;margin-top:2px}}
.sm-kpi-sub{{font-size:9px;color:#94A3B8;margin-top:2px}}
.sm-funnel-row{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
.sm-ftable{{background:#F8FAFC;border-radius:8px;padding:10px;border:1px solid #E2E8F0;overflow-x:auto}}
.sm-ft-title{{font-size:12px;font-weight:700;color:#1E293B;margin-bottom:6px}}
.sm-ft{{width:100%;border-collapse:collapse;font-size:11px}}
.sm-ft th,.sm-ft td{{padding:4px 8px;text-align:center;white-space:nowrap;border-bottom:1px solid #E2E8F0}}
.sm-ft thead th{{background:#F1F5F9;color:#64748B;font-size:9px;font-weight:600}}
.sm-ft th:first-child,.sm-ft td:first-child{{text-align:left;font-weight:600;min-width:60px}}
.sm-ft tr.ft-bold td{{font-weight:700}}
.sm-ft tr.ft-rate td{{color:#475569;font-size:10px}}
.sm-wip{{display:flex;align-items:center;gap:12px;padding:20px;background:#F8FAFC;border-radius:8px;border:1px dashed #CBD5E1}}
.sm-wip-icon{{font-size:24px}}
.sm-wip-text{{font-size:12px;color:#94A3B8;font-weight:500}}
.sm-pagos-tbl{{width:100%;border-collapse:collapse;font-size:11px}}
.sm-pagos-tbl th,.sm-pagos-tbl td{{padding:6px 10px;text-align:center;border-bottom:1px solid #F1F5F9}}
.sm-pagos-tbl thead th{{background:#F8FAFC;color:#94A3B8;font-size:9px;font-weight:600;text-transform:uppercase}}
@media(max-width:599px){{
  .sm-kpi-row{{grid-template-columns:repeat(2,1fr)}}
  .sm-funnel-row{{grid-template-columns:1fr}}
  .sm-kpi-val{{font-size:20px}}
}}

/* Charts grid */
.grid{{display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:calc(50vh - 40px);gap:8px;padding:8px}}
.card{{background:#fff;border-radius:8px;padding:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);display:flex;flex-direction:column;overflow:hidden}}
.card h2{{font-size:11px;font-weight:700;color:var(--dark);margin-bottom:6px;line-height:1.3;flex-shrink:0}}
.card .cw{{flex:1;position:relative;min-height:0}}
.card canvas{{position:absolute;top:0;left:0;width:100%!important;height:100%!important}}

@media(max-width:599px){{
  .grid{{grid-template-columns:1fr;grid-auto-rows:40vh;gap:8px;padding:8px}}
  .card h2{{font-size:10px}}
}}
@media(max-width:480px){{
  .hdr{{padding:10px 10px 8px}}
  .hdr h1{{font-size:15px}}
  .hdr .sub{{font-size:9px}}
  .tab-bar{{top:42px;padding:6px 8px 5px}}
  .tab-btn{{font-size:11px;padding:6px 10px}}
  table{{font-size:10px}}
  th,td{{padding:4px 5px;min-width:46px}}
  th:first-child,td:first-child{{min-width:85px;max-width:105px;font-size:9px}}
  tr.grp td{{font-size:11px;padding:5px 6px}}
  .card{{padding:8px}}
  .grid{{grid-auto-rows:40vh}}
}}
</style>
</head>
<body>
<div class="hdr">
<h1>TAS Weekly Consolidated — W13</h1>
<div class="sub">W12 (16–21 Mar 2026) · Datos: 21/Mar · Días háb: Lun–Sáb excl. feriados · Feriado: 16/Mar · Mar = 25 d.h. ({MAR_ELAPSED_BIZ} MTD)</div>
</div>
<div class="tab-bar">
<button class="tab-btn active" onclick="switchTab('summary',this)">Summary Semanal</button>
<button class="tab-btn" onclick="switchTab('ftas',this)">Funnel TAS</button>
<button class="tab-btn" onclick="switchTab('fali',this)">Funnel Aliado</button>
<button class="tab-btn" onclick="switchTab('charts',this)">Charts</button>
<button class="tab-btn" onclick="switchTab('scorecard',this)">Scorecard</button>
</div>
<div class="tab-content active" id="tab-summary">
{summary_html}
<div class="ft">Summary Semanal · W12 (Mar 16–21) · Datos: 21/Mar · WIP = Work in Progress, datos pendientes de integración</div>
</div>
<div class="tab-content" id="tab-ftas">
<div class="wrap">
{funnel_tas_table}
</div>
<div class="ft">Snapshot (7)+(8) · Solo TAS · Mar* = parcial 21/Mar · Grupos ordenados por compras Mar MTD desc</div>
</div>
<div class="tab-content" id="tab-fali">
<div class="wrap">
{funnel_ali_table}
</div>
<div class="ft">Snapshot (7)+(8) · Solo ALIADO · Mar* = parcial 21/Mar · Grupos ordenados por compras Mar MTD desc</div>
</div>
<div class="tab-content" id="tab-charts">
<div class="grid">
<div class="card"><h2>Compras Mensuales — TAS vs Aliado (Oct → Mar RR: +{growth_pct}%)</h2>
<div class="cw"><canvas id="c1"></canvas></div></div>
<div class="card"><h2>Waterfall Total — Feb MTD → Mar MTD</h2>
<div class="cw"><canvas id="c2"></canvas></div></div>
<div class="card"><h2>Waterfall TAS — Por Etapa del Funnel</h2>
<div class="cw"><canvas id="c3"></canvas></div></div>
<div class="card"><h2>Waterfall Aliado — Por Etapa del Funnel</h2>
<div class="cw"><canvas id="c4"></canvas></div></div>
<div class="card"><h2>Conversiones TAS (mensual)</h2>
<div class="cw"><canvas id="c5"></canvas></div></div>
<div class="card"><h2>Conversiones Aliado (mensual)</h2>
<div class="cw"><canvas id="c6"></canvas></div></div>
<div class="card"><h2>Aporte por Región — Mar MTD</h2>
<div class="cw"><canvas id="c7"></canvas></div></div>
</div>
<div class="ft">Snapshot (7)+(8) · Regiones: Premier/Plasencia→GDL, León/AGS/SLP/Potosina→QRO, Soni/Pachuca→CDMX, WeCars/Misol/Autopolis→MTY, Resto→Lerma · RR = Runrate 7d</div>
</div>

<style>
#sc-root{{padding:10px;display:flex;flex-direction:column;height:100%}}
.sc-top{{display:flex;justify-content:space-between;align-items:flex-start;padding:6px 4px 10px;flex-shrink:0}}
.sc-title{{font-size:20px;font-weight:800;color:#1E293B}}
.sc-subtitle{{font-size:11px;color:#94A3B8;margin-top:2px}}
.sc-filters{{display:flex;gap:0;border:2px solid var(--brd);border-radius:8px;overflow:hidden}}
.sc-ch-btn{{padding:7px 18px;font-size:12px;font-weight:600;border:none;background:#fff;color:#64748B;cursor:pointer;transition:all .15s}}
.sc-ch-btn:not(:last-child){{border-right:1px solid var(--brd)}}
.sc-ch-btn.active{{background:var(--hdr);color:#fff}}
.sc-bar{{display:flex;align-items:center;gap:12px;padding:8px 10px;background:#fff;border:1px solid var(--brd);border-radius:8px;margin-bottom:8px;flex-shrink:0;overflow-x:auto}}
.sc-view-toggle{{display:flex;gap:0;border:1.5px solid var(--brd);border-radius:6px;overflow:hidden;flex-shrink:0}}
.sc-vbtn{{padding:5px 14px;font-size:11px;font-weight:700;border:none;background:#fff;color:#64748B;cursor:pointer}}
.sc-vbtn.active{{background:#475569;color:#fff}}
.sc-metric-pills{{display:flex;gap:4px;flex-wrap:nowrap}}
.sc-mpill{{padding:5px 12px;font-size:11px;font-weight:600;border:1.5px solid var(--brd);border-radius:18px;background:#fff;color:#64748B;cursor:pointer;white-space:nowrap;transition:all .15s}}
.sc-mpill.active{{background:var(--hdr);color:#fff;border-color:var(--hdr)}}
.sc-mpill:hover:not(.active){{border-color:var(--hdr);color:var(--hdr)}}
.sc-table-wrap{{flex:1;overflow:auto;-webkit-overflow-scrolling:touch}}
.sc-tbl2{{border-collapse:collapse;width:max-content;min-width:100%;font-size:13px}}
.sc-tbl2 th,.sc-tbl2 td{{padding:10px 16px;text-align:right;white-space:nowrap;border-bottom:1px solid #F1F5F9}}
.sc-tbl2 thead th{{background:#F8FAFC;color:#94A3B8;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;position:sticky;top:0;z-index:5}}
.sc-tbl2 thead th.sc-th-name{{text-align:left!important;min-width:160px;position:sticky;left:0;z-index:11;background:#F8FAFC!important}}
.sc-tbl2 thead th.sc-th-rr{{background:#DBEAFE!important;color:#2563EB!important}}
.sc-tbl2 tbody td:first-child{{text-align:left;font-weight:600;color:#1E293B;position:sticky;left:0;z-index:2;background:inherit;min-width:160px}}
.sc-tbl2 tbody tr{{background:#fff}}
.sc-tbl2 tbody tr:nth-child(even){{background:#FAFBFC}}
.sc-tbl2 tbody tr:hover{{background:#F0F4FF}}
.sc-tbl2 .rr-cell{{color:#2563EB;font-weight:700}}
.sc-tbl2 .zv{{color:#CBD5E1}}
.sc-tbl2 .rv-pos{{color:#22C55E}}
.sc-tbl2 .rv-neg{{color:#EF4444}}
.sc-abs{{font-size:9px;color:#94A3B8;font-weight:400;margin-top:1px}}
@media(max-width:480px){{
  .sc-top{{flex-direction:column;gap:8px}}
  .sc-title{{font-size:16px}}
  .sc-bar{{gap:6px;padding:6px 8px}}
  .sc-mpill{{padding:4px 8px;font-size:10px}}
  .sc-tbl2{{font-size:11px}}
  .sc-tbl2 th,.sc-tbl2 td{{padding:6px 10px}}
}}
</style>

<div class="tab-content" id="tab-scorecard">
{scorecard_html}
<div class="ft">Scorecard mensual · Mar* = parcial 21/Mar · RR = Mar* / {MAR_ELAPSED_BIZ} d.h. × {MAR_BIZ_DAYS} · Tasas coloreadas verde/rojo vs Feb · Ordenado por RR desc</div>
</div>

<script>
(function(){{
  var curM='3', curCh='all', curView='mo';

  function val(arr,m){{
    if(!arr) return 0;
    if(m==='0'||m==='1'||m==='2'||m==='3') return arr[parseInt(m)];
    if(m==='qm') return arr[0]>0?arr[1]/arr[0]*100:0;
    if(m==='ma') return arr[1]>0?arr[2]/arr[1]*100:0;
    if(m==='ap') return arr[2]>0?arr[3]/arr[2]*100:0;
    if(m==='qp') return arr[0]>0?arr[3]/arr[0]*100:0;
    return 0;
  }}
  function isR(m){{return m==='qm'||m==='ma'||m==='ap'||m==='qp';}}
  function fV(v,m){{
    if(isR(m)) return v===0?'0.0%':v.toFixed(1)+'%';
    return v===0?'0':Math.round(v).toLocaleString('es-MX');
  }}
  // Get absolute numerator/denominator for rate metrics
  function getAbs(arr,m){{
    if(!arr) return [0,0];
    if(m==='qm') return [arr[1],arr[0]]; // Made/Quotes
    if(m==='ma') return [arr[2],arr[1]]; // Approved/Made
    if(m==='ap') return [arr[3],arr[2]]; // Purchased/Approved
    if(m==='qp') return [arr[3],arr[0]]; // Purchased/Quotes
    return [0,0];
  }}
  function fmtN(v){{return v===0?'0':Math.round(v).toLocaleString('es-MX');}}

  function getVal(g,col,m){{
    var src=curView==='mo'?g.mo:g.wk;
    return val(src[col],m);
  }}
  function getRawArr(g,col){{
    var src=curView==='mo'?g.mo:g.wk;
    return src[col]||[0,0,0,0];
  }}
  function rrVal(g,m){{
    if(curView==='mo'){{
      var v=val(g.mo['Mar 26'],m);
      return isR(m)?v:Math.round(v/_scElapsed*_scTotal);
    }}
    return 0;
  }}
  function sortKey(g){{
    if(curView==='mo') return rrVal(g,curM);
    return getVal(g,_scWkCols[1],curM); // sort by S12 (latest complete week)
  }}

  function mergeArr(a,b){{return [a[0]+b[0],a[1]+b[1],a[2]+b[2],a[3]+b[3]];}}
  function consolidate(list){{
    // Merge entries with same grupo name (sum TAS+ALIADO)
    var map={{}};
    list.forEach(function(g){{
      if(!map[g.n]){{
        map[g.n]={{n:g.n,ch:'ALL',mo:{{}},wk:{{}}}};
        _scMoCols.forEach(function(k){{map[g.n].mo[k]=[0,0,0,0];}});
        _scWkCols.forEach(function(k){{map[g.n].wk[k]=[0,0,0,0];}});
      }}
      var m=map[g.n];
      _scMoCols.forEach(function(k){{if(g.mo[k]) m.mo[k]=mergeArr(m.mo[k],g.mo[k]);}});
      _scWkCols.forEach(function(k){{if(g.wk[k]) m.wk[k]=mergeArr(m.wk[k],g.wk[k]);}});
    }});
    return Object.values(map);
  }}

  function render(){{
    var raw=_scD.filter(function(g){{return curCh==='all'||g.ch===curCh;}});
    var data=curCh==='all'?consolidate(raw):raw;
    data.sort(function(a,b){{return sortKey(b)-sortKey(a);}});
    var cols=curView==='mo'?_scMoCols:_scWkCols;

    // Header
    var th='<tr><th class="sc-th-name">ENTIDAD</th>';
    if(curView==='mo') th+='<th class="sc-th-rr">RUN RATE</th>';
    cols.forEach(function(c,i){{
      var star=(curView==='mo'&&i===0)||(curView==='wk'&&i===0)?'*':'';
      th+='<th>'+c+star+'</th>';
    }});
    th+='</tr>';
    document.getElementById('sc-thead').innerHTML=th;

    // Body
    var html='',tQ=0,tP=0;
    data.forEach(function(g){{
      tQ+=val(g.mo['Mar 26'],'0');
      tP+=val(g.mo['Mar 26'],'3');
      html+='<tr><td>'+g.n+'</td>';
      if(curView==='mo'){{
        var rv=rrVal(g,curM);
        var rrContent=rv===0?'<span class="zv">0</span>':fV(rv,curM);
        if(isR(curM)){{
          var ab=getAbs(g.mo['Mar 26'],curM);
          var rrNum=Math.round(ab[0]/_scElapsed*_scTotal);
          var rrDen=Math.round(ab[1]/_scElapsed*_scTotal);
          rrContent+='<div class="sc-abs">'+fmtN(rrNum)+'/'+fmtN(rrDen)+'</div>';
        }}
        html+='<td class="rr-cell">'+rrContent+'</td>';
      }}
      cols.forEach(function(col,ci){{
        var v=getVal(g,col,curM);
        var cls='';
        if(v===0){{ cls=' class="zv"'; }}
        else if(ci<cols.length-1){{ // has a previous period to compare
          var prevCol=cols[ci+1]; // next in array = previous in time
          var pv=getVal(g,prevCol,curM);
          if(pv!==0){{
            cls=v>pv?' class="rv-pos"':v<pv?' class="rv-neg"':'';
          }}
        }}
        var cellContent=fV(v,curM);
        if(isR(curM)){{
          var ab=getAbs(getRawArr(g,col),curM);
          cellContent+='<div class="sc-abs">'+fmtN(ab[0])+'/'+fmtN(ab[1])+'</div>';
        }}
        html+='<td'+cls+'>'+cellContent+'</td>';
      }});
      html+='</tr>';
    }});
    document.getElementById('sc-tbody2').innerHTML=html;
    document.getElementById('sc-sub').textContent=data.length+' grupos · '+Math.round(tQ).toLocaleString('es-MX')+' quotes · '+Math.round(tP).toLocaleString('es-MX')+' compras · Mar 2026';
  }}

  // Channel filter
  document.querySelectorAll('.sc-ch-btn').forEach(function(b){{
    b.addEventListener('click',function(){{
      document.querySelectorAll('.sc-ch-btn').forEach(function(x){{x.classList.remove('active');}});
      b.classList.add('active'); curCh=b.dataset.ch; render();
    }});
  }});
  // Metric pills
  document.querySelectorAll('.sc-mpill').forEach(function(b){{
    b.addEventListener('click',function(){{
      document.querySelectorAll('.sc-mpill').forEach(function(x){{x.classList.remove('active');}});
      b.classList.add('active'); curM=b.dataset.m; render();
    }});
  }});
  // View toggle (Mensual/Semanal)
  document.querySelectorAll('.sc-vbtn').forEach(function(b){{
    b.addEventListener('click',function(){{
      document.querySelectorAll('.sc-vbtn').forEach(function(x){{x.classList.remove('active');}});
      b.classList.add('active'); curView=b.dataset.v; render();
    }});
  }});

  window._scInit=false;
  window.renderScorecard=function(){{ render(); }};
}})();
</script>

<script>
function switchTab(id, btn) {{
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  btn.classList.add('active');
  if (id === 'charts' && !window._chartsInit) {{
    window._chartsInit = true;
    initCharts();
  }}
  if (id === 'scorecard' && !window._scInit) {{
    window._scInit = true;
    renderScorecard();
  }}
}}
</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script>
function initCharts() {{
  Chart.defaults.font.family = "-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif";
  Chart.defaults.font.size = 11;

  const dataLabels = {{
    id:'dLbl',
    afterDatasetsDraw(chart) {{
      const ctx = chart.ctx;
      chart.data.datasets.forEach((ds,di) => {{
        const meta = chart.getDatasetMeta(di);
        if(!meta.hidden) meta.data.forEach((bar,i) => {{
          const val = ds.data[i];
          if(val === undefined || val === null) return;
          if(chart.options.scales?.x?.stacked && di < chart.data.datasets.length - 1) return;
          ctx.save();
          ctx.font = 'bold 10px -apple-system,system-ui,sans-serif';
          ctx.textAlign = 'center';
          ctx.fillStyle = '#1E293B';
          let total = 0;
          chart.data.datasets.forEach((d2,d2i) => {{
            if(!chart.getDatasetMeta(d2i).hidden) total += (d2.data[i]||0);
          }});
          ctx.fillText(total, bar.x, bar.y - 6);
          ctx.restore();
        }});
      }});
    }}
  }};

  const wfLbl = {{
    id:'wfLbl',
    afterDatasetsDraw(chart) {{
      const ctx=chart.ctx, meta=chart.getDatasetMeta(0);
      ctx.save();
      ctx.font='bold 11px -apple-system,system-ui,sans-serif';
      ctx.textAlign='center';
      ctx.fillStyle='#1E293B';
      const labels = chart.data._wfLabels || [];
      meta.data.forEach((b,i) => {{
        if(labels[i] !== undefined) ctx.fillText(labels[i], b.x, b.y - 8);
      }});
      ctx.restore();
    }}
  }};

  // 1. Stacked bar (canvas c1)
  const labels1 = ["Oct", "Nov", "Dic", "Ene", "Feb", "Mar MTD", "Mar RR"];
  const tasP = {json.dumps(chart_data['tas_p'] + [chart_data['tas_rr']])};
  const aliP = {json.dumps(chart_data['ali_p'] + [chart_data['ali_rr']])};
  new Chart('c1', {{
    type:'bar',
    data:{{
      labels: labels1,
      datasets:[
        {{label:'TAS',data:tasP,backgroundColor:labels1.map((l,i)=>i===labels1.length-1?'rgba(27,42,74,0.4)':'#1B2A4A'),borderColor:labels1.map((l,i)=>i===labels1.length-1?'#1B2A4A':'transparent'),borderWidth:labels1.map((l,i)=>i===labels1.length-1?2:0),borderDash:labels1.map((l,i)=>i===labels1.length-1?[4,3]:undefined),borderRadius:3}},
        {{label:'Aliado',data:aliP,backgroundColor:labels1.map((l,i)=>i===labels1.length-1?'rgba(59,125,221,0.4)':'#3B7DDD'),borderColor:labels1.map((l,i)=>i===labels1.length-1?'#3B7DDD':'transparent'),borderWidth:labels1.map((l,i)=>i===labels1.length-1?2:0),borderRadius:3}}
      ]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'top',labels:{{boxWidth:12,padding:8,font:{{size:10}}}}}},
        tooltip:{{mode:'index',intersect:false,callbacks:{{afterBody:function(items){{return 'Total: '+items.reduce((s,i)=>s+i.raw,0)}}}}}}}},
      scales:{{x:{{stacked:true,grid:{{display:false}}}},y:{{stacked:true,beginAtZero:true,title:{{display:true,text:'Compras',font:{{size:10}}}},grid:{{color:'#E2E8F0'}}}}}}
    }},
    plugins:[dataLabels, {{
      id:'annotations',
      afterDraw(chart) {{
        const ctx = chart.ctx;
        const meta = chart.getDatasetMeta(1);
        const febBar = meta.data[4];
        const rrBar = meta.data[6];
        if(!febBar || !rrBar) return;
        ctx.save();
        const febTotal = {chart_data['tas_p'][4]} + {chart_data['ali_p'][4]};
        const febY = chart.scales.y.getPixelForValue(febTotal);
        ctx.strokeStyle = '#F59E0B';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([4,4]);
        ctx.beginPath();
        ctx.moveTo(febBar.x, febY);
        ctx.lineTo(rrBar.x + 20, febY);
        ctx.stroke();
        ctx.setLineDash([]);
        const rrTotal = {chart_data['tas_rr']} + {chart_data['ali_rr']};
        const rrY = chart.scales.y.getPixelForValue(rrTotal);
        const bracketX = rrBar.x + 18;
        const pctVsFeb = Math.round((rrTotal - febTotal) / febTotal * 100);
        ctx.strokeStyle = '#F59E0B';
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(bracketX, febY);
        ctx.lineTo(bracketX, rrY - 6);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(bracketX - 4, febY); ctx.lineTo(bracketX + 4, febY);
        ctx.moveTo(bracketX - 4, rrY - 6); ctx.lineTo(bracketX + 4, rrY - 6);
        ctx.stroke();
        ctx.font = 'bold 10px -apple-system,system-ui,sans-serif';
        ctx.fillStyle = '#F59E0B';
        ctx.textAlign = 'left';
        const labelY = (febY + rrY) / 2;
        ctx.fillText('+' + pctVsFeb + '% vs Feb', bracketX + 6, labelY + 3);
        ctx.restore();
      }}
    }}]
  }});

  // 2. Waterfall Total (canvas c2)
  const feb={chart_data['feb_total']},tasD={chart_data['tas_delta']},aliD={chart_data['ali_delta']},mar={chart_data['mar_total']};
  {{
    const d = {{
      labels:['Feb MTD','Δ TAS','Δ Aliado','Mar MTD'],
      datasets:[{{
        data:[[0,feb],[feb,feb+tasD],[feb+tasD,feb+tasD+aliD],[0,mar]],
        backgroundColor:['#94A3B8',tasD>=0?'#22C55E':'#EF4444',aliD>=0?'#22C55E':'#EF4444','#1B2A4A'],
        borderRadius:3
      }}],
      _wfLabels:[feb,(tasD>=0?'+':'')+tasD,(aliD>=0?'+':'')+aliD,mar]
    }};
    new Chart('c2', {{
      type:'bar',data:d,
      options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{grid:{{display:false}}}},y:{{beginAtZero:true,title:{{display:true,text:'Compras',font:{{size:10}}}}}}}}}},
      plugins:[wfLbl]
    }});
  }}

  // 3. TAS waterfall by funnel stage (canvas c3)
  {{
    const q={chart_data['tas_deltas'][0]},m={chart_data['tas_deltas'][1]},a={chart_data['tas_deltas'][2]},p={chart_data['tas_deltas'][3]};
    const fQ={chart_data['tas_feb_funnel'][0]},fM={chart_data['tas_feb_funnel'][1]},fA={chart_data['tas_feb_funnel'][2]},fP={chart_data['tas_feb_funnel'][3]};
    const mQ={chart_data['tas_mar_funnel'][0]},mM={chart_data['tas_mar_funnel'][1]},mA={chart_data['tas_mar_funnel'][2]},mP={chart_data['tas_mar_funnel'][3]};
    const d = {{
      labels:['Quotes','Made','Approved','Purchased'],
      datasets:[
        {{label:'Feb MTD',data:[fQ,fM,fA,fP],backgroundColor:'rgba(148,163,184,0.5)',borderColor:'#94A3B8',borderWidth:1,borderRadius:3}},
        {{label:'Mar MTD',data:[mQ,mM,mA,mP],backgroundColor:'#1B2A4A',borderRadius:3}}
      ],
    }};
    new Chart('c3', {{
      type:'bar',data:d,
      options:{{responsive:true,maintainAspectRatio:false,
        layout:{{padding:{{bottom:35}}}},
        plugins:{{legend:{{position:'top',labels:{{boxWidth:12,padding:8,font:{{size:10}}}}}}}},
        scales:{{x:{{grid:{{display:false}}}},y:{{beginAtZero:true,title:{{display:true,text:'Volumen',font:{{size:10}}}}}}}}
      }},
      plugins:[{{
        id:'deltaLbl',
        afterDatasetsDraw(chart) {{
          const ctx=chart.ctx;
          const meta=chart.getDatasetMeta(1);
          const deltas=[q,m,a,p];
          ctx.save();ctx.font='bold 10px -apple-system,system-ui,sans-serif';ctx.textAlign='center';
          meta.data.forEach((bar,i) => {{
            const v=deltas[i];
            ctx.fillStyle=v>=0?'#22C55E':'#EF4444';
            ctx.fillText((v>=0?'+':'')+v, bar.x, bar.y-6);
          }});
          const area = chart.chartArea;
          ctx.font='bold 10px -apple-system,system-ui,sans-serif';
          ctx.fillStyle='#2B478B';ctx.textAlign='left';
          ctx.fillText('{chart_data["feb_tas_cells"]} cél.→{chart_data["mar_tas_cells"]} cél. ({chart_data["mar_tas_cells"]-chart_data["feb_tas_cells"]:+d})  |  Q/cél: {chart_data["feb_q_per_cell"]}→{chart_data["mar_q_per_cell"]} ({chart_data["q_per_cell_pct"]:+d}%)', area.left, area.bottom + 28);
          ctx.restore();
        }}
      }}]
    }});
  }}

  // 4. Aliado waterfall by funnel stage (canvas c4)
  {{
    const q={chart_data['ali_deltas'][0]},m={chart_data['ali_deltas'][1]},a={chart_data['ali_deltas'][2]},p={chart_data['ali_deltas'][3]};
    const fQ={chart_data['ali_feb_funnel'][0]},fM={chart_data['ali_feb_funnel'][1]},fA={chart_data['ali_feb_funnel'][2]},fP={chart_data['ali_feb_funnel'][3]};
    const mQ={chart_data['ali_mar_funnel'][0]},mM={chart_data['ali_mar_funnel'][1]},mA={chart_data['ali_mar_funnel'][2]},mP={chart_data['ali_mar_funnel'][3]};
    const d = {{
      labels:['Quotes','Made','Approved','Purchased'],
      datasets:[
        {{label:'Feb MTD',data:[fQ,fM,fA,fP],backgroundColor:'rgba(148,163,184,0.5)',borderColor:'#94A3B8',borderWidth:1,borderRadius:3}},
        {{label:'Mar MTD',data:[mQ,mM,mA,mP],backgroundColor:'#3B7DDD',borderRadius:3}}
      ],
    }};
    new Chart('c4', {{
      type:'bar',data:d,
      options:{{responsive:true,maintainAspectRatio:false,
        layout:{{padding:{{bottom:35}}}},
        plugins:{{legend:{{position:'top',labels:{{boxWidth:12,padding:8,font:{{size:10}}}}}}}},
        scales:{{x:{{grid:{{display:false}}}},y:{{beginAtZero:true,title:{{display:true,text:'Volumen',font:{{size:10}}}}}}}}
      }},
      plugins:[{{
        id:'deltaLbl2',
        afterDatasetsDraw(chart) {{
          const ctx=chart.ctx;
          const meta=chart.getDatasetMeta(1);
          const deltas=[q,m,a,p];
          ctx.save();ctx.font='bold 10px -apple-system,system-ui,sans-serif';ctx.textAlign='center';
          meta.data.forEach((bar,i) => {{
            const v=deltas[i];
            ctx.fillStyle=v>=0?'#22C55E':'#EF4444';
            ctx.fillText((v>=0?'+':'')+v, bar.x, bar.y-6);
          }});
          const area = chart.chartArea;
          ctx.font='bold 10px -apple-system,system-ui,sans-serif';
          ctx.fillStyle='#3B7DDD';ctx.textAlign='left';
          ctx.fillText('{chart_data["feb_ali_cells"]} aliados→{chart_data["mar_ali_cells"]} ({chart_data["mar_ali_cells"]-chart_data["feb_ali_cells"]:+d})  |  Q/aliado: {chart_data["feb_q_per_ali"]}→{chart_data["mar_q_per_ali"]} ({chart_data["q_per_ali_pct"]:+d}%)', area.left, area.bottom + 28);
          ctx.restore();
        }}
      }}]
    }});
  }}

  // 5. Conversion TAS (canvas c5)
  const moLbl = ["Oct", "Nov", "Dic", "Ene", "Feb", "Mar MTD"];
  new Chart('c5', {{
    type:'line',
    data:{{
      labels:moLbl,
      datasets:[
        {{label:'Q→M%',data:{json.dumps(chart_data['tas_qm'])},borderColor:'#1B2A4A',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'M→A%',data:{json.dumps(chart_data['tas_ma'])},borderColor:'#64748B',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'A→P%',data:{json.dumps(chart_data['tas_ap'])},borderColor:'#22C55E',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'Q→P%',data:{json.dumps(chart_data['tas_qp'])},borderColor:'#8B5CF6',borderWidth:2.5,pointRadius:5,tension:.3,fill:false}}
      ]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,padding:8,font:{{size:10}},usePointStyle:true}}}}}},
      scales:{{x:{{grid:{{display:false}}}},y:{{beginAtZero:true,title:{{display:true,text:'%',font:{{size:10}}}},ticks:{{callback:v=>v+'%'}}}}}},
      interaction:{{mode:'index',intersect:false}}
    }},
    plugins:[{{
      id:'ptLbl6',
      afterDatasetsDraw(chart) {{
        const ctx=chart.ctx;
        chart.data.datasets.forEach((ds,di) => {{
          const meta=chart.getDatasetMeta(di);
          if(meta.hidden) return;
          meta.data.forEach((pt,i) => {{
            ctx.save();ctx.font='9px -apple-system,system-ui,sans-serif';ctx.textAlign='center';
            ctx.fillStyle=ds.borderColor;
            ctx.fillText(ds.data[i]+'%', pt.x, pt.y-8);
            ctx.restore();
          }});
        }});
      }}
    }}]
  }});

  // 6. Conversion Aliado (canvas c6)
  new Chart('c6', {{
    type:'line',
    data:{{
      labels:moLbl,
      datasets:[
        {{label:'Q→M%',data:{json.dumps(chart_data['ali_qm'])},borderColor:'#3B7DDD',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'M→A%',data:{json.dumps(chart_data['ali_ma'])},borderColor:'#64748B',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'A→P%',data:{json.dumps(chart_data['ali_ap'])},borderColor:'#00B48A',borderWidth:2,pointRadius:4,tension:.3,fill:false}},
        {{label:'Q→P%',data:{json.dumps(chart_data['ali_qp'])},borderColor:'#D946EF',borderWidth:2.5,pointRadius:5,tension:.3,fill:false}}
      ]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,padding:8,font:{{size:10}},usePointStyle:true}}}}}},
      scales:{{x:{{grid:{{display:false}}}},y:{{beginAtZero:true,title:{{display:true,text:'%',font:{{size:10}}}},ticks:{{callback:v=>v+'%'}}}}}},
      interaction:{{mode:'index',intersect:false}}
    }},
    plugins:[{{
      id:'ptLbl7',
      afterDatasetsDraw(chart) {{
        const ctx=chart.ctx;
        chart.data.datasets.forEach((ds,di) => {{
          const meta=chart.getDatasetMeta(di);
          if(meta.hidden) return;
          meta.data.forEach((pt,i) => {{
            ctx.save();ctx.font='9px -apple-system,system-ui,sans-serif';ctx.textAlign='center';
            ctx.fillStyle=ds.borderColor;
            ctx.fillText(ds.data[i]+'%', pt.x, pt.y-8);
            ctx.restore();
          }});
        }});
      }}
    }}]
  }});

  // 7. Region horizontal bar (canvas c7)
  new Chart('c7', {{
    type:'bar',
    data:{{
      labels:["Kavak Lerma", "Kavak GDL", "Kavak QRO", "Kavak MTY", "Kavak CDMX"],
      datasets:[
        {{label:'TAS',data:{json.dumps(chart_data['region_tas'])},backgroundColor:'#1B2A4A',borderRadius:3}},
        {{label:'Aliado',data:{json.dumps(chart_data['region_ali'])},backgroundColor:'#3B7DDD',borderRadius:3}}
      ]
    }},
    options:{{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'top',labels:{{boxWidth:12,padding:8,font:{{size:10}}}}}}}},
      scales:{{x:{{stacked:true,beginAtZero:true,title:{{display:true,text:'Compras Mar MTD',font:{{size:10}}}}}},y:{{stacked:true,grid:{{display:false}}}}}}
    }},
    plugins:[{{
      id:'hLbl',
      afterDatasetsDraw(chart) {{
        const ctx=chart.ctx;
        const lastDs = chart.data.datasets.length-1;
        const meta=chart.getDatasetMeta(lastDs);
        ctx.save();ctx.font='bold 10px -apple-system,system-ui,sans-serif';ctx.fillStyle='#1E293B';ctx.textBaseline='middle';
        meta.data.forEach((bar,i) => {{
          let total=0;
          chart.data.datasets.forEach((d,di) => {{if(!chart.getDatasetMeta(di).hidden)total+=d.data[i]||0}});
          ctx.fillText(total, bar.x+6, bar.y);
        }});
        ctx.restore();
      }}
    }}]
  }});
}}
</script>
</body></html>'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(full_html)

print(f"\n✅ Generated: {OUTPUT}")
print(f"   Records: {len(all_records)}")
print(f"   TAS groups: {len(groups_tas)} | Aliado groups: {len(groups_ali)}")
print(f"   Data through: Mar 21, 2026 (purchases from CSV8)")
print(f"   Total Mar MTD purchases: {agg_total['Mar MTD'][3]}")
print(f"   TAS: {agg_tas_total['Mar MTD'][3]} | Aliado: {agg_ali_total['Mar MTD'][3]}")
