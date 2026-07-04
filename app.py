"""
Streamlit UI for Trivedi <-> Reliance reconciliation (TCN-driven).
Run: streamlit run app.py
"""
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO, StringIO
import csv
import re

st.set_page_config(page_title="Trivedi Reconciliation", layout="wide")


# -------- loaders --------

def fix_corrupted_xlsx(data: bytes) -> bytes:
    eocd = data.rfind(b'PK\x05\x06')
    if eocd < 0:
        return data
    comment_len = int.from_bytes(data[eocd + 20:eocd + 22], 'little')
    return data[:eocd + 22 + comment_len]


def _sheet_to_dicts(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else '' for h in rows[0]]
    return [dict(zip(headers, r)) for r in rows[1:] if any(c is not None and c != '' for c in r)]


def load_billing(data: bytes, sheet='Annexure'):
    wb = load_workbook(BytesIO(fix_corrupted_xlsx(data)), data_only=True)
    return _sheet_to_dicts(wb[sheet])


def load_reliance(data: bytes, sheet='Sheet1'):
    wb = load_workbook(BytesIO(data), data_only=True)
    return _sheet_to_dicts(wb[sheet])


def load_vendor_summary(data: bytes):
    return list(csv.DictReader(StringIO(data.decode('utf-8'))))


def load_detention(data: bytes):
    """Load all sheets from detention file, combine into one list indexed by LR number."""
    wb = load_workbook(BytesIO(data), data_only=True)
    all_rows = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # Find header row containing LR
        header_idx = None
        for i, row in enumerate(rows):
            if any(cell and 'lr' in str(cell).lower() for cell in row):
                header_idx = i
                break
        if header_idx is None:
            continue
        headers = [str(h).strip().lower() if h is not None else '' for h in rows[header_idx]]
        # Find LR column index
        lr_col = next((i for i, h in enumerate(headers)
                       if h in ('lr n.', 'lr no', 'lr no.', 'lr n', 'lr number', 'lr n. ')), None)
        # Find amount column index (unloading/amount)
        amt_col = next((i for i, h in enumerate(headers)
                        if any(k in h for k in ('unlod', 'unload', 'amount'))), None)
        if lr_col is None:
            continue
        for row in rows[header_idx + 1:]:
            lr_val = row[lr_col] if lr_col < len(row) else None
            amt_val = row[amt_col] if amt_col is not None and amt_col < len(row) else None
            if lr_val is None or str(lr_val).strip() == '':
                continue
            all_rows.append({
                'sheet': sheet_name,
                'lr_number': str(lr_val).strip(),
                'detention_unload_amount': amt_val,
            })
    return all_rows


# -------- helpers --------

def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == '')


def norm_str(v):
    return '' if v is None else str(v).strip().upper()


def norm_num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(',', ''))
    except ValueError:
        return None


def parse_mt_to_kg(v):
    if v is None:
        return None
    m = re.search(r'([\d.]+)\s*MT', str(v), re.IGNORECASE)
    return float(m.group(1)) * 1000 if m else None


# -------- vendor fallback --------

VENDOR_FALLBACK = {
    'SO Number/Ref No.': 'Planned Consignment',
    'Vehicle Number':    'Vehicle Number',
    'Freight Cost':      'Freight Cost',
    'Weight':            'Vehicle Make Name',
}


def build_vendor_index(rows):
    idx = {}
    for row in rows:
        tn = norm_str(row.get('Trip Number'))
        if tn and tn not in idx:
            idx[tn] = row
    return idx


def apply_vendor_fallback(rel, vendor_index):
    filled = dict(rel)
    used = []
    tn = norm_str(rel.get('TCN/Trip Number'))
    if not tn or tn not in vendor_index:
        return filled, used
    v = vendor_index[tn]
    for rel_col, vendor_col in VENDOR_FALLBACK.items():
        if is_empty(filled.get(rel_col)):
            val = v.get(vendor_col)
            if is_empty(val):
                continue
            if rel_col == 'Weight':
                kg = parse_mt_to_kg(val)
                if kg is not None:
                    filled[rel_col] = kg
                    used.append(rel_col)
            else:
                filled[rel_col] = val
                used.append(rel_col)
    return filled, used


FIELD_MAP = [
    ('SO Number/Ref No.',   'SO Number',                        'str'),
    ('Vehicle Number',      'Vehicle Number',                   'str'),
    ('Weight',              'Vehicle Type',                     'weight'),
    ('Freight Cost',        'Frieght as per Emptoris Contract', 'num'),
    ('Loading Charges',     'Load Charges',                     'num'),
    ('Unloading Charges',   'Unload charges',                   'num'),
    ('Other Charges',       'Other charges',                    'num'),
    ('Detentation Charges', 'Detention Charges',                'num'),
    ('Invoice No.',         'Invoice Number',                   'str'),
    ('Internal Ref No.',    'Internal Ref No',                  'str'),
]


def compare_row(rel, bill):
    matched, mismatched = [], []
    for r_col, b_col, kind in FIELD_MAP:
        r_val, b_val = rel.get(r_col), bill.get(b_col)
        if kind == 'str':
            ok = norm_str(r_val) == norm_str(b_val)
        elif kind == 'num':
            rn, bn = norm_num(r_val), norm_num(b_val)
            ok = rn is not None and bn is not None and abs(rn - bn) < 0.01
        else:
            rn = norm_num(r_val)
            bn = parse_mt_to_kg(b_val) if not isinstance(b_val, (int, float)) else float(b_val)
            ok = rn is not None and bn is not None and abs(rn - bn) < 0.01
        item = {'field': r_col, 'reliance': r_val, 'billing': b_val}
        (matched if ok else mismatched).append(item)
    return matched, mismatched


# -------- main reconcile (BILLING-driven, TCN match) --------

def reconcile(rel_rows, bill_rows, vendor_rows):
    rel_index = {}
    for row in rel_rows:
        tcn = norm_str(row.get('TCN/Trip Number'))
        if tcn:
            rel_index[tcn] = row

    vendor_index = build_vendor_index(vendor_rows) if vendor_rows else {}

    results = []
    for b in bill_rows:
        tcn = norm_str(b.get('TCN/Trip Number'))
        entry = {
            'TCN/Trip': b.get('TCN/Trip Number'),
            'SO Number': b.get('SO Number'),
            'Billing Invoice': b.get('Invoice Number'),
            'Tms Status': '',
            'RCPL Remark': '',
            'Fallback Used': '',
            'Matched Fields': '',
            'Mismatches': '',
        }

        if not tcn or tcn not in rel_index:
            entry['Result'] = '❌ NOT IN RELIANCE'
            results.append(entry)
            continue

        rel_row = rel_index[tcn]
        filled, used = apply_vendor_fallback(rel_row, vendor_index)
        entry['Tms Status'] = filled.get('Tms Status') or ''
        entry['RCPL Remark'] = filled.get('RCPL Remark') or ''
        entry['Fallback Used'] = ', '.join(used) if used else ''

        if norm_str(filled.get('Tms Status')) == 'REJECTED':
            entry['Result'] = '🚫 REJECTED'
            results.append(entry)
            continue

        matched, mismatches = compare_row(filled, b)
        entry['Matched Fields'] = ', '.join(m['field'] for m in matched)
        if not mismatches:
            entry['Result'] = '✅ MATCHED'
        else:
            entry['Result'] = '⚠️ MISMATCH'
            entry['Mismatches'] = ' | '.join(
                f"{m['field']}: R={m['reliance']} vs B={m['billing']}"
                for m in mismatches
            )
        results.append(entry)
    return results


# -------- detention reconcile (LR-driven) --------

def reconcile_detention(bill_rows, detention_rows):
    """Match billing LR No. with detention LR number, compare Unload charges."""
    # Index detention by LR number (normalized)
    det_index = {}
    for row in detention_rows:
        lr = norm_str(row.get('lr_number'))
        if lr:
            det_index.setdefault(lr, []).append(row)

    results = []
    for b in bill_rows:
        lr = norm_str(b.get('LR No.'))
        billing_unload = norm_num(b.get('Unload charges'))
        entry = {
            'LR No': b.get('LR No.'),
            'SO Number': b.get('SO Number'),
            'TCN/Trip': b.get('TCN/Trip Number'),
            'Billing Unload Amt': billing_unload,
            'Detention Unload Amt': '',
            'Detention Sheet': '',
        }

        if not lr or lr not in det_index:
            entry['Result'] = '❌ LR NOT IN DETENTION'
            entry['Diff'] = ''
            results.append(entry)
            continue

        det_row = det_index[lr][0]
        det_amt = norm_num(det_row.get('detention_unload_amount'))
        entry['Detention Unload Amt'] = det_amt
        entry['Detention Sheet'] = det_row.get('sheet', '')

        if billing_unload is not None and det_amt is not None and abs(billing_unload - det_amt) < 0.01:
            entry['Result'] = '✅ MATCHED'
            entry['Diff'] = 0
        else:
            entry['Result'] = '⚠️ MISMATCH'
            entry['Diff'] = (billing_unload or 0) - (det_amt or 0)
        results.append(entry)
    return results


# -------- UI --------

st.title("📋 Trivedi ↔ Reliance Reconciliation")

# File uploaders
c1, c2, c3, c4 = st.columns(4)
with c1:
    bill_file = st.file_uploader("Billing Portal (.xlsx)", type=['xlsx'], key='bill')
with c2:
    rel_file = st.file_uploader("Reliance Portal (.xlsx)", type=['xlsx'], key='rel')
with c3:
    vendor_file = st.file_uploader("Vendor Summary (.csv) — optional", type=['csv'], key='vendor')
with c4:
    detention_file = st.file_uploader("Detention Sheet (.xlsx) — optional", type=['xlsx'], key='detention')

if rel_file and bill_file:
    if st.button("🔍 Reconcile", type='primary'):
        try:
            bill_bytes = bill_file.read()
            bill_rows = load_billing(bill_bytes)
            rel_rows = load_reliance(rel_file.read())
            vendor_rows = load_vendor_summary(vendor_file.read()) if vendor_file else None
            detention_rows = load_detention(detention_file.read()) if detention_file else None

            msg = f"Loaded: {len(bill_rows)} Billing | {len(rel_rows)} Reliance"
            if vendor_rows:
                msg += f" | {len(vendor_rows)} Vendor Summary"
            if detention_rows is not None:
                msg += f" | {len(detention_rows)} Detention rows"
            st.success(msg)

            # ---- Tabs ----
            tab1, tab2 = st.tabs(["📊 Reliance Reconciliation", "🚛 Detention LR Check"])

            with tab1:
                results = reconcile(rel_rows, bill_rows, vendor_rows)
                df = pd.DataFrame(results)

                counts = df['Result'].value_counts().to_dict()
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("✅ Matched", counts.get('✅ MATCHED', 0))
                m2.metric("⚠️ Mismatch", counts.get('⚠️ MISMATCH', 0))
                m3.metric("🚫 Rejected", counts.get('🚫 REJECTED', 0))
                m4.metric("❌ Not in Reliance", counts.get('❌ NOT IN RELIANCE', 0))

                if vendor_rows:
                    fb_count = df['Fallback Used'].astype(bool).sum()
                    st.info(f"🔁 Vendor Summary fallback used in **{fb_count}** rows")

                filter_opt = st.multiselect(
                    "Filter by result",
                    options=df['Result'].unique().tolist(),
                    default=df['Result'].unique().tolist(),
                    key='filter1'
                )
                st.dataframe(df[df['Result'].isin(filter_opt)], use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download (CSV)", df.to_csv(index=False).encode(),
                                   "reconciliation_result.csv", "text/csv")

            with tab2:
                if detention_rows is None:
                    st.info("👆 Detention Sheet upload karo (4th uploader).")
                else:
                    det_results = reconcile_detention(bill_rows, detention_rows)
                    ddf = pd.DataFrame(det_results)

                    dc1, dc2, dc3 = st.columns(3)
                    dcounts = ddf['Result'].value_counts().to_dict()
                    dc1.metric("✅ Matched", dcounts.get('✅ MATCHED', 0))
                    dc2.metric("⚠️ Mismatch", dcounts.get('⚠️ MISMATCH', 0))
                    dc3.metric("❌ LR Not Found", dcounts.get('❌ LR NOT IN DETENTION', 0))

                    filter_opt2 = st.multiselect(
                        "Filter by result",
                        options=ddf['Result'].unique().tolist(),
                        default=ddf['Result'].unique().tolist(),
                        key='filter2'
                    )
                    st.dataframe(ddf[ddf['Result'].isin(filter_opt2)], use_container_width=True, hide_index=True)
                    st.download_button("⬇️ Download Detention Results (CSV)",
                                       ddf.to_csv(index=False).encode(),
                                       "detention_result.csv", "text/csv")

        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
else:
    st.info("👆 Billing + Reliance files upload karo.")