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


# -------- UI --------

st.title("📋 Trivedi ↔ Reliance Reconciliation")
st.caption("Billing Annexure ki TCN/Trip Number Reliance me dhundhi jayegi. Phir TMS status check, phir field-by-field compare. Reliance ke empty fields Vendor Summary se fill honge.")

c1, c2, c3 = st.columns(3)
with c1:
    bill_file = st.file_uploader("Billing Portal (.xlsx)", type=['xlsx'], key='bill')
with c2:
    rel_file = st.file_uploader("Reliance Portal (.xlsx)", type=['xlsx'], key='rel')
with c3:
    vendor_file = st.file_uploader("Vendor Summary (.csv) — optional", type=['csv'], key='vendor')

if rel_file and bill_file:
    if st.button("🔍 Reconcile", type='primary'):
        try:
            bill_rows = load_billing(bill_file.read())
            rel_rows = load_reliance(rel_file.read())
            vendor_rows = load_vendor_summary(vendor_file.read()) if vendor_file else None

            msg = f"Loaded: {len(bill_rows)} Billing | {len(rel_rows)} Reliance"
            if vendor_rows is not None:
                msg += f" | {len(vendor_rows)} Vendor Summary"
            st.success(msg)

            results = reconcile(rel_rows, bill_rows, vendor_rows)
            df = pd.DataFrame(results)

            counts = df['Result'].value_counts().to_dict()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("✅ Matched", counts.get('✅ MATCHED', 0))
            m2.metric("⚠️ Mismatch", counts.get('⚠️ MISMATCH', 0))
            m3.metric("🚫 Rejected", counts.get('🚫 REJECTED', 0))
            m4.metric("❌ Not in Reliance", counts.get('❌ NOT IN RELIANCE', 0))

            if vendor_rows is not None:
                fb_count = df['Fallback Used'].astype(bool).sum()
                st.info(f"🔁 Vendor Summary fallback used in **{fb_count}** rows")

            filter_opt = st.multiselect(
                "Filter by result",
                options=df['Result'].unique().tolist(),
                default=df['Result'].unique().tolist(),
            )
            view = df[df['Result'].isin(filter_opt)]
            st.dataframe(view, use_container_width=True, hide_index=True)

            csv_bytes = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "⬇️ Download Results (CSV)",
                csv_bytes,
                file_name="reconciliation_result.csv",
                mime="text/csv",
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
else:
    st.info("👆 Billing + Reliance files upload karo (Vendor Summary optional).")