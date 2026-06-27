# Trivedi ↔ Reliance Reconciliation Portal

Internal tool — matches Reliance portal export with Trivedi billing export, with optional Vendor Summary fallback for empty fields.

## Setup (one-time)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Run UI

```bash
source venv/bin/activate
streamlit run app.py
```

Opens at `http://localhost:8501`. Upload 2 (or 3) files → Reconcile.

## CLI

```bash
python reconcile.py reliance.xlsx billing.xlsx vendor_summary.csv
# vendor_summary.csv is optional
```

## Logic

Per Reliance row:

1. **Vendor Summary fallback** — if any of `[SO, Vehicle Number, Weight, Freight Cost]` is empty in Reliance, look up Vendor Summary by `TCN/Trip Number` and fill from there.
2. **SO match** with Billing Annexure → if missing: `NOT_FOUND_IN_BILLING`
3. **TMS Status** → if `Rejected`: stop, mark `REJECTED`
4. **Field-by-field compare** → `MATCHED` / `MISMATCH` (with per-field diff)

## Field mapping

### Reliance → Billing Annexure (comparison)
| Reliance | Billing | Type |
|---|---|---|
| SO Number/Ref No. | SO Number | match key |
| TCN/Trip Number | TCN/Trip Number | string |
| Vehicle Number | Vehicle Number | string |
| Weight | Vehicle Type | MT → kg |
| Freight Cost | Frieght as per Emptoris Contract | numeric |
| Loading Charges | Load Charges | numeric |
| Unloading Charges | Unload charges | numeric |
| Other Charges | Other charges | numeric |
| Detentation Charges | Detention Charges | numeric |
| Invoice No. | Invoice Number | string |
| Internal Ref No. | Internal Ref No | string |

### Reliance ← Vendor Summary (fallback only, when Reliance value is empty)
Looked up via `TCN/Trip Number` ↔ `Trip Number`.

| Reliance empty field | Vendor Summary source |
|---|---|
| SO Number/Ref No. | Planned Consignment |
| Vehicle Number | Vehicle Number |
| Weight | Vehicle Make Name (parses `9MT 22FT Open` → 9000 kg) |
| Freight Cost | Freight Cost |

Vendor Summary has duplicate Trip Numbers (~200 trips appear with multiple SOs), but trip-level fields (Vehicle, Freight, Weight) are identical across duplicates — first occurrence is used.

## Notes

- Trivedi billing portal exports contain HTML appended after the xlsx (corrupted download). Code auto-truncates at ZIP EOCD before parsing.
- Reliance sheet expected name: `Sheet1`. Billing sheet expected name: `Annexure`.

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `reconcile.py` | CLI / shared logic |
| `requirements.txt` | Dependencies |
