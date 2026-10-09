import pandas as pd
from openpyxl import load_workbook
import shutil, os
import sys

# 17 metrics to extract for every block
METRICS = [
    "YTM", "Modified Duration", "Macaulay Duration",
    "3 Months", "6 Months", "9 Months",
    "1 Year", "2 Years", "3 Years", "5 Years", "10 Years",
    "Sovereign", "AAA/A1+", "AA+", "AA/AA-", "Below AA-", "Cash and Fixed Deposits"
]

# Configuration for each duration block: (Kotak sheet name, schemes list, header row, first output row)
# Kotak renamed its debt categories in Sep 2026 (e.g. "Low Duration" -> "Ultra Short to Short Term").
# Scheme order must match the fund rows in the Estimation "Fund Ranking_Debt" sheet.
BLOCKS = [
    ("Ultra Short Term", [
        "Aditya Birla Sun Life Ultra Short Term Fund - Reg - Growth",
        "HDFC Ultra Short Term Fund - Reg - Growth",
        "SBI Ultra Short Term Fund - Growth",
        "ICICI Prudential Ultra Short Term Fund - Growth",
        "Kotak Ultra Short Term Fund - Reg - Growth",
        "Nippon India Ultra Short Term Fund - Reg - Growth",
        "UTI Ultra Short Term Fund - Growth"
    ], 6, 7),
    ("Ultra Short to Short Term", [
        "ICICI Prudential Ultra Short to Short Term Fund - Reg - Growth",
        "HDFC Ultra Short to Short Term Fund - Growth",
        "SBI Ultra Short to Short Term Fund - Growth",
        "Aditya Birla Sun Life Ultra Short to Short Term Fund - Reg - Growth",
        "Kotak Ultra Short to Short Term Fund - Reg - Growth",
        "Nippon India Ultra Short to Short Term Fund - Reg - Growth",
        "UTI Ultra Short to Short Term Fund - Reg - Growth"
    ], 18, 19),
    ("Short Term", [
        "ICICI Prudential Short Term Fund - Growth",
        "Kotak Short Term Fund - Reg - Growth",
        "HDFC Short Term Fund - Growth",
        "SBI Short Term Fund - Growth",
        "Aditya Birla Sun Life Short Term Fund - Reg - Growth",
        "Nippon India Short Term Fund - Reg - Growth",
        "UTI Short Term Fund - Reg - Growth"
    ], 30, 31),
    ("Medium Term", [
        "SBI Medium Term Fund - Growth",
        "ICICI Prudential Medium Term Fund - Growth",
        "HDFC Medium Term Fund - Growth",
        "Aditya Birla Sun Life Medium Term Fund - Reg - Growth",
        "Kotak Medium Term Fund - Reg - Growth",
        "Nippon India Medium Term Fund - Reg - Growth",
        "UTI Medium Term Fund - Reg - Growth"
    ], 41, 42),
    ("Medium to Long Term", [
        "ICICI Prudential Medium to Long Term Fund - Growth",
        "Aditya Birla Sun Life Medium to Long Term Fund - Reg - Growth",
        "Kotak Medium to Long Term Fund - Reg - Growth",
        "SBI Medium to Long Term Fund - Growth",
        "HDFC Medium to Long Term Fund - Growth",
        "Nippon India Medium to Long Term Fund - Reg - G P - Growth",
        "UTI Medium to Long Term Fund - Growth"
    ], 52, 53)
]

def flatten_kotak(kotak_path, sheet_name):
    print(f"[DEBUG] Reading Kotak sheet: {sheet_name}")
    raw = pd.read_excel(kotak_path, sheet_name=sheet_name, header=[4, 5], dtype=str)
    raw.columns = [f"{a or ''}{' ' if a and b else ''}{b or ''}".strip() for a, b in raw.columns]
    print(f"[DEBUG] Raw columns: {raw.columns.tolist()}")

    df = raw.rename(columns={raw.columns[0]: "Scheme"})
    df["Scheme"] = df["Scheme"].str.strip()
    df = df.map(lambda x: 0 if isinstance(x, str) and x.strip() == "--" else x)
    print(f"[DEBUG] Cleaned columns: {df.columns.tolist()}")
    return df

def process(est_path, kotak_path):
    print(f"[DEBUG] Starting process with Estimation: {est_path}, Kotak: {kotak_path}")
    base = os.path.dirname(est_path)
    output = os.path.join(base, "Estimation_V4_updated.xlsx")
    shutil.copy(est_path, output)
    print(f"[DEBUG] Copied estimation to output: {output}")

    wb = load_workbook(output)
    ws = wb['Fund Ranking_Debt']

    for sheet, schemes, header_row, first_row in BLOCKS:
        print(f"[DEBUG] Processing block: {sheet}")
        df_k = flatten_kotak(kotak_path, sheet)
        print(f"[DEBUG] df_k columns after flatten: {df_k.columns.tolist()}")

        # Build metric→column mapping
        col_map = {}
        for m in METRICS:
            matches = [c for c in df_k.columns if m in c]
            if not matches:
                print(f"[DEBUG] No column match for metric: {m}")
            exact = [c for c in matches if c.endswith(m)]
            col_map[m] = exact[0] if exact else (matches[0] if matches else None)
        print(f"[DEBUG] Column map: {col_map}")

        # Extract and round values
        data = {}
        for sc in schemes:
            if sheet == "Short Term":
                subset = df_k[df_k["Scheme"].str.contains(sc, na=False)]
            else:
                subset = df_k[df_k["Scheme"] == sc]
            if subset.empty:
                print(f"[DEBUG] Warning: '{sc}' not found in {sheet}")
                continue
            row = subset.iloc[0]
            data[sc] = {}
            for m in METRICS:
                raw_val = row[col_map[m]] if col_map[m] else None
                try:
                    val = round(float(raw_val), 2)
                except:
                    val = raw_val
                data[sc][m] = val
        print(f"[DEBUG] Extracted data keys: {list(data.keys())}")

        # Locate Estimation columns
        est_cols = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=header_row, column=col).value
            if val in METRICS:
                est_cols[val] = col
        print(f"[DEBUG] Estimation columns map: {est_cols}")

        # Write values into Estimation sheet
        for idx, sc in enumerate(schemes):
            row_idx = first_row + idx
            if sc in data:
                for m, col_idx in est_cols.items():
                    ws.cell(row=row_idx, column=col_idx, value=data[sc][m])
        print(f"[DEBUG] Finished writing block: {sheet}")

    wb.save(output)
    print(f"[DEBUG] Saved updated estimation file: {output}")
    return output

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python logic.py <Estimation_V4.xlsx> <Kotak_Report.xlsx>")
    else:
        process(sys.argv[1], sys.argv[2])
