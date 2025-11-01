#!/usr/bin/env python3
"""
Extract tables from the GAP PDF, normalize them, and produce:
 - gap-2.5.2-full.xlsx  (sheet: 'gap-2.5.2')
 - assets/data/equipment-distances.json (mapping for the UI)

Usage:
  python3 scripts/extract_gap_tables.py path/to/gap.pdf

Requirements:
  - Python 3.8+
  - pip install pandas openpyxl tabula-py camelot-py[cv]
  - Java (for tabula)
  - (optional) Ghostscript and Tk/Poppler if Camelot lattice requires it

Notes:
  - The script tries Camelot (lattice, then stream) first then falls back to tabula.
  - Table layout in the PDF can vary; check the produced Excel and do manual cleanup if needed.
"""
import sys
import os
import json
from pathlib import Path
import pandas as pd

# Output paths
OUT_XLSX = Path("gap-2.5.2-full.xlsx")
OUT_JSON = Path("assets/data/equipment-distances.json")
OUT_JSON.parent.mkdir(parents=True, exist_ok=True)

def try_camelot(path):
    try:
        import camelot
    except Exception:
        print("camelot not available", file=sys.stderr)
        return []
    tables = []
    # try lattice then stream
    for flavor in ("lattice", "stream"):
        try:
            print(f"Trying Camelot flavor={{flavor}} ...")
            tlist = camelot.read_pdf(path, pages="all", flavor=flavor, strip_text="\n")
            print(f"Camelot found {{len(tlist)}} tables with flavor={{flavor}}")
            for t in tlist:
                try:
                    df = t.df
                    tables.append(df)
                except Exception as e:
                    print("Camelot table -> df failed:", e)
            if tables:
                break
        except Exception as e:
            print("Camelot read_pdf error:", e)
    return tables

def try_tabula(path):
    try:
        import tabula
    except Exception:
        print("tabula-py not available", file=sys.stderr)
        return []
    tables = []
    try:
        print("Trying tabula.read_pdf (multiple_tables=True) ...")
        dfs = tabula.read_pdf(path, pages="all", multiple_tables=True)
        print(f"tabula returned {{len(dfs)}} tables")
        for df in dfs:
            if isinstance(df, pd.DataFrame):
                tables.append(df)
    except Exception as e:
        print("tabula.read_pdf error:", e)
    return tables

def normalize_tables(dfs):
    """
    Heuristic normalization:
    - For each dataframe, try to detect columns that correspond to equipment pair and distance.
    - Produce a unified DataFrame with columns: Equipment A, Equipment B, Distance, Note
    """
    rows = []
    for df in dfs:
        # drop completely empty rows
        df = df.dropna(how="all")
        if df.shape[0] == 0 or df.shape[1] == 0:
            continue

        # standardize column names
        cols = [str(c).strip() for c in df.columns.tolist()]

        # heuristic 1: if there are exactly 4 columns assume [A, B, Distance, Note]
        if df.shape[1] == 4:
            for _, r in df.iterrows():
                rows.append({
                    "Equipment A": str(r.iloc[0]).strip(),
                    "Equipment B": str(r.iloc[1]).strip(),
                    "Distance": str(r.iloc[2]).strip(),
                    "Note": str(r.iloc[3]).strip(),
                })
            continue

        # heuristic 2: if 3 columns assume [A, B, Distance/Note]
        if df.shape[1] == 3:
            for _, r in df.iterrows():
                rows.append({
                    "Equipment A": str(r.iloc[0]).strip(),
                    "Equipment B": str(r.iloc[1]).strip(),
                    "Distance": str(r.iloc[2]).strip(),
                    "Note": "",
                })
            continue

        # heuristic 3: if 2 columns assume "A - B" in first column, distance in second
        if df.shape[1] == 2:
            for _, r in df.iterrows():
                left = str(r.iloc[0]).strip()
                right = str(r.iloc[1]).strip()
                # try to split left into A and B (by common separators)
                sep = None
                for s in [" - ", " — ", "–", "/", " with ", " and ", ";"]:
                    if s in left:
                        sep = s
                        break
                if sep:
                    parts = [p.strip() for p in left.split(sep, 1)]
                    a = parts[0]
                    b = parts[1] if len(parts) > 1 else ""
                else:
                    a = left
                    b = ""
                rows.append({
                    "Equipment A": a,
                    "Equipment B": b,
                    "Distance": right,
                    "Note": ""
                })
            continue

        # fallback: try to locate columns containing 'dist' or 'm' in header
        text = " ".join(cols).lower()
        if "distance" in text or "dist" in text or "m" in text:
            for _, r in df.iterrows():
                a = str(r.iloc[0]).strip() if df.shape[1] >= 1 else ""
                b = str(r.iloc[1]).strip() if df.shape[1] >= 2 else ""
                dist = str(r.iloc[-1]).strip()
                rows.append({"Equipment A": a, "Equipment B": b, "Distance": dist, "Note": ""})
            continue

        # last fallback: read each row as single description in Note
        for _, r in df.iterrows():
            joined = " | ".join([str(x).strip() for x in r.tolist() if str(x).strip()])
            rows.append({"Equipment A": "", "Equipment B": "", "Distance": "", "Note": joined})
    out_df = pd.DataFrame(rows, columns=["Equipment A", "Equipment B", "Distance", "Note"])
    out_df = out_df.dropna(how="all").reset_index(drop=True)
    return out_df

def to_json_map(df):
    """
    Convert rows to mapping "equipment a|equipment b" -> {distance, note}
    - lowercased keys
    """
    m = {}
    for _, r in df.iterrows():
        a = str(r["Equipment A"]).strip()
        b = str(r["Equipment B"]).strip()
        dist = str(r["Distance"]).strip()
        note = str(r["Note"]).strip()
        if not a and not b and not dist and not note:
            continue
        key_a = a.lower()
        key_b = b.lower() if b else a.lower()
        key = f"{{key_a}}|{{key_b}}"
        if not key.strip("|"):
            continue
        m[key] = {"distance": dist, "note": note}
    return m

def main(pdf_path):
    if not os.path.exists(pdf_path):
        print("PDF not found:", pdf_path)
        sys.exit(2)
    # try camelot
    dfs = try_camelot(pdf_path)
    if not dfs:
        dfs = try_tabula(pdf_path)
    if not dfs:
        print("No tables extracted by Camelot or Tabula. You may need to try a different tool or check the PDF.")
        sys.exit(1)

    out_df = normalize_tables(dfs)
    if out_df.empty:
        print("No usable table rows after normalization.")
        sys.exit(1)

    # write excel
    print(f"Writing Excel to {{OUT_XLSX}}")
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="gap-2.5.2", index=False)

    # write JSON mapping for UI
    mapping = to_json_map(out_df)
    print(f"Rows -> JSON mapping size: {{len(mapping)}}")
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2, ensure_ascii=False)

    print("Done. Files created:")
    print("  -", OUT_XLSX.resolve())
    print("  -", OUT_JSON.resolve())

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 scripts/extract_gap_tables.py path/to/gap.pdf")
        sys.exit(1)
    main(sys.argv[1])