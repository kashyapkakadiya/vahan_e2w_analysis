#!/usr/bin/env python3
"""
extract_clean.py  -  Step 1 of 2
--------------------------------
Parse the Excel file downloaded from the Vahan dashboard and write a clean
"long" CSV (one row per maker per month) for analyse.py.

Dashboard settings used to produce the export
    Report      : Maker  x  Month Wise,  Calendar year 2026,  State = All India
    Vehicle cls : TWO WHEELER (Invalid Carriage) + TWO WHEELER (NT) + TWO WHEELER (T)
    Fuel        : PURE EV + ELECTRIC(BOV)
    URL         : https://vahan.parivahan.gov.in/vahan4dashboard/

Usage
    python src/extract_clean.py \
        --input data/raw/vahan_reportTable_2026.xlsx \
        --extracted-on 2026-10-31 \
        --out data/processed/vahan_e2w_long.csv

What it does
    1. finds the header row (the row holding JAN, FEB, ... month names)
    2. strips non-breaking spaces / extra whitespace from names
    3. converts Indian-style numbers ("4,70,369") to integers
    4. drops summary rows (any TOTAL row) and non-data rows
    5. checks every maker's reported TOTAL against the sum of its months
    6. adds a short display name for the well-known brands
    7. flags the month that contains the extraction date as PARTIAL
"""
import argparse
import calendar
import datetime as dt
import re
from pathlib import Path

import pandas as pd

MONTH_ABBR = [m.upper() for m in calendar.month_abbr[1:]]      # JAN ... DEC

# Vahan "maker" legal-entity name  ->  display brand.  Anything not listed keeps
# its Vahan name.  Only unambiguous renames / merges are done here.
BRAND_NAMES = {
    "TVS MOTOR COMPANY LTD": "TVS Motor",
    "BAJAJ AUTO LTD": "Bajaj Auto",
    "ATHER ENERGY LTD": "Ather Energy",
    "HERO MOTOCORP LTD": "Hero MotoCorp (Vida)",
    "OLA ELECTRIC TECHNOLOGIES PVT LTD": "Ola Electric",
    "GREAVES ELECTRIC MOBILITY LTD": "Greaves Electric (Ampere)",
    "RIVER MOBILITY PVT LTD": "River Mobility",
    "BGAUSS AUTO PRIVATE LIMITED": "BGauss Auto",
    "SIMPLEENERGY PVT LTD": "Simple Energy",
    "E-SPRINTO GREEN ENERGY PVT LTD": "E-Sprinto",
    "BOUNCE ELECTRIC 1 PVT LTD": "Bounce Electric",
    "REVOLT INTELLICORP PVT LTD": "Revolt Motors",
    "LECTRIX E VEHICLES PVT LTD": "Lectrix EV",
    "MOTOVOLT MOBILITY PVT LTD": "Motovolt Mobility",
    "OBEN ELECTRIC VEHICLES PVT LTD": "Oben Electric",
    "HONDA MOTORCYCLE AND SCOOTER INDIA (P) LTD": "Honda (HMSI)",
    "QUANTUM ENERGY LTD.": "Quantum Energy",
    "KINETIC GREEN ENERGY & POWER SOLUTIONS LTD": "Kinetic Green",
    "KINETIC GREEN ENERGY P SOL LTD": "Kinetic Green",          # same company, 13 units
    "ULTRAVIOLETTE AUTOMOTIVE PVT LTD": "Ultraviolette",
    "SUZUKI MOTORCYCLE INDIA PVT LTD": "Suzuki Motorcycle",
    "KLB KOMAKI PVT LTD": "Komaki (KLB)",
    "KINETIC WATTS AND VOLTS LTD": "Kinetic Watts & Volts",
    "ODYSSE ELECTRIC VEHICLES PVT LTD": "Odysse Electric",
    "OTHERS": "OTHERS (Vahan category)",
}


def clean_text(v) -> str:
    """Replace non-breaking spaces and collapse whitespace."""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def to_int(v):
    """'4,70,369' -> 470369 ; blanks / text -> None."""
    s = clean_text(v).replace(",", "")
    return int(s) if re.fullmatch(r"-?\d+", s) else None


def find_header_row(raw: pd.DataFrame) -> int:
    for i in range(min(len(raw), 15)):
        hits = sum(clean_text(x).upper() in MONTH_ABBR for x in raw.iloc[i])
        if hits >= 3:
            return i
    raise ValueError("No header row with month names (JAN, FEB, ...) found.")


def parse_vahan_export(path: Path, year: int) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    h = find_header_row(raw)
    header = [clean_text(x).upper() for x in raw.iloc[h]]
    month_cols = {j: c for j, c in enumerate(header) if c in MONTH_ABBR}

    # The maker column is the text column to the left of the first month column.
    first_month = min(month_cols)
    maker_col = first_month - 1
    total_col = next((j for j, c in enumerate(header) if c == "TOTAL"), None)
    if total_col is None:                       # TOTAL header sits in an earlier row
        for i in range(h):
            for j, x in enumerate(raw.iloc[i]):
                if clean_text(x).upper() == "TOTAL":
                    total_col = j

    rows, mismatches = [], []
    for _, r in raw.iloc[h + 1:].iterrows():
        sno, maker = clean_text(r.iloc[0]), clean_text(r.iloc[maker_col])
        if not maker or "TOTAL" in maker.upper() or not re.fullmatch(r"\d+", sno):
            continue                            # blank / summary / footer row
        vals = {m: to_int(r.iloc[j]) for j, m in month_cols.items()}
        if any(v is None or v < 0 for v in vals.values()):
            raise ValueError(f"Bad numeric value for maker '{maker}': {vals}")
        if total_col is not None:
            rep = to_int(r.iloc[total_col])
            if rep is not None and rep != sum(vals.values()):
                mismatches.append((maker, rep, sum(vals.values())))
        for m, u in vals.items():
            rows.append({"month_num": MONTH_ABBR.index(m) + 1, "maker": maker, "units": u})

    if mismatches:
        print(f"WARNING: {len(mismatches)} makers whose TOTAL != sum of months, e.g. {mismatches[:3]}")
    df = pd.DataFrame(rows)
    df["month"] = df["month_num"].apply(lambda n: f"{year}-{n:02d}")
    df["brand"] = df["maker"].map(lambda m: BRAND_NAMES.get(m.upper(), m))
    return df.drop(columns="month_num")


def flag_partial(df: pd.DataFrame, extracted_on: dt.date) -> pd.DataFrame:
    """A month is PARTIAL if the extraction happened on or before its last day."""
    def is_partial(month: str) -> bool:
        y, m = map(int, month.split("-"))
        return extracted_on <= dt.date(y, m, calendar.monthrange(y, m)[1])
    df["is_partial_month"] = df["month"].map(is_partial)
    return df


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="data/raw/vahan_reportTable_2026.xlsx")
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--extracted-on", default="2026-10-31", help="YYYY-MM-DD date the file was downloaded")
    ap.add_argument("--out", default="data/processed/vahan_e2w_long.csv")
    a = ap.parse_args()

    extracted_on = dt.date.fromisoformat(a.extracted_on)
    df = parse_vahan_export(Path(a.input), a.year)
    df = flag_partial(df, extracted_on)

    # Merge makers that map to the same brand (e.g. the two Kinetic Green entries).
    df = (df.groupby(["month", "brand", "is_partial_month"], as_index=False)
            .agg(units=("units", "sum"), vahan_maker_entries=("maker", lambda s: " | ".join(sorted(set(s))))))
    df["source"] = "Vahan dashboard (vahan.parivahan.gov.in), MoRTH"
    df["extracted_on"] = extracted_on.isoformat()
    df = df.sort_values(["month", "units"], ascending=[True, False]).reset_index(drop=True)

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(a.out, index=False)
    print(f"{df['brand'].nunique()} brands x {df['month'].nunique()} months = {len(df)} rows -> {a.out}")
    print("Partial month(s):", sorted(df.loc[df.is_partial_month, 'month'].unique()))


if __name__ == "__main__":
    main()
