#!/usr/bin/env python3
"""
analyse.py  -  Step 2 of 2
--------------------------
Read data/processed/vahan_e2w_long.csv (made by extract_clean.py) and build

  * month-wise, brand-wise registrations
  * month-wise industry totals (+ MoM growth)
  * brand-wise market share (%) for every month
  * top brands of each month
  * month-on-month growth (%) for each brand and for the industry
  * a short list of share changes (Jan -> last full month)

and write CSV files plus one formatted Excel workbook to outputs/.

Usage
    python src/analyse.py
    python src/analyse.py --input data/processed/vahan_e2w_long.csv --out outputs

Definitions
    industry total  = sum of ALL makers in the Vahan export (no total row is supplied)
    market share    = brand units / industry total in the same month x 100
    MoM growth      = (units this month / units previous month - 1) x 100
                      left blank when the previous month is 0 (undefined)
    low_base flag   = previous-month units < 100 (growth % is not meaningful)
    partial month   = month in which the data was extracted; figures are
                      month-to-date and will rise as Vahan back-fills
    OTHERS          = a catch-all maker category inside Vahan; it counts in the
                      total but is never ranked as a brand
"""
import argparse
import calendar
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OTHERS_LABEL = "OTHERS (Vahan category)"
LOW_BASE = 100          # previous-month units below this => MoM flagged as low base
TOP_N_TABLE = 10        # brands shown individually in the summary tables


def month_label(m: str, partial: bool = False) -> str:
    y, n = map(int, m.split("-"))
    return f"{calendar.month_abbr[n]} {y}" + (" (partial)" if partial else "")


def build(long: pd.DataFrame):
    months = sorted(long["month"].unique())
    partial = long.groupby("month")["is_partial_month"].first().to_dict()

    units = long.pivot(index="brand", columns="month", values="units").reindex(columns=months)
    units["YTD"] = units[months].sum(axis=1)
    units = units.sort_values("YTD", ascending=False)

    totals = units[months].sum()
    share = units[months].div(totals, axis=1) * 100

    prev = units[months].shift(1, axis=1)
    mom = (units[months] / prev - 1) * 100
    mom = mom.where(prev > 0)                       # undefined when previous month = 0
    low_base = (prev < LOW_BASE) & (prev > 0)

    # ---- industry table --------------------------------------------------
    ranked = units.drop(index=OTHERS_LABEL, errors="ignore")
    ind = pd.DataFrame(index=months)
    ind["industry_total"] = totals
    ind["mom_growth_pct"] = (totals / totals.shift(1) - 1) * 100
    ind["brands_with_registrations"] = (units[months] > 0).sum()
    ind["top1_brand"] = [ranked[m].idxmax() for m in months]
    ind["top3_share_pct"] = [ranked[m].nlargest(3).sum() / totals[m] * 100 for m in months]
    ind["top5_share_pct"] = [ranked[m].nlargest(5).sum() / totals[m] * 100 for m in months]
    ind["is_partial_month"] = [partial[m] for m in months]

    # ---- top brands per month ---------------------------------------------
    rows = []
    for m in months:
        r = {"month": m, "is_partial_month": partial[m], "industry_total": int(totals[m])}
        for i, (b, u) in enumerate(ranked[m].nlargest(5).items(), 1):
            r[f"rank{i}_brand"], r[f"rank{i}_units"] = b, int(u)
            r[f"rank{i}_share_pct"] = round(u / totals[m] * 100, 2)
        rows.append(r)
    top = pd.DataFrame(rows)

    # ---- consolidated long table -------------------------------------------
    rank_m = ranked[months].rank(ascending=False, method="min")
    cons = (units[months].reset_index().melt(id_vars="brand", var_name="month", value_name="units"))
    cons["rank_in_month"] = [rank_m.loc[b, m] if b in rank_m.index else np.nan
                             for b, m in zip(cons.brand, cons.month)]
    cons["market_share_pct"] = [share.loc[b, m] for b, m in zip(cons.brand, cons.month)]
    cons["mom_growth_pct"] = [mom.loc[b, m] for b, m in zip(cons.brand, cons.month)]
    cons["mom_low_base"] = [bool(low_base.loc[b, m]) for b, m in zip(cons.brand, cons.month)]
    cons["is_partial_month"] = cons["month"].map(partial)
    cons = cons.sort_values(["month", "units"], ascending=[True, False]).reset_index(drop=True)

    # ---- compact top-N table (top brands + all other brands) ---------------
    top_brands = [b for b in ranked.index[:TOP_N_TABLE]]
    comp = units.loc[top_brands, months].copy()
    comp.loc["All other brands (incl. OTHERS)"] = totals - comp.sum()
    comp_share = comp.div(totals, axis=1) * 100
    comp_mom = (comp / comp.shift(1, axis=1) - 1) * 100
    comp_mom = comp_mom.where(comp.shift(1, axis=1) > 0)

    # ---- share change: first month -> last FULL month ----------------------
    full = [m for m in months if not partial[m]]
    first, last = months[0], full[-1]
    chg = pd.DataFrame({
        f"units_{first}": units.loc[top_brands, first],
        f"units_{last}": units.loc[top_brands, last],
        "unit_growth_pct": (units.loc[top_brands, last] / units.loc[top_brands, first] - 1) * 100,
        f"share_{first}_pct": share.loc[top_brands, first],
        f"share_{last}_pct": share.loc[top_brands, last],
    })
    chg["share_change_pp"] = chg[f"share_{last}_pct"] - chg[f"share_{first}_pct"]
    peak = share.loc[top_brands, months].idxmax(axis=1)
    chg["peak_share_month"] = peak
    chg["peak_share_pct"] = [share.loc[b, peak[b]] for b in top_brands]

    return dict(months=months, partial=partial, units=units, share=share, mom=mom, ind=ind,
                top=top, cons=cons, comp=comp, comp_share=comp_share, comp_mom=comp_mom, chg=chg)


# ------------------------------- Excel helpers ---------------------------------
HEAD_FILL = PatternFill("solid", start_color="1F3864")
PART_FILL = PatternFill("solid", start_color="FCE4D6")


def style_sheet(ws, int_cols=(), pct_cols=(), partial_cols=(), freeze="B2"):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.iter_cols(min_row=2):
        idx = col[0].column
        for cell in col:
            if idx in int_cols and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
            elif idx in pct_cols and isinstance(cell.value, (int, float)):
                cell.number_format = "0.0"
    for idx in partial_cols:
        ws.cell(row=1, column=idx).fill = PatternFill("solid", start_color="C55A11")
    for col in ws.columns:
        w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, 11), 46)
    ws.freeze_panes = freeze


def write_excel(path: Path, R: dict, about: list):
    months, partial = R["months"], R["partial"]
    labels = {m: month_label(m, partial[m]) for m in months}
    pcols = [i + 2 for i, m in enumerate(months) if partial[m]]      # partial month columns (1-based)

    def wide(df, name, fmt):
        d = df.rename(columns=labels).reset_index()
        d.columns = [str(c) for c in d.columns]
        d.to_excel(xw, name, index=False)
        ncol = d.shape[1]
        style_sheet(xw.sheets[name],
                    int_cols=range(2, ncol + 1) if fmt == "int" else (),
                    pct_cols=range(2, ncol + 1) if fmt == "pct" else (),
                    partial_cols=pcols)

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        # About
        pd.DataFrame(about, columns=["Item", "Detail"]).to_excel(xw, "About", index=False)
        ws = xw.sheets["About"]
        ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 26, 120
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), HEAD_FILL
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")

        # Compact summary tables
        wide(R["comp"], "Summary_Units", "int")
        wide(R["comp_share"], "Summary_Share_pct", "pct")
        wide(R["comp_mom"], "Summary_MoM_pct", "pct")

        # Industry
        ind = R["ind"].copy()
        ind.insert(0, "month", [labels[m] for m in ind.index])
        ind.drop(columns="is_partial_month").round(2).to_excel(xw, "Industry_Monthly", index=False)
        style_sheet(xw.sheets["Industry_Monthly"], int_cols=(2, 4), pct_cols=(3, 6, 7))

        # Top brands
        t = R["top"].copy()
        t["month"] = [labels[m] for m in R["top"]["month"]]
        t.drop(columns="is_partial_month").to_excel(xw, "Top_Brands", index=False)
        style_sheet(xw.sheets["Top_Brands"])

        # Share change
        R["chg"].round(2).reset_index().to_excel(xw, "Share_Change", index=False)
        style_sheet(xw.sheets["Share_Change"])

        # Full brand tables (all brands)
        u = R["units"].copy()
        wide(u.rename(columns={"YTD": "YTD (Jan-Oct, partial)"}), "All_Brands_Units", "int")
        wide(R["share"], "All_Brands_Share_pct", "pct")
        wide(R["mom"], "All_Brands_MoM_pct", "pct")

        # Consolidated long
        c = R["cons"].copy()
        c["month"] = c["month"].map(labels)
        c.round(2).to_excel(xw, "Consolidated_Long", index=False)
        style_sheet(xw.sheets["Consolidated_Long"], int_cols=(3,), pct_cols=(5, 6))
        xw.sheets["Consolidated_Long"].auto_filter.ref = xw.sheets["Consolidated_Long"].dimensions


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="data/processed/vahan_e2w_long.csv")
    ap.add_argument("--out", default="outputs")
    a = ap.parse_args()

    long = pd.read_csv(a.input)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    R = build(long)
    months = R["months"]

    # CSV outputs
    R["units"].to_csv(out / "brand_month_units.csv")
    R["share"].round(3).to_csv(out / "brand_month_market_share_pct.csv")
    R["mom"].round(2).to_csv(out / "brand_month_mom_growth_pct.csv")
    R["ind"].round(3).rename_axis("month").to_csv(out / "industry_monthly_totals.csv")
    R["top"].to_csv(out / "top_brands_per_month.csv", index=False)
    R["cons"].round(3).to_csv(out / "consolidated_month_brand_summary.csv", index=False)
    R["chg"].round(2).rename_axis("brand").to_csv(out / "share_change_summary.csv")

    extracted = long["extracted_on"].iloc[0]
    partial_m = [month_label(m) for m in months if R["partial"][m]]
    about = [
        ("Source", "Vahan dashboard, Ministry of Road Transport & Highways - https://vahan.parivahan.gov.in/vahan4dashboard/"),
        ("Extraction date", f"{extracted} (file downloaded from the dashboard by the user; exact time not stored in the file)"),
        ("Report layout", "Y-axis: Maker; X-axis: Month Wise; Calendar year 2026; State: All India"),
        ("Vehicle classes", "TWO WHEELER (Invalid Carriage) + TWO WHEELER (NT) + TWO WHEELER (T)"),
        ("Fuel", "PURE EV + ELECTRIC(BOV)"),
        ("Period covered", f"{month_label(months[0])} to {month_label(months[-1])}"),
        ("PARTIAL MONTH", f"{', '.join(partial_m)} is month-to-date at extraction; Vahan back-fills registrations, so it will "
                          "increase. Its MoM growth and share are provisional (orange header)."),
        ("Industry total", "Sum of all makers in the export (the export has no total row; each maker's TOTAL column was verified)."),
        ("Brands", "Each Vahan maker entity is one brand; a few names were shortened and two Kinetic Green entries merged. "
                   "Importer-led entries (e.g. 'ZAP (IMPORTER: ...)') stay separate. Hero MotoCorp (Vida) and Hero Electric "
                   "are different companies. 'OTHERS' is a Vahan catch-all and is not ranked."),
        ("Market share", "Brand units / industry total in the same month x 100"),
        ("MoM growth", f"(this month / previous month - 1) x 100; blank if previous month = 0; 'mom_low_base' = previous month < {LOW_BASE} units"),
        ("Sheets", "Summary_* = top 10 brands + all others; Industry_Monthly; Top_Brands; Share_Change; All_Brands_* = every brand; Consolidated_Long = tidy table"),
        ("Code", "src/extract_clean.py -> src/analyse.py (see README.md)"),
    ]
    write_excel(out / "vahan_e2w_2026_analysis.xlsx", R, about)

    pd.set_option("display.width", 220)
    print(R["ind"][["industry_total", "mom_growth_pct", "top3_share_pct", "top5_share_pct"]].round(1))
    print(R["chg"].round(1).to_string())
    print(f"Wrote outputs to {out}/")


if __name__ == "__main__":
    main()
