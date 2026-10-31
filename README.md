# Vahan – Electric Two-Wheeler Registrations, 2026

```
vahan_e2w_analysis/
├── README.md
├── ANALYSIS.md                              findings, source, extraction date, limitations
├── requirements.txt
├── sql/task1_queries.sql                    Task 1 (Q1, Q2) with expected output
├── data/
│   ├── raw/vahan_reportTable_2026.xlsx      original dashboard export (untouched)
│   └── processed/vahan_e2w_long.csv         cleaned: month, brand, units, partial flag, source
├── src/
│   ├── extract_clean.py                     step 1: parse + clean the dashboard export
│   └── analyse.py                           step 2: totals, share, MoM, top brands, Excel/CSV
└── outputs/
    ├── vahan_e2w_2026_analysis.xlsx         all tables in one workbook
    └── *.csv                                the same tables as CSV
```

## Requirements

Python 3.9+ and the packages in `requirements.txt`:

```bash
pip install -r requirements.txt     # pandas, numpy, openpyxl
```

## Run

From the project folder:

```bash
python src/extract_clean.py --input data/raw/vahan_reportTable_2026.xlsx --extracted-on 2026-10-31
python src/analyse.py
```

Results are written to `outputs/`. Run both steps in that order. The first one produces
`data/processed/vahan_e2w_long.csv`, which the second one reads.

## Getting a fresh export from Vahan

1. Open https://vahan.parivahan.gov.in/vahan4dashboard/
2. Y-axis: **Maker**; X-axis: **Month Wise**; Year: **2026** (calendar year); State: **All India**
3. Vehicle class: **TWO WHEELER (Invalid Carriage)**, **TWO WHEELER (NT)**, **TWO WHEELER (T)**
4. Fuel: **PURE EV** and **ELECTRIC(BOV)**
5. Refresh and download the Excel file, save it into `data/raw/`
6. Run the two commands above, passing the new file name and the download date as
   `--extracted-on YYYY-MM-DD`. The month containing that date is automatically marked **partial**.

The Excel file's layout (title row, header row with month names, maker column, TOTAL column) is detected
automatically, and non-breaking spaces and Indian-style numbers ("4,70,369") are handled. The script stops
with an error if a month value is non-numeric or negative, and warns if a maker's TOTAL does not equal the
sum of its months.

## Workbook sheets

`About` (source, extraction date, filters, definitions) · `Summary_Units` / `Summary_Share_pct` / `Summary_MoM_pct`
(top 10 brands + all others) · `Industry_Monthly` · `Top_Brands` · `Share_Change` · `All_Brands_Units` /
`All_Brands_Share_pct` / `All_Brands_MoM_pct` (every brand) · `Consolidated_Long` (one row per brand per month,
filterable). Partial-month column headers are orange.

## Definitions

* Industry total = sum of all makers in the export. Market share = brand units ÷ industry total × 100.
* MoM growth = (this month ÷ previous month − 1) × 100; blank when the previous month is 0. Previous month below
  100 units is flagged as low base in `Consolidated_Long`.
* `OTHERS` is a Vahan catch-all maker; it counts in totals but is never ranked.

## Limitations

See `ANALYSIS.md`. The main ones: October 2026 is month-to-date, Vahan back-fills registrations, and the data are
registrations (not wholesale sales) for the filter scope listed above.
