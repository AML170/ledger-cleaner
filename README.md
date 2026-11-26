# ledger-cleaner

A small Python script that takes a messy accounting CSV (like a bank or bookkeeping export) and turns it into a clean file plus a simple Excel report.

I built it as a learning project. It uses only the `csv` module and `openpyxl`, and it is meant to be easy to read.

## What it does

- Reads a CSV with the columns `Date, Description, Category, Amount`
- Fixes dates written in different styles (`2026-01-03`, `03/01/2026`, `03-01-2026`) into `YYYY-MM-DD`
- Fixes amounts such as `"-4,500.00"` into numbers
- Cleans category names (`  office supplies ` becomes `Office Supplies`)
- Skips empty lines and removes exact duplicates
- Reports lines it cannot read instead of hiding them
- Writes `output/clean_transactions.csv` and `output/report.xlsx` (sheets: Transactions, Summary with totals per category, income, expenses, balance)

## Run it

```bash
pip install -r requirements.txt
python clean_and_report.py messy_transactions.csv
python -m unittest test_cleaning
```

## Example output

```
Read 15 lines, kept 11 clean rows.
  Skipped line 14: {... 'Amount': 'abc'}
Totals by category:
  Income             25,250.75
  Rent               -4,500.00
  ...
```

## Files

- `clean_and_report.py`: the script, one small function per job
- `messy_transactions.csv`: made-up sample data
- `sample_clean_transactions.csv`: what the script produces from the sample (the Excel report is created when you run it)
- `test_cleaning.py`: basic tests for the cleaning functions

## Limits

Only three date formats are supported (add more in `DATE_FORMATS`). Day comes before month in `03/01/2026`. The sample data is invented.

Author: AML Ashraf
