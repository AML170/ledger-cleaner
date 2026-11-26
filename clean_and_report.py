"""Clean a messy accounting CSV and build a small Excel report.

Beginner project. Uses only the csv module and openpyxl.
Run:  python clean_and_report.py messy_transactions.csv
"""
import csv
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

# Date formats we expect to meet in messy files.
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]


def clean_date(text):
    """Turn many date styles into YYYY-MM-DD. Return None if impossible."""
    text = text.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def clean_amount(text):
    """Turn '-4,500.00' into -4500.0. Return None if it is not a number."""
    text = text.strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def clean_category(text):
    """Remove extra spaces and unify capital letters: ' rent ' -> 'Rent'."""
    return " ".join(text.split()).title()


def clean_rows(raw_rows):
    """Return (good_rows, problems). Skips empty lines, bad values, duplicates."""
    good, problems, seen = [], [], set()
    for number, row in enumerate(raw_rows, start=2):  # line 1 is the header
        if not any(value.strip() for value in row.values()):
            continue  # completely empty line
        date = clean_date(row["Date"])
        amount = clean_amount(row["Amount"])
        if date is None or amount is None:
            problems.append((number, row))
            continue
        cleaned = {
            "Date": date,
            "Description": row["Description"].strip(),
            "Category": clean_category(row["Category"]),
            "Amount": amount,
        }
        key = tuple(cleaned.values())
        if key in seen:
            continue  # exact duplicate
        seen.add(key)
        good.append(cleaned)
    return good, problems


def totals_by_category(rows):
    """Add up the amounts for each category."""
    totals = {}
    for row in rows:
        totals[row["Category"]] = totals.get(row["Category"], 0) + row["Amount"]
    return dict(sorted(totals.items()))


def write_clean_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["Date", "Description", "Category", "Amount"])
        writer.writeheader()
        writer.writerows(rows)


def write_report(rows, totals, path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Transactions"
    sheet.append(["Date", "Description", "Category", "Amount"])
    for row in sorted(rows, key=lambda r: r["Date"]):
        sheet.append([row["Date"], row["Description"], row["Category"], row["Amount"]])

    summary = book.create_sheet("Summary")
    summary.append(["Category", "Total"])
    for category, total in totals.items():
        summary.append([category, round(total, 2)])
    income = sum(r["Amount"] for r in rows if r["Amount"] > 0)
    expenses = sum(r["Amount"] for r in rows if r["Amount"] < 0)
    summary.append([])
    summary.append(["Total income", round(income, 2)])
    summary.append(["Total expenses", round(expenses, 2)])
    summary.append(["Balance", round(income + expenses, 2)])

    for ws in (sheet, summary):
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 14
    book.save(path)


def main():
    if len(sys.argv) != 2:
        print("Usage: python clean_and_report.py <input.csv>")
        sys.exit(1)
    with open(sys.argv[1], newline="", encoding="utf-8") as file:
        raw_rows = list(csv.DictReader(file))

    rows, problems = clean_rows(raw_rows)
    totals = totals_by_category(rows)

    Path("output").mkdir(exist_ok=True)
    write_clean_csv(rows, "output/clean_transactions.csv")
    write_report(rows, totals, "output/report.xlsx")

    print(f"Read {len(raw_rows)} lines, kept {len(rows)} clean rows.")
    for number, row in problems:
        print(f"  Skipped line {number}: {dict(row)}")
    print("Totals by category:")
    for category, total in totals.items():
        print(f"  {category:<16}{total:>12,.2f}")
    print("Saved output/clean_transactions.csv and output/report.xlsx")


if __name__ == "__main__":
    main()
