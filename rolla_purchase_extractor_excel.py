import re
import argparse
from datetime import datetime
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo


DATE_RE = re.compile(
    r'(?P<date>(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)'
    r'\s+\d{1,2},\s+\d{4})',
    re.IGNORECASE
)

TIME_RE = re.compile(
    r'(?P<time>\d{1,2}:\d{2}:\d{2})'
)

DOLLAR_RE = re.compile(
    r'\$([0-9,]+(?:\.[0-9]{1,2})?)'
)

DISCOUNT_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*%\s*OFF',
    re.IGNORECASE
)

ID_RE = re.compile(
    r'\b(\d{5,})\b'
)


def money(value):
    """
    Round monetary values to the nearest cent.
    """
    return Decimal(value).quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP
    )


def split_transaction_blocks(text):
    """
    Rolla transaction history usually looks like:

        transaction date
        transaction time
        transaction information
        expiry date
        expiry time

    The first date is the transaction date.
    The second date is the expiry date.
    """

    date_matches = list(DATE_RE.finditer(text))
    blocks = []

    i = 0

    while i < len(date_matches):
        transaction_date_match = date_matches[i]

        start = transaction_date_match.start()

        if i + 2 < len(date_matches):
            end = date_matches[i + 2].start()
            i += 2
        else:
            end = len(text)
            i += 1

        blocks.append(
            (
                transaction_date_match.group("date"),
                text[start:end]
            )
        )

    return blocks


def clean_block(block):
    return [
        line.strip()
        for line in block.splitlines()
        if line.strip()
    ]


def calculate_actual_purchase_price(description):
    """
    Calculate how much was actually paid.

    Examples:

    Purchase - $99.99 - 5,250,000 GC
        -> $99.99

    MEGA PACK 25% OFF $999.99
        -> $749.99

    MEGA STORE 33% OFF $149.99 PACK
        -> $100.49

    REAC 90% OFF $9.99
        -> $1.00

    MEGA STORE 25% EXTRA $999.99
        -> $999.99
    """

    dollar_matches = DOLLAR_RE.findall(description)

    if not dollar_matches:
        return None

    # Use the last dollar amount appearing in the description.
    listed_price = Decimal(
        dollar_matches[-1].replace(",", "")
    )

    discount_match = DISCOUNT_RE.search(description)

    if discount_match:

        discount_percent = Decimal(
            discount_match.group(1)
        )

        multiplier = (
            Decimal("1")
            - discount_percent / Decimal("100")
        )

        actual_price = listed_price * multiplier

        return money(actual_price)

    # No "% OFF" means the displayed dollar amount
    # is treated as the actual amount paid.
    return money(listed_price)


def parse_rolla_purchases(text):

    rows = []

    seen_transaction_ids = set()

    blocks = split_transaction_blocks(text)

    for date_text, block in blocks:

        lines = clean_block(block)

        if not lines:
            continue

        # Find a line containing a dollar amount.
        transaction_line = ""

        for line in lines:
            if DOLLAR_RE.search(line):
                transaction_line = line
                break

        if not transaction_line:
            continue

        # Find transaction time.
        time_match = TIME_RE.search(block)

        time_text = (
            time_match.group("time")
            if time_match
            else ""
        )

        # --------------------------------------------------
        # DESCRIPTION
        # --------------------------------------------------

        tab_parts = [
            part.strip()
            for part in transaction_line.split("\t")
            if part.strip()
        ]

        if len(tab_parts) >= 3:
            description = tab_parts[2]
        else:
            description = transaction_line

        # --------------------------------------------------
        # TRANSACTION ID
        # --------------------------------------------------

        possible_ids = ID_RE.findall(transaction_line)

        transaction_id = ""

        if possible_ids:
            transaction_id = possible_ids[-1]

        else:
            remaining_block = block[
                block.find(transaction_line)
                + len(transaction_line):
            ]

            possible_ids = ID_RE.findall(
                remaining_block
            )

            if possible_ids:
                transaction_id = possible_ids[0]

        if not transaction_id:
            print(
                f"WARNING: No transaction ID found: "
                f"{date_text} | {description}"
            )
            continue

        # Don't count the same transaction twice.
        if transaction_id in seen_transaction_ids:
            continue

        seen_transaction_ids.add(
            transaction_id
        )

        # --------------------------------------------------
        # ACTUAL PURCHASE AMOUNT
        # --------------------------------------------------

        purchase_amount = (
            calculate_actual_purchase_price(
                description
            )
        )

        if purchase_amount is None:
            continue

        # --------------------------------------------------
        # DATE
        # --------------------------------------------------

        transaction_date = datetime.strptime(
            date_text.replace(
                "Sept ",
                "Sep "
            ),
            "%b %d, %Y"
        )

        rows.append({
            "date": transaction_date,
            "time": time_text,
            "purchase_amount": purchase_amount,
            "description": description,
            "transaction_id": transaction_id,
        })

    return rows


def write_excel(rows, output_path):

    wb = Workbook()

    ws = wb.active

    ws.title = "Rolla Purchases"

    headers = [
        "Date",
        "Time",
        "Actual Purchase Amount",
        "Description",
        "Transaction ID",
    ]

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:

        ws.append([
            row["date"],
            row["time"],
            float(row["purchase_amount"]),
            row["description"],
            row["transaction_id"],
        ])

    # Date formatting
    for cell in ws["A"][1:]:
        cell.number_format = "mmm d, yyyy"

    # Currency formatting
    for cell in ws["C"][1:]:
        cell.number_format = '$#,##0.00'

    # Transaction IDs as text
    for cell in ws["E"][1:]:
        cell.number_format = "@"

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 45
    ws.column_dimensions["E"].width = 18

    ws.freeze_panes = "A2"

    if rows:

        last_row = len(rows) + 1

        table = Table(
            displayName="RollaPurchasesTable",
            ref=f"A1:E{last_row}"
        )

        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        table.tableStyleInfo = style

        ws.add_table(table)

    # ======================================================
    # SUMMARY
    # ======================================================

    summary = wb.create_sheet(
        "Summary"
    )

    summary["A1"] = "Rolla Purchase Summary"

    summary["A1"].font = Font(
        bold=True,
        size=14
    )

    summary["A3"] = "Purchase Count"
    summary["B3"] = len(rows)

    summary["A4"] = "Total Amount Paid"

    if rows:

        summary["B4"] = (
            f"=SUM('Rolla Purchases'!"
            f"C2:C{len(rows) + 1})"
        )

    else:
        summary["B4"] = 0

    summary["B4"].number_format = (
        '$#,##0.00'
    )

    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 18

    # ======================================================
    # SAVE
    # ======================================================

    wb.save(output_path)

    python_total = sum(
        row["purchase_amount"]
        for row in rows
    )

    print()
    print("-----------------------------------")
    print("Rolla Purchase Extraction")
    print("-----------------------------------")

    print(
        f"Purchases found: {len(rows)}"
    )

    print(
        f"Total amount paid: "
        f"${python_total:,.2f}"
    )

    print(
        f"Excel file: {output_path}"
    )

    print("-----------------------------------")


def main():

    parser = argparse.ArgumentParser(
        description=(
            "Extract Rolla purchases and "
            "calculate actual amount paid "
            "after percentage-off promotions."
        )
    )

    parser.add_argument(
        "input_file",
        help=(
            "Text file containing copied "
            "Rolla transaction history"
        )
    )

    parser.add_argument(
        "-o",
        "--output",
        default="rolla_purchases.xlsx",
        help="Output Excel filename"
    )

    args = parser.parse_args()

    text = Path(
        args.input_file
    ).read_text(
        encoding="utf-8",
        errors="replace"
    )

    rows = parse_rolla_purchases(text)

    write_excel(
        rows,
        Path(args.output)
    )


if __name__ == "__main__":
    main()