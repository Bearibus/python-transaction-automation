import re
import argparse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo


DATE_RE = re.compile(
    r'(?P<date>(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})',
    re.IGNORECASE
)

TIME_RE = re.compile(r'(?P<time>\d{1,2}:\d{2}:\d{2})')

STATUS_RE = re.compile(
    r'\b(COMPLETED|PENDING|PROCESSING|CANCELLED|CANCELED|REJECTED|FAILED)\b',
    re.IGNORECASE
)

AMOUNT_RE = re.compile(
    r'\bSC\s*(-?[0-9,]+(?:\.[0-9]{1,2})?)',
    re.IGNORECASE
)


def split_redemption_blocks(text):
    matches = list(DATE_RE.finditer(text))
    blocks = []

    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        blocks.append((match.group("date"), text[start:end]))

    return blocks


def parse_rolla_redemptions(text):
    rows = []

    for date_text, block in split_redemption_blocks(text):
        status_match = STATUS_RE.search(block)
        amount_match = AMOUNT_RE.search(block)

        if not status_match or not amount_match:
            continue

        tm = TIME_RE.search(block)
        time_text = tm.group("time") if tm else ""

        raw_amount = float(amount_match.group(1).replace(",", ""))
        redemption_amount = abs(raw_amount)

        method = block[amount_match.end():status_match.start()]
        method = re.sub(r'[\|\t\r\n]+', ' ', method)
        method = re.sub(r'\s+', ' ', method).strip(' :-|')

        after_status = block[status_match.end():]
        id_match = re.search(r'\b(\d{5,})\b', after_status)
        transaction_id = id_match.group(1) if id_match else ""

        requested_date = datetime.strptime(
            date_text.replace("Sept ", "Sep "),
            "%b %d, %Y"
        )

        rows.append({
            "date_requested": requested_date,
            "time_requested": time_text,
            "redemption_amount": redemption_amount,
            "method": method,
            "status": status_match.group(1).upper(),
            "transaction_id": transaction_id,
        })

    return rows


def write_excel(rows, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rolla Redemptions"

    headers = [
        "Date Requested",
        "Time",
        "Redemption Amount",
        "Method",
        "Status",
        "Transaction ID",
    ]

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([
            row["date_requested"],
            row["time_requested"],
            row["redemption_amount"],
            row["method"],
            row["status"],
            row["transaction_id"],
        ])

    for cell in ws["A"][1:]:
        cell.number_format = "mmm d, yyyy"

    for cell in ws["C"][1:]:
        cell.number_format = '$#,##0.00'

    for cell in ws["F"][1:]:
        cell.number_format = "@"

    widths = {
        "A": 18,
        "B": 12,
        "C": 20,
        "D": 28,
        "E": 14,
        "F": 18,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A2"

    if len(rows) > 0:
        table_ref = f"A1:F{len(rows) + 1}"
        table = Table(
            displayName="RollaRedemptionsTable",
            ref=table_ref
        )

        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        table.tableStyleInfo = style
        ws.add_table(table)

    wb.save(output_path)

    print(f"Wrote {len(rows)} redemption rows -> {output_path}")


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "input_file",
        help="Text file containing copied Rolla redemption history"
    )

    parser.add_argument(
        "-o",
        "--output",
        default="rolla_redemptions.xlsx"
    )

    args = parser.parse_args()

    text = Path(args.input_file).read_text(
        encoding="utf-8",
        errors="replace"
    )

    rows = parse_rolla_redemptions(text)

    write_excel(
        rows,
        Path(args.output)
    )


if __name__ == "__main__":
    main()