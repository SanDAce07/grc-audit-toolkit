#!/usr/bin/env python3
"""Analyze accounts receivable aging from a CSV export.

The script groups open invoice balances into aging buckets and highlights
customers with overdue exposure to support audit and GRC review work.
"""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

DEFAULT_BUCKETS = [0, 30, 60, 90, 120]
STATUS_OPEN_VALUES = {"", "open", "outstanding", "unpaid", "pending"}
STATUS_CLOSED_VALUES = {"paid", "closed", "settled", "void", "cancelled", "canceled"}


@dataclass
class InvoiceRecord:
    customer: str
    invoice_number: str
    invoice_date: date | None
    due_date: date
    amount: Decimal
    status: str


@dataclass
class Summary:
    invoice_count: int = 0
    total_amount: Decimal = Decimal("0.00")
    bucket_totals: dict[str, Decimal] | None = None

    def __post_init__(self) -> None:
        if self.bucket_totals is None:
            self.bucket_totals = defaultdict(lambda: Decimal("0.00"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate an aging analysis from an accounts receivable CSV export."
    )
    parser.add_argument("csv_file", help="Path to the AR aging CSV file")
    parser.add_argument(
        "--as-of",
        default=date.today().isoformat(),
        help="As-of date for the analysis in YYYY-MM-DD format (default: today)",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=10,
        help="Number of highest-risk customers to display (default: 10)",
    )
    return parser.parse_args()


def parse_date(value: str) -> date | None:
    value = value.strip()
    if not value:
        return None

    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    raise ValueError(f"Unsupported date format: {value}")


def normalize_amount(value: str) -> Decimal:
    cleaned = value.strip().replace(",", "").replace("$", "")
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = f"-{cleaned[1:-1]}"

    try:
        return Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid amount: {value}") from exc


def normalize_status(value: str) -> str:
    return value.strip().lower()


def is_open_item(status: str) -> bool:
    if status in STATUS_CLOSED_VALUES:
        return False
    return status in STATUS_OPEN_VALUES or bool(status)


def find_column(fieldnames: list[str], aliases: Iterable[str]) -> str:
    lowered = {name.strip().lower(): name for name in fieldnames}
    for alias in aliases:
        if alias in lowered:
            return lowered[alias]
    raise KeyError(f"Missing required column. Expected one of: {', '.join(aliases)}")


def load_invoices(path: Path) -> list[InvoiceRecord]:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError("CSV file is missing a header row")

        customer_col = find_column(reader.fieldnames, ["customer", "customer name", "client", "account"])
        invoice_col = find_column(reader.fieldnames, ["invoice", "invoice number", "invoice #", "document number"])
        due_col = find_column(reader.fieldnames, ["due date", "due", "payment due date"])
        amount_col = find_column(reader.fieldnames, ["amount", "balance", "open amount", "invoice amount"])

        invoice_date_col = None
        status_col = None
        for aliases, target in ((["invoice date", "date", "document date"], "invoice_date"), (["status", "invoice status", "payment status"], "status")):
            try:
                column = find_column(reader.fieldnames, aliases)
            except KeyError:
                column = None
            if target == "invoice_date":
                invoice_date_col = column
            else:
                status_col = column

        invoices = []
        for row in reader:
            status = normalize_status(row.get(status_col, "") if status_col else "")
            if status and not is_open_item(status):
                continue

            amount = normalize_amount(row[amount_col])
            if amount <= Decimal("0"):
                continue

            invoices.append(
                InvoiceRecord(
                    customer=row[customer_col].strip() or "Unknown Customer",
                    invoice_number=row[invoice_col].strip() or "Unknown Invoice",
                    invoice_date=parse_date(row[invoice_date_col]) if invoice_date_col and row.get(invoice_date_col) else None,
                    due_date=parse_date(row[due_col]) or date.today(),
                    amount=amount,
                    status=status or "open",
                )
            )

    return invoices


def bucket_name(days_past_due: int) -> str:
    if days_past_due <= 0:
        return "Current"
    if days_past_due <= 30:
        return "1-30"
    if days_past_due <= 60:
        return "31-60"
    if days_past_due <= 90:
        return "61-90"
    if days_past_due <= 120:
        return "91-120"
    return "120+"


def format_currency(amount: Decimal) -> str:
    return f"${amount:,.2f}"


def analyze(invoices: list[InvoiceRecord], as_of: date) -> tuple[Summary, dict[str, Summary]]:
    summary = Summary()
    by_customer: dict[str, Summary] = defaultdict(Summary)

    for invoice in invoices:
        days_past_due = (as_of - invoice.due_date).days
        bucket = bucket_name(days_past_due)

        summary.invoice_count += 1
        summary.total_amount += invoice.amount
        summary.bucket_totals[bucket] += invoice.amount

        customer_summary = by_customer[invoice.customer]
        customer_summary.invoice_count += 1
        customer_summary.total_amount += invoice.amount
        customer_summary.bucket_totals[bucket] += invoice.amount

    return summary, by_customer


def print_summary(summary: Summary, by_customer: dict[str, Summary], top_n: int) -> None:
    print("Aging Report Analyzer")
    print("=" * 72)
    print(f"Open invoices analyzed: {summary.invoice_count}")
    print(f"Total open balance:     {format_currency(summary.total_amount)}")
    print()
    print("Aging bucket totals")
    print("-" * 72)
    for bucket in ("Current", "1-30", "31-60", "61-90", "91-120", "120+"):
        print(f"{bucket:>8}: {format_currency(summary.bucket_totals.get(bucket, Decimal('0.00')))}")

    ranked_customers = sorted(
        by_customer.items(),
        key=lambda item: (
            item[1].bucket_totals.get("120+", Decimal("0.00"))
            + item[1].bucket_totals.get("91-120", Decimal("0.00"))
            + item[1].bucket_totals.get("61-90", Decimal("0.00")),
            item[1].total_amount,
        ),
        reverse=True,
    )

    print()
    print(f"Top {top_n} customers by overdue exposure")
    print("-" * 72)
    for customer, customer_summary in ranked_customers[:top_n]:
        overdue = (
            customer_summary.bucket_totals.get("1-30", Decimal("0.00"))
            + customer_summary.bucket_totals.get("31-60", Decimal("0.00"))
            + customer_summary.bucket_totals.get("61-90", Decimal("0.00"))
            + customer_summary.bucket_totals.get("91-120", Decimal("0.00"))
            + customer_summary.bucket_totals.get("120+", Decimal("0.00"))
        )
        print(
            f"{customer[:28]:28} "
            f"total={format_currency(customer_summary.total_amount):>14} "
            f"overdue={format_currency(overdue):>14} "
            f"120+={format_currency(customer_summary.bucket_totals.get('120+', Decimal('0.00'))):>14}"
        )


def main() -> None:
    args = parse_args()
    as_of = parse_date(args.as_of)
    if as_of is None:
        raise ValueError("As-of date is required")

    invoices = load_invoices(Path(args.csv_file))
    if not invoices:
        print("No open invoice records found in the supplied CSV file.")
        return

    summary, by_customer = analyze(invoices, as_of)
    print_summary(summary, by_customer, args.top)


if __name__ == "__main__":
    main()
