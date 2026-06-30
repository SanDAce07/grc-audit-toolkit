from datetime import datetime

import pandas as pd

from conftest import load_audit_script


aging_report = load_audit_script("aging-report-analyzer.py")


def test_overdue_exposure_excludes_credit_balances():
    receivables = pd.DataFrame([
        {
            "customer": "Example Customer",
            "invoice_number": "INV-1",
            "invoice_date": pd.Timestamp("2026-01-01"),
            "due_date": pd.Timestamp("2026-02-01"),
            "balance": 100.0,
            "status": "open",
        },
        {
            "customer": "Example Customer",
            "invoice_number": "CM-1",
            "invoice_date": pd.Timestamp("2026-01-01"),
            "due_date": pd.Timestamp("2026-02-01"),
            "balance": -90.0,
            "status": "open",
        },
    ])

    analyzed, customer_summary = aging_report.analyze_aging(
        receivables, datetime(2026, 6, 30)
    )
    findings = aging_report.detect_concentration_risk(analyzed, customer_summary)

    assert customer_summary.loc[0, "overdue_balance"] == 100.0
    assert "100.0% of positive AR" in findings.loc[0, "Detail"]
