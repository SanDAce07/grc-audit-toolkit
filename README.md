# GRC & IT Audit Toolkit

A structured repository for Governance, Risk, and Compliance (GRC) and IT Audit work.

## Structure

```
grc-audit-toolkit/
├── control-frameworks/     # COBIT, SOC 2, CISA control mappings
├── risk-register/          # Risk identification and scoring
├── audit-evidence/         # Evidence checklists and templates
├── compliance-checklists/  # SOC 2, ITGC, and regulatory checklists
├── audit-scripts/          # Python/Excel scripts for audit automation
└── reports/                # Audit report templates
```

## Audit Scripts

### Aging Report Analyzer
`audit-scripts/aging-report-analyzer.py` analyzes accounts receivable aging exports and produces a formatted Excel workpaper for audit analytics.

The script includes:
- aging buckets from `Current` through `120+`
- customer concentration analysis
- exception checks for missing due dates, credit balances, and severely aged receivables
- an Excel workbook with `Invoice Aging`, `Customer Summary`, `Exceptions`, and `Executive Summary` sheets

Expected columns:
- `customer` or `customer name`
- `invoice_number` or `invoice`
- `invoice_date` or `date`
- `due_date` or `due date`
- `balance` or `amount`
- `status`

Example:
```bash
python audit-scripts/aging-report-analyzer.py --input ar_aging.csv --output aging_report.xlsx --as-of 2026-06-30 --top 15
```

## Frameworks Covered
- COBIT 2019
- SOC 2 (Trust Service Criteria)
- CISA IT Audit Standards
- ITGC (IT General Controls)

## Author
Sandesh | Accounting & CIS | University of Louisiana Monroe
