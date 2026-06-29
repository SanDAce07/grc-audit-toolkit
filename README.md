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
`audit-scripts/aging_report_analyzer.py` analyzes accounts receivable aging exports in CSV format and summarizes balances by aging bucket for audit review.

Expected columns:
- `customer` or `customer name`
- `invoice` or `invoice number`
- `due date`
- `amount` or `balance`
- optional `status`

Example:
```bash
python3 audit-scripts/aging_report_analyzer.py ar_aging.csv --as-of 2026-06-30 --top 15
```

The script reports:
- total open balance
- balances by aging bucket
- customers with the highest overdue exposure

## Frameworks Covered
- COBIT 2019
- SOC 2 (Trust Service Criteria)
- CISA IT Audit Standards
- ITGC (IT General Controls)

## Author
Sandesh | Accounting & CIS | University of Louisiana Monroe
