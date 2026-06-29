# GRC & IT Audit Toolkit

A structured repository for Governance, Risk, and Compliance (GRC) and IT Audit work.

## Project Areas

- [control-frameworks](./control-frameworks) - COBIT, SOC 2, and CISA control mappings
- [risk-register](./risk-register) - Risk identification and scoring materials
- [audit-evidence](./audit-evidence) - Evidence checklists and templates
- [compliance-checklists](./compliance-checklists) - SOC 2, ITGC, and regulatory checklists
- [audit-scripts](./audit-scripts) - Python and Excel audit automation tools
- [reports](./reports) - Audit report templates

## Audit Scripts

- [Access Review Analyzer](./audit-scripts/access-review-analyzer.py) - Flags user access anomalies for IT audit testing
- [Change Log Sampler](./audit-scripts/change-log-sampler.py) - Selects change samples and pre-flags exceptions
- [Risk Score Calculator](./audit-scripts/risk-score-calculator.py) - Calculates inherent and residual risk scores
- [Aging Report Analyzer](./audit-scripts/aging-report-analyzer.py) - Produces an AR aging workpaper and exception summary
- [Audit Scripts Overview](./audit-scripts/README.md) - Landing page for the full script catalog

### Aging Report Analyzer
[Open the script](./audit-scripts/aging-report-analyzer.py)

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
