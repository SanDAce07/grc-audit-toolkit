# GRC & IT Audit Toolkit

A structured repository for Governance, Risk, and Compliance (GRC) and IT Audit work.

## Implemented Toolkit

- [control-frameworks](./control-frameworks) - COBIT, SOC 2, and CISA control mappings
- [risk-register](./risk-register) - Risk identification and scoring materials
- [audit-evidence](./audit-evidence) - Evidence checklists and templates
- [compliance-checklists](./compliance-checklists) - SOC 2, ITGC, and regulatory checklists
- [audit-scripts](./audit-scripts) - Python and Excel audit automation tools
- [reports](./reports) - Audit report templates

## Audit Scripts

- [Access Review Analyzer](./audit-scripts/access-review-analyzer.py) - Flags user access anomalies for IT audit testing
- [Change Log Sampler](./audit-scripts/change-log-sampler.py) - Selects change samples and pre-flags exceptions
- [Risk Score Calculator](./audit-scripts/risk-score-calculator.py) - Calculates inherent and residual risk, with an inherent-risk heat map and residual-risk summary
- [Aging Report Analyzer](./audit-scripts/aging-report-analyzer.py) - Produces an AR aging workpaper and exception summary
- [Audit Scripts Overview](./audit-scripts/README.md) - Landing page for the full script catalog

All four scripts run with generated sample data when `--input` is omitted, making the repository directly demonstrable after dependency installation.

### Aging Report Analyzer
[Open the script](./audit-scripts/aging-report-analyzer.py)

The script includes:
- aging buckets from `Current` through `120+`
- customer concentration analysis based on positive receivable exposure
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

## Setup and Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest
```

The automated tests cover invalid access-review dates, change-sampler workpaper classification, positive-AR concentration logic, and inherent-risk heat-map labeling.

## Planned Additions

- SOC 2 gap assessment scorer
- ITGC control-effectiveness dashboard
- Audit evidence tracker

## Frameworks Covered
- COBIT 2019
- SOC 2 (Trust Service Criteria)
- CISA IT Audit Standards
- ITGC (IT General Controls)

## Author
Sandesh | Accounting & CIS | University of Louisiana Monroe

## License

This project is available under the [MIT License](./LICENSE).
