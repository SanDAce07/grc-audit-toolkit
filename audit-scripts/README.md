# Audit Scripts

Automation scripts to support IT audit and GRC work.

## Scripts

| Script | Purpose | Language |
|--------|---------|----------|
| [access-review-analyzer.py](./access-review-analyzer.py) | Parse user access exports and flag anomalies | Python |
| [change-log-sampler.py](./change-log-sampler.py) | Random sample selection from change logs | Python |
| [risk-score-calculator.py](./risk-score-calculator.py) | Rank residual risk and visualize inherent risk | Python |
| [aging-report-analyzer.py](./aging-report-analyzer.py) | Analyze AR aging and positive receivable exposure | Python |

## Aging Report Analyzer
[aging-report-analyzer.py](./aging-report-analyzer.py) reads AR aging data, classifies balances into aging buckets, highlights severe overdue exposure, and writes a formatted Excel workbook for audit review.

## Coming Soon
- SOC 2 gap assessment scorer
- ITGC control effectiveness dashboard
- Audit evidence tracker (Excel output)

## Requirements
```bash
python -m pip install -r ../requirements.txt
```

## Tests

From the repository root:

```bash
python -m pip install -r requirements-dev.txt
python -m pytest
```
