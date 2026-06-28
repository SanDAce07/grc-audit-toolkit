# Audit Scripts

Automation scripts to support IT audit and GRC work.

## Scripts

| Script | Purpose | Language |
|--------|---------|----------|
| `access-review-analyzer.py` | Parse user access exports and flag anomalies | Python |
| `change-log-sampler.py` | Random sample selection from change logs | Python |
| `risk-score-calculator.py` | Calculate and rank risk scores from register | Python |
| `aging-report-analyzer.py` | AR aging analysis (audit analytics) | Python |

## Coming Soon
- SOC 2 gap assessment scorer
- ITGC control effectiveness dashboard
- Audit evidence tracker (Excel output)

## Requirements
```
pip install pandas openpyxl xlsxwriter
```

