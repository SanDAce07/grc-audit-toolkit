import pandas as pd
from openpyxl import Workbook

from conftest import load_audit_script


risk_calculator = load_audit_script("risk-score-calculator.py")


def test_heat_map_is_explicitly_labeled_as_inherent_risk():
    risks = pd.DataFrame([
        {"risk_id": "R-001", "likelihood": 4, "impact": 5},
    ])
    heat_map = risk_calculator.build_heat_map_data(risks)
    worksheet = Workbook().active

    risk_calculator.write_heat_map_sheet(worksheet, risks, heat_map)

    assert worksheet["A1"].value == "INHERENT IT RISK HEAT MAP — Likelihood × Impact"
