import pandas as pd

from conftest import load_audit_script


change_sampler = load_audit_script("change-log-sampler.py")


def test_approval_status_only_uses_missing_approval_findings():
    sample = pd.DataFrame([
        {"change_id": "CHG-001", "approved_by": "manager.one"},
        {"change_id": "CHG-002", "approved_by": ""},
        {"change_id": "CHG-003", "approved_by": "manager.two"},
    ])
    exceptions = pd.DataFrame([
        {"Change ID": "CHG-001", "Exception Type": "Change Not Tested"},
        {"Change ID": "CHG-002", "Exception Type": "Missing Approval"},
    ])

    workpaper = change_sampler.add_workpaper_columns(sample, exceptions)

    assert workpaper["Approval Documented?"].tolist() == ["Yes", "No ⚠️", "Yes"]
    assert workpaper["Auditor Notes"].tolist() == [
        "⚠️ See Exceptions Sheet",
        "⚠️ See Exceptions Sheet",
        "",
    ]
