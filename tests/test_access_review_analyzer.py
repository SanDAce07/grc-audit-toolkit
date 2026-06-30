import pandas as pd

from conftest import load_audit_script


access_review = load_audit_script("access-review-analyzer.py")


def test_invalid_last_login_is_reported_without_crashing():
    users = pd.DataFrame([
        {
            "user_id": "U999",
            "name": "Test User",
            "email": "test@example.com",
            "department": "Finance",
            "role": "Analyst",
            "system": "ERP",
            "access_level": "Read",
            "last_login": "not-a-date",
            "status": "Active",
            "mfa_enabled": "Yes",
        }
    ])

    findings = access_review.analyze_access(users)

    assert findings.loc[0, "Finding Type"] == "Dormant Account"
    assert findings.loc[0, "Last Login"] == ""
    assert findings.loc[0, "Days Since Login"] == 9999
