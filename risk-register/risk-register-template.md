# IT Risk Register

## Risk Scoring Matrix

**Likelihood:** 1 (Rare) | 2 (Unlikely) | 3 (Possible) | 4 (Likely) | 5 (Almost Certain)
**Impact:** 1 (Negligible) | 2 (Minor) | 3 (Moderate) | 4 (Major) | 5 (Critical)
**Risk Score = Likelihood × Impact**

| Risk Score | Rating |
|------------|--------|
| 1–4 | Low |
| 5–9 | Medium |
| 10–19 | High |
| 20–25 | Critical |

---

## Risk Register

| Risk ID | Risk Description | Category | Likelihood | Impact | Score | Rating | Owner | Mitigation | Residual Risk | Status |
|---------|-----------------|----------|------------|--------|-------|--------|-------|------------|---------------|--------|
| R-001 | Unauthorized access to financial systems | Access Control | 3 | 5 | 15 | High | IT Manager | MFA, RBAC, quarterly access reviews | Medium | Open |
| R-002 | Unpatched vulnerabilities in critical systems | Vulnerability Mgmt | 4 | 4 | 16 | High | IT Security | Monthly patching cycle, vulnerability scans | Low | In Progress |
| R-003 | Data loss due to failed backups | Business Continuity | 2 | 5 | 10 | High | Ops Manager | Automated backups, monthly restore tests | Low | Open |
| R-004 | Insider threat / privilege misuse | Access Control | 2 | 4 | 8 | Medium | CISO | Least privilege, audit logging, DLP | Low | Open |
| R-005 | Phishing / social engineering | Cybersecurity | 4 | 3 | 12 | High | CISO | Security awareness training, email filtering | Medium | Open |
| R-006 | Change management failures | ITGC | 3 | 3 | 9 | Medium | Change Manager | Change approval board, rollback procedures | Low | Open |

---

## Risk Categories
- Access Control
- Vulnerability Management
- Business Continuity / DR
- Cybersecurity
- ITGC (IT General Controls)
- Compliance / Regulatory
- Third-Party / Vendor
- Data Integrity

