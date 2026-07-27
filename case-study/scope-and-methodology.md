# Scope and Methodology

## Engagement Overview

**Entity:** Bayou Cloud Services  
**Audit Type:** IT General Controls and GRC Readiness Review  
**Audit Period:** January 1, 2026 through June 30, 2026  
**Prepared By:** Sandesh Lama Tamang  
**Data Classification:** Synthetic portfolio data

The objective of this review was to evaluate whether selected IT general controls were designed and operating effectively to support secure system access, reliable change management, logging, vulnerability remediation, and operational resilience.

## Systems In Scope

| System | Purpose | Risk Relevance |
|---|---|---|
| Active Directory | Identity and privileged access management | Central authentication and administrator access |
| ERP | Financial transaction processing | Supports financial reporting and business operations |
| Financial System | General ledger and reporting | Sensitive financial data and reporting integrity |
| Payroll | Payroll processing | Sensitive employee and compensation data |
| Change Manager | Change ticketing and approval | Supports change authorization and audit trail |
| Backup Platform | Backup scheduling and restore evidence | Supports business continuity |

## Control Areas

| Control Area | Audit Focus |
|---|---|
| Access to Programs and Data | User access review, terminated users, MFA, privileged access, least privilege |
| Change Management | Approval before implementation, testing evidence, emergency changes, segregation of duties |
| Computer Operations | Audit logging, monitoring, backups, restore testing |
| Risk and Remediation Governance | Risk ownership, overdue remediation, control effectiveness, POA&M tracking |

## Criteria and Framework Mapping

The case study maps controls and findings to commonly recognized audit and security frameworks:

- SOC 2 Trust Services Criteria: CC6 and CC7
- COBIT 2019: APO12, BAI06, DSS01, DSS05
- NIST SP 800-53 Rev. 5: AC-2, AC-6, IA-2, AU-2, AU-6, CM-3, CM-6, RA-5, SI-2, CP-9
- NIST Cybersecurity Framework 2.0: Govern, Identify, Protect, Detect, Respond, Recover
- CISA IT audit domains: audit process, IT governance, operations/resilience, and protection of information assets

## Testing Approach

### Access Review Testing

The access review analyzer tested the full synthetic user access population for:

- Terminated users retaining active access
- Dormant active accounts
- Privileged and sensitive-system users without MFA
- Non-IT users with administrative access
- Contractor elevated access
- Duplicate user/system access records

### Change Management Testing

The change log sampler evaluated a population of 75 changes. The sample included all emergency changes as targeted selections and a seeded random sample from the remaining normal changes.

Testing focused on:

- Approval prior to implementation
- Evidence of testing
- Segregation of duties between requester, approver, and implementer
- Emergency change identification
- Completed changes with complete implementation dates

### Risk and Remediation Testing

The risk score calculator assessed inherent and residual risk using a 5x5 likelihood and impact model. Control effectiveness was applied to calculate residual exposure. Exceptions were identified for missing risk owners, weak controls, overdue mitigations, and stale open risks.

### Evidence Reliability

Evidence was rated as:

- **Reliable:** complete, dated, system-generated, or independently supportable
- **Partially Reliable:** useful but missing a field, sign-off, date, or clear owner
- **Insufficient:** not received, incomplete, or not enough to support a control conclusion

## Sampling Note

This portfolio uses transparent demonstration sampling assumptions. A real engagement would document sample size based on the governing audit methodology, control frequency, population size, expected deviation, tolerable deviation, sampling risk, and auditor judgment.

## Limitations

This is a synthetic portfolio project. It does not represent an actual audit opinion, SOC 2 report, DoD authorization package, or client engagement. The purpose is to demonstrate practical IT audit execution and reporting skills.
