# Application Security Risk Assessment Tool

A Python command-line tool that conducts structured security risk 
assessments for applications and generates colour-coded Excel reports 
with weighted risk scores, findings, and remediation recommendations.

## What it does

- Walks an assessor through 12 security questions across 5 categories
- Calculates a weighted risk score reflecting real-world risk priorities
- Generates a two-sheet Excel report:
  - Executive Summary with overall risk score and category breakdown
  - Detailed Findings with colour-coded scores and specific recommendations
- Automatically names output files by application name and date

## Risk categories and weights

| Category | Weight | Why |
|---|---|---|
| Access Control | 25% | Directly enables or prevents unauthorised access |
| Data Protection | 25% | Core requirement for GDPR and data breach prevention |
| Vulnerability Management | 20% | Unpatched systems are the most common attack vector |
| Incident Response | 15% | Determines how quickly breaches are detected and contained |
| Compliance & Certifications | 15% | Administrative assurance — important but not directly preventive |

## Risk scoring

| Score | Risk Level | Meaning |
|---|---|---|
| 0 – 9 | Low | Strong security posture — proceed with standard monitoring |
| 10 – 19 | Medium | Gaps identified — remediation required within 90 days |
| 20 – 29 | High | Significant gaps — escalation required before approval |
| 30 – 40 | Critical | Critical failures — do not approve without immediate remediation |

## Frameworks referenced

- **ISO 27001 Annex A** — control categories and requirements
- **NIST Cybersecurity Framework** — Identify, Protect, Detect, Respond, Recover
- **DORA (Digital Operational Resilience Act)** — EU regulation effective 
  January 2025, governing ICT risk management for financial institutions
- **GDPR / DSGVO** — data protection compliance questions

## Tech stack

- Python 3.11
- openpyxl — Excel report generation
- json — structured question and recommendation data
- No external API or internet connection required

## How to run

```bash
pip install openpyxl
python assessor.py
```

Answer A, B, C, or D for each question. The Excel report is saved 
automatically in the current folder.

## Sample output

Running with all D answers (worst case) produces a Critical risk report.  
Running with all A answers (best case) produces a Low risk report.  
Both sample reports are included in this repository for reference.

## Project structure
app-risk-assessment/
├── assessor.py          — main script, runs assessment and builds report
├── questions.json       — all assessment questions and answer scores
├── recommendations.json — remediation recommendations per answer
└── README.md

## Relevance to corporate IT security

This tool replicates the application security risk assessment process 
used in corporate IT security teams — particularly in regulated industries 
like financial services. The weighted scoring model, ISO 27001 alignment, 
and DORA reference make it directly applicable to the risk management 
workflows at financial institutions.

## What I would add next

- CSV input mode so multiple applications can be assessed in batch
- Trend tracking across multiple assessments of the same application
- Integration with a SQLite database for historical reporting
- A Streamlit web interface to replace the command-line input
