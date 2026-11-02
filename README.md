# VAPT Report Generator

Automated tool that produces professional **Vulnerability Assessment & Penetration Testing (VAPT)** reports in Microsoft Word (`.docx`) format.

The generated structure closely mirrors industry-standard reports and includes:

- Cover / title page with classification banner  
- Document control, version history, contacts & test team  
- Table of contents  
- Executive summary with overall risk rating  
- Scope & methodology  
- Vulnerability summary & distribution tables  
- Mitigation tracker  
- End-to-end attack narrative (kill chain) with commands  
- Detailed findings (severity, CWE/CVE, CVSS, steps-to-reproduce, evidence, remediation)  
- Business impact analysis  
- OWASP Top 10 2021 & CWE coverage matrices  
- Consolidated remediation roadmap (Immediate / Medium / Long-term)  
- Secure coding guidelines  
- Risk-rating interpretation matrix  
- Observations & abbreviations  

## Requirements

```bash
pip install python-docx
```

(Already available in most Kali / assessment environments.)

## Quick Start

```bash
# Generate the sample GoldenEye report
python generate_vapt_report.py sample_input.json -o GoldenEye_VAPT_Report.docx

# Or with your own data
python generate_vapt_report.py my_engagement.json -o ClientName_VAPT_Report.docx
```

## Input Format

All content is driven by a single JSON file.  See `sample_input.json` for a complete, working example based on a real lab engagement (GoldenEye:1).

### Top-level keys

| Key | Purpose |
|-----|---------|
| `meta` | Report title, target, dates, version, classification, legal note |
| `version_history` | Array of `{version, date, author, description}` |
| `contacts` / `test_team` | People involved |
| `executive_summary` | Introduction, objective, results, overall risk, recommendations |
| `scope` | Target, constraints, approach |
| `methodology` | PTES-style phases |
| `tools_used` | Free-text list |
| `severity_counts` | `{critical, high, medium, low, informational}` |
| `findings` | Array of finding objects (see below) |
| `attack_narrative` | Ordered kill-chain steps with commands |
| `business_impact` | Summary + CIA statement |
| `owasp_coverage` / `cwe_coverage` | Status matrices |
| `remediation_roadmap` | `immediate`, `medium_term`, `long_term` arrays |
| `secure_coding_guidelines` | Bullet list |
| `observations` | Informational items |
| `abbreviations` | Glossary |

### Finding object schema

```json
{
  "id": "VAPT-01",
  "title": "OS Command Injection → RCE …",
  "severity": "CRITICAL",          // CRITICAL | HIGH | MEDIUM | LOW | INFO
  "ease": "Moderate …",
  "cwe": "CWE-78 — …",
  "cve": "CVE-2013-3630 — …",
  "owasp": "A03:2021 — Injection",
  "cvss": "9.8 (AV:N/…)",
  "affected": "Moodle LMS 2.2.3 …",
  "description": "…",
  "risk": "…",
  "preconditions": "…",
  "steps": "1. …\n2. …",          // multi-line, rendered as code block
  "evidence": "$ id\nuid=33…",
  "recommendation": "• Upgrade …\n• …",
  "references": "CVE-… · CWE-…"
}
```

## Customisation tips

1. **Copy `sample_input.json`** and replace the content with your engagement data.  
2. Keep severity values exactly as `CRITICAL`, `HIGH`, `MEDIUM`, `LOW`, `INFO` (case-insensitive) so colour coding works.  
3. Multi-line fields (`steps`, `evidence`, `recommendation`) use `\n` for line breaks.  
4. The generator automatically colours severity cells, status cells (`OPEN`/`UNSAFE`/`SAFE`), and applies professional styling (headers, alternating row colours, monospace command blocks, page headers/footers with page numbers).

## Extending the tool

- Add new sections by writing a `build_*` function and calling it from `generate_report()`.  
- Change colours in the `SEVERITY_COLORS` / `HEADER_BG` constants at the top of the script.  
- For screenshots: after generation you can insert images manually or extend the script with `python-docx` `InlineShape` support.

## Licence

Provided as-is for authorised security assessments and report production.  
Do not use against systems without explicit permission.
