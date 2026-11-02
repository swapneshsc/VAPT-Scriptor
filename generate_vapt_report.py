#!/usr/bin/env python3
"""
VAPT Report Generator
=====================
Automated tool to generate professional Vulnerability Assessment & Penetration
Testing (VAPT) reports in Microsoft Word (.docx) format, modelled on industry
best-practice structure (cover, control, executive summary, findings with PoC,
attack narrative, remediation roadmap, etc.).

Usage:
    python generate_vapt_report.py sample_input.json -o GoldenEye_VAPT_Report.docx
    python generate_vapt_report.py my_findings.json --output report.docx

Input is a JSON file describing the engagement, findings, narrative, etc.
See sample_input.json for the full schema.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Inches, Pt, RGBColor, Cm, Twips
from docx.enum.style import WD_STYLE_TYPE


# ---------------------------------------------------------------------------
# Colour palette (severity + branding)
# ---------------------------------------------------------------------------
SEVERITY_COLORS = {
    "CRITICAL": RGBColor(0xC0, 0x00, 0x00),   # dark red
    "HIGH":     RGBColor(0xE0, 0x60, 0x00),   # orange-red
    "MEDIUM":   RGBColor(0xD4, 0xA0, 0x00),   # amber
    "LOW":      RGBColor(0x2E, 0x7D, 0x32),   # green
    "INFO":     RGBColor(0x15, 0x65, 0xC0),   # blue
    "INFORMATIONAL": RGBColor(0x15, 0x65, 0xC0),
}

HEADER_BG     = "1A237E"   # deep indigo
HEADER_FG     = "FFFFFF"
TABLE_ALT_BG  = "F5F5F5"
BORDER_COLOR  = "B0BEC5"
ACCENT        = RGBColor(0x1A, 0x23, 0x7E)


def set_cell_shading(cell, hex_color: str) -> None:
    """Set background colour of a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell, **kwargs) -> None:
    """Set borders on a cell. kwargs: top, bottom, left, right = (sz, color)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge, (sz, color) in kwargs.items():
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(sz))
        element.set(qn("w:color"), color)
        tcBorders.append(element)
    tcPr.append(tcBorders)


def add_horizontal_line(paragraph) -> None:
    """Add a bottom border line under a paragraph."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), HEADER_BG)
    pBdr.append(bottom)
    pPr.append(pBdr)


def set_run_font(run, name="Calibri", size=11, bold=False, italic=False,
                 color: Optional[RGBColor] = None) -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def add_heading_custom(doc: Document, text: str, level: int = 1) -> None:
    """Add a styled heading."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, size=16, bold=True, color=ACCENT)
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        add_horizontal_line(p)
    elif level == 2:
        set_run_font(run, size=13, bold=True, color=ACCENT)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
    else:
        set_run_font(run, size=11, bold=True, color=RGBColor(0x33, 0x33, 0x33))
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)


def add_body(doc: Document, text: str, size: int = 10, space_after: int = 6) -> None:
    """Add a body paragraph; supports simple newlines."""
    for block in text.split("\n\n"):
        p = doc.add_paragraph()
        run = p.add_run(block.strip())
        set_run_font(run, size=size)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.15


def add_code_block(doc: Document, text: str) -> None:
    """Monospace block for commands / evidence."""
    for line in text.strip().splitlines():
        p = doc.add_paragraph()
        run = p.add_run(line)
        set_run_font(run, name="Consolas", size=8, color=RGBColor(0x21, 0x21, 0x21))
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Inches(0.15)
        # light background via shading on the paragraph is harder; keep simple
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "F0F4F8")
        shd.set(qn("w:val"), "clear")
        pPr.append(shd)


def make_table(doc: Document, headers: List[str], rows: List[List[str]],
               col_widths: Optional[List[float]] = None,
               header_bg: str = HEADER_BG) -> None:
    """Create a clean professional table."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        set_run_font(run, size=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, header_bg)

    # Body
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            # colour severity cells
            text = str(val)
            color = None
            bold = False
            if text.upper() in SEVERITY_COLORS:
                color = SEVERITY_COLORS[text.upper()]
                bold = True
            elif text.upper() in ("UNSAFE", "OPEN"):
                color = SEVERITY_COLORS["CRITICAL"]
                bold = True
            elif text.upper() in ("SAFE",):
                color = SEVERITY_COLORS["LOW"]
                bold = True
            run = p.add_run(text)
            set_run_font(run, size=8, bold=bold, color=color)
            if r_idx % 2 == 1:
                set_cell_shading(cell, TABLE_ALT_BG)

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)

    doc.add_paragraph()  # spacer


def severity_badge_paragraph(doc: Document, severity: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(f"  {severity.upper()}  ")
    set_run_font(run, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    # approximate badge with shading on paragraph
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    hex_map = {
        "CRITICAL": "C00000",
        "HIGH": "E06000",
        "MEDIUM": "D4A000",
        "LOW": "2E7D32",
        "INFO": "1565C0",
        "INFORMATIONAL": "1565C0",
    }
    shd.set(qn("w:fill"), hex_map.get(severity.upper(), "666666"))
    shd.set(qn("w:val"), "clear")
    pPr.append(shd)
    p.paragraph_format.space_after = Pt(6)


# ---------------------------------------------------------------------------
# Report builders
# ---------------------------------------------------------------------------

def build_cover(doc: Document, data: Dict[str, Any]) -> None:
    meta = data["meta"]

    # Spacer
    for _ in range(3):
        doc.add_paragraph()

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta["report_title"].upper())
    set_run_font(run, size=22, bold=True, color=ACCENT)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta.get("target_description", ""))
    set_run_font(run, size=12, italic=True, color=RGBColor(0x55, 0x55, 0x55))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta.get("assessment_type", ""))
    set_run_font(run, size=11, color=RGBColor(0x55, 0x55, 0x55))

    doc.add_paragraph()

    # Target box
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Target System:  {meta['target_system']}")
    set_run_font(run, size=14, bold=True)

    doc.add_paragraph()

    # Meta table
    rows = [
        ["Prepared by", meta["prepared_by"]],
        ["Assessment Window", meta["assessment_window"]],
        ["Report Date", meta["report_date"]],
        ["Version", meta["version"]],
        ["Classification", meta["classification"]],
    ]
    make_table(doc, ["Field", "Value"], rows, col_widths=[2.2, 4.5])

    # Confidential banner
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("CONFIDENTIAL — NOT TO BE CIRCULATED OR REPRODUCED WITHOUT AUTHORIZATION")
    set_run_font(run, size=9, bold=True, color=SEVERITY_COLORS["CRITICAL"])

    doc.add_page_break()


def build_document_control(doc: Document, data: Dict[str, Any]) -> None:
    meta = data["meta"]
    add_heading_custom(doc, "Document Control", 1)

    rows = [
        ["Document Title", meta["report_title"]],
        ["Target / System Under Test", f"{meta['target_system']} ({meta.get('os', '')})"],
        ["Classification", "Confidential"],
        ["Document Type", "Report"],
        ["Application / Service Name", meta.get("application", "")],
        ["Audit Type", meta.get("audit_type", "")],
        ["Engagement Objective", meta.get("objective", "")],
        ["Testing Standard", meta.get("standards", "")],
        ["Report Version", meta["version"]],
        ["Distribution", meta.get("distribution", "")],
    ]
    make_table(doc, ["Attribute", "Value"], rows, col_widths=[2.5, 4.2])

    add_heading_custom(doc, "Version History", 2)
    vh_rows = [[v["version"], v["date"], v["author"], v["description"]]
               for v in data.get("version_history", [])]
    make_table(doc, ["Version", "Date", "Author", "Description"], vh_rows,
               col_widths=[0.9, 1.3, 1.5, 3.0])

    add_heading_custom(doc, "Contact Details", 2)
    contact_rows = [[c["name"], c["role"], c.get("email", "-"), c.get("organization", "-")]
                    for c in data.get("contacts", [])]
    make_table(doc, ["Name", "Role", "Email ID", "Organization"], contact_rows,
               col_widths=[1.8, 2.2, 1.8, 1.5])

    add_heading_custom(doc, "Test Team", 2)
    team = data.get("test_team", {})
    make_table(doc, ["Role", "Name"], [
        ["Security Analyst", team.get("security_analyst", "")],
        ["Reviewed / QA by", team.get("reviewed_by", "")],
    ], col_widths=[2.5, 4.2])

    p = doc.add_paragraph()
    run = p.add_run("Legal & authorization: ")
    set_run_font(run, size=9, bold=True)
    run = p.add_run(meta.get("legal_note", ""))
    set_run_font(run, size=9)

    doc.add_page_break()


def build_toc(doc: Document) -> None:
    add_heading_custom(doc, "Table of Contents", 1)
    toc_items = [
        "1.  Executive Summary",
        "2.  Scope & Methodology",
        "3.  Vulnerability Summary & Distribution",
        "4.  Mitigation Tracker",
        "5.  Attack Narrative — End-to-End Kill Chain (with commands)",
        "6.  Detailed Findings (with Step-by-Step Reproduction & PoC)",
        "7.  Business Impact Analysis",
        "8.  OWASP Top 10 — 2021 Coverage",
        "9.  SANS / CWE Top 25 Coverage",
        "10. Consolidated Remediation Roadmap",
        "11. Secure Coding Guidelines",
        "12. Interpreting Risk Ratings",
        "13. Observations",
        "14. Abbreviations",
    ]
    for item in toc_items:
        p = doc.add_paragraph()
        run = p.add_run(item)
        set_run_font(run, size=11)
        p.paragraph_format.space_after = Pt(4)
    doc.add_page_break()


def build_executive_summary(doc: Document, data: Dict[str, Any]) -> None:
    es = data["executive_summary"]
    counts = data.get("severity_counts", {})
    add_heading_custom(doc, "1. Executive Summary", 1)

    add_heading_custom(doc, "Introduction", 2)
    add_body(doc, es["introduction"])

    add_heading_custom(doc, "Objective", 2)
    add_body(doc, es["objective"])

    add_heading_custom(doc, "Summary of Results", 2)
    add_body(doc, es["summary_of_results"])

    # Overall risk callout
    p = doc.add_paragraph()
    run = p.add_run(f"Overall Risk Rating: {es['overall_risk']}")
    set_run_font(run, size=12, bold=True, color=SEVERITY_COLORS.get(es["overall_risk"], ACCENT))
    add_body(doc, es["overall_risk_rationale"])

    add_heading_custom(doc, "High-Level Recommendation", 2)
    add_body(doc, es["high_level_recommendation"])

    # Quick counts table
    make_table(doc, ["Critical", "High", "Medium", "Low", "Informational"],
               [[str(counts.get("critical", 0)),
                 str(counts.get("high", 0)),
                 str(counts.get("medium", 0)),
                 str(counts.get("low", 0)),
                 str(counts.get("informational", 0))]],
               col_widths=[1.3, 1.3, 1.3, 1.3, 1.5])
    doc.add_page_break()


def build_scope(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "2. Scope & Methodology", 1)
    add_heading_custom(doc, "2.1 Scope", 2)
    scope = data["scope"]
    rows = [
        ["Target", scope["target"]],
        ["Target Address", scope["target_address"]],
        ["In-Scope Services", scope["in_scope_services"]],
        ["Objective", scope["objective"]],
        ["Constraints", scope["constraints"]],
        ["Testing Approach", scope["testing_approach"]],
    ]
    make_table(doc, ["Attribute", "Detail"], rows, col_widths=[2.0, 4.7])

    add_heading_custom(doc, "2.2 Methodology", 2)
    add_body(doc, "The assessment followed a structured methodology aligned with the Penetration Testing Execution Standard (PTES) and OWASP Testing Guide principles, mapped to MITRE ATT&CK where relevant:")
    meth_rows = [[m["phase"], m["activities"]] for m in data.get("methodology", [])]
    make_table(doc, ["Phase", "Activities"], meth_rows, col_widths=[2.2, 4.5])

    add_heading_custom(doc, "2.3 Tools Used", 2)
    add_body(doc, data.get("tools_used", ""))
    doc.add_page_break()


def build_vuln_summary(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "3. Vulnerability Summary & Distribution", 1)
    counts = data.get("severity_counts", {})

    add_heading_custom(doc, "3.1 Severity Count", 2)
    make_table(doc, ["Critical", "High", "Medium", "Low", "Informational"],
               [[str(counts.get("critical", 0)),
                 str(counts.get("high", 0)),
                 str(counts.get("medium", 0)),
                 str(counts.get("low", 0)),
                 str(counts.get("informational", 0))]],
               col_widths=[1.3, 1.3, 1.3, 1.3, 1.5])

    add_heading_custom(doc, "3.2 Findings at a Glance", 2)
    findings = data.get("findings", [])
    glance = []
    for i, f in enumerate(findings, 1):
        glance.append([
            str(i),
            f["id"],
            f["title"][:55] + ("…" if len(f["title"]) > 55 else ""),
            f.get("affected", "")[:30],
            f.get("cwe", "").split("—")[0].strip()[:12],
            f["severity"],
            f.get("cvss", "").split()[0] if f.get("cvss") else "",
        ])
    make_table(doc,
               ["Sr.", "ID", "Vulnerability", "Affected Service", "CWE", "Risk", "CVSS"],
               glance,
               col_widths=[0.5, 0.9, 2.2, 1.5, 0.9, 0.9, 0.6])
    doc.add_page_break()


def build_mitigation_tracker(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "4. Mitigation Tracker", 1)
    add_body(doc, "Update the Status column as fixes are applied and re-tested.")
    findings = data.get("findings", [])
    rows = []
    for f in findings:
        owner = "<App Team>" if "Moodle" in f.get("affected", "") or "HTTP" in f.get("affected", "") else "<Infra Team>"
        if "Credential" in f["title"] or "password" in f["title"].lower():
            owner = "<IAM Team>"
        target = "Immediate" if f["severity"] == "CRITICAL" else ("< 1 week" if f["severity"] == "HIGH" else "< 3 weeks")
        rows.append([f["id"], f["title"][:45], f["severity"], owner, target, "OPEN"])
    make_table(doc,
               ["ID", "Vulnerability Name", "Risk", "Owner", "Target Date", "Status"],
               rows,
               col_widths=[0.9, 2.4, 0.9, 1.2, 1.1, 0.8])
    doc.add_page_break()


def build_attack_narrative(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "5. Attack Narrative — End-to-End Kill Chain", 1)
    add_body(doc, "This section demonstrates the assessor's full thought process and the exact commands used, showing how each finding fed the next and how a chain of individually \"lower\" issues culminated in root compromise.")

    for step in data.get("attack_narrative", []):
        add_heading_custom(doc, f"Step {step['step']}  ({step['phase']})", 2)
        add_body(doc, step["description"])
        if step.get("commands"):
            add_code_block(doc, step["commands"])

    p = doc.add_paragraph()
    run = p.add_run("Result: ")
    set_run_font(run, size=10, bold=True, color=SEVERITY_COLORS["CRITICAL"])
    run = p.add_run("unauthenticated external attacker → www-data (RCE) → root, with the objective flag captured. Complete loss of confidentiality, integrity and availability.")
    set_run_font(run, size=10)
    doc.add_page_break()


def build_detailed_findings(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "6. Detailed Findings", 1)
    add_body(doc, "Each finding records technical detail, a defensible risk rating, full step-by-step reproduction / proof of concept, captured evidence, and remediation.")

    for f in data.get("findings", []):
        # Title with ID
        p = doc.add_paragraph()
        run = p.add_run(f"{f['id']} — {f['title']}")
        set_run_font(run, size=12, bold=True, color=ACCENT)
        p.paragraph_format.space_before = Pt(16)
        add_horizontal_line(p)

        severity_badge_paragraph(doc, f["severity"])

        # Meta table
        meta_rows = [
            ["Severity", f["severity"]],
            ["Ease of Exploitation", f.get("ease", "")],
            ["CWE / SANS", f.get("cwe", "")],
            ["CVE", f.get("cve") or "—"],
            ["OWASP Category", f.get("owasp", "")],
            ["CVSS v3.1", f.get("cvss", "")],
            ["Affected System / App", f.get("affected", "")],
        ]
        make_table(doc, ["Attribute", "Value"], meta_rows, col_widths=[2.0, 4.7])

        add_heading_custom(doc, "Description", 3)
        add_body(doc, f.get("description", ""))

        add_heading_custom(doc, "Risk", 3)
        add_body(doc, f.get("risk", ""))

        if f.get("preconditions"):
            add_heading_custom(doc, "Pre-conditions", 3)
            add_body(doc, f["preconditions"])

        add_heading_custom(doc, "Steps to Reproduce / PoC", 3)
        if f.get("steps"):
            add_code_block(doc, f["steps"])

        add_heading_custom(doc, "Evidence", 3)
        if f.get("evidence"):
            add_code_block(doc, f["evidence"])

        add_heading_custom(doc, "Recommendation", 3)
        add_body(doc, f.get("recommendation", ""))

        if f.get("references"):
            p = doc.add_paragraph()
            run = p.add_run("References: ")
            set_run_font(run, size=9, bold=True)
            run = p.add_run(f["references"])
            set_run_font(run, size=9, italic=True)

    doc.add_page_break()


def build_business_impact(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "7. Business Impact Analysis", 1)
    bi = data.get("business_impact", {})
    add_body(doc, bi.get("summary", ""))
    p = doc.add_paragraph()
    run = p.add_run(f"CIA impact: {bi.get('cia', '')}")
    set_run_font(run, size=11, bold=True)
    doc.add_page_break()


def build_owasp(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "8. OWASP Top 10 — 2021 Coverage", 1)
    rows = [[o["id"], o["name"], o["status"], o["related"]]
            for o in data.get("owasp_coverage", [])]
    make_table(doc, ["#", "OWASP Top 10 — 2021", "Status", "Related Findings"],
               rows, col_widths=[0.6, 2.8, 1.0, 2.3])
    doc.add_page_break()


def build_cwe(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "9. SANS / CWE Top 25 Coverage", 1)
    add_body(doc, "The weaknesses exploited map to the following CWE entries. Entries not listed were not found exploitable within scope.")
    rows = [[c["id"], c["name"], c["status"], c["finding"]]
            for c in data.get("cwe_coverage", [])]
    make_table(doc, ["ID", "Name", "Status", "Finding"],
               rows, col_widths=[1.4, 2.8, 1.0, 1.5])
    doc.add_page_break()


def build_roadmap(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "10. Consolidated Remediation Roadmap", 1)
    rm = data.get("remediation_roadmap", {})

    add_heading_custom(doc, "Immediate (Critical / High)", 2)
    for i, item in enumerate(rm.get("immediate", []), 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{i}. {item}")
        set_run_font(run, size=10)
        p.paragraph_format.space_after = Pt(4)

    add_heading_custom(doc, "Medium-Term", 2)
    for i, item in enumerate(rm.get("medium_term", []), 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{i + 5}. {item}")
        set_run_font(run, size=10)
        p.paragraph_format.space_after = Pt(4)

    add_heading_custom(doc, "Long-Term", 2)
    for i, item in enumerate(rm.get("long_term", []), 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{i + 10}. {item}")
        set_run_font(run, size=10)
        p.paragraph_format.space_after = Pt(4)
    doc.add_page_break()


def build_secure_coding(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "11. Secure Coding Guidelines", 1)
    add_body(doc, "The following practices provide multiple layers of defence and directly address the root causes behind this engagement's findings.")
    for i, g in enumerate(data.get("secure_coding_guidelines", []), 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{i}. {g}")
        set_run_font(run, size=10)
        p.paragraph_format.space_after = Pt(3)
    doc.add_page_break()


def build_risk_ratings(doc: Document) -> None:
    add_heading_custom(doc, "12. Interpreting Risk Ratings", 1)
    add_body(doc, "This matrix shows how a risk rating is derived for a threat/vulnerability combination, based on three dimensions: Impact, Ease of exploitation, and Access level required.")

    add_heading_custom(doc, "Impact", 2)
    add_body(doc, "• High: significant security threat — full administrative access, user-level access, or DoS of a critical service. Fix immediately.\n• Medium: significant limitations on direct impact — e.g. significant information leakage, DoS, or significantly limited access.\n• Low: low-level threat — typically information leakage useful to an attacker (e.g. a user list or descriptive error messages).")

    add_heading_custom(doc, "Ease of Exploitation", 2)
    add_body(doc, "• Easy: little-to-no attacker knowledge; exploitable with standard OS tools.\n• Moderate: requires specific knowledge; exploitable using a combination of OS/public tools.\n• Difficult: requires significant effort and knowledge, and may require a combination of attacks.")

    add_heading_custom(doc, "Risk Rating Matrix", 2)
    rows = [
        ["No Physical Access Required", "Easy", "Critical", "High", "Medium"],
        ["No Physical Access Required", "Moderate", "High", "Medium", "Low"],
        ["No Physical Access Required", "Difficult", "Medium", "Medium", "Low"],
        ["Physical Access Required", "Easy", "Medium", "Medium", "Low"],
        ["Physical Access Required", "Moderate", "Medium", "Low", "Low"],
        ["Physical Access Required", "Difficult", "Low", "Low", "Low"],
    ]
    make_table(doc,
               ["Level of Access Required", "Ease of Exploitation", "Impact: High", "Impact: Medium", "Impact: Low"],
               rows,
               col_widths=[2.0, 1.5, 1.1, 1.2, 1.0])
    add_body(doc, "Critical is applied where a 'No Physical Access / Easy / High Impact' combination additionally results in full system or administrative compromise.")
    doc.add_page_break()


def build_observations(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "13. Observations", 1)
    add_body(doc, "Weaknesses in this section are informational best-practice items. They cannot be exploited directly, or are difficult to exploit due to constraints, but should be addressed to improve the overall posture.")
    rows = [[o["id"], o["observation"], o["recommendation"]]
            for o in data.get("observations", [])]
    make_table(doc, ["#", "Observation", "Recommendation"],
               rows, col_widths=[0.8, 3.0, 3.0])
    doc.add_page_break()


def build_abbreviations(doc: Document, data: Dict[str, Any]) -> None:
    add_heading_custom(doc, "14. Abbreviations", 1)
    rows = [[a["term"], a["definition"]] for a in data.get("abbreviations", [])]
    make_table(doc, ["Term", "Definition"], rows, col_widths=[1.8, 5.0])

    # Footer note
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    run = p.add_run("CONFIDENTIAL. ")
    set_run_font(run, size=8, bold=True, color=SEVERITY_COLORS["CRITICAL"])
    run = p.add_run(
        "This report documents testing performed, with authorization, against an isolated laboratory target "
        "for security assessment and remediation purposes only. Reproduction steps, credentials and flag values "
        "relate solely to that lab instance."
    )
    set_run_font(run, size=8)


def add_header_footer(doc: Document, data: Dict[str, Any]) -> None:
    meta = data["meta"]
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    # Header
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"{meta['target_system']}  |  VAPT Report  |  {meta['version']}")
    set_run_font(run, size=8, color=RGBColor(0x88, 0x88, 0x88))

    # Footer with page numbers
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("CONFIDENTIAL  •  Page ")
    set_run_font(run, size=8, color=RGBColor(0x88, 0x88, 0x88))

    # PAGE field
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run2 = p.add_run()
    run2._r.append(fldChar1)
    run2._r.append(instrText)
    run2._r.append(fldChar2)
    set_run_font(run2, size=8, color=RGBColor(0x88, 0x88, 0x88))

    run3 = p.add_run(" of ")
    set_run_font(run3, size=8, color=RGBColor(0x88, 0x88, 0x88))

    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "begin")
    instrText2 = OxmlElement("w:instrText")
    instrText2.text = "NUMPAGES"
    fldChar4 = OxmlElement("w:fldChar")
    fldChar4.set(qn("w:fldCharType"), "end")
    run4 = p.add_run()
    run4._r.append(fldChar3)
    run4._r.append(instrText2)
    run4._r.append(fldChar4)
    set_run_font(run4, size=8, color=RGBColor(0x88, 0x88, 0x88))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def generate_report(data: Dict[str, Any], output_path: Path) -> None:
    doc = Document()

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    add_header_footer(doc, data)

    build_cover(doc, data)
    build_document_control(doc, data)
    build_toc(doc)
    build_executive_summary(doc, data)
    build_scope(doc, data)
    build_vuln_summary(doc, data)
    build_mitigation_tracker(doc, data)
    build_attack_narrative(doc, data)
    build_detailed_findings(doc, data)
    build_business_impact(doc, data)
    build_owasp(doc, data)
    build_cwe(doc, data)
    build_roadmap(doc, data)
    build_secure_coding(doc, data)
    build_risk_ratings(doc)
    build_observations(doc, data)
    build_abbreviations(doc, data)

    doc.save(str(output_path))
    print(f"[+] Report written to: {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a professional VAPT report (DOCX) from a JSON description.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input_json", help="Path to the engagement / findings JSON file")
    parser.add_argument("-o", "--output", default=None,
                        help="Output DOCX path (default: <target>_VAPT_Report.docx)")
    args = parser.parse_args()

    input_path = Path(args.input_json)
    if not input_path.exists():
        print(f"[!] Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    with open(input_path, encoding="utf-8") as fh:
        data = json.load(fh)

    if args.output:
        output_path = Path(args.output)
    else:
        target = data.get("meta", {}).get("target_system", "VAPT").replace(" ", "_").replace(":", "")
        output_path = Path(f"{target}_VAPT_Report.docx")

    generate_report(data, output_path)


if __name__ == "__main__":
    main()
