#!/usr/bin/env python3
"""
Master assembly script for:
  Pollution Report Section 5.3 — FY 2025
  Re-Typographed & Corrected Edition
  Compliance: RULE_01 through RULE_06
"""

import sys
sys.path.insert(0, '/home/user/jcode')

from docx import Document
from docx.shared import Pt, Inches
from build_report_part1 import build_part1
from build_report_part2 import build_part2
from build_report_part3 import build_part3
from build_report_part4 import build_part4
from build_report_part5 import build_part5

OUTPUT = '/home/user/jcode/Pollution_Report_5_3_FY2025_Corrected.docx'

doc = Document()

# ── Page margins ──────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.2)
    section.right_margin  = Inches(1.2)

# ── Assemble all parts ────────────────────────────────────────
print("Building Part 1: Introduction + Sections 5.3.1–5.3.4 …")
build_part1(doc)

print("Building Part 2: Section 5.3.5 — Pollution of Air …")
build_part2(doc)

print("Building Part 3: Sections 5.3.6–5.3.7 — Water & Soil …")
build_part3(doc)

print("Building Part 4: Sections 5.3.8–5.3.11 …")
build_part4(doc)

print("Building Part 5: Sections 5.3.12–5.3.14 — KPI Dashboard, Mgmt Approach, Content Index …")
build_part5(doc)

# ── Save ──────────────────────────────────────────────────────
doc.save(OUTPUT)
print(f"\n✓ Corrected report saved → {OUTPUT}")
