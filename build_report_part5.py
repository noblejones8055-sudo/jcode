#!/usr/bin/env python3
"""Part 5: Sections 5.3.12 KPI Dashboard, 5.3.13 Management Approach, 5.3.14 Content Index."""

import sys
sys.path.insert(0, '/home/user/jcode')
from build_report_helpers import *

def build_part5(doc):

    # ═══════════════════════════════════════════════════════════
    # 5.3.12  KPI Dashboard
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.12  KPI Dashboard — Key Pollution Performance Indicators — FY 2025', level=2)
    code_para(doc, 'ESRS E2-4 | ESRS E2-5 | ESRS E2-6 | GRI 305-7')

    tbl(doc,
        'Table 5.3.12.1 — Pollution KPI Dashboard: Summary of Key Performance Indicators '
        '(FY 2024 vs FY 2025)',
        'ESRS E2-4 | ESRS E2-5 | ESRS E2-6 | GRI 305-7',
        ['KPI Category', 'Key Performance Indicator', 'FY 2024', 'FY 2025', 'YoY Change'],
        [
            ('Air Emissions',      'Total Emissions to Air (metric tonnes)', '2,209.5',  '2,023.0', '−8.4%'),
            ('Air Emissions',      'NOx (metric tonnes)',                    '1,245.8',  '1,141.0', '−8.4%'),
            ('Air Emissions',      'SOx (metric tonnes)',                    '387.4',    '343.9',   '−11.2%'),
            ('Air Emissions',      'PM10 (metric tonnes)',                   '92.6',     '86.4',    '−6.7%'),
            ('Air Emissions',      'VOC (metric tonnes)',                    '234.1',    '218.6',   '−6.6%'),
            ('Water Emissions',    'Total Emissions to Water (metric tonnes)','382.2',   '372.2',   '−2.6%'),
            ('Water Emissions',    'TSS (metric tonnes)',                    '48.3',     '43.9',    '−9.1%'),
            ('Water Emissions',    'COD (metric tonnes)',                    '234.6',    '237.7',   '+1.3%'),
            ('Soil Emissions',     'Total Emissions to Soil (metric tonnes)', '44.6',    '42.2',    '−5.4%'),
            ('Microplastics',      'Microplastics Generated (kg)',           '—',        '12,840',  'N/A'),
            ('Microplastics',      'Microplastics Used (metric tonnes)',     '—',        '4.7',     'N/A'),
            ('SoC',                'Total SoC Generated / Used / Procured (kg)', '100,660','87,710','−12.8%'),
            ('SoC',                'Total SoC Leaving Facilities — Combined (kg)','63,662','53,538','−15.9%'),
            ('SVHC',               'Total SVHC Generated / Used / Procured (kg)','22,500','18,340','−18.5%'),
            ('SVHC',               'Total SVHC Leaving Facilities — Combined (kg)','13,653','11,091','−18.8%'),
            ("Financial Effects",  "Net Anticipated Financial Effect (USD '000)", '—',   '17,440',  'N/A'),
            ("Financial Effects",  "Total Pollution-Related Expenditure and Provisions (USD '000)", '—', '25,240', 'N/A'),
        ])

    # ═══════════════════════════════════════════════════════════
    # 5.3.13  Management Approach
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.13  Management Approach to Pollution Prevention and Control', level=2)
    code_para(doc, 'ESRS E2 | ESRS E2-IRO-1 | ESRS E2-1 | ESRS E2-2 | ESRS E2-3')

    body(doc,
         "XYZ Logistics LLC's management approach to pollution prevention and control is "
         "anchored in a formal governance framework that integrates environmental risk "
         "management, regulatory compliance, and strategic sustainability commitments. The "
         "Board of Directors, through its Environmental and Sustainability Committee, maintains "
         "oversight of the Company's pollution-related risks and targets. The Chief Sustainability "
         "Officer holds executive accountability for pollution performance, supported by Regional "
         "Environmental Leads at each primary operating site.")

    body(doc,
         "The Company's Pollution Prevention and Environmental Stewardship Policy establishes the "
         "overarching framework for pollution management across all operational boundaries. This "
         "policy is subject to annual review and is aligned with the EU Zero Pollution Action "
         "Plan, REACH Regulation, the Industrial Emissions Directive, and applicable national "
         "environmental legislation in all jurisdictions where the Company operates. The policy "
         "requires all operating entities to maintain certified Environmental Management Systems "
         "aligned with ISO 14001, conduct annual environmental risk register updates, and "
         "implement site-level pollution prevention action plans.")

    body(doc,
         "Material pollution-related impacts, risks, and opportunities are identified through the "
         "Company's annual double materiality assessment, which combines impact materiality "
         "evaluation (effects on environment and communities) with financial materiality analysis "
         "(effects on enterprise value). The outcomes of this assessment inform target-setting, "
         "resource allocation, and disclosure priorities. Stakeholder engagement — including "
         "community consultations, regulatory dialogue, and supplier assessments — is embedded "
         "in the assessment process to ensure that external perspectives are incorporated into "
         "pollution management priorities.")

    body(doc,
         "Performance against pollution targets is tracked quarterly through the environmental "
         "KPI dashboard and reported to the Board's Environmental and Sustainability Committee. "
         "All material pollution incidents are subject to root cause analysis, corrective action "
         "planning, and regulatory reporting in accordance with applicable permit conditions. "
         "The Company's Environmental Management Information System (EMIS) provides the "
         "centralised platform for data collection, quality assurance, and reporting across all "
         "pollution parameters.")

    # ═══════════════════════════════════════════════════════════
    # 5.3.14  Content Index
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.14  Content Index — ESRS E2 and GRI Framework Mapping', level=2)
    code_para(doc, 'ESRS E2 | GRI 305-7')

    body(doc,
         "The table below provides a cross-reference between the disclosure requirements of "
         "ESRS E2 (Pollution) and GRI 305 (Emissions), and the corresponding sections of this "
         "report where each requirement is addressed.")

    tbl(doc,
        'Table 5.3.14.1 — Content Index: ESRS E2 and GRI 305-7 Disclosure Requirements vs '
        'Report Section Cross-Reference',
        'ESRS E2 | GRI 305-7',
        ['ESRS / GRI Reference', 'Disclosure Requirement', 'Report Section'],
        [
            ('ESRS E2-IRO-1-11-(a)', 'Process to identify actual and potential pollution-related impacts, risks and opportunities', '5.3.1.1'),
            ('ESRS E2-IRO-1-11-(a)', 'Sites and activities screened; methodologies, assumptions and tools used', '5.3.1.2'),
            ('ESRS E2-IRO-1-11-(b)', 'Stakeholder and community consultations in pollution impact assessment', '5.3.1.3'),
            ('ESRS E2-1-15-(a)',     'Policy coverage: mitigation of air, water and soil pollution impacts', '5.3.2.1'),
            ('ESRS E2-1-15-(b)',     'Policy coverage: substitution/minimisation of SoC; phase-out of SVHC', '5.3.2.2'),
            ('ESRS E2-1-15-(c)',     'Policy coverage: prevention and response to pollution incidents', '5.3.2.3'),
            ('ESRS E2-1-AR-12',     'Contextual alignment with EU Zero Pollution Action Plan', '5.3.2.4'),
            ('ESRS E2-2-19',        'Mitigation hierarchy allocation — pollution actions', '5.3.3.1'),
            ('ESRS E2-2-AR-13',     'Pollution actions extending to value chain', '5.3.3.2'),
            ('ESRS E2-2-19',        'Mitigation hierarchy allocation — pollution resources', '5.3.3.3'),
            ('ESRS E2-2-AR-15',     'Site-level pollution prevention action plans', '5.3.3.4'),
            ('ESRS E2-3-23-(a)',    'Target: prevention and control of air pollutants', '5.3.4.1'),
            ('ESRS E2-3-23-(b)',    'Target: prevention and control of water emissions', '5.3.4.2'),
            ('ESRS E2-3-23-(c)',    'Target: prevention and control of soil pollution', '5.3.4.3'),
            ('ESRS E2-3-23-(d)',    'Target: prevention and control of SoC and SVHC', '5.3.4.4'),
            ('ESRS E2-3-24 | E2-3-AR-16-(a)', 'Ecological thresholds and entity-specific allocations', '5.3.4.5'),
            ('ESRS E2-3-24 | E2-3-AR-16-(a)-(b)', 'Methodology for ecological threshold identification', '5.3.4.6'),
            ('ESRS E2-3-24 | E2-3-25 | E2-3-AR-16-(c) | E2-3-AR-17', 'Responsibility, mandatory/voluntary status, Substantial Contribution criteria', '5.3.4.7'),
            ('ESRS E2-3-AR-18',    'Pollution targets implemented at site level', '5.3.4.8'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a) | GRI 305-7', 'Emissions to air by pollutant — summary', '5.3.5.1'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to air by pollutant — by sector', '5.3.5.2'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to air by pollutant — by geographical area', '5.3.5.3'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to air by pollutant — by source', '5.3.5.4'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to air by pollutant — by site location', '5.3.5.5'),
            ('ESRS E2-4-28 | E2-4-AR-20-(b)', 'Microplastics generated and used — air context', '5.3.5.6'),
            ('ESRS E2-4-30-(a)',    'Description of changes over time — air', '5.3.5.7'),
            ('ESRS E2-4-30-(b)',    'Measurement methodologies — air', '5.3.5.8'),
            ('ESRS E2-4-30-(c)',    'Data collection processes — air', '5.3.5.9'),
            ('ESRS E2-4-31-(c)',    'Inferior methodology — air emissions', '5.3.5.10'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to water by pollutant — summary', '5.3.6.1'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to water — by sector, geo area, source, site', '5.3.6.2–5.3.6.5'),
            ('ESRS E2-4-28 | E2-4-AR-20-(b)', 'Microplastics — water context', '5.3.6.6'),
            ('ESRS E2-4-AR-23-(c)', 'Emissions to water in areas with water stress', '5.3.6.7'),
            ('ESRS E2-4-30 | E2-4-31', 'Changes over time, methodologies, data collection, inferior methodology — water', '5.3.6.8–5.3.6.11'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to soil by pollutant — summary', '5.3.7.1'),
            ('ESRS E2-4-28 | E2-4-AR-22-(a)', 'Emissions to soil — by sector, geo area, source, site', '5.3.7.2–5.3.7.5'),
            ('ESRS E2-4-28 | E2-4-AR-20-(b)', 'Microplastics — soil context', '5.3.7.6'),
            ('ESRS E2-4-AR-23-(c)', 'Emissions to soil in areas with water stress', '5.3.7.7'),
            ('ESRS E2-4-30 | E2-4-31', 'Changes over time, methodologies, data collection, inferior methodology — soil', '5.3.7.8–5.3.7.11'),
            ('ESRS E2-4-AR-25',    'IED applicability and relevant BREFs', '5.3.8.1'),
            ('ESRS E2-4-AR-25-(a)','List of IED-regulated installations and applicable BAT Conclusions', '5.3.8.2'),
            ('ESRS E2-4-AR-25-(b)','Non-compliance incidents and enforcement actions', '5.3.8.3'),
            ('ESRS E2-4-AR-25-(c)','Actual emission performance vs BAT-AEL levels', '5.3.8.4'),
            ('ESRS E2-4-AR-25-(d)','Actual environmental performance vs BAT-AEPL benchmarks', '5.3.8.5'),
            ('ESRS E2-4-AR-25-(e)','Compliance schedules and derogations under Article 15(4) IED', '5.3.8.6'),
            ('ESRS E2-5-34',       'Amounts of SoC leaving facilities as emissions, products, part of products, services; total generated/used/procured', '5.3.9.1–5.3.9.6'),
            ('ESRS E2-5-35',       'Amounts of SVHC leaving facilities as emissions, products, part of products, services; total generated/used/procured', '5.3.10.1–5.3.10.6'),
            ('ESRS E2-6-40-(a)',   'Quantitative anticipated financial effects — pollution risks and opportunities', '5.3.11.1–5.3.11.3'),
            ('ESRS E2-6-39-(a)',   'Qualitative anticipated financial effects', '5.3.11.4'),
            ('ESRS E2-6-39-(b)',   'Effects considered, related impacts and time horizons', '5.3.11.5'),
            ('ESRS E2-6-39-(c)',   'Critical assumptions, sources and level of uncertainty', '5.3.11.6'),
            ('ESRS E2-6-41',       'Material incidents and deposits with negative pollution financial impacts', '5.3.11.7'),
            ('ESRS E2-6-AR-33',    'Assessment of related products and services at risk', '5.3.11.8'),
        ],
        first_col_hdr=True)
