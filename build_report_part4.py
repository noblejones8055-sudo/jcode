#!/usr/bin/env python3
"""Part 4: Sections 5.3.8 – 5.3.11."""

import sys
sys.path.insert(0, '/home/user/jcode')
from build_report_helpers import *

def build_part4(doc):

    # ═══════════════════════════════════════════════════════════
    # 5.3.8  Integrated Pollution Control & BAT Performance
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.8  Integrated Pollution Control & BAT Performance Tracking', level=2)
    code_para(doc, 'ESRS E2-4 | ESRS E2-4-AR-25 | ESRS E2-4-AR-25-(a) | ESRS E2-4-AR-25-(b) | '
              'ESRS E2-4-AR-25-(c) | ESRS E2-4-AR-25-(d) | ESRS E2-4-AR-25-(e)')

    # 5.3.8.1 – IED Applicability
    heading(doc, '5.3.8.1  Disclose Applicability of Directive 2010/75/EU (IED) '
            'and Relevant BREFs', level=3)
    code_para(doc, 'ESRS E2-4-AR-25')
    bool_table(doc,
               "State whether the undertaking's activities are subject to Directive "
               "2010/75/EU (IED) and relevant Best Available Techniques Reference Documents "
               "(BREFs)",
               'Yes')
    body(doc,
         "XYZ Logistics LLC confirms that certain of its operations are subject to Directive "
         "2010/75/EU of the European Parliament and of the Council on industrial emissions (IED) "
         "and the relevant Best Available Techniques Reference Documents (BREFs). Four "
         "installations within the Company's operating estate fall within the scope of the IED "
         "and are subject to site-specific IED permits granted by competent authorities in the "
         "respective jurisdictions. Applicable BAT Conclusions are referenced in each "
         "installation's IED permit and form the basis for emission limit values and environmental "
         "performance standards applicable to those installations.")

    # 5.3.8.2 – IED Installations list
    heading(doc, '5.3.8.2  Disclose List of Installations Falling Under IED and '
            'EU BAT Conclusions — FY 2025', level=3)
    code_para(doc, 'ESRS E2-4-AR-25-(a)')
    tbl(doc,
        'Table 5.3.8.2 — IED-Regulated Installations and Applicable BAT Conclusions — FY 2025',
        'ESRS E2-4-AR-25-(a)',
        ['Site / Installation', 'IED Permit Reference', 'Applicable BAT Conclusion',
         'BAT Review Status'],
        [
            ('Dubai (Site 1) — Process Facility',  'UAE-IED-2019-0041',
             'Large Combustion Plants (2021)', 'Current — Next Review FY 2027'),
            ('Mumbai (Site 2) — Chemical Handling', 'IN-ENV-PERMIT-2020-112',
             'Industrial Cooling Systems (2022)', 'Current — Next Review FY 2028'),
            ('Frankfurt (Site 3) — Logistics Hub',  'DE-IED-2018-FRA-083',
             'Surface Treatment using Solvents (2020)', 'Current — Next Review FY 2026'),
            ('Singapore (Site 4) — Refinery Support','SG-EP-2021-047',
             'Waste Treatment (2018)', 'Under Review — Target Completion Q3 FY 2025'),
        ])

    # 5.3.8.3 – Non-Compliance
    heading(doc, '5.3.8.3  Disclose Non-Compliance Incidents and Enforcement Actions '
            'for IED Permit Breaches — FY 2025', level=3)
    code_para(doc, 'ESRS E2-4-AR-25-(b)')
    body(doc,
         "No non-compliance incidents or enforcement actions were recorded during FY 2025 in "
         "relation to permit condition breaches across IED-regulated installations. One minor "
         "permit deviation was self-reported at Site 4 (Singapore) relating to a temporary "
         "exceedance of a COD emission limit during unplanned equipment maintenance. This was "
         "rectified within 72 hours, reported to the competent authority in accordance with "
         "regulatory requirements, and subject to a full root cause analysis. No formal "
         "enforcement action was initiated by any competent authority during the reporting period. "
         "A corrective action plan has been implemented at Site 4 to prevent recurrence.")

    # 5.3.8.4 – BAT-AEL performance
    heading(doc, '5.3.8.4  Disclose Actual Emission Performance vs BAT-AEL Levels '
            '(EU BAT Conclusions) — FY 2025', level=3)
    code_para(doc, 'ESRS E2-4-AR-25-(c)')
    body(doc,
         "Actual emission performance at all IED-regulated installations has been assessed "
         "against emission levels associated with best available techniques (BAT-AEL) as "
         "described in applicable EU BAT Conclusions. All four IED installations report emissions "
         "within or below the applicable BAT-AEL ranges for the relevant pollutant categories. "
         "Site 3 (Frankfurt) achieved performance approximately 18% below the lower end of the "
         "applicable BAT-AEL range for SOx emissions, reflecting the deployment of enhanced "
         "abatement technology exceeding BAT requirements. Site 4 (Singapore) returned to full "
         "BAT-AEL compliance following rectification of the minor permit deviation.")

    # 5.3.8.5 – BAT-AEPL
    heading(doc, '5.3.8.5  Disclose Actual Environmental Performance vs BAT-AEPL '
            'Benchmarks (Sector and Installation) — FY 2025', level=3)
    code_para(doc, 'ESRS E2-4-AR-25-(d)')
    body(doc,
         "Actual environmental performance has been assessed against environmental performance "
         "levels associated with best available techniques (BAT-AEPLs) applicable to the sector "
         "and installation. All installations demonstrate alignment with or outperformance of the "
         "relevant BAT-AEPL benchmarks. The most significant outperformance is observed at "
         "Site 1 (Dubai) for energy-related air emission intensity, where process efficiency "
         "improvements implemented in FY 2024 have reduced specific emission intensity to "
         "approximately 22% below the sector BAT-AEPL benchmark. Site 2 (Mumbai) meets the "
         "BAT-AEPL benchmark for cooling water thermal discharge intensity.")

    # 5.3.8.6 – Compliance Schedules
    heading(doc, '5.3.8.6  Disclose Compliance Schedules and Derogations Under '
            'Article 15(4) IED for BAT-AEL Implementation — FY 2025', level=3)
    code_para(doc, 'ESRS E2-4-AR-25-(e)')
    body(doc,
         "No compliance schedules or derogations under Article 15(4) of the Industrial "
         "Emissions Directive (IED) associated with the implementation of BAT-AELs have been "
         "granted by competent authorities in respect of any XYZ Logistics LLC installation "
         "during FY 2025. The Company confirms that all installations subject to IED permitting "
         "are in compliance with applicable BAT-AEL requirements and no derogation applications "
         "are pending.")

    # ═══════════════════════════════════════════════════════════
    # 5.3.9  Substances of Concern
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.9  Substances of Concern', level=2)
    code_para(doc, 'ESRS E2-5 | ESRS E2-5-34')

    body(doc,
         "XYZ Logistics LLC tracks and discloses the amounts of substances of concern (SoC) "
         "generated, procured, or used in production, and the amounts leaving its facilities as "
         "emissions, products, components of products, or services. All data in this section is "
         "reported in kilograms (kg) unless otherwise stated.")

    soc_hdrs = ['Pollutant / Category', 'FY 2024 (kg)', 'FY 2025 (kg)']

    # 5.3.9.1
    heading(doc, '5.3.9.1  Quantify Amount of Substances of Concern Leaving Facilities '
            'as Emissions — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.1 — SoC: Amount Leaving Facilities as Emissions (FY 2024 vs FY 2025)',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                        '1,240', '980'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)',  '410',   '340'),
            ('Aromatic Hydrocarbons (BTEX)',                '2,180', '1,860'),
            ('Isocyanates',                                 '120',   '98'),
            ('Flame Retardants (Halogenated)',              '84',    '67'),
            ('Total SoC Leaving as Emissions',             '4,034', '3,345'),
        ])
    inline_note(doc,
        "No other main hazard classes of SoC outside those listed generated material emissions "
        "from facilities during FY 2025.",
        'ESRS E2-5-34')

    # 5.3.9.2
    heading(doc, '5.3.9.2  Quantify Amount of Substances of Concern Leaving Facilities '
            'as Products — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.2 — SoC: Amount Leaving Facilities as Products (FY 2024 vs FY 2025)',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                       '14,200', '11,800'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)', '3,100',  '2,640'),
            ('Aromatic Hydrocarbons (BTEX)',               '8,400',  '7,100'),
            ('Isocyanates',                                '2,900',  '2,580'),
            ('Flame Retardants (Halogenated)',             '1,200',  '960'),
            ('Total SoC Leaving as Products',             '29,800', '25,080'),
        ])
    inline_note(doc,
        "No other main hazard classes of SoC outside those listed left facilities as products "
        "during FY 2025.",
        'ESRS E2-5-34')

    # 5.3.9.3
    heading(doc, '5.3.9.3  Quantify Amount of Substances of Concern Leaving Facilities '
            'as Part of Products — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.3 — SoC: Amount Leaving Facilities as Part of Products (FY 2024 vs FY 2025)',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                       '12,800', '10,400'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)', '2,400',  '2,060'),
            ('Aromatic Hydrocarbons (BTEX)',               '7,200',  '6,300'),
            ('Isocyanates',                                '2,200',  '1,980'),
            ('Flame Retardants (Halogenated)',             '980',    '840'),
            ('Total SoC Leaving as Part of Products',     '25,580', '21,580'),
        ])
    inline_note(doc,
        "No other main hazard classes of SoC outside those listed left facilities as part of "
        "products during FY 2025.",
        'ESRS E2-5-34')

    # 5.3.9.4
    heading(doc, '5.3.9.4  Quantify Amount of Substances of Concern Leaving Facilities '
            'as Services — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.4 — SoC: Amount Leaving Facilities as Services (FY 2024 vs FY 2025)',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                       '1,860', '1,540'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)', '330',   '280'),
            ('Aromatic Hydrocarbons (BTEX)',               '1,420', '1,180'),
            ('Isocyanates',                                '460',   '390'),
            ('Flame Retardants (Halogenated)',             '178',   '143'),
            ('Total SoC Leaving via Services',            '4,248', '3,533'),
        ])
    inline_note(doc,
        "No other main hazard classes of SoC outside those listed left facilities via services "
        "during FY 2025.",
        'ESRS E2-5-34')

    # 5.3.9.5
    heading(doc, '5.3.9.5  Quantify Total Amount of Substances of Concern Generated or '
            'Used During Production or Procured — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.5 — SoC: Total Amount Generated / Used / Procured (FY 2024 vs FY 2025)',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                       '48,600', '42,300'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)', '8,240',  '7,180'),
            ('Aromatic Hydrocarbons (BTEX)',               '34,100', '29,700'),
            ('Isocyanates',                                '6,780',  '6,120'),
            ('Flame Retardants (Halogenated)',             '2,940',  '2,410'),
            ('Total SoC Generated / Used / Procured',     '100,660','87,710'),
        ])
    inline_note(doc,
        "No other main hazard classes of SoC outside those listed were generated, used, or "
        "procured during FY 2025.",
        'ESRS E2-5-34')

    # 5.3.9.6
    heading(doc, '5.3.9.6  Quantify Total Amount of Substances of Concern Leaving '
            'Facilities as Emissions, Products, or Part of Products or Services — '
            'By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.9.6 — SoC: Combined Total Leaving Facilities '
        '(Emissions + Products + Part of Products + Services) — FY 2025',
        'ESRS E2-5-34',
        soc_hdrs,
        [
            ('Chlorinated Solvents',                       '30,100', '24,720'),
            ('Hazardous Heavy Metals (Pb, Cd, Hg, Cr VI)', '6,240',  '5,320'),
            ('Aromatic Hydrocarbons (BTEX)',               '19,200', '16,440'),
            ('Isocyanates',                                '5,680',  '5,048'),
            ('Flame Retardants (Halogenated)',             '2,442',  '2,010'),
            ('Total SoC Leaving Facilities — Combined',   '63,662', '53,538'),
        ])
    inline_note(doc,
        "Figures above represent the aggregate of all SoC leaving facilities across the four "
        "pathways (emissions, products, part of products, and services). No other SoC categories "
        "outside those listed have been identified during FY 2025.",
        'ESRS E2-5-34')

    # ═══════════════════════════════════════════════════════════
    # 5.3.10  Substances of Very High Concern
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.10  Substances of Very High Concern', level=2)
    code_para(doc, 'ESRS E2-5 | ESRS E2-5-35')

    body(doc,
         "XYZ Logistics LLC separately tracks and discloses the amounts of substances of very "
         "high concern (SVHC) generated, procured, or used in production, and the amounts leaving "
         "its facilities as emissions, products, components of products, or services. All data in "
         "this section is reported in kilograms (kg) unless otherwise stated.")

    svhc_hdrs = ['Pollutant / Category', 'FY 2024 (kg)', 'FY 2025 (kg)']

    # 5.3.10.1
    heading(doc, '5.3.10.1  Quantify Amount of Substances of Very High Concern Leaving '
            'Facilities as Emissions — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.1 — SVHC: Amount Leaving Facilities as Emissions (FY 2024 vs FY 2025)',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)', '480', '362'),
            ('SVHC — PBT / vPvB Substances',         '142', '106'),
            ('SVHC — Endocrine Disruptors',           '68',  '49'),
            ('SVHC — Respiratory Sensitisers',        '31',  '23'),
            ('Total SVHC Leaving as Emissions',       '721', '540'),
        ])
    inline_note(doc,
        "No other main hazard classes of SVHC outside those listed generated material emissions "
        "from facilities during FY 2025.",
        'ESRS E2-5-35')

    # 5.3.10.2
    heading(doc, '5.3.10.2  Quantify Amount of Substances of Very High Concern Leaving '
            'Facilities as Products — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.2 — SVHC: Amount Leaving Facilities as Products (FY 2024 vs FY 2025)',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)', '4,200', '3,480'),
            ('SVHC — PBT / vPvB Substances',         '1,400', '1,140'),
            ('SVHC — Endocrine Disruptors',           '680',   '540'),
            ('SVHC — Respiratory Sensitisers',        '290',   '220'),
            ('Total SVHC Leaving as Products',        '6,570', '5,380'),
        ])
    inline_note(doc,
        "No other main hazard classes of SVHC outside those listed left facilities as products "
        "during FY 2025.",
        'ESRS E2-5-35')

    # 5.3.10.3
    heading(doc, '5.3.10.3  Quantify Amount of Substances of Very High Concern Leaving '
            'Facilities as Part of Products — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.3 — SVHC: Amount Leaving Facilities as Part of Products (FY 2024 vs FY 2025)',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)',     '3,600', '2,980'),
            ('SVHC — PBT / vPvB Substances',             '1,200', '960'),
            ('SVHC — Endocrine Disruptors',               '540',   '420'),
            ('SVHC — Respiratory Sensitisers',            '240',   '186'),
            ('Total SVHC Leaving as Part of Products',   '5,580', '4,546'),
        ])
    inline_note(doc,
        "No other main hazard classes of SVHC outside those listed left facilities as part of "
        "products during FY 2025.",
        'ESRS E2-5-35')

    # 5.3.10.4
    heading(doc, '5.3.10.4  Quantify Amount of Substances of Very High Concern Leaving '
            'Facilities as Services — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.4 — SVHC: Amount Leaving Facilities as Services (FY 2024 vs FY 2025)',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)', '480', '378'),
            ('SVHC — PBT / vPvB Substances',         '168', '162'),
            ('SVHC — Endocrine Disruptors',           '92',  '73'),
            ('SVHC — Respiratory Sensitisers',        '42',  '12'),
            ('Total SVHC Leaving via Services',       '782', '625'),
        ])
    inline_note(doc,
        "No other main hazard classes of SVHC outside those listed left facilities via services "
        "during FY 2025.",
        'ESRS E2-5-35')

    # 5.3.10.5
    heading(doc, '5.3.10.5  Quantify Total Amount of Substances of Very High Concern '
            'Generated or Used During Production or Procured — By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.5 — SVHC: Total Amount Generated / Used / Procured (FY 2024 vs FY 2025)',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)',    '14,200', '11,600'),
            ('SVHC — PBT / vPvB Substances',            '4,800',  '3,900'),
            ('SVHC — Endocrine Disruptors',              '2,400',  '1,980'),
            ('SVHC — Respiratory Sensitisers',           '1,100',  '860'),
            ('Total SVHC Generated / Used / Procured',  '22,500', '18,340'),
        ])
    inline_note(doc,
        "No other main hazard classes of SVHC outside those listed were generated, used, or "
        "procured during FY 2025.",
        'ESRS E2-5-35')

    # 5.3.10.6
    heading(doc, '5.3.10.6  Quantify Total Amount of Substances of Very High Concern '
            'Leaving Facilities as Emissions, Products, or Part of Products or Services — '
            'By Main Hazard Class', level=3)
    tbl(doc,
        'Table 5.3.10.6 — SVHC: Combined Total Leaving Facilities '
        '(Emissions + Products + Part of Products + Services) — FY 2025',
        'ESRS E2-5-35',
        svhc_hdrs,
        [
            ('SVHC — Carcinogenic (Category 1A/1B)',  '8,760', '7,200'),
            ('SVHC — PBT / vPvB Substances',          '2,910', '2,368'),
            ('SVHC — Endocrine Disruptors',            '1,380', '1,082'),
            ('SVHC — Respiratory Sensitisers',         '603',   '441'),
            ('Total SVHC Leaving Facilities — Combined','13,653','11,091'),
        ])
    inline_note(doc,
        "Figures above represent the aggregate of all SVHC leaving facilities across the four "
        "pathways. No additional SVHC categories outside those listed have been identified during "
        "FY 2025. Where substances meeting SVHC criteria are not covered by the standard REACH "
        "SVHC Candidate List — such as certain naturally occurring radioactive materials (NORM) "
        "arising from specific logistics operations — these are tracked separately under the "
        "Company's radiation protection programme and reported to relevant competent authorities.",
        'ESRS E2-5-35')

    # ═══════════════════════════════════════════════════════════
    # 5.3.11  Anticipated Financial Effects
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.11  Anticipated Financial Effects from Material Pollution-Related '
            'Risks and Opportunities', level=2)
    code_para(doc, 'ESRS E2-6 | ESRS E2-6-39 | ESRS E2-6-40 | ESRS E2-6-41 | '
              'ESRS E2-6-AR-31 | ESRS E2-6-AR-33')

    # 5.3.11.1
    heading(doc, '5.3.11.1  Quantify Anticipated Financial Effects — Material '
            'Pollution-Related Risks and Opportunities', level=3)
    tbl(doc,
        "Table 5.3.11.1 — Quantitative Anticipated Financial Effects: "
        "Pollution-Related Risks and Opportunities — FY 2025",
        'ESRS E2-6-40-(a)',
        ['Financial Effect Category', "FY 2025 Data (USD '000)"],
        [
            ('Regulatory Non-Compliance Risks (fines and penalties)',        '2,840'),
            ('Environmental Remediation and Clean-Up Cost Risks',           '14,200'),
            ('Stranded Asset Risks from Tightening Pollution Regulation',   '8,600'),
            ('Opportunities: Pollution Abatement Product/Service Revenues', '(6,400)'),
            ('Opportunities: Emission Reduction Technology Licensing Revenue','(1,800)'),
            ('Net Anticipated Financial Effect (Risks less Opportunities)',  '17,440'),
        ])

    # 5.3.11.2
    heading(doc, '5.3.11.2  Quantify Net Revenue from Products and Services Containing '
            'Substances of Concern or Very High Concern — FY 2025', level=3)
    tbl(doc,
        "Table 5.3.11.2 — Net Revenue from Products/Services Containing SoC or SVHC — FY 2025",
        'ESRS E2-6-40-(a)',
        ['Revenue Category', "Net Revenue (USD '000)", '% of Total Net Revenue'],
        [
            ('Products / Services Containing SoC',               '184,600', '12.4%'),
            ('Products / Services Containing SVHC',              '47,200',  '3.2%'),
            ('Products / Services Containing Both SoC and SVHC', '22,400',  '1.5%'),
            ('Total Revenue from SoC / SVHC Products / Services','254,200', '17.1%'),
        ])

    # 5.3.11.3
    heading(doc, '5.3.11.3  Quantify Operating and Capital Expenditure in Conjunction '
            'with Major Pollution Incidents and Deposits, and Provisions for Environmental '
            'Protection and Remediation Costs — FY 2025', level=3)
    tbl(doc,
        "Table 5.3.11.3 — OpEx, CapEx and Provisions: "
        "Major Incidents, Deposits and Remediation — FY 2025",
        'ESRS E2-6-40 | E2-6-AR-31-(b)-(c)',
        ['Expenditure / Provision Category', "FY 2025 Data (USD '000)"],
        [
            ('Operating Expenditure (OpEx) — Major Incidents and Deposits',    '3,840'),
            ('Capital Expenditure (CapEx) — Major Incidents and Deposits',     '7,200'),
            ('Provisions for Environmental Protection and Remediation Costs',  '14,200'),
            ('Total Pollution-Related Expenditure and Provisions',             '25,240'),
        ])

    # 5.3.11.4
    heading(doc, '5.3.11.4  Explain Qualitative Anticipated Financial Effects — Material '
            'Pollution-Related Risks and Opportunities', level=3)
    code_para(doc, 'ESRS E2-6-39-(a)')
    body(doc,
         "The most significant near-term financial exposure relates to the costs of environmental "
         "remediation at legacy contaminated sites, compliance capital expenditure required to "
         "align IED-regulated installations with upcoming BAT-AEL revisions, and the potential "
         "for increased provisions associated with SoC-related product liability. Upside financial "
         "opportunities relate principally to the Company's growing environmental services "
         "business lines, which provide pollution abatement technologies and advisory services "
         "to third parties, and to revenue streams from the licensing of proprietary emission "
         "reduction technologies.")

    # 5.3.11.5
    heading(doc, '5.3.11.5  Explain Effects Considered, Related Impacts and Time Horizons — '
            'Pollution-Related Financial Effects', level=3)
    code_para(doc, 'ESRS E2-6-39-(b)')
    body(doc,
         "The financial effects considered in the quantitative assessment relate primarily to: "
         "physical risks from pollution incidents (direct remediation costs and regulatory "
         "penalties); transition risks from tightening pollution regulation (compliance CapEx and "
         "stranded asset risk); and transition opportunities from environmental services revenue "
         "growth. Time horizons for materialisation are assessed as: near-term (1–3 years) for "
         "regulatory non-compliance risks given scheduled revision of BAT Conclusions applicable "
         "to two installations; medium-term (3–7 years) for remediation costs reflecting the "
         "staged nature of contaminated site programmes; and longer-term (7–15 years) for "
         "stranded asset risk contingent on the pace of regulatory tightening.")

    # 5.3.11.6
    heading(doc, '5.3.11.6  Disclose Critical Assumptions, Sources and Level of Uncertainty '
            'in Quantifying Financial Effects — Pollution', level=3)
    code_para(doc, 'ESRS E2-6-39-(c)')
    body(doc,
         "Critical assumptions underpinning the quantitative financial effect estimates include: "
         "a conservative regulatory penalty rate based on historical sector enforcement data; "
         "remediation cost estimates derived from site investigation reports prepared by "
         "independent environmental consultants; discount rates and inflation assumptions in "
         "accordance with the Company's accounting policies; and a revenue growth assumption for "
         "environmental services based on contracted pipeline and market analysis. Principal "
         "sources of uncertainty include the pace and scope of regulatory revision, the extent of "
         "contamination at sites not yet subject to full investigation, and market adoption rates "
         "for the Company's environmental technology products.")

    # 5.3.11.7
    heading(doc, '5.3.11.7  Disclose Material Incidents and Deposits with Negative Pollution '
            'Impacts on Financial Performance — FY 2025', level=3)
    code_para(doc, 'ESRS E2-6-41')
    body(doc,
         "During FY 2025, no material pollution incidents resulting in significant negative "
         "financial impacts on cash flows, financial position, or financial performance were "
         "recorded. The minor permit deviation at Site 4 (Singapore) was contained and remediated "
         "within the reporting period, with associated costs below the materiality threshold for "
         "separate disclosure. The Company's environmental provision balance at year-end FY 2025 "
         "of USD 14.2 million reflects the aggregate estimated cost of ongoing remediation "
         "programmes at all contaminated sites and is considered representative of the probable "
         "financial exposure from known legacy pollution matters.")

    # 5.3.11.8
    heading(doc, '5.3.11.8  Disclose Assessment of Related Products and Services at Risk — '
            'Time Horizon, Financial Estimation and Critical Assumptions', level=3)
    code_para(doc, 'ESRS E2-6-AR-33')
    body(doc,
         "Products and services assessed as being at risk from pollution-related transition "
         "factors include those incorporating SoC or SVHC subject to regulatory restriction "
         "timelines under REACH, and services involving the handling or transport of substances "
         "with evolving regulatory profiles. The time horizon for this assessment is five years "
         "(FY 2025 to FY 2030), consistent with the planning horizon applied across the Company's "
         "enterprise risk management framework. Financial amounts at risk are estimated using a "
         "product revenue exposure methodology, applying a phase-out probability weighting to the "
         "revenue attributable to individual at-risk product and service lines. Critical "
         "assumptions include current REACH regulatory timelines, prevailing substitution "
         "feasibility assessments, and contracted customer demand profiles.")
