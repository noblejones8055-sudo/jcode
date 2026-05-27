#!/usr/bin/env python3
"""Part 3: Sections 5.3.6 (Water) and 5.3.7 (Soil)."""

import sys
sys.path.insert(0, '/home/user/jcode')
from build_report_helpers import *

def build_part3(doc):

    # ═══════════════════════════════════════════════════════════
    # 5.3.6  Pollution of Water
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.6  Pollution of Water', level=2)
    code_para(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
              'ESRS E2-4-AR-23-(c) | ESRS E2-4-30 | ESRS E2-4-31')

    # 5.3.6.1 – Summary
    heading(doc, '5.3.6.1  Quantify Emissions to Water by Pollutant — Summary', level=3)
    tbl(doc,
        'Table 5.3.6.1 — Emissions to Water by Pollutant (FY 2024 vs FY 2025)',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant / Category', 'FY 2024 (metric tonnes)', 'FY 2025 (metric tonnes)'],
        [
            ('Total Suspended Solids (TSS)',       '48.3',  '43.9'),
            ('Chemical Oxygen Demand (COD)',       '234.6', '237.7'),
            ('Biochemical Oxygen Demand (BOD)',    '78.4',  '72.1'),
            ('Nitrates (NO3-N)',                   '14.2',  '12.8'),
            ('Phosphorus (Total-P)',               '3.6',   '3.1'),
            ('Heavy Metals (total)',               '0.87',  '0.74'),
            ('Hydrocarbons (TPH)',                 '2.14',  '1.89'),
            ('Persistent Organic Pollutants (POP)','0.012', '0.009'),
            ('Total Emissions to Water',           '382.2', '372.2'),
        ])

    # 5.3.6.2 – By Sector
    heading(doc, '5.3.6.2  Quantify Emissions to Water by Pollutant — By Sector', level=3)
    tbl(doc,
        'Table 5.3.6.2 — Emissions to Water by Pollutant (By Sector) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10', 'Total (t)'],
        [
            ('TSS',   '4.8', '3.9','3.1','4.2','2.8','2.4','2.1','3.4','3.0','2.7','32.4'),
            ('COD',   '26.1','21.2','16.4','22.4','14.8','12.7','11.2','17.8','15.9','14.3','172.8'),
            ('BOD',   '7.9', '6.4','5.0','6.8','4.5','3.9','3.4','5.4','4.8','4.3','52.4'),
            ('Total', '38.8','31.5','24.5','33.4','22.1','19.0','16.7','26.6','23.7','21.3','257.6'),
        ])
    inline_note(doc,
        "No other sector categories outside those listed generated material water emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.6.3 – By Geographical Area
    heading(doc, '5.3.6.3  Quantify Emissions to Water by Pollutant — By Geographical Area', level=3)
    tbl(doc,
        'Table 5.3.6.3 — Emissions to Water by Pollutant (By Geographical Area) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'UAE', 'India', 'Germany', 'France', 'Singapore', 'Japan', 'UK',
         'Brazil', 'Australia', 'Canada', 'Total (t)'],
        [
            ('TSS',   '5.8','6.8','2.4','2.0','3.4','2.2','1.6','2.9','2.6','2.2','31.9'),
            ('COD',   '31.2','36.4','12.8','10.6','18.2','11.8','8.6','15.4','13.8','11.8','170.6'),
            ('BOD',   '9.4','11.2','3.9','3.2','5.5','3.6','2.6','4.7','4.2','3.6','51.9'),
            ('Total', '46.4','54.4','19.1','15.8','27.1','17.6','12.8','23.0','20.6','17.6','254.4'),
        ])
    inline_note(doc,
        "No other geographical areas outside those listed generated material water emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.6.4 – By Source
    heading(doc, '5.3.6.4  Quantify Emissions to Water by Pollutant — By Source', level=3)
    tbl(doc,
        'Table 5.3.6.4 — Emissions to Water by Pollutant (By Source) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Direct', 'Runoff', 'Cooling', 'Process', 'Sewage',
         'Leachate', 'Atm. Dep.', 'Spill', 'Sludge', 'Other', 'Total (t)'],
        [
            ('TSS',   '8.2','7.4','4.8','6.4','4.2','3.8','2.6','2.1','2.0','1.4','42.9'),
            ('COD',   '44.2','39.8','26.4','34.2','22.8','20.4','14.2','11.4','10.8','7.4','231.6'),
            ('BOD',   '13.4','12.2','8.0','10.4','6.8','6.2','4.2','3.4','3.2','2.2','70.0'),
            ('Total', '65.8','59.4','39.2','51.0','33.8','30.4','21.0','16.9','16.0','11.0','344.5'),
        ])
    inline_note(doc,
        "No other source categories outside those listed generated material water emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.6.5 – By Site Location
    heading(doc, '5.3.6.5  Quantify Emissions to Water by Pollutant — By Site Location', level=3)
    tbl(doc,
        'Table 5.3.6.5 — Emissions to Water by Pollutant (By Site Location) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Dubai', 'Mumbai', 'Frankfurt', 'Paris', 'Singapore',
         'Tokyo', 'London', 'São Paulo', 'Sydney', 'Toronto', 'Total (t)'],
        [
            ('TSS',   '5.6','6.6','2.3','1.9','3.2','2.1','1.5','2.8','2.5','2.1','30.6'),
            ('COD',   '30.4','35.2','12.4','10.2','17.6','11.4','8.2','14.8','13.4','11.4','165.0'),
            ('BOD',   '9.2','10.8','3.8','3.1','5.3','3.5','2.5','4.5','4.1','3.5','50.3'),
            ('Total', '45.2','52.6','18.5','15.2','26.1','17.0','12.2','22.1','20.0','17.0','245.9'),
        ])
    inline_note(doc,
        "No other site locations outside those listed generated material water emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.6.6 – Microplastics
    heading(doc, '5.3.6.6  Quantify Microplastics Generated and Used — Water Context', level=3)
    tbl(doc,
        'Table 5.3.6.6 — Microplastics Generated and Used (Water-Related Operations) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-20-(b)',
        ['Microplastics KPI', 'FY 2025 Data'],
        [
            ('Microplastics Generated', '12,840 kg'),
            ('Microplastics Used',       '4.7 metric tonnes'),
        ])

    # 5.3.6.7 – Water Stress
    heading(doc, '5.3.6.7  Quantify Emissions of Pollutants to Water in Areas '
            'with Water Stress — FY 2025', level=3)
    tbl(doc,
        'Table 5.3.6.7 — Emissions of Pollutants to Water by Area with Water Stress — FY 2025',
        'ESRS E2-4-AR-23-(c)',
        ['Area Classification', 'Emissions to Water — Absolute (metric tonnes)',
         '% of Total Emissions'],
        [
            ('Areas at Water Risk',       '47.8',  '12.8%'),
            ('Areas of High Water Stress','89.4',  '24.0%'),
            ('Total Emissions to Water',  '372.2', '100.0%'),
        ])

    # 5.3.6.8 – Changes Over Time
    heading(doc, '5.3.6.8  Explain Description of Changes Over Time — Pollution of Water', level=3)
    code_para(doc, 'ESRS E2-4-30-(a)')
    body(doc,
         "Total emissions to water declined by 2.6% year-on-year to 372.2 metric tonnes in "
         "FY 2025. TSS declined by 9.1% to 43.9 metric tonnes and BOD declined by 8.0% to "
         "72.1 metric tonnes, reflecting the full-year effect of upgraded biological treatment "
         "systems installed at Sites 2 and 4 in late FY 2024. COD increased marginally by 1.3% "
         "to 237.7 metric tonnes due to a temporary increase in production throughput at two "
         "facilities during H1 FY 2025. Nitrate and phosphorus loadings continued to decline in "
         "line with effluent management improvement programmes. No uncontrolled releases or "
         "spills to water were recorded during FY 2025.")

    # 5.3.6.9 – Measurement Methodologies
    heading(doc, '5.3.6.9  Explain Description of Measurement Methodologies — '
            'Pollution of Water', level=3)
    code_para(doc, 'ESRS E2-4-30-(b)')
    body(doc,
         "Water emission quantification uses a combination of continuous on-line effluent "
         "monitoring, composite sample analysis by accredited environmental laboratories, and "
         "mass balance methodologies for cooling water and storm water discharges. All "
         "IED-regulated installations with direct water discharges are subject to effluent "
         "monitoring in accordance with permit conditions. Monitoring frequencies range from "
         "continuous (for CEMS-equipped outfalls) to monthly composite sampling for non-CEMS "
         "discharge points.")

    # 5.3.6.10 – Data Collection Processes
    heading(doc, '5.3.6.10  Explain Description of Data Collection Processes for '
            'Water Pollution Accounting and Reporting', level=3)
    code_para(doc, 'ESRS E2-4-30-(c)')
    body(doc,
         "Water pollution data is collected via the Company's EMIS, integrating on-line "
         "monitoring data, laboratory analysis certificates, and process data inputs from "
         "site-level environmental officers. Data quality controls include laboratory "
         "accreditation requirements (ISO 17025), automated validation checks for on-line "
         "monitoring data, and annual third-party data assurance. Where sampling gaps arise due "
         "to equipment maintenance, conservative estimation is applied in accordance with "
         "documented data gap-filling protocols.")

    # 5.3.6.11 – Inferior Methodology
    heading(doc, '5.3.6.11  Disclose Inferior Methodology for Quantifying Water Emissions — '
            'Applicability and Disclosure', level=3)
    code_para(doc, 'ESRS E2-4-31-(c)')
    bool_table(doc,
               'State whether an inferior methodology compared to direct measurement was '
               'chosen to quantify water emissions',
               'No')
    body(doc,
         "No inferior methodology compared to direct measurement was chosen to quantify water "
         "emissions during FY 2025. All water emission quantification is based on direct effluent "
         "monitoring or recognised mass balance and calculation standards. No disclosure of "
         "reasons for choosing an inferior methodology is required for this reporting period.")

    # ═══════════════════════════════════════════════════════════
    # 5.3.7  Pollution of Soil
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.7  Pollution of Soil', level=2)
    code_para(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
              'ESRS E2-4-AR-23-(c) | ESRS E2-4-30 | ESRS E2-4-31')

    # 5.3.7.1 – Summary
    heading(doc, '5.3.7.1  Quantify Emissions to Soil by Pollutant — Summary', level=3)
    tbl(doc,
        'Table 5.3.7.1 — Emissions to Soil by Pollutant (FY 2024 vs FY 2025)',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant / Category', 'FY 2024 (metric tonnes)', 'FY 2025 (metric tonnes)'],
        [
            ('Heavy Metals',          '6.9',  '6.5'),
            ('Hydrocarbons (TPH)',    '24.9', '23.7'),
            ('Pesticides/Biocides',  '1.8',  '1.7'),
            ('Nitrates — Diffuse',   '11.0', '10.3'),
            ('Total Emissions to Soil','44.6', '42.2'),
        ])

    # 5.3.7.2 – By Sector
    heading(doc, '5.3.7.2  Quantify Emissions to Soil by Pollutant — By Sector', level=3)
    tbl(doc,
        'Table 5.3.7.2 — Emissions to Soil by Pollutant (By Sector) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10', 'Total (t)'],
        [
            ('Heavy Metals',   '1.2','0.8','0.6','0.9','0.5','0.4','0.3','0.5','0.4','0.3','5.9'),
            ('Hydrocarbons',   '3.4','2.8','2.1','2.9','1.8','1.6','1.4','2.2','1.9','1.6','21.7'),
            ('Pesticides',     '0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','2.0'),
            ('Nitrates—Diffuse','1.4','1.0','0.9','1.2','0.8','1.0','0.9','1.3','1.0','1.5','11.0'),
            ('Total',          '6.2','4.8','3.8','5.2','3.3','3.2','2.8','4.2','3.5','3.6','40.6'),
        ])
    inline_note(doc,
        "No other sector categories outside those listed generated material soil emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.7.3 – By Geographical Area
    heading(doc, '5.3.7.3  Quantify Emissions to Soil by Pollutant — By Geographical Area', level=3)
    tbl(doc,
        'Table 5.3.7.3 — Emissions to Soil by Pollutant (By Geographical Area) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'UAE', 'India', 'Germany', 'France', 'Singapore', 'Japan', 'UK',
         'Brazil', 'Australia', 'Canada', 'Total (t)'],
        [
            ('Heavy Metals',    '1.8','2.1','0.4','0.3','0.6','0.3','0.2','0.4','0.4','0.3','6.8'),
            ('Hydrocarbons',    '4.8','5.6','2.1','1.8','3.0','1.9','1.4','1.7','1.8','1.6','25.7'),
            ('Pesticides',      '0.1','0.4','0.2','0.1','0.2','0.1','0.1','0.2','0.2','0.2','1.8'),
            ('Nitrates—Diffuse','1.0','2.6','1.2','0.9','0.9','0.8','0.7','1.0','0.9','1.0','11.0'),
            ('Total',           '7.7','10.7','3.9','3.1','4.7','3.1','2.4','3.3','3.3','3.1','45.3'),
        ])
    inline_note(doc,
        "No other geographical areas outside those listed generated material soil emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.7.4 – By Source
    heading(doc, '5.3.7.4  Quantify Emissions to Soil by Pollutant — By Source', level=3)
    tbl(doc,
        'Table 5.3.7.4 — Emissions to Soil by Pollutant (By Source) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Fugitive', 'Spill', 'Diffuse', 'Leachate', 'Runoff',
         'Process', 'Atm. Dep.', 'Waste', 'Sludge', 'Remediation', 'Total (t)'],
        [
            ('Heavy Metals',    '0.8','1.4','0.5','0.6','0.4','0.6','0.3','0.5','0.5','0.3','5.9'),
            ('Hydrocarbons',    '3.2','6.4','3.8','2.2','2.6','2.9','1.2','2.4','2.2','1.4','28.3'),
            ('Pesticides',      '0.1','0.2','0.6','0.1','0.2','0.1','0.1','0.1','0.1','0.1','1.7'),
            ('Nitrates—Diffuse','0.8','0.7','4.2','0.8','1.4','0.7','0.4','0.8','0.8','0.4','11.0'),
            ('Total',           '4.9','8.7','9.1','3.7','4.6','4.3','2.0','3.8','3.6','2.2','46.9'),
        ])
    inline_note(doc,
        "No other source categories outside those listed generated material soil emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.7.5 – By Site Location
    heading(doc, '5.3.7.5  Quantify Emissions to Soil by Pollutant — By Site Location', level=3)
    tbl(doc,
        'Table 5.3.7.5 — Emissions to Soil by Pollutant (By Site Location) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Dubai', 'Mumbai', 'Frankfurt', 'Paris', 'Singapore',
         'Tokyo', 'London', 'São Paulo', 'Sydney', 'Toronto', 'Total (t)'],
        [
            ('Heavy Metals',    '1.8','2.0','0.4','0.2','0.5','0.3','0.2','0.4','0.4','0.3','6.5'),
            ('Hydrocarbons',    '4.7','5.5','2.1','1.7','2.9','1.8','1.3','1.7','1.7','1.5','24.9'),
            ('Pesticides',      '0.1','0.4','0.2','0.1','0.2','0.1','0.1','0.2','0.2','0.2','1.8'),
            ('Nitrates—Diffuse','1.0','2.5','1.2','0.9','0.9','0.7','0.6','0.9','0.9','0.7','10.3'),
            ('Total',           '7.6','10.4','3.9','2.9','4.5','2.9','2.2','3.2','3.2','2.7','37.5'),
        ])
    inline_note(doc,
        "No other site locations outside those listed generated material soil emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.7.6 – Microplastics
    heading(doc, '5.3.7.6  Quantify Microplastics Generated and Used — Soil Context', level=3)
    tbl(doc,
        'Table 5.3.7.6 — Microplastics Generated and Used (Soil-Related Operations) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-20-(b)',
        ['Microplastics KPI', 'FY 2025 Data'],
        [
            ('Microplastics Generated', '12,840 kg'),
            ('Microplastics Used',       '4.7 metric tonnes'),
        ])

    # 5.3.7.7 – Soil Water Stress
    heading(doc, '5.3.7.7  Quantify Emissions of Pollutants to Soil in Areas '
            'with Water Stress — FY 2025', level=3)
    tbl(doc,
        'Table 5.3.7.7 — Emissions of Pollutants to Soil by Area with Water Stress — FY 2025',
        'ESRS E2-4-AR-23-(c)',
        ['Area Classification', 'Emissions to Soil — Absolute (metric tonnes)',
         '% of Total Emissions'],
        [
            ('Areas at Water Risk',      '4.2',  '9.5%'),
            ('Areas of High Water Stress','9.7',  '21.9%'),
            ('Total Emissions to Soil',  '42.2', '100.0%'),
        ])

    # 5.3.7.8 – Changes Over Time
    heading(doc, '5.3.7.8  Explain Description of Changes Over Time — Pollution of Soil', level=3)
    code_para(doc, 'ESRS E2-4-30-(a)')
    body(doc,
         "Total soil emissions declined by 5.4% year-on-year to 42.2 metric tonnes in FY 2025. "
         "Heavy metal loadings to soil declined by 5.8% and hydrocarbon (TPH) loadings declined "
         "by 4.8%, reflecting improved secondary containment measures and spill prevention "
         "controls implemented across Sites 2 and 8. Nitrate diffuse loadings remained broadly "
         "stable. No new uncontrolled releases to soil were recorded during FY 2025. Ongoing "
         "remediation at three legacy contaminated sites progressed in line with approved site "
         "remediation plans.")

    # 5.3.7.9 – Measurement Methodologies
    heading(doc, '5.3.7.9  Explain Description of Measurement Methodologies — '
            'Pollution of Soil', level=3)
    code_para(doc, 'ESRS E2-4-30-(b)')
    body(doc,
         "Soil pollution quantification applies a combination of direct measurement (site "
         "investigation sampling and analysis for legacy contamination assessment), emission "
         "factor calculations (for diffuse agricultural nitrate loading estimates), mass balance "
         "methodologies (for secondary containment and spill quantification), and incident-based "
         "quantification where applicable. Site investigation sampling follows accredited "
         "protocols aligned with ISO 18400 series standards. Measurement approaches are reviewed "
         "annually in conjunction with site environmental management plan updates.")

    # 5.3.7.10 – Data Collection Processes
    heading(doc, '5.3.7.10  Explain Description of Data Collection Processes for '
            'Soil Pollution Accounting and Reporting', level=3)
    code_para(doc, 'ESRS E2-4-30-(c)')
    body(doc,
         "Soil pollution data is collected via the Company's EMIS, integrating site investigation "
         "reports, incident records, process data inputs, and regulatory submission data. Data "
         "quality controls include peer review of site investigation data by Senior Environmental "
         "Managers, cross-validation against regulatory permit data, and annual third-party data "
         "assurance. Where data gaps arise, conservative estimation is applied in accordance with "
         "documented protocols, and such estimates are clearly flagged in the EMIS audit trail.")

    # 5.3.7.11 – Inferior Methodology
    heading(doc, '5.3.7.11  Disclose Inferior Methodology for Quantifying Soil Emissions — '
            'Applicability and Disclosure', level=3)
    code_para(doc, 'ESRS E2-4-31-(c)')
    bool_table(doc,
               'State whether an inferior methodology compared to direct measurement was '
               'chosen to quantify soil emissions',
               'Partial — Emission factor calculations applied for diffuse nitrate loading '
               'estimates only; all other soil emission categories measured directly')
    body(doc,
         "No inferior methodology compared to direct measurement was chosen to quantify soil "
         "emissions during FY 2025, with the exception of diffuse nitrate loading estimates "
         "where emission factor calculations are applied in the absence of direct measurement "
         "feasibility for diffuse agricultural sources. This approach is consistent with accepted "
         "scientific and regulatory practice for diffuse pollution quantification, and does not "
         "constitute an inferior methodology in the context of the available measurement options "
         "for this emission type.")
