#!/usr/bin/env python3
"""Part 2: Section 5.3.5 – Pollution of Air (all sub-sections)."""

import sys
sys.path.insert(0, '/home/user/jcode')
from build_report_helpers import *

def build_part2(doc):

    # ═══════════════════════════════════════════════════════════
    # 5.3.5  Pollution of Air
    # ═══════════════════════════════════════════════════════════
    heading(doc, '5.3.5  Pollution of Air', level=2)
    code_para(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
              'ESRS E2-4-30 | ESRS E2-4-31 | GRI 305-7')

    # 5.3.5.1 – Summary table
    heading(doc, '5.3.5.1  Quantify Emissions to Air by Pollutant — Summary', level=3)
    tbl(doc,
        'Table 5.3.5.1 — Emissions to Air by Pollutant (FY 2024 vs FY 2025)',
        'ESRS E2-4-28 | E2-4-AR-22-(a) | GRI 305-7',
        ['Pollutant / Category', 'FY 2024 (metric tonnes)', 'FY 2025 (metric tonnes)'],
        [
            ('Nitrogen Oxides (NOx)',             '1,245.8', '1,141.0'),
            ('Sulphur Oxides (SOx)',               '387.4',   '343.9'),
            ('Particulate Matter (PM10)',           '92.6',    '86.4'),
            ('Particulate Matter (PM2.5)',          '61.4',    '55.7'),
            ('Volatile Organic Compounds (VOC)',   '234.1',   '218.6'),
            ('Carbon Monoxide (CO)',               '156.3',   '148.2'),
            ('Ammonia (NH3)',                      '18.9',    '17.4'),
            ('Persistent Organic Pollutants (POP)','0.34',    '0.29'),
            ('Hazardous Air Pollutants (HAP)',     '12.7',    '11.3'),
            ('Total Emissions to Air',            '2,209.5', '2,023.0'),
        ])

    # 5.3.5.2 – By Sector
    heading(doc, '5.3.5.2  Quantify Emissions to Air by Pollutant — By Sector', level=3)
    tbl(doc,
        'Table 5.3.5.2 — Emissions to Air by Pollutant (By Sector) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10', 'Total (t)'],
        [
            ('NOx',   '127', '114', '98', '124', '87', '81', '72', '104', '95', '89',  '991'),
            ('SOx',   '38',  '34',  '28', '37',  '25', '24', '21', '31',  '28', '25',  '291'),
            ('PM10',  '9.8', '8.6', '7.4','9.2', '6.4','6.0','5.4','7.8', '7.0','6.3', '73.9'),
            ('VOC',   '24',  '21',  '18', '23',  '16', '15', '13', '19',  '17', '15',  '181'),
            ('Total', '198.8','177.6','151.4','193.2','134.4','126.0','111.4','161.8','147.0','135.3','1,536.9'),
        ])
    inline_note(doc,
        "Sector classifications correspond to the Company's internal sector coding framework. "
        "No additional sector categories outside those listed generated material air emissions "
        "during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.5.3 – By Geographical Area
    heading(doc, '5.3.5.3  Quantify Emissions to Air by Pollutant — By Geographical Area', level=3)
    tbl(doc,
        'Table 5.3.5.3 — Emissions to Air by Pollutant (By Geographical Area) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'UAE', 'India', 'Germany', 'France', 'Singapore', 'Japan', 'UK',
         'Brazil', 'Australia', 'Canada', 'Total (t)'],
        [
            ('NOx',   '189','168','84','72','114','89','68','102','96','87', '1,069'),
            ('SOx',   '58', '52', '25','21','34', '27','20','31', '29','26', '323'),
            ('PM10',  '14.2','12.8','6.1','5.2','8.4','6.7','4.9','7.6','7.1','6.4','79.4'),
            ('VOC',   '36', '32', '15','13','21', '17','12','19', '18','16', '199'),
            ('Total', '297.2','264.8','130.1','111.2','177.4','139.7','104.9','159.6','150.1','135.4','1,670.4'),
        ])
    inline_note(doc,
        "Geographical area classifications are based on the Company's primary operating "
        "geographies. No other geographical areas outside those listed generated material "
        "air emissions during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.5.4 – By Source
    heading(doc, '5.3.5.4  Quantify Emissions to Air by Pollutant — By Source', level=3)
    tbl(doc,
        'Table 5.3.5.4 — Emissions to Air by Pollutant (By Source) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Fugitive', 'Stack', 'Process', 'Combustion', 'Transport',
         'Storage', 'Cooling', 'Waste', 'Maintenance', 'Other', 'Total (t)'],
        [
            ('NOx',   '84', '312','198','174','126','82','54','47','38','26','1,141'),
            ('SOx',   '24', '98', '62', '54', '38', '24','16','14','10','4',  '344'),
            ('PM10',  '6.2','24.1','15.2','13.4','9.2','5.8','3.9','3.4','3.1','2.1','86.4'),
            ('VOC',   '34', '58', '42', '36', '18', '12','8', '6', '3', '2',  '219'),
            ('Total', '148.2','492.1','317.2','277.4','191.2','123.8','81.9','70.4','54.1','34.1','1,790.4'),
        ])
    inline_note(doc,
        "Source categories encompass all primary emission source types identified across "
        "IED-regulated and non-regulated installations. No other source categories outside "
        "those listed generated material air emissions during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.5.5 – By Site Location
    heading(doc, '5.3.5.5  Quantify Emissions to Air by Pollutant — By Site Location', level=3)
    tbl(doc,
        'Table 5.3.5.5 — Emissions to Air by Pollutant (By Site Location) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-22-(a)',
        ['Pollutant', 'Dubai', 'Mumbai', 'Frankfurt', 'Paris', 'Singapore',
         'Tokyo', 'London', 'São Paulo', 'Sydney', 'Toronto', 'Total (t)'],
        [
            ('NOx',   '184','162','82','71','112','87','66','101','94','82', '1,041'),
            ('SOx',   '56', '50', '24','20','33', '26','19','30', '28','25', '311'),
            ('PM10',  '13.8','12.4','5.9','5.0','8.1','6.5','4.7','7.3','6.9','6.2','76.8'),
            ('VOC',   '35', '31', '14','12','20', '16','11','18', '17','15', '189'),
            ('Total', '288.8','255.4','125.9','108.0','173.1','135.5','100.7','156.3','145.9','128.2','1,617.8'),
        ])
    inline_note(doc,
        "Site locations listed correspond to all primary operating facilities. No other site "
        "locations outside those listed generated material air emissions during FY 2025.",
        'ESRS E2-4-28 | E2-4-AR-22-(a)')

    # 5.3.5.6 – Microplastics
    heading(doc, '5.3.5.6  Quantify Microplastics Generated and Used — Air Context', level=3)
    tbl(doc,
        'Table 5.3.5.6 — Microplastics Generated and Used (Air-Related Operations) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-20-(b)',
        ['Microplastics KPI', 'FY 2025 Data'],
        [
            ('Microplastics Generated', '12,840 kg'),
            ('Microplastics Used',       '4.7 metric tonnes'),
        ])

    # 5.3.5.7 – Changes Over Time
    heading(doc, '5.3.5.7  Explain Description of Changes Over Time — Pollution of Air', level=3)
    code_para(doc, 'ESRS E2-4-30-(a)')
    body(doc,
         "Air emission performance improved across all pollutant categories during FY 2025 "
         "relative to FY 2024. NOx emissions declined by 8.4% year-on-year to 1,141.0 metric "
         "tonnes; SOx declined by 11.2% to 343.9 metric tonnes; and PM10 declined by 6.7% to "
         "86.4 metric tonnes. VOC emissions declined by 6.6% to 218.6 metric tonnes. These "
         "improvements reflect the progressive deployment of abatement technology upgrades "
         "initiated in FY 2023 and the full-year effect of process optimisation measures at "
         "Sites 1 and 3. All air emission trends are consistent with the Company's FY 2030 "
         "reduction targets.")

    # 5.3.5.8 – Measurement Methodologies
    heading(doc, '5.3.5.8  Explain Description of Measurement Methodologies — '
            'Pollution of Air', level=3)
    code_para(doc, 'ESRS E2-4-30-(b)')
    body(doc,
         "Air emissions are quantified using a combination of direct continuous emission "
         "monitoring systems (CEMS), periodic stack testing, mass balance methodologies, and "
         "emission factor-based calculations, applied in accordance with the hierarchy of "
         "measurement approaches specified in applicable EU BAT Conclusions and IED permit "
         "conditions. Measurement methodologies are reviewed annually to ensure alignment with "
         "evolving regulatory and scientific standards. All IED-regulated installations use CEMS "
         "as the primary measurement approach. Non-IED facilities apply mass balance and emission "
         "factor methodologies appropriate to the emission source type and substance involved.")

    # 5.3.5.9 – Data Collection Processes
    heading(doc, '5.3.5.9  Explain Description of Data Collection Processes for '
            'Air Pollution Accounting and Reporting', level=3)
    code_para(doc, 'ESRS E2-4-30-(c)')
    body(doc,
         "Air pollution data collection is managed through the Company's centralised "
         "Environmental Management Information System (EMIS), which aggregates site-level "
         "monitoring data, CEMS outputs, meter readings, process data inputs, and laboratory "
         "analysis results on a continuous basis. Data quality controls include automated "
         "range-check validation, manual verification by site environmental officers, and annual "
         "third-party data assurance reviews. Where measurement gaps occur due to equipment "
         "downtime or data quality failures, conservative estimation procedures are applied in "
         "accordance with documented data gap-filling protocols.")

    # 5.3.5.10 – Inferior Methodology
    heading(doc, '5.3.5.10  Disclose Inferior Methodology for Quantifying Air Emissions — '
            'Applicability and Disclosure', level=3)
    code_para(doc, 'ESRS E2-4-31-(c)')
    bool_table(doc,
               'State whether an inferior methodology compared to direct measurement of '
               'emissions was chosen to quantify air emissions',
               'No')
    body(doc,
         "No inferior methodology compared to direct measurement was chosen to quantify air "
         "emissions during FY 2025. All emission quantification is based on direct measurement "
         "via CEMS or recognised mass balance and emission factor calculation standards as "
         "described in Section 5.3.5.8. In the absence of any deviation from primary "
         "methodologies, no disclosure of reasons for choosing an inferior methodology is "
         "required for this reporting period.")
