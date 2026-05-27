#!/usr/bin/env python3
"""Sections 5.3.5 (Air), 5.3.6 (Water), 5.3.7 (Soil)."""

import sys
sys.path.insert(0, '/home/user/jcode')
from rb_helpers import heading, body, section_codes, note, data_table, status_table

AR22 = 'ESRS E2-4-28 | E2-4-AR-22-(a)'


def _emissions_block(doc, num, medium, codes_summary, summary_rows,
                     sector_rows, geo_rows, source_rows, site_rows,
                     micro_label, stress_rows=None,
                     changes=None, methods=None, datacol=None, inferior=None):
    """Generic builder shared by Air / Water / Soil sections."""
    pol = {'Air': 'air', 'Water': 'water', 'Soil': 'soil'}[medium]

    # .1 Summary
    heading(doc, f'{num}.1  Quantify Emissions to {medium} by Pollutant — Summary', 3)
    data_table(doc,
        f'Table {num}.1 — Emissions to {medium} by Pollutant (FY 2024 vs FY 2025)',
        codes_summary,
        ['Pollutant / Category', 'FY 2024 (metric tonnes)', 'FY 2025 (metric tonnes)'],
        summary_rows)

    # .2 By Sector
    heading(doc, f'{num}.2  Quantify Emissions to {medium} by Pollutant — By Sector', 3)
    data_table(doc,
        f'Table {num}.2 — Emissions to {medium} by Pollutant (By Sector) — FY 2025',
        AR22,
        ['Pollutant', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10', 'Total (t)'],
        sector_rows)
    note(doc, f"No other sector categories outside those listed generated material {pol} "
              f"emissions during FY 2025.", code=AR22)

    # .3 By Geographical Area
    heading(doc, f'{num}.3  Quantify Emissions to {medium} by Pollutant — By Geographical Area', 3)
    data_table(doc,
        f'Table {num}.3 — Emissions to {medium} by Pollutant (By Geographical Area) — FY 2025',
        AR22,
        ['Pollutant', 'UAE', 'India', 'Germany', 'France', 'Singapore', 'Japan', 'UK',
         'Brazil', 'Australia', 'Canada', 'Total (t)'],
        geo_rows)
    note(doc, f"No other geographical areas outside those listed generated material {pol} "
              f"emissions during FY 2025.", code=AR22)

    # .4 By Source
    heading(doc, f'{num}.4  Quantify Emissions to {medium} by Pollutant — By Source', 3)
    data_table(doc,
        f'Table {num}.4 — Emissions to {medium} by Pollutant (By Source) — FY 2025',
        AR22,
        source_rows[0], source_rows[1])
    note(doc, f"No other source categories outside those listed generated material {pol} "
              f"emissions during FY 2025.", code=AR22)

    # .5 By Site Location
    heading(doc, f'{num}.5  Quantify Emissions to {medium} by Pollutant — By Site Location', 3)
    data_table(doc,
        f'Table {num}.5 — Emissions to {medium} by Pollutant (By Site Location) — FY 2025',
        AR22,
        ['Pollutant', 'Dubai', 'Mumbai', 'Frankfurt', 'Paris', 'Singapore',
         'Tokyo', 'London', 'São Paulo', 'Sydney', 'Toronto', 'Total (t)'],
        site_rows)
    note(doc, f"No other site locations outside those listed generated material {pol} "
              f"emissions during FY 2025.", code=AR22)

    # .6 Microplastics
    heading(doc, f'{num}.6  Quantify Microplastics Generated and Used — {medium} Context', 3)
    data_table(doc,
        f'Table {num}.6 — Microplastics Generated and Used ({micro_label}) — FY 2025',
        'ESRS E2-4-28 | E2-4-AR-20-(b)',
        ['Microplastics KPI', 'FY 2025 Data'],
        [('Microplastics Generated', '12,840 kg'),
         ('Microplastics Used', '4.7 metric tonnes')])

    sub = 7
    # optional water-stress table (Water & Soil)
    if stress_rows is not None:
        heading(doc, f'{num}.{sub}  Quantify Emissions of Pollutants to {medium} in Areas '
                f'with Water Stress — FY 2025', 3)
        data_table(doc,
            f'Table {num}.{sub} — Emissions of Pollutants to {medium} by Area with Water '
            f'Stress — FY 2025',
            'ESRS E2-4-AR-23-(c)',
            ['Area Classification', f'Emissions to {medium} — Absolute (metric tonnes)',
             '% of Total Emissions'],
            stress_rows)
        sub += 1

    # Changes over time
    heading(doc, f'{num}.{sub}  Explain Description of Changes Over Time — '
            f'Pollution of {medium}', 3)
    body(doc, changes, code='ESRS E2-4-30-(a)')
    sub += 1

    # Measurement methodologies
    heading(doc, f'{num}.{sub}  Explain Description of Measurement Methodologies — '
            f'Pollution of {medium}', 3)
    body(doc, methods, code='ESRS E2-4-30-(b)')
    sub += 1

    # Data collection
    heading(doc, f'{num}.{sub}  Explain Description of Data Collection Processes for '
            f'{medium} Pollution Accounting and Reporting', 3)
    body(doc, datacol, code='ESRS E2-4-30-(c)')
    sub += 1

    # Inferior methodology (boolean + narrative)
    heading(doc, f'{num}.{sub}  Disclose Inferior Methodology for Quantifying '
            f'{medium} Emissions — Applicability and Disclosure', 3)
    status_table(doc, [
        (f'Inferior methodology compared to direct measurement chosen to quantify '
         f'{pol} emissions', 'ESRS E2-4-31-(c)', inferior[0]),
    ])
    body(doc, inferior[1], code='ESRS E2-4-31-(c)')


def build(doc):
    # ── 5.3.5  AIR ──────────────────────────────────────────
    heading(doc, '5.3.5  Pollution of Air', 2)
    section_codes(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
                  'ESRS E2-4-30 | ESRS E2-4-31 | GRI 305-7')
    _emissions_block(doc, '5.3.5', 'Air',
        'ESRS E2-4-28 | E2-4-AR-22-(a) | GRI 305-7',
        [
            ('Nitrogen Oxides (NOx)', '1,245.8', '1,141.0'),
            ('Sulphur Oxides (SOx)', '387.4', '343.9'),
            ('Particulate Matter (PM10)', '92.6', '86.4'),
            ('Particulate Matter (PM2.5)', '61.4', '55.7'),
            ('Volatile Organic Compounds (VOC)', '234.1', '218.6'),
            ('Carbon Monoxide (CO)', '156.3', '148.2'),
            ('Ammonia (NH3)', '18.9', '17.4'),
            ('Persistent Organic Pollutants (POP)', '0.34', '0.29'),
            ('Hazardous Air Pollutants (HAP)', '12.7', '11.3'),
            ('Total Emissions to Air', '2,209.5', '2,023.0'),
        ],
        [   # sector
            ('NOx', '127','114','98','124','87','81','72','104','95','89','991'),
            ('SOx', '38','34','28','37','25','24','21','31','28','25','291'),
            ('PM10','9.8','8.6','7.4','9.2','6.4','6.0','5.4','7.8','7.0','6.3','73.9'),
            ('VOC','24','21','18','23','16','15','13','19','17','15','181'),
            ('Total','198.8','177.6','151.4','193.2','134.4','126.0','111.4','161.8','147.0','135.3','1,536.9'),
        ],
        [   # geo
            ('NOx','189','168','84','72','114','89','68','102','96','87','1,069'),
            ('SOx','58','52','25','21','34','27','20','31','29','26','323'),
            ('PM10','14.2','12.8','6.1','5.2','8.4','6.7','4.9','7.6','7.1','6.4','79.4'),
            ('VOC','36','32','15','13','21','17','12','19','18','16','199'),
            ('Total','297.2','264.8','130.1','111.2','177.4','139.7','104.9','159.6','150.1','135.4','1,670.4'),
        ],
        (   # source (custom headers)
            ['Pollutant','Fugitive','Stack','Process','Combustion','Transport','Storage',
             'Cooling','Waste','Maintenance','Other','Total (t)'],
            [
                ('NOx','84','312','198','174','126','82','54','47','38','26','1,141'),
                ('SOx','24','98','62','54','38','24','16','14','10','4','344'),
                ('PM10','6.2','24.1','15.2','13.4','9.2','5.8','3.9','3.4','3.1','2.1','86.4'),
                ('VOC','34','58','42','36','18','12','8','6','3','2','219'),
                ('Total','148.2','492.1','317.2','277.4','191.2','123.8','81.9','70.4','54.1','34.1','1,790.4'),
            ]),
        [   # site
            ('NOx','184','162','82','71','112','87','66','101','94','82','1,041'),
            ('SOx','56','50','24','20','33','26','19','30','28','25','311'),
            ('PM10','13.8','12.4','5.9','5.0','8.1','6.5','4.7','7.3','6.9','6.2','76.8'),
            ('VOC','35','31','14','12','20','16','11','18','17','15','189'),
            ('Total','288.8','255.4','125.9','108.0','173.1','135.5','100.7','156.3','145.9','128.2','1,617.8'),
        ],
        'Air-Related Operations',
        stress_rows=None,
        changes="Air emission performance improved across all pollutant categories during FY 2025 "
                "relative to FY 2024. NOx emissions declined by 8.4% year-on-year to 1,141.0 "
                "metric tonnes; SOx declined by 11.2% to 343.9 metric tonnes; and PM10 declined by "
                "6.7% to 86.4 metric tonnes. VOC emissions declined by 6.6% to 218.6 metric tonnes. "
                "These improvements reflect the progressive deployment of abatement technology "
                "upgrades initiated in FY 2023 and the full-year effect of process optimisation "
                "measures at Sites 1 and 3. All air emission trends are consistent with the "
                "Company's FY 2030 reduction targets.",
        methods="Air emissions are quantified using a combination of direct continuous emission "
                "monitoring systems (CEMS), periodic stack testing, mass balance methodologies, and "
                "emission factor-based calculations, applied in accordance with the hierarchy of "
                "measurement approaches specified in applicable EU BAT Conclusions and IED permit "
                "conditions. All IED-regulated installations use CEMS as the primary measurement "
                "approach. Non-IED facilities apply mass balance and emission factor methodologies "
                "appropriate to the emission source type and substance involved.",
        datacol="Air pollution data collection is managed through the Company's centralised "
                "Environmental Management Information System (EMIS), which aggregates site-level "
                "monitoring data, CEMS outputs, meter readings, process data inputs, and laboratory "
                "analysis results on a continuous basis. Data quality controls include automated "
                "range-check validation, manual verification by site environmental officers, and "
                "annual third-party data assurance reviews. Where measurement gaps occur, "
                "conservative estimation procedures are applied in accordance with documented data "
                "gap-filling protocols.",
        inferior=('No. All air emission quantification is based on direct measurement via CEMS or '
                  'recognised mass balance and emission factor calculation standards.',
                  "No inferior methodology compared to direct measurement was chosen to quantify "
                  "air emissions during FY 2025. In the absence of any deviation from primary "
                  "methodologies, no disclosure of reasons for choosing an inferior methodology is "
                  "required for this reporting period."))

    # ── 5.3.6  WATER ────────────────────────────────────────
    heading(doc, '5.3.6  Pollution of Water', 2)
    section_codes(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
                  'ESRS E2-4-AR-23-(c) | ESRS E2-4-30 | ESRS E2-4-31')
    _emissions_block(doc, '5.3.6', 'Water',
        AR22,
        [
            ('Total Suspended Solids (TSS)', '48.3', '43.9'),
            ('Chemical Oxygen Demand (COD)', '234.6', '237.7'),
            ('Biochemical Oxygen Demand (BOD)', '78.4', '72.1'),
            ('Nitrates (NO3-N)', '14.2', '12.8'),
            ('Phosphorus (Total-P)', '3.6', '3.1'),
            ('Heavy Metals (total)', '0.87', '0.74'),
            ('Hydrocarbons (TPH)', '2.14', '1.89'),
            ('Persistent Organic Pollutants (POP)', '0.012', '0.009'),
            ('Total Emissions to Water', '382.2', '372.2'),
        ],
        [
            ('TSS','4.8','3.9','3.1','4.2','2.8','2.4','2.1','3.4','3.0','2.7','32.4'),
            ('COD','26.1','21.2','16.4','22.4','14.8','12.7','11.2','17.8','15.9','14.3','172.8'),
            ('BOD','7.9','6.4','5.0','6.8','4.5','3.9','3.4','5.4','4.8','4.3','52.4'),
            ('Total','38.8','31.5','24.5','33.4','22.1','19.0','16.7','26.6','23.7','21.3','257.6'),
        ],
        [
            ('TSS','5.8','6.8','2.4','2.0','3.4','2.2','1.6','2.9','2.6','2.2','31.9'),
            ('COD','31.2','36.4','12.8','10.6','18.2','11.8','8.6','15.4','13.8','11.8','170.6'),
            ('BOD','9.4','11.2','3.9','3.2','5.5','3.6','2.6','4.7','4.2','3.6','51.9'),
            ('Total','46.4','54.4','19.1','15.8','27.1','17.6','12.8','23.0','20.6','17.6','254.4'),
        ],
        (
            ['Pollutant','Direct','Runoff','Cooling','Process','Sewage','Leachate',
             'Atm. Dep.','Spill','Sludge','Other','Total (t)'],
            [
                ('TSS','8.2','7.4','4.8','6.4','4.2','3.8','2.6','2.1','2.0','1.4','42.9'),
                ('COD','44.2','39.8','26.4','34.2','22.8','20.4','14.2','11.4','10.8','7.4','231.6'),
                ('BOD','13.4','12.2','8.0','10.4','6.8','6.2','4.2','3.4','3.2','2.2','70.0'),
                ('Total','65.8','59.4','39.2','51.0','33.8','30.4','21.0','16.9','16.0','11.0','344.5'),
            ]),
        [
            ('TSS','5.6','6.6','2.3','1.9','3.2','2.1','1.5','2.8','2.5','2.1','30.6'),
            ('COD','30.4','35.2','12.4','10.2','17.6','11.4','8.2','14.8','13.4','11.4','165.0'),
            ('BOD','9.2','10.8','3.8','3.1','5.3','3.5','2.5','4.5','4.1','3.5','50.3'),
            ('Total','45.2','52.6','18.5','15.2','26.1','17.0','12.2','22.1','20.0','17.0','245.9'),
        ],
        'Water-Related Operations',
        stress_rows=[
            ('Areas at Water Risk', '47.8', '12.8%'),
            ('Areas of High Water Stress', '89.4', '24.0%'),
            ('Total Emissions to Water', '372.2', '100.0%'),
        ],
        changes="Total emissions to water declined by 2.6% year-on-year to 372.2 metric tonnes in "
                "FY 2025. TSS declined by 9.1% to 43.9 metric tonnes and BOD declined by 8.0% to "
                "72.1 metric tonnes, reflecting the full-year effect of upgraded biological "
                "treatment systems installed at Sites 2 and 4 in late FY 2024. COD increased "
                "marginally by 1.3% to 237.7 metric tonnes due to a temporary increase in "
                "production throughput at two facilities during H1 FY 2025. Nitrate and phosphorus "
                "loadings continued to decline. No uncontrolled releases or spills to water were "
                "recorded during FY 2025.",
        methods="Water emission quantification uses a combination of continuous on-line effluent "
                "monitoring, composite sample analysis by accredited environmental laboratories, "
                "and mass balance methodologies for cooling water and storm water discharges. All "
                "IED-regulated installations with direct water discharges are subject to effluent "
                "monitoring in accordance with permit conditions. Monitoring frequencies range from "
                "continuous (for CEMS-equipped outfalls) to monthly composite sampling for non-CEMS "
                "discharge points.",
        datacol="Water pollution data is collected via the Company's EMIS, integrating on-line "
                "monitoring data, laboratory analysis certificates, and process data inputs from "
                "site-level environmental officers. Data quality controls include laboratory "
                "accreditation requirements (ISO 17025), automated validation checks for on-line "
                "monitoring data, and annual third-party data assurance. Where sampling gaps arise, "
                "conservative estimation is applied in accordance with documented data gap-filling "
                "protocols.",
        inferior=('No. All water emission quantification is based on direct effluent monitoring or '
                  'recognised mass balance and calculation standards.',
                  "No inferior methodology compared to direct measurement was chosen to quantify "
                  "water emissions during FY 2025. No disclosure of reasons for choosing an "
                  "inferior methodology is required for this reporting period."))

    # ── 5.3.7  SOIL ─────────────────────────────────────────
    heading(doc, '5.3.7  Pollution of Soil', 2)
    section_codes(doc, 'ESRS E2-4 | ESRS E2-4-28 | ESRS E2-4-AR-20 | ESRS E2-4-AR-22 | '
                  'ESRS E2-4-AR-23-(c) | ESRS E2-4-30 | ESRS E2-4-31')
    _emissions_block(doc, '5.3.7', 'Soil',
        AR22,
        [
            ('Heavy Metals', '6.9', '6.5'),
            ('Hydrocarbons (TPH)', '24.9', '23.7'),
            ('Pesticides/Biocides', '1.8', '1.7'),
            ('Nitrates — Diffuse', '11.0', '10.3'),
            ('Total Emissions to Soil', '44.6', '42.2'),
        ],
        [
            ('Heavy Metals','1.2','0.8','0.6','0.9','0.5','0.4','0.3','0.5','0.4','0.3','5.9'),
            ('Hydrocarbons','3.4','2.8','2.1','2.9','1.8','1.6','1.4','2.2','1.9','1.6','21.7'),
            ('Pesticides','0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','0.2','2.0'),
            ('Nitrates—Diffuse','1.4','1.0','0.9','1.2','0.8','1.0','0.9','1.3','1.0','1.5','11.0'),
            ('Total','6.2','4.8','3.8','5.2','3.3','3.2','2.8','4.2','3.5','3.6','40.6'),
        ],
        [
            ('Heavy Metals','1.8','2.1','0.4','0.3','0.6','0.3','0.2','0.4','0.4','0.3','6.8'),
            ('Hydrocarbons','4.8','5.6','2.1','1.8','3.0','1.9','1.4','1.7','1.8','1.6','25.7'),
            ('Pesticides','0.1','0.4','0.2','0.1','0.2','0.1','0.1','0.2','0.2','0.2','1.8'),
            ('Nitrates—Diffuse','1.0','2.6','1.2','0.9','0.9','0.8','0.7','1.0','0.9','1.0','11.0'),
            ('Total','7.7','10.7','3.9','3.1','4.7','3.1','2.4','3.3','3.3','3.1','45.3'),
        ],
        (
            ['Pollutant','Fugitive','Spill','Diffuse','Leachate','Runoff','Process',
             'Atm. Dep.','Waste','Sludge','Remediation','Total (t)'],
            [
                ('Heavy Metals','0.8','1.4','0.5','0.6','0.4','0.6','0.3','0.5','0.5','0.3','5.9'),
                ('Hydrocarbons','3.2','6.4','3.8','2.2','2.6','2.9','1.2','2.4','2.2','1.4','28.3'),
                ('Pesticides','0.1','0.2','0.6','0.1','0.2','0.1','0.1','0.1','0.1','0.1','1.7'),
                ('Nitrates—Diffuse','0.8','0.7','4.2','0.8','1.4','0.7','0.4','0.8','0.8','0.4','11.0'),
                ('Total','4.9','8.7','9.1','3.7','4.6','4.3','2.0','3.8','3.6','2.2','46.9'),
            ]),
        [
            ('Heavy Metals','1.8','2.0','0.4','0.2','0.5','0.3','0.2','0.4','0.4','0.3','6.5'),
            ('Hydrocarbons','4.7','5.5','2.1','1.7','2.9','1.8','1.3','1.7','1.7','1.5','24.9'),
            ('Pesticides','0.1','0.4','0.2','0.1','0.2','0.1','0.1','0.2','0.2','0.2','1.8'),
            ('Nitrates—Diffuse','1.0','2.5','1.2','0.9','0.9','0.7','0.6','0.9','0.9','0.7','10.3'),
            ('Total','7.6','10.4','3.9','2.9','4.5','2.9','2.2','3.2','3.2','2.7','37.5'),
        ],
        'Soil-Related Operations',
        stress_rows=[
            ('Areas at Water Risk', '4.2', '9.5%'),
            ('Areas of High Water Stress', '9.7', '21.9%'),
            ('Total Emissions to Soil', '42.2', '100.0%'),
        ],
        changes="Total soil emissions declined by 5.4% year-on-year to 42.2 metric tonnes in "
                "FY 2025. Heavy metal loadings to soil declined by 5.8% and hydrocarbon (TPH) "
                "loadings declined by 4.8%, reflecting improved secondary containment measures and "
                "spill prevention controls implemented across Sites 2 and 8. Nitrate diffuse "
                "loadings remained broadly stable. No new uncontrolled releases to soil were "
                "recorded during FY 2025. Ongoing remediation at three legacy contaminated sites "
                "progressed in line with approved site remediation plans.",
        methods="Soil pollution quantification applies a combination of direct measurement (site "
                "investigation sampling and analysis for legacy contamination assessment), emission "
                "factor calculations (for diffuse agricultural nitrate loading estimates), mass "
                "balance methodologies (for secondary containment and spill quantification), and "
                "incident-based quantification where applicable. Site investigation sampling "
                "follows accredited protocols aligned with ISO 18400 series standards. Measurement "
                "approaches are reviewed annually in conjunction with site environmental management "
                "plan updates.",
        datacol="Soil pollution data is collected via the Company's EMIS, integrating site "
                "investigation reports, incident records, process data inputs, and regulatory "
                "submission data. Data quality controls include peer review of site investigation "
                "data by Senior Environmental Managers, cross-validation against regulatory permit "
                "data, and annual third-party data assurance. Where data gaps arise, conservative "
                "estimation is applied in accordance with documented protocols, and such estimates "
                "are clearly flagged in the EMIS audit trail.",
        inferior=('Partial. Emission factor calculations are applied for diffuse nitrate loading '
                  'estimates only; all other soil emission categories are measured directly.',
                  "No inferior methodology compared to direct measurement was chosen to quantify "
                  "soil emissions during FY 2025, with the exception of diffuse nitrate loading "
                  "estimates where emission factor calculations are applied in the absence of "
                  "direct measurement feasibility for diffuse agricultural sources. This approach "
                  "is consistent with accepted scientific and regulatory practice for diffuse "
                  "pollution quantification."))

    print("Sections 5.3.5–5.3.7 built.")
