#!/usr/bin/env python3
"""Pollution Report Section 5.3 — FY 2025 — Reference-Aligned Edition."""

import sys
sys.path.insert(0, '/home/user/jcode')
from docx import Document
from docx.shared import Inches
from rb_helpers import (heading, body, section_codes, note, data_table,
                        status_table, kpi_table, governance_table,
                        overview_table, content_index_table)

OUT = '/home/user/jcode/Pollution_Report_5_3_FY2025_Corrected.docx'

doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Inches(1)
    s.left_margin = s.right_margin = Inches(1.1)

# ════════════════════════════════════════════════════════════
# 5.3  POLLUTION  (title + intro)
# ════════════════════════════════════════════════════════════
heading(doc, '5.3  Pollution', 1)

body(doc,
     "Pollution prevention and environmental stewardship are core commitments within "
     "XYZ Logistics LLC's environmental strategy. The Company recognises that effective "
     "pollution management is fundamental to protecting ecosystems, human health, and the "
     "communities in which it operates. XYZ Logistics LLC maintains a comprehensive Pollution "
     "Prevention and Control programme designed to minimise emissions to air, water, and soil; "
     "responsibly manage substances of concern (SoC) and substances of very high concern (SVHC); "
     "and proactively mitigate pollution-related financial and operational risks across all "
     "levels of the organisation.",
     code='GRI 305-7')

body(doc,
     "This section presents XYZ Logistics LLC's Pollution disclosure for FY 2025, aligned with "
     "the relevant European Sustainability Reporting Standards (ESRS) and applicable GRI "
     "Standards.",
     code='ESRS E2 | ESRS E2-IRO-1 | ESRS E2-4 | ESRS E2-5 | ESRS E2-6')

# ── Pollution at a Glance ────────────────────────────────────
heading(doc, 'Pollution at a Glance', 2)
body(doc, "Overview: Reporting Topics & GRI / ESRS Alignment", bold=True)
overview_table(doc, [
    ('Processes to Identify and Assess Material Pollution-Related Impacts, Risks and Opportunities',
     '—', 'ESRS E2 | E2-IRO-1'),
    ('Policies Related to Pollution', '—', 'ESRS E2 | E2-1'),
    ('Actions and Resources Related to Pollution', '—', 'ESRS E2 | E2-2'),
    ('Targets Related to Pollution', '—', 'ESRS E2 | E2-3'),
    ('Pollution of Air', 'GRI 305-7', 'ESRS E2 | E2-4'),
    ('Pollution of Water', '—', 'ESRS E2 | E2-4'),
    ('Pollution of Soil', '—', 'ESRS E2 | E2-4'),
    ('Integrated Pollution Control & BAT Performance Tracking', '—', 'ESRS E2 | E2-4'),
    ('Substances of Concern', '—', 'ESRS E2 | E2-5'),
    ('Substances of Very High Concern', '—', 'ESRS E2 | E2-5'),
    ('Anticipated Financial Effects from Material Pollution-Related Risks and Opportunities',
     '—', 'ESRS E2 | E2-6'),
])

# ── Management Approach ──────────────────────────────────────
heading(doc, 'Management Approach', 2)
body(doc,
     "XYZ Logistics LLC's management approach to pollution prevention and control is anchored in "
     "a formal governance framework that integrates environmental risk management, regulatory "
     "compliance, and strategic sustainability commitments. The Pollution Prevention and "
     "Environmental Stewardship Policy establishes the overarching framework across all "
     "operational boundaries and is reviewed annually against the EU Zero Pollution Action Plan, "
     "the REACH Regulation, the Industrial Emissions Directive, and applicable national "
     "environmental legislation. Material pollution-related impacts, risks, and opportunities are "
     "identified through the Company's annual double materiality assessment, the outcomes of "
     "which inform target-setting, resource allocation, and disclosure priorities. Performance "
     "against pollution targets is tracked quarterly through the environmental KPI dashboard and "
     "reported to the Board's Environmental and Sustainability Committee.",
     code='ESRS E2-1 | ESRS E2-2 | ESRS E2-3')
governance_table(doc, [
    ('Board Environmental & Sustainability Committee',
     'Oversight of pollution-related risks and targets; quarterly review of performance against targets.'),
    ('Chief Sustainability Officer (CSO)',
     'Executive accountability for pollution performance, policy ownership, and regulatory monitoring.'),
    ('Regional Environmental Leads',
     'Local implementation, site-level pollution prevention action plans, and monitoring data integrity.'),
    ('Operations & Facilities',
     'On-site emission control, abatement equipment operation, and incident response documentation.'),
    ('Internal Audit',
     'Independent assurance, testing of environmental controls, and policy adherence verification.'),
])

# ════════════════════════════════════════════════════════════
# 5.3.1  Processes to Identify and Assess
# ════════════════════════════════════════════════════════════
heading(doc, '5.3.1  Processes to Identify and Assess Material Pollution-Related '
        'Impacts, Risks and Opportunities', 2)
section_codes(doc, 'ESRS E2-IRO-1 | ESRS E2-IRO-1-11-(a) | ESRS E2-IRO-1-11-(b)')

heading(doc, '5.3.1.1  Information About the Process to Identify Actual and Potential '
        'Pollution-Related Impacts, Risks and Opportunities', 3)
body(doc,
     "XYZ Logistics LLC maintains a formal, structured process to identify and assess material "
     "pollution-related impacts, risks, and opportunities across its operations and value chain. "
     "The Company applies a systematic double materiality assessment methodology, combining an "
     "impact materiality lens — the actual and potential effects of the Company's operations on "
     "ecosystems, water bodies, soils, and communities — with a financial materiality lens "
     "covering risks and opportunities arising from pollution-related regulatory, reputational, "
     "and transition factors. This integrated assessment is conducted annually and forms the "
     "foundation of the Company's environmental reporting and strategic planning cycle.",
     code='ESRS E2-IRO-1-11-(a)')

heading(doc, '5.3.1.2  Sites and Activities Screened for Pollution-Related Impacts, '
        'Risks and Opportunities — and Methodologies, Assumptions and Tools Used', 3)
status_table(doc, [
    ('Sites and activities screened for pollution-related impacts, risks, and opportunities',
     'ESRS E2-IRO-1-11-(a)', 'Yes. All operating sites and business activities were screened '
     'during FY 2025 as part of the annual environmental risk register review.'),
])
body(doc,
     "The screening process applies internationally recognised methodologies, including "
     "ISO 14001-aligned environmental aspect and impact evaluation, EU Industrial Emissions "
     "Directive (IED) applicability assessments, and site-specific environmental baseline studies "
     "where elevated risk is identified. Assumptions underpinning the screening include the "
     "application of conservative emission thresholds, worst-case scenario modelling for incident "
     "probability assessments, and reference to current EU BAT Conclusions for sector-relevant "
     "installations. Geospatial mapping tools are deployed to identify proximity to sensitive "
     "receptors including water bodies, Natura 2000 sites, and populated areas.",
     code='ESRS E2-IRO-1-11-(a)')

heading(doc, '5.3.1.3  Stakeholder and Community Consultations Conducted in the '
        'Pollution Impact Assessment Process', 3)
status_table(doc, [
    ('Consultations conducted, particularly with affected communities',
     'ESRS E2-IRO-1-11-(b)', 'Yes. Structured consultations were carried out at all operating '
     'sites adjacent to residential or environmentally sensitive areas during FY 2025.'),
])
body(doc,
     "Activities included community liaison meetings, engagement with local environmental "
     "authorities, and formalised grievance mechanisms accessible to affected parties. The "
     "outcomes of these consultations are documented, reviewed by the Pollution Prevention "
     "Council, and incorporated into site-level environmental management plans. Where material "
     "community-level concerns are identified, corrective actions are tracked through the "
     "Company's environmental incident management system and reported to the Board's "
     "Environmental and Sustainability Committee.",
     code='ESRS E2-IRO-1-11-(b)')

# ════════════════════════════════════════════════════════════
# 5.3.2  Policies Related to Pollution
# ════════════════════════════════════════════════════════════
heading(doc, '5.3.2  Policies Related to Pollution', 2)
section_codes(doc, 'ESRS E2-1 | ESRS E2-1-15-(a) | ESRS E2-1-15-(b) | '
              'ESRS E2-1-15-(c) | ESRS E2-1-AR-12')

heading(doc, '5.3.2.1  Policy Coverage: Mitigation of Negative Impacts on Air, '
        'Water and Soil', 3)
status_table(doc, [
    ('Policy addresses mitigating negative impacts related to pollution of air, water and soil',
     'ESRS E2-1-15-(a)', 'Yes. The Pollution Prevention and Environmental Stewardship Policy '
     'explicitly addresses mitigation across all operating entities and the value chain.'),
])
body(doc,
     "The Policy establishes clear emission reduction targets, mandatory best available technique "
     "(BAT) adoption requirements for regulated installations, and site-level environmental "
     "management plan obligations. Policy commitments on air, water, and soil pollution are "
     "reviewed annually in conjunction with updated regulatory standards, monitoring data, and "
     "materiality assessment outcomes.",
     code='ESRS E2-1-15-(a)')

heading(doc, '5.3.2.2  Policy Coverage: Substitution and Minimisation of Substances '
        'of Concern and Phase-Out of Substances of Very High Concern', 3)
status_table(doc, [
    ('Policy addresses substituting and minimising use of substances of concern and phasing '
     'out substances of very high concern',
     'ESRS E2-1-15-(b)', 'Yes. The commitment is operative across all business functions, '
     'supported by a Restricted Substances List aligned with REACH Annex XVII.'),
])
body(doc,
     "The Company maintains a Restricted Substances List (RSL) aligned with REACH Regulation "
     "Annex XVII and the SVHC Candidate List, and applies a substitution programme requiring "
     "annual review of all SoC and SVHC in use, with mandatory substitution timelines for SVHC "
     "where technically and economically feasible alternatives exist.",
     code='ESRS E2-1-15-(b)')

heading(doc, '5.3.2.3  Policy Coverage: Prevention of and Response to Incidents and '
        'Emergency Situations Resulting in Pollution', 3)
status_table(doc, [
    ('Policy addresses avoiding incidents and emergency situations resulting in pollution, '
     'and controlling or limiting their impact',
     'ESRS E2-1-15-(c)', 'Yes. The obligation is reflected across all operating sites and is '
     'subject to mandatory annual testing.'),
])
body(doc,
     "Emergency response procedures are maintained at all sites, incorporating spill containment "
     "protocols, emergency shutdown procedures, regulatory notification requirements, and "
     "post-incident environmental impact assessments. All relevant personnel receive mandatory "
     "emergency response training. The Company maintains a centralised incident register, and all "
     "pollution-related incidents are subject to root cause analysis and corrective action "
     "tracking.",
     code='ESRS E2-1-15-(c)')

heading(doc, '5.3.2.4  Contextual Alignment of Pollution Policy with the '
        'EU Zero Pollution Action Plan', 3)
body(doc,
     "XYZ Logistics LLC's pollution-related policy commitments are explicitly contextualised "
     "against the EU Zero Pollution Action Plan (ZPAP) and its 2030 targets. The Company has "
     "mapped its internal emission reduction targets, SoC substitution commitments, and water "
     "quality protection objectives against the ZPAP's headline goals, including the reduction of "
     "air pollution health impacts, the reduction of microplastic pollution, and the improvement "
     "of water and soil quality. This alignment is reviewed as part of the annual policy review "
     "cycle and disclosed in the Company's Annual Sustainability Report.",
     code='ESRS E2-1-AR-12')

# ════════════════════════════════════════════════════════════
# 5.3.3  Actions and Resources Related to Pollution
# ════════════════════════════════════════════════════════════
heading(doc, '5.3.3  Actions and Resources Related to Pollution', 2)
section_codes(doc, 'ESRS E2-2 | ESRS E2-2-19 | ESRS E2-2-AR-13 | ESRS E2-2-AR-15')

heading(doc, '5.3.3.1  Layer in Mitigation Hierarchy to Which Actions '
        'Can Be Allocated — Pollution', 3)
body(doc,
     "XYZ Logistics LLC's pollution-related actions are structured in accordance with the "
     "environmental mitigation hierarchy, prioritising: (i) Avoidance — process redesign to "
     "eliminate emission-generating steps, substitution of high-emission materials, and "
     "installation of enclosed handling systems for volatile substances; (ii) Reduction — "
     "deployment of advanced emission abatement technologies, energy efficiency improvements, and "
     "real-time emission monitoring systems; (iii) Restoration — targeted remediation programmes "
     "at historically contaminated sites; and (iv) Compensation — compensatory environmental "
     "measures applied where residual impacts cannot otherwise be addressed. The majority of "
     "FY 2025 pollution actions are allocated to the Avoidance and Reduction layers of the "
     "mitigation hierarchy.",
     code='ESRS E2-2-19')

heading(doc, '5.3.3.2  Pollution Actions Extending to Upstream and Downstream '
        'Value Chain Engagements', 3)
status_table(doc, [
    ('Pollution-related actions extend to upstream and downstream value chain engagements',
     'ESRS E2-2-AR-13', 'Yes. Supplier Environmental Assessments are conducted for all Tier 1 '
     'suppliers presenting elevated pollution risk, with downstream customer engagement.'),
])
body(doc,
     "Supplier Environmental Assessments (SEAs) are conducted for all Tier 1 suppliers assessed as "
     "presenting elevated pollution risk. Downstream, XYZ Logistics LLC engages customers on the "
     "environmental attributes of its products and services, including guidance on responsible "
     "end-of-life management for products containing substances of concern. Value chain "
     "engagement outcomes are tracked through the Company's supplier sustainability programme and "
     "reported annually.",
     code='ESRS E2-2-AR-13')

heading(doc, '5.3.3.3  Layer in Mitigation Hierarchy to Which Resources '
        'Can Be Allocated — Pollution', 3)
body(doc,
     "Financial and human resources allocated to pollution prevention and control are prioritised "
     "in accordance with risk ranking from the annual environmental risk register. Resource "
     "allocations are directed to the Avoidance and Reduction layers of the mitigation hierarchy "
     "as a primary priority, with Restoration-layer resources committed for active contaminated "
     "site remediation programmes. The Company's EHS Capital Expenditure (CapEx) budget includes "
     "dedicated allocations for emission abatement technology upgrades, site remediation "
     "activities, and environmental monitoring infrastructure. Resource allocation decisions are "
     "reviewed semi-annually by the Chief Sustainability Officer and presented to the Board's "
     "Environmental and Sustainability Committee on an annual basis.",
     code='ESRS E2-2-19')

heading(doc, '5.3.3.4  Site-Level Pollution Prevention Action Plans', 3)
status_table(doc, [
    ('Action plans implemented at site level',
     'ESRS E2-2-AR-15', 'Yes. Pollution prevention action plans are implemented at site level '
     'across all primary operating facilities for the FY 2025 reporting period.'),
])
body(doc,
     "These plans cover emission monitoring protocols, substance management procedures, emergency "
     "response plans, and maintenance schedules for abatement equipment. Site-level action plans "
     "are reviewed annually by Regional Environmental Leads and updated to reflect current risk "
     "assessments, regulatory developments, and performance monitoring outcomes. Facilities "
     "subject to IED permitting maintain site-specific environmental improvement programmes "
     "aligned with applicable BAT Conclusions.",
     code='ESRS E2-2-AR-15')

# ════════════════════════════════════════════════════════════
# 5.3.4  Targets Related to Pollution
# ════════════════════════════════════════════════════════════
heading(doc, '5.3.4  Targets Related to Pollution', 2)
section_codes(doc, 'ESRS E2-3 | ESRS E2-3-23 | ESRS E2-3-24 | ESRS E2-3-25 | '
              'ESRS E2-3-AR-16 | ESRS E2-3-AR-17 | ESRS E2-3-AR-18')

heading(doc, '5.3.4.1  Target: Prevention and Control of Air Pollutants and '
        'Respective Specific Loads', 3)
status_table(doc, [
    ('Target relates to the prevention and control of air pollutants and respective specific loads',
     'ESRS E2-3-23-(a)', 'Yes. Quantitative air quality targets are operative for FY 2025.'),
])
body(doc,
     "The Company's air quality target for FY 2030 commits to a 40% reduction in total NOx and "
     "SOx emissions relative to the FY 2022 baseline, and a 35% reduction in particulate matter "
     "(PM2.5 and PM10) over the same period. Annual interim milestones are tracked through the "
     "environmental KPI dashboard and reported quarterly to the Board's Environmental and "
     "Sustainability Committee. The FY 2025 interim milestone of a 16% reduction in total NOx/SOx "
     "emissions relative to baseline was achieved, with a reported 17.2% reduction.",
     code='ESRS E2-3-23-(a)')

heading(doc, '5.3.4.2  Target: Prevention and Control of Emissions to Water and '
        'Respective Specific Loads', 3)
status_table(doc, [
    ('Target relates to the prevention and control of emissions to water and respective specific loads',
     'ESRS E2-3-23-(b)', 'Yes. Water emission targets are operative for FY 2025.'),
])
body(doc,
     "The water emission target for FY 2030 commits to a 30% reduction in total suspended solids "
     "(TSS) and chemical oxygen demand (COD) loadings to receiving water bodies relative to the "
     "FY 2022 baseline, with a specific sub-target of zero direct discharge of untreated process "
     "water. Interim FY 2025 performance indicates a 13.5% reduction in TSS and COD loadings. A "
     "supplementary target for the elimination of discharges to water bodies in areas of high "
     "water stress has been set for FY 2027.",
     code='ESRS E2-3-23-(b)')

heading(doc, '5.3.4.3  Target: Prevention and Control of Pollution to Soil and '
        'Respective Specific Loads', 3)
status_table(doc, [
    ('Target relates to the prevention and control of pollution to soil and respective specific loads',
     'ESRS E2-3-23-(c)', 'Yes. Soil pollution targets are operative for FY 2025.'),
])
body(doc,
     "These targets are centred on the elimination of uncontrolled releases of hazardous "
     "substances to land, the progressive remediation of historically contaminated sites, and the "
     "reduction of diffuse soil contamination. The FY 2030 soil pollution target requires all "
     "active contaminated site remediation programmes to have achieved regulatory closure or be "
     "at advanced remediation stage. As of FY 2025, three of seven identified legacy contaminated "
     "sites are at advanced remediation stage, and one has achieved regulatory closure.",
     code='ESRS E2-3-23-(c)')

heading(doc, '5.3.4.4  Target: Prevention and Control of Substances of Concern and '
        'Substances of Very High Concern', 3)
status_table(doc, [
    ('Target relates to the prevention and control of substances of concern and substances of '
     'very high concern',
     'ESRS E2-3-23-(d)', 'Yes. SoC and SVHC targets are operative for FY 2025.'),
])
body(doc,
     "The SVHC phase-out target for FY 2028 requires the substitution or elimination of all SVHC "
     "for which technically and economically feasible alternatives have been identified through "
     "the annual RSL review. The SoC minimisation target requires a 25% reduction in total SoC "
     "generated, procured, or used in production by FY 2030 relative to the FY 2022 baseline. As "
     "of FY 2025, 4 of 11 identified SVHC substitution projects have been completed.",
     code='ESRS E2-3-23-(d)')

heading(doc, '5.3.4.5  Ecological Thresholds and Entity-Specific Allocations '
        'Considered When Setting Pollution-Related Targets', 3)
status_table(doc, [
    ('Ecological thresholds and entity-specific allocations taken into consideration when '
     'setting pollution-related targets',
     'ESRS E2-3-24 | ESRS E2-3-AR-16-(a)', 'Yes. Ecological thresholds were identified at local, '
     'national, and global levels and incorporated into the target-setting process.'),
])
body(doc,
     "Ecological thresholds — including applicable environmental quality standards (EQS) for water "
     "bodies and air quality limit values under EU Directive 2008/50/EC — have been incorporated "
     "into the target-setting process. Ecological thresholds were identified through regulatory "
     "limit benchmarking, scientific literature review, and consultation with competent "
     "environmental authorities. Entity-specific emission allocations were determined through a "
     "top-down allocation approach distributing the Company's permitted emission budget across "
     "operating sites on the basis of site throughput, receptor sensitivity, and applicable "
     "permit conditions.",
     code='ESRS E2-3-24 | ESRS E2-3-AR-16-(a)')

heading(doc, '5.3.4.6  Methodology for Identifying Ecological Thresholds and '
        'Determining Entity-Specific Allocations', 3)
body(doc,
     "The methodology applied to identify ecological thresholds combines: (i) a regulatory "
     "benchmark approach referencing applicable EU environmental quality standards, IED permit "
     "conditions, and EU BAT Conclusions; (ii) a scientific literature review drawing on current "
     "ecological risk assessment frameworks; and (iii) site-specific environmental carrying "
     "capacity assessments prepared by accredited environmental consultants. Entity-specific "
     "thresholds for each operating site were determined through proportional allocation "
     "modelling, applying site emission intensity ratios adjusted for receptor sensitivity "
     "classifications established through geospatial environmental sensitivity mapping.",
     code='ESRS E2-3-24 | ESRS E2-3-AR-16-(a) | ESRS E2-3-AR-16-(b)')

heading(doc, '5.3.4.7  Responsibility for Respecting Ecological Thresholds, '
        'Mandatory/Voluntary Status, and Substantial Contribution Criteria', 3)
status_table(doc, [
    ('Pollution-related target is mandatory (required by legislation) or voluntary',
     'ESRS E2-3-25', 'Mandatory and voluntary. Targets comprise mandatory components required by '
     'IED and REACH and voluntary components aligned with the EU Zero Pollution Action Plan.'),
    ('Target addresses shortcomings related to the Substantial Contribution criteria for '
     'Pollution Prevention and Control',
     'ESRS E2-3-AR-17', 'Yes. The mandatory components address shortcomings identified under the '
     'EU Taxonomy Substantial Contribution criteria.'),
])
body(doc,
     "Responsibility for monitoring performance against ecological thresholds rests with Regional "
     "Environmental Leads, with escalation obligations to the Chief Sustainability Officer where "
     "site-level allocations are at risk of being exceeded. Pollution reduction targets comprise "
     "both mandatory targets required by applicable environmental permitting conditions and EU "
     "regulatory frameworks (including IED and REACH), and voluntary commitments adopted in "
     "alignment with the EU Zero Pollution Action Plan. The voluntary components represent "
     "additional ambition beyond regulatory minimum requirements.",
     code='ESRS E2-3-24 | ESRS E2-3-AR-16-(c) | ESRS E2-3-25 | ESRS E2-3-AR-17')

heading(doc, '5.3.4.8  Pollution Targets Implemented at Site Level', 3)
body(doc,
     "Site-level implementation of all pollution targets is tracked through the EHS incident and "
     "performance management system, with quarterly progress reporting to the Board's "
     "Environmental and Sustainability Committee. Each operating site maintains a site-level "
     "pollution target implementation plan, reviewed annually by Regional Environmental Leads and "
     "updated to reflect current risk assessments, regulatory developments, and monitoring data "
     "outcomes. As of year-end FY 2025, all primary operating sites have operational site-level "
     "pollution target implementation plans in place.",
     code='ESRS E2-3-AR-18')

print("Sections 5.3–5.3.4 built.")
import rb_build_air_water_soil as p2
p2.build(doc)
import rb_build_tail as p3
p3.build(doc)

doc.save(OUT)
print(f"✓ saved {OUT}")
