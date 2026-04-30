"""04 — Series D Investor Pitch: Helio Health.

The classic ten-slide pitch deck — used to raise growth-stage capital
from institutional investors. The numbers are illustrative.

Slides:
    1. Cover
    2. Problem
    3. Solution
    4. Market opportunity (column chart)
    5. Traction & growth (line chart)
    6. Business model (KPIs)
    7. Go-to-market (3-phase)
    8. Competitive moat (table)
    9. Team
    10. The ask
    11. Closing
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from power_pptx.chart.data import CategoryChartData
from power_pptx.design.recipes import kpi_slide
from power_pptx.enum.chart import XL_CHART_TYPE
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from power_pptx.util import Inches, Pt

from _brand import INVESTOR as TOKENS, INVESTOR_PALETTE as PALETTE
from _common import (
    closing_slide,
    cover_slide,
    divider,
    eyebrow,
    footer,
    hex_rgb,
    lint_or_die,
    new_deck,
    section_title,
    styled_card,
)

FOOTER = "Helio Health — Series D  |  Confidential"


def build(out: Path) -> None:
    prs = new_deck()

    cover_slide(
        prs,
        eyebrow_text="Series D — $180M",
        title="Helio Health.\nThe operating system for cardiac care.",
        subtitle="A clinically-validated platform that has detected "
                 "atrial fibrillation in 1.4M patients and reduced "
                 "stroke incidence by 31% in deployed populations.",
        presenter="Dr. Anya Patel, Co-founder & CEO  |  David Park, CFO",
        date="May 2026",
        tokens=TOKENS,
    )

    _problem(prs)
    _solution(prs)
    _market(prs)
    _traction(prs)

    kpi_slide(
        prs, title="Compounding unit economics",
        kpis=[
            {"label": "ARR",                "value": "$162M", "delta": +1.43},
            {"label": "Net dollar retention", "value": "138%", "delta": +0.07},
            {"label": "Gross margin",       "value": "78%",   "delta": +0.06},
            {"label": "CAC payback",        "value": "11 mo", "delta": -0.18},
        ],
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER, "6", TOKENS)

    _go_to_market(prs)
    _moat(prs)
    _team(prs)
    _the_ask(prs)

    closing_slide(
        prs,
        headline="Cardiac care, run as software.",
        sub="anya@heliohealth.com  |  david@heliohealth.com",
        tokens=TOKENS,
    )

    lint_or_die(prs)
    prs.save(out)
    print(f"wrote {out}")


def _problem(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "The problem", TOKENS)
    section_title(slide, "Atrial fibrillation is the silent epidemic of the developed world.", TOKENS)
    divider(slide, TOKENS)

    stats = [
        ("38M", "people living with AFib globally"),
        ("1 in 4", "adults will develop AFib in their lifetime"),
        ("$26B", "annual U.S. cost of AFib-related strokes"),
        ("70%", "of cases are caught only after a stroke"),
    ]
    width = 2.95
    gap = 0.18
    left0 = 0.6
    for i, (num, body) in enumerate(stats):
        left = left0 + i * (width + gap)
        card = styled_card(slide, left, 1.95, width, 2.6, tokens=TOKENS,
                           fill_hex="#FFFFFF", stroke_hex="#E5E7EB")
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(16)
        tf.margin_top = Pt(18)
        tf.text = num
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(40)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
        p1 = tf.add_paragraph()
        p1.text = body
        p1.space_before = Pt(6)
        p1.font.size = Pt(13)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])

    summary = slide.shapes.add_textbox(
        Inches(0.6), Inches(4.95), Inches(12.1), Inches(1.4),
    )
    tf = summary.text_frame
    tf.word_wrap = True
    tf.text = ("AFib is asymptomatic, episodic, and easily missed by "
               "annual visits. Existing screening — 12-lead ECGs and "
               "Holter monitors — is expensive, episodic, and "
               "inaccessible at population scale.")
    tf.fit_text(font_family=TOKENS.typography["body"].family,
                max_size=18, bold=False)
    tf.paragraphs[0].font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
    footer(slide, FOOTER, "2", TOKENS)


def _solution(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "The solution", TOKENS)
    section_title(slide, "Continuous, FDA-cleared, ambient cardiac monitoring", TOKENS)
    divider(slide, TOKENS)

    pillars = [
        ("Wearable",
         "Adhesive patch worn 14 days. Class II, single-lead. 99.1% sensitivity.",
         TOKENS.palette["primary"]),
        ("Platform",
         "On-device ML triages, cloud confirms, clinician adjudicates. HIPAA / SOC 2 II.",
         TOKENS.palette["accent"]),
        ("Reimbursement",
         "CPT 93241–93248 reimbursable in 32 payers. Average revenue $812 per study.",
         PALETTE[2]),
    ]
    width = 4.1
    gap = 0.1
    for i, (head, body, color) in enumerate(pillars):
        left = 0.6 + i * (width + gap)
        card = styled_card(slide, left, 1.85, width, 2.7,
                           tokens=TOKENS, fill_hex="#FFFFFF",
                           stroke_hex="#E5E7EB")
        # Color band at top of card
        band = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(left), Inches(1.85),
            Inches(width), Inches(0.16),
        )
        band.fill.solid()
        band.fill.fore_color.rgb = hex_rgb(color)
        band.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(18)
        tf.margin_top = Pt(28)
        tf.margin_bottom = Pt(18)
        tf.text = head
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(22)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p1 = tf.add_paragraph()
        p1.text = body
        p1.space_before = Pt(6)
        p1.font.size = Pt(13)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])

    # Outcome banner
    banner = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(4.85), Inches(12.1), Inches(1.7),
    )
    banner.fill.linear_gradient(
        TOKENS.palette["primary"], TOKENS.palette["accent"], angle=0,
    )
    banner.line.fill.background()
    banner.adjustments[0] = 0.08
    tf = banner.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(28)
    tf.margin_top = Pt(20)
    tf.text = "Clinically demonstrated outcome"
    p0 = tf.paragraphs[0]
    p0.font.name = TOKENS.typography["body"].family
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = hex_rgb("#FFFFFF")
    p0.font.color.alpha = 0.9
    p1 = tf.add_paragraph()
    p1.text = "31% reduction in stroke incidence over 24 months in 41,000 deployed patients (peer-reviewed, JACC 2025)."
    p1.space_before = Pt(4)
    p1.font.name = TOKENS.typography["heading"].family
    p1.font.size = Pt(20)
    p1.font.bold = True
    p1.font.color.rgb = hex_rgb("#FFFFFF")
    footer(slide, FOOTER, "3", TOKENS)


def _market(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Market opportunity", TOKENS)
    section_title(slide, "$48B SAM in the U.S. alone, growing 18% annually", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["TAM (global cardiac monitoring)",
                       "SAM (U.S. AFib screening)",
                       "SOM (5 yr target)"]
    data.add_series("Market ($B)", (148, 48, 6.2))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(8.4), Inches(5.0),
        data,
    ).chart
    series = chart.series[0]
    fill = series.format.fill
    fill.gradient(kind="linear")
    fill.gradient_stops.replace([(0.0, TOKENS.palette["primary"]),
                                  (1.0, TOKENS.palette["accent"])])
    chart.apply_quick_layout("title_no_legend")
    chart.chart_title.text_frame.text = "Cardiac monitoring opportunity ($B)"

    card = styled_card(slide, 9.3, 1.7, 3.5, 5.0, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(18)
    tf.margin_top = tf.margin_bottom = Pt(18)
    tf.text = "Why now"
    p = tf.paragraphs[0]
    p.font.name = TOKENS.typography["heading"].family
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
    rows = [
        ("USPSTF guidance",
         "AFib screening recommended for adults 65+ as of Jan 2025."),
        ("Reimbursement parity",
         "32 payers now cover at parity with hospital telemetry."),
        ("Wearable ubiquity",
         "78% of adults 60+ own a smartphone — distribution at hand."),
    ]
    for h, b in rows:
        p = tf.add_paragraph()
        p.text = h
        p.space_before = Pt(10)
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p2 = tf.add_paragraph()
        p2.text = b
        p2.font.size = Pt(11)
        p2.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "4", TOKENS)


def _traction(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Traction", TOKENS)
    section_title(slide, "ARR has 7×'d in three years", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["FY23", "FY24", "FY25", "FY26"]
    data.add_series("ARR ($M)", (23, 56, 108, 162))
    data.add_series("Patients on platform (k)", (210, 480, 940, 1410))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Helio Health — annual recurring revenue & active patients"
    footer(slide, FOOTER, "5", TOKENS)


def _go_to_market(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Go-to-market", TOKENS)
    section_title(slide, "Three motions, sequenced over the next four years", TOKENS)
    divider(slide, TOKENS)

    motions = [
        ("Now — Cardiology clinics",
         "Direct sales to 1,800 cardiology practices in N. America. ACV $185k.",
         "FY26 — FY27"),
        ("Next — Health systems",
         "Enterprise contracts with IDNs and ACOs. Outcomes-based pricing.",
         "FY27 — FY28"),
        ("Then — Payer partnerships",
         "Capitated screening for Medicare Advantage cohorts. PMPM model.",
         "FY28 — FY29"),
    ]
    width = 4.05
    gap = 0.15
    for i, (head, body, when) in enumerate(motions):
        left = 0.6 + i * (width + gap)
        card = styled_card(slide, left, 1.85, width, 4.7,
                           tokens=TOKENS, fill_hex="#FFFFFF",
                           stroke_hex="#E5E7EB")
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = tf.margin_right = Pt(18)
        tf.margin_top = tf.margin_bottom = Pt(20)
        tf.text = when.upper()
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["body"].family
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(PALETTE[i])
        p1 = tf.add_paragraph()
        p1.text = head
        p1.space_before = Pt(4)
        p1.font.name = TOKENS.typography["heading"].family
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p2 = tf.add_paragraph()
        p2.text = body
        p2.space_before = Pt(8)
        p2.font.name = TOKENS.typography["body"].family
        p2.font.size = Pt(13)
        p2.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "7", TOKENS)


def _moat(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Moat", TOKENS)
    section_title(slide, "Five compounding advantages", TOKENS)
    divider(slide, TOKENS)

    rows = [
        ("Advantage",      "Detail",                                                   "Years to replicate"),
        ("Regulatory",     "Two FDA clearances (K223104, K231881) and CE mark.",       "3–4 years"),
        ("Clinical",       "Peer-reviewed JACC outcomes study with 41k patients.",     "4–5 years"),
        ("Reimbursement",  "32 payer contracts at parity rates negotiated since 2022.", "3+ years"),
        ("Data",           "1.4M annotated rhythm strips — proprietary training set.",  "5+ years"),
        ("Distribution",   "1,800 cardiology relationships, NPS 71.",                  "3+ years"),
    ]
    shape = slide.shapes.add_table(
        rows=len(rows), cols=3,
        left=Inches(0.6), top=Inches(1.7),
        width=Inches(12.1), height=Inches(4.4),
    )
    table = shape.table
    table.columns[0].width = Inches(2.7)
    table.columns[1].width = Inches(6.7)
    table.columns[2].width = Inches(2.7)

    for c, label in enumerate(rows[0]):
        cell = table.cell(0, c)
        cell.text = label
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["primary"])
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = hex_rgb(TOKENS.palette["on_primary"])
        cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    for r, row in enumerate(rows[1:], start=1):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(12)
            p.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
            if c == 0:
                p.font.bold = True
                p.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            if r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["surface"])
            cell.borders.bottom.color.rgb = hex_rgb("#E5E7EB")
            cell.borders.bottom.width = Pt(0.5)
    footer(slide, FOOTER, "8", TOKENS)


def _team(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Leadership", TOKENS)
    section_title(slide, "Operators who have built and exited in this market", TOKENS)
    divider(slide, TOKENS)

    team = [
        ("Dr. Anya Patel",   "Co-founder & CEO",
         "Practicing electrophysiologist. Ex-VP, Clinical at Abbott (Verily acquired)."),
        ("Marcus Lee",       "Co-founder & CTO",
         "Ex-Principal ML, Apple Health. Architected Apple Watch ECG."),
        ("David Park",       "CFO",
         "Took GeniusCardio public in 2019 ($2.4B exit)."),
        ("Dr. Lia Hernandez","Chief Medical Officer",
         "Past President, Heart Rhythm Society. 140+ publications."),
        ("Aisha Khan",       "Chief Commercial Officer",
         "Built and led Boston Scientific's diagnostics field org."),
        ("James Okafor",     "Chief Regulatory Officer",
         "Ex-FDA reviewer, Class II/III cardiology devices."),
    ]
    cols = 3
    card_w = 4.05
    card_h = 2.45
    gap_x = 0.15
    gap_y = 0.2
    for i, (name, role, bio) in enumerate(team):
        col = i % cols
        row = i // cols
        left = 0.6 + col * (card_w + gap_x)
        top = 1.8 + row * (card_h + gap_y)
        card = styled_card(slide, left, top, card_w, card_h,
                           tokens=TOKENS, fill_hex="#FFFFFF",
                           stroke_hex="#E5E7EB")
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(16)
        tf.margin_top = tf.margin_bottom = Pt(14)
        tf.text = name
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(17)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p1 = tf.add_paragraph()
        p1.text = role
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
        p2 = tf.add_paragraph()
        p2.text = bio
        p2.space_before = Pt(6)
        p2.font.name = TOKENS.typography["body"].family
        p2.font.size = Pt(11)
        p2.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "9", TOKENS)


def _the_ask(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5),
    )
    bg.fill.linear_gradient(
        TOKENS.palette["primary"], TOKENS.palette["neutral"], angle=135,
    )
    bg.line.fill.background()

    eb = slide.shapes.add_textbox(
        Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.4),
    )
    tf = eb.text_frame
    tf.text = "THE ASK"
    p = tf.paragraphs[0]
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["accent"])

    head = slide.shapes.add_textbox(
        Inches(0.8), Inches(1.3), Inches(11.7), Inches(2.0),
    )
    tf = head.text_frame
    tf.word_wrap = True
    tf.text = "$180M Series D at a $1.6B post-money."
    tf.fit_text(font_family=TOKENS.typography["heading"].family,
                max_size=52, bold=True)
    tf.paragraphs[0].font.color.rgb = hex_rgb("#FFFFFF")

    use = [
        ("$70M",  "Sales & marketing — health-system motion build-out"),
        ("$55M",  "R&D — atrial flutter + heart-failure expansion"),
        ("$30M",  "International — UK, Germany, Japan launches"),
        ("$25M",  "Working capital & balance sheet"),
    ]
    for i, (amt, body) in enumerate(use):
        left = 0.8 + (i % 4) * 3.05
        top = 4.0
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(2.9), Inches(2.4),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = hex_rgb("#FFFFFF")
        card.fill.fore_color.alpha = 0.10
        card.line.color.rgb = hex_rgb("#FFFFFF")
        card.line.color.alpha = 0.45
        card.line.width = Pt(0.75)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(16)
        tf.margin_top = Pt(18)
        tf.text = amt
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(28)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["accent"])
        p1 = tf.add_paragraph()
        p1.text = body
        p1.space_before = Pt(8)
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(12)
        p1.font.color.rgb = hex_rgb("#FFFFFF")
    # No footer on closing-style slide.


if __name__ == "__main__":
    out = HERE / "_out" / "04_investor_pitch.pptx"
    out.parent.mkdir(exist_ok=True)
    build(out)
