"""02 — FY27–FY29 Annual Strategic Plan.

Three-year strategic plan presented to the executive committee.

Slides:
    1. Cover
    2. The story so far (KPIs)
    3. Where the market is going (line chart, TAM)
    4. Strategic priorities (3 pillars)
    5. North-star metrics (table)
    6. FY27 investment plan (bar chart)
    7. Operating model changes (bullets)
    8. Risks and mitigations (2-column)
    9. CEO commitment (quote)
    10. Closing
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from power_pptx.chart.data import CategoryChartData
from power_pptx.design.recipes import kpi_slide, quote_slide
from power_pptx.enum.chart import XL_CHART_TYPE
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from power_pptx.util import Inches, Pt

from _brand import STRATEGY as TOKENS, STRATEGY_PALETTE as PALETTE
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
    write_card_text,
)

FOOTER = "Strategic Plan FY27–FY29  |  Confidential — Executive Committee"


def build(out: Path) -> None:
    prs = new_deck()

    cover_slide(
        prs,
        eyebrow_text="Annual Strategic Plan",
        title="Three Horizons.\nOne Compounding Engine.",
        subtitle="Our path from $19B to $30B in revenue, anchored in "
                 "software-led platform expansion and durable margin.",
        presenter="Marcus Chen, Chief Executive Officer",
        date="Executive Committee Offsite  •  May 2026",
        tokens=TOKENS,
    )

    kpi_slide(
        prs, title="Where we are today (FY26)",
        kpis=[
            {"label": "Revenue",        "value": "$18.93B", "delta": +0.083},
            {"label": "Operating margin", "value": "26.0%",  "delta": +0.021},
            {"label": "Software ARR",   "value": "$4.6B",   "delta": +0.240},
            {"label": "Customer NPS",   "value": "62",      "delta": +0.135},
        ],
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER, "2", TOKENS)

    _market_outlook(prs)
    _three_pillars(prs)
    _north_star_table(prs)
    _investment_plan(prs)
    _operating_model(prs)
    _risks_and_mitigations(prs)

    quote_slide(
        prs,
        quote="Strategy without execution is decoration. We will measure "
              "ourselves quarterly against the same five north-star "
              "metrics from now until the end of FY29.",
        attribution="Marcus Chen, Chief Executive Officer",
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER, "9", TOKENS)

    closing_slide(
        prs,
        headline="From plan to motion.",
        sub="Functional plans due to COO by June 30, 2026.",
        tokens=TOKENS,
    )

    lint_or_die(prs)
    prs.save(out)
    print(f"wrote {out}")


def _market_outlook(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Market context", TOKENS)
    section_title(slide, "Industrial software TAM doubles by FY29", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["FY24", "FY25", "FY26", "FY27", "FY28", "FY29"]
    data.add_series("Industrial software TAM ($B)", (62, 71, 82, 96, 112, 131))
    data.add_series("Industrial hardware TAM ($B)", (148, 152, 157, 162, 167, 172))
    data.add_series("Aftermarket services TAM ($B)", (94, 98, 103, 108, 113, 119))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.6), Inches(1.7), Inches(8.4), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Addressable market ($B)"

    card = styled_card(slide, 9.3, 1.7, 3.5, 5.0, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    write_card_text(
        card,
        eyebrow_text="Implication",
        heading="Pivot mix to software",
        body=("If we hold today's mix, we capture a shrinking share of a "
              "growing market. We must shift 700 bps of revenue mix into "
              "software and recurring services by FY29."),
        tokens=TOKENS, heading_size=20,
    )
    footer(slide, FOOTER, "3", TOKENS)


def _three_pillars(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Strategy", TOKENS)
    section_title(slide, "Three pillars, one platform", TOKENS)
    divider(slide, TOKENS)

    pillars = [
        ("01", "Industrial Cloud",
         "Become the system-of-record for asset performance.",
         ["Unify Equipment + Software + Services data",
          "Open APIs to ERP and CMMS systems",
          "Customer 360 and predictive maintenance"]),
        ("02", "Outcome-based contracts",
         "Sell uptime and yield, not boxes.",
         ["Pilot 12 strategic accounts in FY27",
          "Move 20% of segment revenue to subscription by FY29",
          "Tie commercial terms to operational KPIs"]),
        ("03", "Geographic depth",
         "Win share in MENA, India, and Southeast Asia.",
         ["Localised manufacturing in Pune and Riyadh",
          "Tripled regional sales footprint",
          "Sovereign data residency by Q3 FY27"]),
    ]
    width = 4.05
    gap = 0.15
    left0 = 0.6
    for i, (num, head, sub, items) in enumerate(pillars):
        left = left0 + i * (width + gap)
        card = styled_card(slide, left, 1.7, width, 5.0,
                           tokens=TOKENS, fill_hex="#FFFFFF",
                           stroke_hex="#E5E7EB")

        # Number ribbon at top
        ribbon = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left), Inches(1.7), Inches(width), Inches(0.45),
        )
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = hex_rgb(PALETTE[i])
        ribbon.line.fill.background()
        rt = ribbon.text_frame
        rt.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        rt.margin_left = Pt(14)
        rt.text = f"PILLAR {num}"
        p = rt.paragraphs[0]
        p.font.name = TOKENS.typography["body"].family
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["on_primary"])

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(18)
        tf.margin_top = Pt(64)
        tf.margin_bottom = Pt(18)
        tf.text = head
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(22)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p1 = tf.add_paragraph()
        p1.text = sub
        p1.space_before = Pt(6)
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(13)
        p1.font.italic = True
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
        for it in items:
            pi = tf.add_paragraph()
            pi.text = f"•  {it}"
            pi.space_before = Pt(8)
            pi.font.name = TOKENS.typography["body"].family
            pi.font.size = Pt(13)
            pi.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])

    footer(slide, FOOTER, "4", TOKENS)


def _north_star_table(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Accountability", TOKENS)
    section_title(slide, "North-star metrics — FY29 commitments", TOKENS)
    divider(slide, TOKENS)

    rows = [
        ("Metric",                 "FY26",   "FY27 plan", "FY28 plan", "FY29 commit"),
        ("Revenue",                "$18.9B", "$20.5B",    "$23.0B",    "$30.0B"),
        ("Software ARR",           "$4.6B",  "$5.7B",     "$8.0B",     "$11.0B"),
        ("Software % of revenue",  "24%",    "28%",       "35%",       "37%"),
        ("Operating margin",       "26.0%",  "27.5%",     "29.0%",     "30.0%"),
        ("Customer NPS",           "62",     "65",        "68",        "70"),
    ]
    shape = slide.shapes.add_table(
        rows=len(rows), cols=5,
        left=Inches(0.6), top=Inches(1.7),
        width=Inches(12.1), height=Inches(4.2),
    )
    table = shape.table
    table.columns[0].width = Inches(3.7)
    for c in range(1, 5):
        table.columns[c].width = Inches(2.1)

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
            if c == 4:
                p.font.bold = True
                p.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            if r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["surface"])
            cell.borders.bottom.color.rgb = hex_rgb("#E5E7EB")
            cell.borders.bottom.width = Pt(0.5)

    note = slide.shapes.add_textbox(
        Inches(0.6), Inches(6.1), Inches(12.1), Inches(0.4),
    )
    tf = note.text_frame
    tf.text = "Reviewed quarterly by the Executive Committee. Anchored in our long-term incentive plan."
    p = tf.paragraphs[0]
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "5", TOKENS)


def _investment_plan(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Capital deployment", TOKENS)
    section_title(slide, "FY27 investment envelope: $1.85B", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = [
        "Industrial Cloud platform",
        "Outcome-based GTM",
        "International expansion",
        "Manufacturing automation",
        "Workforce upskilling",
        "Brand & demand gen",
    ]
    data.add_series("FY27 plan ($M)", (560, 380, 320, 280, 180, 130))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(8.6), Inches(5.0),
        data,
    ).chart
    series = chart.series[0]
    for point, hexc in zip(series.points, PALETTE):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = hex_rgb(hexc)
    chart.apply_quick_layout("title_no_legend")
    chart.chart_title.text_frame.text = "FY27 strategic investment ($M)"

    card = styled_card(slide, 9.5, 1.7, 3.3, 5.0, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    write_card_text(
        card,
        eyebrow_text="Funding",
        heading="60% reallocated, 40% incremental",
        body=("$1.1B of FY27 spend comes from sun-setting six legacy "
              "programs and tightening discretionary opex; the remainder "
              "is funded by retained free cash flow."),
        tokens=TOKENS, heading_size=18,
    )
    footer(slide, FOOTER, "6", TOKENS)


def _operating_model(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "How we work", TOKENS)
    section_title(slide, "Operating model — what changes in FY27", TOKENS)
    divider(slide, TOKENS)

    items = [
        ("One global P&L for software",
         "Industrial Software consolidates under a single GM with cross-segment authority."),
        ("Outcome squads, not project teams",
         "Multi-functional squads aligned to a customer outcome with named sponsors and durable funding."),
        ("Quarterly portfolio review",
         "30 strategic programs reviewed every 90 days; anything red two quarters in a row is killed or reset."),
        ("Engineering-led pricing",
         "Pricing committee chaired by the CTO, not the CFO — value-based, telemetry-informed."),
        ("Talent density over headcount",
         "Net-zero growth in non-engineering G&A; double engineering bar with internal mobility."),
    ]
    grid_top = 1.7
    row_h = 1.0
    for i, (head, body) in enumerate(items):
        # Number circle
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(0.7), Inches(grid_top + i * row_h + 0.05),
            Inches(0.55), Inches(0.55),
        )
        circ.fill.solid()
        circ.fill.fore_color.rgb = hex_rgb(TOKENS.palette["primary"])
        circ.line.fill.background()
        tf = circ.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        tf.text = str(i + 1)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["on_primary"])

        tb = slide.shapes.add_textbox(
            Inches(1.4), Inches(grid_top + i * row_h),
            Inches(11.4), Inches(0.95),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = head
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(17)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p1 = tf.add_paragraph()
        p1.text = body
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(12)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "7", TOKENS)


def _risks_and_mitigations(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Honest assessment", TOKENS)
    section_title(slide, "What could derail us — and how we'll respond", TOKENS)
    divider(slide, TOKENS)

    risks = [
        ("Talent attrition",
         "Loss of senior software engineering leaders during the platform rebuild.",
         "Retention grants for top 200; rotational program; engineering brand investment."),
        ("Customer transition",
         "Existing perpetual-license customers resist subscription conversion.",
         "5-year glide-path pricing; cloud-credit incentives; co-development for top 25 logos."),
        ("Geopolitical / trade",
         "Sanctions or tariffs disrupt component flow from one geography.",
         "Dual-source qualification target: 100% by Q4 FY27; 90 days of buffer inventory."),
        ("Execution capacity",
         "Pursuing all three pillars at once exceeds organisational bandwidth.",
         "Quarterly portfolio gate; ruthless killing of legacy programs; outcome-squad model."),
    ]
    col_w = 6.05
    for i, (title, risk, mitigation) in enumerate(risks):
        col = i % 2
        row = i // 2
        left = 0.6 + col * (col_w + 0.15)
        top = 1.7 + row * 2.6
        card = styled_card(slide, left, top, col_w, 2.4, tokens=TOKENS,
                           fill_hex="#FFFFFF", stroke_hex="#E5E7EB")
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(16)
        tf.margin_top = tf.margin_bottom = Pt(14)
        tf.text = title
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(16)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["negative"])
        p1 = tf.add_paragraph()
        p1.text = risk
        p1.space_before = Pt(4)
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(12)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p2 = tf.add_paragraph()
        p2.text = "Mitigation"
        p2.space_before = Pt(8)
        p2.font.name = TOKENS.typography["body"].family
        p2.font.size = Pt(10)
        p2.font.bold = True
        p2.font.color.rgb = hex_rgb(TOKENS.palette["accent"])
        p3 = tf.add_paragraph()
        p3.text = mitigation
        p3.font.name = TOKENS.typography["body"].family
        p3.font.size = Pt(11)
        p3.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "8", TOKENS)


if __name__ == "__main__":
    out = HERE / "_out" / "02_annual_strategic_plan.pptx"
    out.parent.mkdir(exist_ok=True)
    build(out)
