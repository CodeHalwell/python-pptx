"""01 — Q4 FY26 Earnings Review.

A board-level financial review you'd actually present to a
Fortune 500 audit committee or investor relations call.

Slides:
    1. Cover
    2. Agenda
    3. Q4 financial summary (KPIs)
    4. Revenue trend (column chart)
    5. Segment performance (bar chart)
    6. Margin & opex (line chart)
    7. Capital allocation (pie)
    8. Forward guidance (table)
    9. Risk factors (bullets)
    10. CEO commentary (quote)
    11. Closing
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from power_pptx.chart.data import CategoryChartData
from power_pptx.design.recipes import bullet_slide, kpi_slide, quote_slide
from power_pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from power_pptx.util import Inches, Pt

from _brand import EARNINGS as TOKENS, EARNINGS_PALETTE as PALETTE
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

FOOTER_LEFT = "Confidential  |  Northwind Industries, Inc.  |  Q4 FY26"


def build(out: Path) -> None:
    prs = new_deck()

    # 1. Cover
    cover_slide(
        prs,
        eyebrow_text="Q4 FY26 Earnings Review",
        title="Disciplined growth.\nDurable margin.",
        subtitle="A review of fourth-quarter and full-year operating "
                 "results, capital deployment, and FY27 outlook.",
        presenter="Helen Vance, Chief Financial Officer",
        date="February 12, 2026  •  Investor Relations Call",
        tokens=TOKENS,
    )

    # 2. Agenda
    _agenda(prs)

    # 3. KPI summary
    kpi_slide(
        prs,
        title="Q4 FY26 — at a glance",
        kpis=[
            {"label": "Revenue",        "value": "$4.82B", "delta": +0.094},
            {"label": "Operating margin", "value": "27.1%", "delta": +0.021},
            {"label": "Free cash flow", "value": "$1.13B", "delta": +0.166},
            {"label": "Diluted EPS",    "value": "$3.42",  "delta": +0.118},
        ],
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER_LEFT, "3", TOKENS)

    # 4. Revenue trend
    _revenue_trend(prs)

    # 5. Segment performance
    _segment_performance(prs)

    # 6. Margin & opex
    _margin_opex(prs)

    # 7. Capital allocation
    _capital_allocation(prs)

    # 8. Forward guidance table
    _guidance_table(prs)

    # 9. Risk factors
    bullet_slide(
        prs,
        title="Risk factors we are actively managing",
        bullets=[
            "FX volatility — 38% of revenue is non-USD; 60% hedged through Q2 FY27.",
            "Customer concentration — top 10 logos represent 22% of ARR (down from 27%).",
            "Component supply — second-source qualified for the two critical SKUs.",
            "Regulatory — EU AI Act compliance program on track for August 2026.",
            "Cyber — no material incidents in Q4; SOC 2 Type II re-certified.",
        ],
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER_LEFT, "9", TOKENS)

    # 10. Quote
    quote_slide(
        prs,
        quote="We exited the year with the strongest balance sheet in "
              "the company's history and a healthier mix of recurring "
              "revenue than at any point in the past decade.",
        attribution="Marcus Chen, Chief Executive Officer",
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER_LEFT, "10", TOKENS)

    # 11. Closing
    closing_slide(
        prs,
        headline="Questions & Answers",
        sub="investor.relations@northwind.com  |  ir.northwind.com",
        tokens=TOKENS,
    )

    lint_or_die(prs)
    prs.save(out)
    print(f"wrote {out}")


def _agenda(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Today's session", TOKENS)
    section_title(slide, "Agenda", TOKENS)
    divider(slide, TOKENS)

    items = [
        ("01", "Quarterly summary",  "Top-line, margin, cash flow."),
        ("02", "Segment results",    "Industrial, Software, Services."),
        ("03", "Capital allocation", "Buybacks, dividend, M&A."),
        ("04", "FY27 guidance",      "Revenue, margin, capex envelope."),
        ("05", "Risk landscape",     "FX, supply, regulatory, cyber."),
    ]
    for i, (num, head, body) in enumerate(items):
        # Number circle
        num_card = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(0.7), Inches(1.7 + i * 1.0),
            Inches(0.7), Inches(0.7),
        )
        num_card.fill.solid()
        num_card.fill.fore_color.rgb = hex_rgb(TOKENS.palette["primary"])
        num_card.line.fill.background()
        tf = num_card.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        tf.text = num
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = TOKENS.typography["heading"].family
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["on_primary"])

        # Heading text
        tb = slide.shapes.add_textbox(
            Inches(1.6), Inches(1.7 + i * 1.0), Inches(11), Inches(0.95),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = head
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(20)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p1 = tf.add_paragraph()
        p1.text = body
        p1.font.name = TOKENS.typography["body"].family
        p1.font.size = Pt(13)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])

    footer(slide, FOOTER_LEFT, "2", TOKENS)


def _revenue_trend(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Top-line growth", TOKENS)
    section_title(slide, "Revenue rose 9.4% YoY in Q4", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("FY25", (4180, 4290, 4380, 4408))
    data.add_series("FY26", (4520, 4660, 4720, 4822))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(8.5), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Revenue ($M)"

    # Right-rail highlights
    card = styled_card(slide, 9.4, 1.7, 3.4, 5.0, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    write_card_text(
        card,
        eyebrow_text="Drivers",
        heading="Software ARR +24%",
        body=("Industrial Software crossed $1B in run-rate ARR for "
              "the first time, contributing 290 bps of total growth."),
        tokens=TOKENS,
    )
    footer(slide, FOOTER_LEFT, "4", TOKENS)


def _segment_performance(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Segment view", TOKENS)
    section_title(slide, "Software led growth; Industrial held margin", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = [
        "Industrial Equipment",
        "Industrial Software",
        "Aftermarket Services",
        "Energy Solutions",
    ]
    data.add_series("Revenue ($M)", (1980, 1240, 920, 682))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(8.5), Inches(5.0),
        data,
    ).chart
    series = chart.series[0]
    for point, hexc in zip(series.points, PALETTE):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = hex_rgb(hexc)
    chart.apply_quick_layout("title_no_legend")
    chart.chart_title.text_frame.text = "Q4 segment revenue ($M)"

    card = styled_card(slide, 9.4, 1.7, 3.4, 2.4, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    write_card_text(
        card,
        eyebrow_text="Highlight",
        heading="+24% ARR",
        body="Software set a record with 24% ARR growth and 92% GRR.",
        tokens=TOKENS, heading_size=20,
    )
    card2 = styled_card(slide, 9.4, 4.3, 3.4, 2.4, tokens=TOKENS,
                        fill_hex=TOKENS.palette["surface"])
    write_card_text(
        card2,
        eyebrow_text="Watch",
        heading="Energy −3% YoY",
        body="Lower oil-and-gas capex; offset by service backlog +18%.",
        tokens=TOKENS, heading_size=20,
    )
    footer(slide, FOOTER_LEFT, "5", TOKENS)


def _margin_opex(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Profitability", TOKENS)
    section_title(slide, "Operating margin expanded 210 bps in FY26", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("Gross margin (%)", (47.2, 47.9, 48.6, 49.1))
    data.add_series("Operating margin (%)", (24.0, 25.5, 26.4, 27.1))
    data.add_series("Free cash flow margin (%)", (16.8, 18.0, 19.2, 23.4))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Margin progression FY26 (%)"
    footer(slide, FOOTER_LEFT, "6", TOKENS)


def _capital_allocation(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Capital allocation", TOKENS)
    section_title(slide, "$3.6B returned to shareholders in FY26", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["Buybacks", "Dividends", "Tuck-in M&A", "R&D capex", "Debt paydown"]
    data.add_series("FY26 ($M)", (1820, 940, 410, 280, 150))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(0.6), Inches(1.7), Inches(7.0), Inches(5.0),
        data,
    ).chart
    series = chart.series[0]
    for point, hexc in zip(series.points, PALETTE):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = hex_rgb(hexc)
    chart.apply_quick_layout({
        "has_title": True,
        "title_text": "Capital deployed FY26 ($M)",
        "has_legend": True,
        "legend_position": XL_LEGEND_POSITION.RIGHT,
    })

    card = styled_card(slide, 8.0, 1.8, 4.8, 4.8, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(20)
    tf.margin_top = tf.margin_bottom = Pt(18)
    tf.text = "Capital framework"
    p0 = tf.paragraphs[0]
    p0.font.name = TOKENS.typography["heading"].family
    p0.font.size = Pt(20)
    p0.font.bold = True
    p0.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
    points = [
        ("Buybacks", "Opportunistic, $4B authorization remaining."),
        ("Dividend", "Raised 8% — 14th consecutive annual increase."),
        ("M&A", "Disciplined tuck-ins; ROIC > WACC + 400 bps."),
        ("Balance sheet", "Net leverage 1.1x; investment-grade BBB+."),
    ]
    for head, body in points:
        p = tf.add_paragraph()
        p.text = head
        p.space_before = Pt(10)
        p.font.name = TOKENS.typography["body"].family
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.name = TOKENS.typography["body"].family
        p2.font.size = Pt(11)
        p2.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER_LEFT, "7", TOKENS)


def _guidance_table(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "FY27 outlook", TOKENS)
    section_title(slide, "Guidance: balanced growth, continued margin lift", TOKENS)
    divider(slide, TOKENS)

    rows = [
        ("Metric", "FY26 actual", "FY27 guidance", "Implied YoY"),
        ("Revenue",            "$18.93B", "$20.4–20.8B", "+8% – +10%"),
        ("Operating margin",   "26.0%",   "27.0–27.5%",  "+100–150 bps"),
        ("Diluted EPS",        "$12.86",  "$14.10–14.40", "+10% – +12%"),
        ("Free cash flow",     "$3.81B",  "$4.20–4.40B",  "+10% – +15%"),
        ("Capex (% revenue)",  "3.4%",    "3.0–3.5%",     "Within range"),
        ("Effective tax rate", "21.6%",   "21.0–22.0%",   "Stable"),
    ]
    shape = slide.shapes.add_table(
        rows=len(rows), cols=4,
        left=Inches(0.6), top=Inches(1.7),
        width=Inches(12.1), height=Inches(4.5),
    )
    table = shape.table
    table.columns[0].width = Inches(3.5)
    table.columns[1].width = Inches(2.8)
    table.columns[2].width = Inches(3.0)
    table.columns[3].width = Inches(2.8)

    primary = hex_rgb(TOKENS.palette["primary"])
    on_primary = hex_rgb(TOKENS.palette["on_primary"])
    light = hex_rgb("#E5E7EB")
    body_color = hex_rgb(TOKENS.palette["neutral"])

    for c, label in enumerate(rows[0]):
        cell = table.cell(0, c)
        cell.text = label
        cell.fill.solid()
        cell.fill.fore_color.rgb = primary
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.name = TOKENS.typography["body"].family
        p.font.color.rgb = on_primary
        cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        cell.borders.bottom.color.rgb = primary
        cell.borders.bottom.width = Pt(1.5)

    for r, row in enumerate(rows[1:], start=1):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(12)
            p.font.name = TOKENS.typography["body"].family
            p.font.color.rgb = body_color
            if c == 0:
                p.font.bold = True
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            if r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["surface"])
            cell.borders.bottom.color.rgb = light
            cell.borders.bottom.width = Pt(0.5)

    note = slide.shapes.add_textbox(
        Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.5),
    )
    tf = note.text_frame
    tf.text = ("Guidance assumes USD relatively stable vs. FY26 average; "
               "excludes impact of unannounced M&A.")
    p = tf.paragraphs[0]
    p.font.name = TOKENS.typography["body"].family
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER_LEFT, "8", TOKENS)


if __name__ == "__main__":
    out = HERE / "_out" / "01_q4_earnings_review.pptx"
    out.parent.mkdir(exist_ok=True)
    build(out)
