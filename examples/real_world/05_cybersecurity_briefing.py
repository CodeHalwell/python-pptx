"""05 — Cybersecurity Board Briefing.

Quarterly briefing the CISO delivers to the board's risk and audit
committee. Focuses on posture, incidents, and the path forward.

Slides:
    1. Cover
    2. Executive summary (KPIs)
    3. Threat landscape (column chart)
    4. Posture against NIST CSF (bar with conditional color)
    5. Incidents this quarter (table with severity coloring)
    6. Top 5 risks (numbered list)
    7. Investment plan (line chart of spend over time)
    8. People & culture (KPIs + bullets)
    9. Closing — asks of the board
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
from power_pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from power_pptx.util import Inches, Pt

from _brand import SECURITY as TOKENS, SECURITY_PALETTE as PALETTE
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

FOOTER = "Q4 FY26 Risk & Audit Committee  |  Restricted — Board"


def build(out: Path) -> None:
    prs = new_deck()

    cover_slide(
        prs,
        eyebrow_text="Cybersecurity Board Briefing",
        title="Posture, incidents,\nand the road to maturity 4.0.",
        subtitle="Northwind's quarterly assessment against the NIST CSF "
                 "and a candid view of the threat environment.",
        presenter="Lena Whitfield, Chief Information Security Officer",
        date="Q4 FY26  •  Risk & Audit Committee  •  February 18, 2026",
        tokens=TOKENS,
    )

    kpi_slide(
        prs,
        title="Executive summary",
        kpis=[
            {"label": "NIST CSF maturity",   "value": "3.4 / 4", "delta": +0.10},
            {"label": "Sev-1 incidents",     "value": "0",       "delta": -1.00},
            {"label": "Mean time to detect", "value": "9 min",   "delta": -0.42},
            {"label": "Phishing click rate", "value": "1.7%",    "delta": -0.36},
        ],
        tokens=TOKENS,
    )
    footer(prs.slides[-1], FOOTER, "2", TOKENS)

    _threat_landscape(prs)
    _nist_csf(prs)
    _incidents_table(prs)
    _top_risks(prs)
    _investment_plan(prs)
    _people_culture(prs)

    _board_asks(prs)

    lint_or_die(prs)
    prs.save(out)
    print(f"wrote {out}")


def _threat_landscape(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "External environment", TOKENS)
    section_title(slide, "Threats targeting our sector are up 41% YoY", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["Q1 FY25", "Q2 FY25", "Q3 FY25", "Q4 FY25",
                       "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"]
    data.add_series("Industry-wide attempts (k)",
                    (118, 132, 154, 168, 191, 210, 234, 248))
    data.add_series("Northwind blocked (k)",
                    (108, 123, 148, 161, 188, 207, 229, 244))
    data.add_series("Industry breaches (count)",
                    (24,  28,  31,  33,  41,  46,  53,  58))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(8.4), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Northwind threat volume vs. industry"

    card = styled_card(slide, 9.3, 1.7, 3.5, 5.0, tokens=TOKENS,
                       fill_hex=TOKENS.palette["surface"])
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(18)
    tf.margin_top = tf.margin_bottom = Pt(18)
    tf.text = "Most active actor groups"
    p = tf.paragraphs[0]
    p.font.name = TOKENS.typography["heading"].family
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
    actors = [
        ("Akira (criminal)",  "Ransomware, double-extortion."),
        ("Vanir (criminal)",  "Supply-chain compromise."),
        ("Gnomon (state-aligned)", "IP theft from R&D systems."),
        ("Insider (low)",     "Negligent misconfiguration."),
    ]
    for h, b in actors:
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
    footer(slide, FOOTER, "3", TOKENS)


def _nist_csf(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Posture", TOKENS)
    section_title(slide, "NIST CSF maturity — current vs. FY27 target", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["Identify", "Protect", "Detect", "Respond", "Recover", "Govern"]
    data.add_series("Current (FY26)", (3.6, 3.5, 3.7, 3.2, 2.9, 3.4))
    data.add_series("Target (FY27)",  (4.0, 4.0, 4.0, 4.0, 3.8, 4.0))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
        data,
    ).chart
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = hex_rgb(TOKENS.palette["primary"])
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = hex_rgb(TOKENS.palette["accent"])
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "NIST CSF maturity (1–4 scale, higher is better)"
    footer(slide, FOOTER, "4", TOKENS)


def _incidents_table(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Q4 FY26", TOKENS)
    section_title(slide, "Incident register — 6 events, all contained", TOKENS)
    divider(slide, TOKENS)

    rows = [
        ("Date",         "Severity", "Type",                      "Affected",      "MTTD",  "MTTR",  "Status"),
        ("Oct 21, 2025", "S2",       "Credential stuffing",       "Customer portal",   "11m",  "1h 8m", "Closed"),
        ("Nov 04, 2025", "S3",       "Phishing — finance",        "12 inboxes",        "6m",   "32m",   "Closed"),
        ("Nov 27, 2025", "S2",       "DDoS volumetric",           "Public marketing",  "3m",   "47m",   "Closed"),
        ("Dec 14, 2025", "S3",       "Misconfigured S3 bucket",   "Marketing assets",  "2h",   "4h",    "Closed"),
        ("Jan 09, 2026", "S2",       "Vendor compromise",         "SaaS analytics",    "9m",   "2h 3m", "Closed"),
        ("Jan 23, 2026", "S3",       "Lost device (encrypted)",   "1 endpoint",        "8m",   "22m",   "Closed"),
    ]
    shape = slide.shapes.add_table(
        rows=len(rows), cols=7,
        left=Inches(0.6), top=Inches(1.7),
        width=Inches(12.1), height=Inches(4.6),
    )
    table = shape.table
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(1.0)
    table.columns[2].width = Inches(2.6)
    table.columns[3].width = Inches(2.5)
    table.columns[4].width = Inches(1.3)
    table.columns[5].width = Inches(1.7)
    table.columns[6].width = Inches(1.5)

    for c, label in enumerate(rows[0]):
        cell = table.cell(0, c)
        cell.text = label
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["primary"])
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = hex_rgb(TOKENS.palette["on_primary"])
        cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    sev_color = {
        "S1": TOKENS.palette["negative"],
        "S2": TOKENS.palette["accent"],
        "S3": TOKENS.palette["positive"],
    }
    for r, row in enumerate(rows[1:], start=1):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
            if c == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_rgb(sev_color.get(value, "#94A3B8"))
                p.font.bold = True
                p.font.color.rgb = hex_rgb("#FFFFFF")
                p.alignment = PP_ALIGN.CENTER
            elif c == 6 and value == "Closed":
                p.font.color.rgb = hex_rgb(TOKENS.palette["positive"])
                p.font.bold = True
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            if r % 2 == 0 and c not in (1,):
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_rgb(TOKENS.palette["surface"])
            cell.borders.bottom.color.rgb = hex_rgb("#E5E7EB")
            cell.borders.bottom.width = Pt(0.5)

    note = slide.shapes.add_textbox(
        Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.5),
    )
    tf = note.text_frame
    tf.text = ("Zero S1 events. All S2/S3 contained within SLAs. "
               "No regulatory notification thresholds met.")
    p = tf.paragraphs[0]
    p.font.size = Pt(11)
    p.font.italic = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "5", TOKENS)


def _top_risks(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Risk register", TOKENS)
    section_title(slide, "Top 5 enterprise cyber risks — board view", TOKENS)
    divider(slide, TOKENS)

    risks = [
        ("Identity & access compromise",
         "5,800 SaaS apps; non-human identities outnumber humans 12:1.",
         "High",   "#DC2626"),
        ("Ransomware via vendor",
         "Third-party SaaS providers hold or process 41% of regulated data.",
         "High",   "#DC2626"),
        ("AI / data leakage",
         "Sanctioned and shadow LLM usage continues to expand.",
         "Medium", "#F59E0B"),
        ("OT / plant systems",
         "Legacy industrial protocols; 18% of plant assets still on EOL OS.",
         "Medium", "#F59E0B"),
        ("Insider — negligent",
         "Misconfiguration remains our most common root cause.",
         "Medium", "#F59E0B"),
    ]
    grid_top = 1.8
    row_h = 1.0
    for i, (head, body, level, color) in enumerate(risks):
        # Number circle
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(0.7), Inches(grid_top + i * row_h + 0.1),
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

        # Heading + body
        tb = slide.shapes.add_textbox(
            Inches(1.4), Inches(grid_top + i * row_h),
            Inches(9.5), Inches(0.95),
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

        # Severity pill
        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(11.0), Inches(grid_top + i * row_h + 0.18),
            Inches(1.6), Inches(0.4),
        )
        pill.adjustments[0] = 0.5
        pill.fill.solid()
        pill.fill.fore_color.rgb = hex_rgb(color)
        pill.line.fill.background()
        tf = pill.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = Pt(2)
        tf.text = level.upper()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = hex_rgb("#FFFFFF")
    footer(slide, FOOTER, "6", TOKENS)


def _investment_plan(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "Roadmap to maturity 4.0", TOKENS)
    section_title(slide, "Three-year cyber investment plan: $268M", TOKENS)
    divider(slide, TOKENS)

    data = CategoryChartData()
    data.categories = ["FY26", "FY27", "FY28", "FY29"]
    data.add_series("Identity",      (18, 24, 28, 30))
    data.add_series("Detection",     (16, 22, 24, 24))
    data.add_series("Response",      (10, 14, 18, 20))
    data.add_series("OT/IoT",        (6,  10, 16, 22))
    data.add_series("Awareness",     (4,  5,  6,  7))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
        data,
    ).chart
    chart.apply_palette(PALETTE)
    chart.apply_quick_layout("title_legend_bottom")
    chart.chart_title.text_frame.text = "Annual cyber investment by capability ($M)"
    footer(slide, FOOTER, "7", TOKENS)


def _people_culture(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(slide, "People", TOKENS)
    section_title(slide, "Culture: phishing failures down 36%, awareness up", TOKENS)
    divider(slide, TOKENS)

    # Left: KPI cards
    kpis = [
        ("1.7%",  "Q4 phishing click rate"),
        ("96%",   "Mandatory training completion"),
        ("142",   "Internal reporters per 1,000 emp."),
        ("28%",   "Reduction in repeat clickers"),
    ]
    for i, (num, label) in enumerate(kpis):
        col = i % 2
        row = i // 2
        left = 0.6 + col * 3.05
        top = 1.85 + row * 1.7
        card = styled_card(slide, left, top, 2.95, 1.55,
                           tokens=TOKENS, fill_hex="#FFFFFF",
                           stroke_hex="#E5E7EB")
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(14)
        tf.margin_top = Pt(12)
        tf.text = num
        p0 = tf.paragraphs[0]
        p0.font.name = TOKENS.typography["heading"].family
        p0.font.size = Pt(28)
        p0.font.bold = True
        p0.font.color.rgb = hex_rgb(TOKENS.palette["primary"])
        p1 = tf.add_paragraph()
        p1.text = label
        p1.space_before = Pt(2)
        p1.font.size = Pt(11)
        p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])

    # Right: bullets
    box = slide.shapes.add_textbox(
        Inches(7.0), Inches(1.85), Inches(5.7), Inches(4.6),
    )
    tf = box.text_frame
    tf.word_wrap = True
    items = [
        ("Phishing simulation cadence", "Every 30 days, 4 difficulty tiers."),
        ("Just-in-time coaching",       "Repeat clickers receive 5-minute reset training."),
        ("Champions network",           "208 cyber-champions across BUs and geographies."),
        ("Executive tabletop",          "Quarterly C-suite simulation; next session: April 9."),
    ]
    tf.text = items[0][0]
    p0 = tf.paragraphs[0]
    p0.font.name = TOKENS.typography["heading"].family
    p0.font.size = Pt(15)
    p0.font.bold = True
    p0.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
    p1 = tf.add_paragraph()
    p1.text = items[0][1]
    p1.font.size = Pt(12)
    p1.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    p1.space_after = Pt(8)
    for h, b in items[1:]:
        p = tf.add_paragraph()
        p.text = h
        p.space_before = Pt(8)
        p.font.name = TOKENS.typography["heading"].family
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = hex_rgb(TOKENS.palette["neutral"])
        p2 = tf.add_paragraph()
        p2.text = b
        p2.font.size = Pt(12)
        p2.font.color.rgb = hex_rgb(TOKENS.palette["muted"])
    footer(slide, FOOTER, "8", TOKENS)


def _board_asks(prs) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5),
    )
    bg.fill.linear_gradient(
        TOKENS.palette["primary"], TOKENS.palette["neutral"], angle=135,
    )
    bg.line.fill.background()

    eb = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.4))
    tf = eb.text_frame
    tf.text = "ASKS OF THE BOARD"
    p = tf.paragraphs[0]
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = hex_rgb(TOKENS.palette["accent"])

    title = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.7), Inches(1.6))
    tf = title.text_frame
    tf.word_wrap = True
    tf.text = "Three decisions for today's session."
    tf.fit_text(font_family=TOKENS.typography["heading"].family,
                max_size=42, bold=True)
    tf.paragraphs[0].font.color.rgb = hex_rgb("#FFFFFF")

    asks = [
        ("01", "Approve the FY27 cyber budget at $98M (+15% YoY)."),
        ("02", "Endorse the OT modernization program — 18-month plan."),
        ("03", "Reaffirm the cyber-incident disclosure policy and "
               "delegation thresholds."),
    ]
    for i, (num, body) in enumerate(asks):
        top = 3.6 + i * 1.05
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(0.8), Inches(top), Inches(0.6), Inches(0.6),
        )
        circ.fill.solid()
        circ.fill.fore_color.rgb = hex_rgb(TOKENS.palette["accent"])
        circ.line.fill.background()
        tf = circ.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        tf.text = num
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = hex_rgb("#FFFFFF")

        tb = slide.shapes.add_textbox(
            Inches(1.6), Inches(top + 0.05), Inches(11.0), Inches(1.0),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = body
        tf.fit_text(font_family=TOKENS.typography["body"].family,
                    max_size=20, bold=False)
        tf.paragraphs[0].font.color.rgb = hex_rgb("#FFFFFF")


if __name__ == "__main__":
    out = HERE / "_out" / "05_cybersecurity_briefing.pptx"
    out.parent.mkdir(exist_ok=True)
    build(out)
