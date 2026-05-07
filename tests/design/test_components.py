"""Unit tests for :mod:`power_pptx.design.components`."""

from __future__ import annotations

import pytest

from power_pptx import Presentation
from power_pptx.design.components import (
    KpiCard,
    ProgressBar,
    add_kpi_card,
    add_progress_bar,
)
from power_pptx.design.tokens import DesignTokens
from power_pptx.dml.color import RGBColor
from power_pptx.util import Inches


@pytest.fixture
def slide():
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


@pytest.fixture
def tokens():
    return DesignTokens.from_dict(
        {
            "palette": {
                "primary": "#3C2F80",
                "neutral": "#222222",
                "accent": "#FF6600",
                "muted": "#777777",
                "surface": "#F5F5F8",
                "positive": "#119922",
                "negative": "#DD2233",
            }
        }
    )


class DescribeAddKpiCard:
    def it_returns_a_KpiCard_with_card_value_and_label_shapes(self, slide, tokens):
        result = add_kpi_card(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(2),
            height=Inches(2),
            label="ARR",
            value="$182M",
            tokens=tokens,
        )
        assert isinstance(result, KpiCard)
        assert result.card is not None
        assert result.value_box is not None
        assert result.label_box is not None
        # No delta dict supplied → no delta box.
        assert result.delta_box is None

    def it_renders_a_delta_when_supplied(self, slide, tokens):
        result = add_kpi_card(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(2),
            height=Inches(2),
            label="ARR",
            value="$182M",
            delta={"delta": 0.27},
            tokens=tokens,
        )
        assert result.delta_box is not None
        # The fraction 0.27 renders as "+27%".
        rendered = result.delta_box.text_frame.paragraphs[0].runs[0].text
        assert "27" in rendered

    def it_tags_constituent_shapes_with_a_lint_group(self, slide, tokens):
        result = add_kpi_card(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(2),
            height=Inches(2),
            label="ARR",
            value="$182M",
            delta={"delta": 0.27},
            tokens=tokens,
        )
        groups = {
            result.card.lint_group,
            result.value_box.lint_group,
            result.label_box.lint_group,
            result.delta_box.lint_group,
        }
        assert len(groups) == 1
        assert next(iter(groups)).startswith("kpi_card@")

    def it_works_without_tokens(self, slide):
        # Should fall back to PowerPoint defaults rather than raise.
        result = add_kpi_card(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(2),
            height=Inches(2),
            label="ARR",
            value="$182M",
        )
        assert isinstance(result, KpiCard)


class DescribeAddProgressBar:
    def it_creates_a_track_and_fill_shape(self, slide, tokens):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=0.5,
            tokens=tokens,
        )
        assert isinstance(bar, ProgressBar)
        assert bar.track is not None
        assert bar.fill is not None

    def it_sizes_the_fill_proportional_to_fraction(self, slide, tokens):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=0.25,
            tokens=tokens,
        )
        # 25% of 4" (allow ±1 EMU rounding tolerance).
        assert abs(int(bar.fill.width) - int(Inches(1))) <= 1
        assert int(bar.track.width) == int(Inches(4))

    def it_clamps_out_of_range_fractions(self, slide, tokens):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=1.5,  # > 1.0 — clamp to 1.0
            tokens=tokens,
        )
        assert int(bar.fill.width) == int(Inches(4))

    def it_emits_a_zero_width_fill_at_zero_fraction(self, slide, tokens):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=0.0,
            tokens=tokens,
        )
        # Zero fill but the shape is still emitted so callers can
        # mutate / animate it later.
        assert int(bar.fill.width) == 0

    def it_honours_explicit_fill_color(self, slide):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=0.5,
            fill_color="#FF0000",
        )
        assert bar.fill.fill.fore_color.rgb == RGBColor(0xFF, 0x00, 0x00)

    def it_tags_track_and_fill_with_a_lint_group(self, slide, tokens):
        bar = add_progress_bar(
            slide,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(0.3),
            fraction=0.5,
            tokens=tokens,
        )
        assert bar.track.lint_group == bar.fill.lint_group
        assert bar.track.lint_group.startswith("progress_bar@")
