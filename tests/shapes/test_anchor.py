"""Integration tests for ``anchor=`` on add_picture / add_shape / add_textbox.

The anchor keyword places the new shape relative to a container (the
slide by default, or a parent shape). It collapses the
``add → measure → reposition`` pattern that authors otherwise repeat
on every corner-anchored element (logos, watermarks, page numbers).
"""

from __future__ import annotations

import pytest

from power_pptx import Presentation
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.shapes.shapetree import (
    _compute_anchor_left_top,
    _resolve_anchor,
)
from power_pptx.util import Emu, Inches


def _slide():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    return prs, slide


class DescribeResolveAnchor:
    @pytest.mark.parametrize(
        "raw, expected",
        [
            ("top-left", ("top", "left")),
            ("top-center", ("top", "center")),
            ("top-centre", ("top", "center")),
            ("top-right", ("top", "right")),
            ("middle-left", ("middle", "left")),
            ("middle-center", ("middle", "center")),
            ("middle-right", ("middle", "right")),
            ("bottom-left", ("bottom", "left")),
            ("bottom-center", ("bottom", "center")),
            ("bottom-right", ("bottom", "right")),
            ("center", ("middle", "center")),
            ("centre", ("middle", "center")),
            ("CENTER", ("middle", "center")),
            ("center-left", ("middle", "left")),  # tolerant synonym
        ],
    )
    def it_parses_valid_anchor_strings(self, raw, expected):
        assert _resolve_anchor(raw) == expected

    @pytest.mark.parametrize("bad", ["", "weird", "left-top", "top", "north-west"])
    def it_rejects_invalid_anchor_strings(self, bad):
        with pytest.raises(ValueError):
            _resolve_anchor(bad)


class DescribeComputeAnchorLeftTop:
    def it_places_top_left(self):
        # Slide 10x7.5", shape 1x0.5", margin 0.25" → top-left.
        left, top = _compute_anchor_left_top(
            "top-left",
            container_w=int(Inches(10)),
            container_h=int(Inches(7.5)),
            shape_w=int(Inches(1)),
            shape_h=int(Inches(0.5)),
            margin=int(Inches(0.25)),
        )
        assert left == int(Inches(0.25))
        assert top == int(Inches(0.25))

    def it_places_bottom_right(self):
        left, top = _compute_anchor_left_top(
            "bottom-right",
            container_w=int(Inches(10)),
            container_h=int(Inches(7.5)),
            shape_w=int(Inches(1)),
            shape_h=int(Inches(0.5)),
            margin=int(Inches(0.25)),
        )
        # Slide is 10" wide; shape 1"; right margin 0.25" → left = 8.75".
        assert left == int(Inches(10)) - int(Inches(0.25)) - int(Inches(1))
        assert top == int(Inches(7.5)) - int(Inches(0.25)) - int(Inches(0.5))

    def it_centres_on_both_axes(self):
        left, top = _compute_anchor_left_top(
            "center",
            container_w=int(Inches(10)),
            container_h=int(Inches(8)),
            shape_w=int(Inches(2)),
            shape_h=int(Inches(2)),
        )
        assert left == int(Inches(4))
        assert top == int(Inches(3))


class DescribeAnchorOnAddShape:
    def it_repositions_to_bottom_right_after_creation(self):
        prs, slide = _slide()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(2),
            Inches(1),
            anchor="bottom-right",
            margin=Inches(0.25),
        )
        # Bottom-right with 0.25" margin in a 10x7.5" slide for a 2x1" shape:
        assert int(shape.left) == int(Inches(10) - Inches(0.25) - Inches(2))
        assert int(shape.top) == int(Inches(7.5) - Inches(0.25) - Inches(1))

    def it_centres_a_shape_inside_a_parent_shape_container(self):
        # A common pattern: drop a label into the centre of a card.
        prs, slide = _slide()
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1),
            Inches(1),
            Inches(4),
            Inches(2),
        )

        class _BoxLike:
            # Stand-in for a parent-shape-with-coords; the helper only
            # uses .width/.height.
            width = card.width
            height = card.height

        label = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(1),
            Inches(0.4),
            anchor="center",
            container=_BoxLike(),
        )
        # The label is positioned in the *container's* coordinate
        # system (0..card.width), so left/top are in card-local EMU.
        assert int(label.left) == (int(card.width) - int(Inches(1))) // 2
        assert int(label.top) == (int(card.height) - int(Inches(0.4))) // 2

    def it_leaves_position_alone_when_anchor_is_none(self):
        prs, slide = _slide()
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(2),
            Inches(3),
            Inches(1),
            Inches(1),
        )
        assert int(shape.left) == int(Inches(2))
        assert int(shape.top) == int(Inches(3))

    def it_rejects_unknown_anchor_strings(self):
        prs, slide = _slide()
        with pytest.raises(ValueError):
            slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0),
                Inches(0),
                Inches(1),
                Inches(1),
                anchor="south-west",
            )


class DescribeAnchorOnAddTextbox:
    def it_anchors_a_textbox_to_top_center_of_the_slide(self):
        prs, slide = _slide()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        tb = slide.shapes.add_textbox(
            Inches(0),
            Inches(0),
            Inches(2),
            Inches(0.5),
            anchor="top-center",
            margin=Inches(0.5),
        )
        assert int(tb.left) == (int(Inches(10)) - int(Inches(2))) // 2
        # margin only applies on top axis, not the centred horizontal.
        assert int(tb.top) == int(Inches(0.5))


class DescribeAnchorOnAddPicture:
    def it_anchors_a_picture_to_bottom_right_of_the_slide(self):
        prs, slide = _slide()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        pic = slide.shapes.add_picture(
            "tests/test_files/python-powered.png",
            Inches(0),
            Inches(0),
            height=Inches(0.5),
            anchor="bottom-right",
            margin=Inches(0.25),
        )
        # Picture's rendered width depends on aspect ratio; whatever it
        # is, the right edge must sit at slide_w - margin.
        assert int(pic.left + pic.width) == int(Inches(10) - Inches(0.25))
        assert int(pic.top + pic.height) == int(Inches(7.5) - Inches(0.25))
