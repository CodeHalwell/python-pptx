"""Integration tests for the new table API surfaces:

* ``row.borders`` / ``col.borders`` shorthand
* ``Table.banded_rows`` / ``banded_cols`` aliases
* ``Table.fit_to_box``
* ``cell.text_frame.fit_text`` honoring cell bounds
"""

from __future__ import annotations

import pytest

from power_pptx import Presentation
from power_pptx.dml.color import RGBColor
from power_pptx.util import Inches, Pt


def _new_table(rows=2, cols=2, w=Inches(4), h=Inches(2)):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    gf = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), w, h)
    return prs, slide, gf.table


class DescribeRowBordersShorthand:
    def it_applies_a_bottom_border_to_every_cell_in_the_row(self):
        _, _, table = _new_table()
        table.rows[0].borders.bottom(width=Pt(2), color=RGBColor(0, 0, 0))
        for cell in table.rows[0].cells:
            assert cell.borders.bottom.width == Pt(2)

    def it_applies_outer_borders_to_every_cell_in_the_row(self):
        _, _, table = _new_table()
        table.rows[0].borders.outer(width=Pt(1), color=RGBColor(255, 0, 0))
        for cell in table.rows[0].cells:
            for edge in (cell.borders.left, cell.borders.right, cell.borders.top, cell.borders.bottom):
                assert edge.width == Pt(1)

    def it_clears_all_borders_in_the_row(self):
        _, _, table = _new_table()
        table.rows[0].borders.outer(width=Pt(2), color=RGBColor(0, 0, 0))
        # Sanity: the width was set.
        assert table.rows[0].cells[0].borders.bottom.width == Pt(2)
        table.rows[0].borders.none()
        # After clear, the underlying ``<a:ln*>`` elements are removed; the
        # LineFormat returns either ``None`` or a default of 0 EMU depending
        # on which path the proxy takes — what matters is the explicit value
        # from before is gone.
        for cell in table.rows[0].cells:
            for edge in (cell.borders.left, cell.borders.right, cell.borders.top, cell.borders.bottom):
                assert edge.width != Pt(2)


class DescribeColumnBordersShorthand:
    def it_applies_a_right_border_to_every_cell_in_the_column(self):
        _, _, table = _new_table()
        table.columns[0].borders.right(width=Pt(2), color=(0, 128, 0))
        # First column is the cells at col_idx=0 across all rows.
        for row in table.rows:
            assert row.cells[0].borders.right.width == Pt(2)


class DescribeBandedRowsAliases:
    def it_aliases_banded_rows_to_horz_banding(self):
        _, _, table = _new_table()
        table.banded_rows = True
        assert table.banded_rows is True
        assert table.horz_banding is True
        table.banded_rows = False
        assert table.horz_banding is False

    def it_aliases_banded_cols_to_vert_banding(self):
        _, _, table = _new_table()
        table.banded_cols = True
        assert table.vert_banding is True


class DescribeFitToBox:
    def it_keeps_max_size_when_text_already_fits(self):
        _, _, table = _new_table(2, 2, w=Inches(8), h=Inches(4))
        table.cell(0, 0).text = "short"
        table.cell(0, 1).text = "fine"
        size = table.fit_to_box(max_font_pt=18, min_font_pt=8)
        assert size == 18

    def it_shrinks_when_a_cell_overflows(self):
        # Tiny cells (1in × 0.25in) with way too much text.
        _, _, table = _new_table(2, 2, w=Inches(2), h=Inches(0.5))
        table.cell(0, 0).text = (
            "lorem ipsum dolor sit amet consectetur adipiscing elit"
        )
        table.cell(0, 1).text = "x"
        size = table.fit_to_box(max_font_pt=18, min_font_pt=6)
        assert 6 <= size < 18

    def it_clamps_to_min_font_pt(self):
        _, _, table = _new_table(2, 2, w=Inches(0.5), h=Inches(0.25))
        # Even one word won't fit at any size.
        table.cell(0, 0).text = "supercalifragilisticexpialidocious " * 5
        size = table.fit_to_box(max_font_pt=18, min_font_pt=8)
        assert size == 8

    def it_applies_target_size_to_every_cell(self):
        _, _, table = _new_table()
        table.cell(0, 0).text = "a"
        table.cell(0, 1).text = "b"
        table.cell(1, 0).text = "c"
        table.cell(1, 1).text = "d"
        target = table.fit_to_box(max_font_pt=14, min_font_pt=8)
        for cell in table.iter_cells():
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    assert run.font.size == Pt(target)

    def it_rejects_invalid_bounds(self):
        _, _, table = _new_table()
        with pytest.raises(ValueError):
            table.fit_to_box(max_font_pt=4, min_font_pt=8)
        with pytest.raises(ValueError):
            table.fit_to_box(min_font_pt=0)


class DescribeCellExtents:
    def it_exposes_cell_width_and_height(self):
        _, _, table = _new_table(2, 2, w=Inches(4), h=Inches(2))
        cell = table.cell(0, 0)
        # Each column = 2", each row = 1" by default split.
        assert int(cell.width) == int(Inches(2))
        assert int(cell.height) == int(Inches(1))


class DescribeCellFitText:
    def it_now_measures_against_cell_bounds_not_table_bounds(self):
        # Before this fix, ``cell.text_frame.fit_text`` measured against the
        # whole table; the result was meaningless. After: it measures
        # against the cell's own width/height.
        _, _, table = _new_table(1, 1, w=Inches(2), h=Inches(0.5))
        cell = table.cell(0, 0)
        cell.text = "lorem ipsum dolor sit amet consectetur adipiscing"
        cell.text_frame.fit_text(max_size=18)
        # The applied size must be < 18 because the text overflows the cell.
        size = cell.text_frame.paragraphs[0].runs[0].font.size
        assert size is not None
        assert size.pt < 18
