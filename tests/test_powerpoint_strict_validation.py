"""Deck-level "PowerPoint-strict" checks the OOXML schema doesn't catch.

Microsoft PowerPoint's open-time validator is stricter than the
OOXML schemas. The classic example is the area + doughnut chart
``<a:endParaRPr/>`` bug fixed in this branch — schema-valid (the
``lang`` attribute is optional) but rejected by PowerPoint.

Until a Windows CI runner with real PowerPoint roundtrip exists,
the next-best signal is to scan the saved deck's XML for the
specific strict-validator hooks we know about. The list grows as
we discover more.

Each test builds a deck containing the chart / shape kinds most
likely to trigger the relevant rule, saves it to bytes, and walks
every XML part looking for the stricter-than-spec patterns. A
failure here means a writer regressed.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from power_pptx import Presentation
from power_pptx.chart.data import CategoryChartData
from power_pptx.enum.chart import XL_CHART_TYPE
from power_pptx.util import Inches


def _deck_with_every_chart_type():
    prs = Presentation()
    types = [
        XL_CHART_TYPE.AREA,
        XL_CHART_TYPE.AREA_STACKED,
        XL_CHART_TYPE.AREA_STACKED_100,
        XL_CHART_TYPE.BAR_CLUSTERED,
        XL_CHART_TYPE.BAR_STACKED,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        XL_CHART_TYPE.COLUMN_STACKED,
        XL_CHART_TYPE.LINE,
        XL_CHART_TYPE.PIE,
        XL_CHART_TYPE.DOUGHNUT,
        XL_CHART_TYPE.RADAR,
    ]
    for chart_type in types:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        data = CategoryChartData()
        data.categories = ["A", "B", "C"]
        data.add_series("S1", (1.0, 2.0, 3.0))
        data.add_series("S2", (1.5, 1.8, 2.2))
        slide.shapes.add_chart(
            chart_type, Inches(1), Inches(1), Inches(6), Inches(4), data
        )
    return prs


def _save_to_bytes(prs):
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _walk_xml_parts(blob):
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        for info in zf.infolist():
            if not info.filename.endswith(".xml"):
                continue
            yield info.filename, zf.read(info.filename)


@pytest.fixture(scope="module")
def deck_blob():
    return _save_to_bytes(_deck_with_every_chart_type())


class DescribeEndParaRPrLang:
    """Issue 0 (post-2.5 review).

    PowerPoint's open-time validator rejects bare ``<a:endParaRPr/>``;
    every emission must carry a ``lang`` attribute even though the
    schema marks it optional.
    """

    def it_finds_no_bare_endParaRPr_anywhere_in_the_deck(self, deck_blob):
        bare = re.compile(rb"<a:endParaRPr\s*/>")
        for filename, body in _walk_xml_parts(deck_blob):
            assert not bare.search(body), (
                f"{filename} emits a bare <a:endParaRPr/>; PowerPoint "
                "will prompt to 'Repair' the file. Add lang=\"en-US\"."
            )

    def it_finds_lang_on_every_endParaRPr(self, deck_blob):
        with_attrs = re.compile(rb"<a:endParaRPr\b[^>]*>")
        for filename, body in _walk_xml_parts(deck_blob):
            for match in with_attrs.finditer(body):
                assert b"lang=" in match.group(0), (
                    f"{filename} emitted <a:endParaRPr> without lang: "
                    f"{match.group(0)!r}"
                )


class DescribeStrictHookList:
    """Reserve a slot for additional strict-validator rules we discover.

    Each of these is a regression-prevention check: if PowerPoint
    rejects something the schema accepts, encode that here so the
    next bug doesn't ship silently.
    """

    def it_runs_against_a_real_chart_deck(self, deck_blob):
        # Smoke check that the deck has the parts we think it does;
        # if zero charts were registered, the more-targeted tests
        # above would pass vacuously.
        chart_parts = [
            f for f, _ in _walk_xml_parts(deck_blob) if "/charts/chart" in f
        ]
        assert len(chart_parts) >= 11
