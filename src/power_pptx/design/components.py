"""Shape-level building blocks layered on top of the slide recipes.

The :mod:`power_pptx.design.recipes` module covers whole-slide layouts
(``title_slide``, ``kpi_slide``, …); this module exposes the components
those recipes are built from so callers can compose mixed layouts with
brand-consistent components.

Two public callables today, both intentionally small:

* :func:`add_kpi_card` — a single KPI tile (label + headline value +
  optional delta), styled from a :class:`DesignTokens` instance.
* :func:`add_progress_bar` — a track + fill rounded-rectangle pair
  representing a 0..1 fraction.

Both accept an optional ``tokens`` argument that drives palette and
typography. When omitted, sensible defaults derived from the chart
palette are used. Each component is created using the existing
``slide.shapes.add_*`` primitives and returns a small dataclass
exposing the constituent shapes — callers can reach into ``.card`` /
``.value_box`` / ``.fill`` / ``.track`` for further per-deck tweaks
without re-implementing the layout.

The intentional shape stacking inside these components (label box on
top of card, fill bar on top of track) is tagged with
``shape.lint_group`` so :mod:`power_pptx.lint` does not flag them as
overlap warnings.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Mapping, Optional

from power_pptx.design.tokens import DesignTokens
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from power_pptx.util import Inches, Length, Pt

# Re-use the private helpers from recipes rather than reimplement them.
# Same package, same module-private convention.
from power_pptx.design.recipes import (
    _apply_card_styling,
    _delta_color,
    _fill_text_frame,
    _palette,
    _resolve_delta,
    _typography,
)

if TYPE_CHECKING:
    from power_pptx.shapes.autoshape import Shape
    from power_pptx.slide import Slide


__all__ = ("KpiCard", "ProgressBar", "add_kpi_card", "add_progress_bar")


@dataclass
class KpiCard:
    """Bundle of shapes produced by :func:`add_kpi_card`.

    Use ``card`` for global tweaks (border colour, fill, shadow), the
    text boxes for typography or content edits.
    """

    card: Any
    value_box: Any
    label_box: Any
    delta_box: Optional[Any] = None


@dataclass
class ProgressBar:
    """Bundle of shapes produced by :func:`add_progress_bar`.

    ``track`` is the full-width background; ``fill`` is the
    proportionally-sized foreground.
    """

    track: Any
    fill: Any


def add_kpi_card(
    slide: "Slide",
    *,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    label: str,
    value: str,
    delta: Optional[Mapping[str, Any]] = None,
    tokens: Optional[DesignTokens] = None,
) -> KpiCard:
    """Add a single KPI card (label + value + optional delta) to *slide*.

    `delta`, when supplied, is a mapping with the same shape consumed
    by :func:`power_pptx.design.recipes.kpi_slide` — ``{"delta": 0.27}``
    renders as ``+27%``, ``{"delta_text": "+14 pts"}`` renders verbatim.
    Tinted from the palette's ``positive`` / ``negative`` slot.

    Returns a :class:`KpiCard` bundle so callers can reach into the
    constituent shapes for per-deck tweaks without re-implementing
    the layout.
    """
    fill_color = _palette(tokens, ("surface", "lt2"))
    border_color = _palette(tokens, ("muted", "lt1"))
    value_color = _palette(tokens, ("primary", "neutral"))
    label_color = _palette(tokens, ("muted",))

    value_token = _typography(tokens, "heading", default_size=Pt(30), default_bold=True)
    label_token = _typography(tokens, "body", default_size=Pt(12))
    delta_token = _typography(tokens, "body", default_size=Pt(11), default_bold=True)

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    if fill_color is not None:
        card.fill.solid()
        card.fill.fore_color.rgb = fill_color
    else:
        card.fill.background()
    if border_color is not None:
        card.line.color.rgb = border_color
        card.line.width = Pt(0.75)
    _apply_card_styling(card, tokens)
    card.text_frame.text = ""

    # Vertical layout inside the card. Heights chosen to mirror
    # kpi_slide's existing recipe so the visual matches when used in
    # the same deck.
    value_h = Inches(0.85)
    label_h = Inches(0.4)
    delta_h = Inches(0.35)
    inner_top = Length(top + Inches(0.25))
    label_top = Length(top + Inches(1.10))
    delta_top = Length(top + Inches(1.50))

    value_box = slide.shapes.add_textbox(left, inner_top, width, value_h)
    _fill_text_frame(
        value_box.text_frame,
        str(value),
        token=value_token,
        color=value_color,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    label_box = slide.shapes.add_textbox(left, label_top, width, label_h)
    _fill_text_frame(
        label_box.text_frame,
        str(label),
        token=label_token,
        color=label_color,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.TOP,
    )

    delta_box = None
    if delta is not None:
        d_text, d_sign = _resolve_delta(delta)
        if d_text is not None:
            delta_box = slide.shapes.add_textbox(left, delta_top, width, delta_h)
            _fill_text_frame(
                delta_box.text_frame,
                d_text,
                token=delta_token,
                color=_delta_color(tokens, d_sign),
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.TOP,
            )

    # Tag the stack so the lint pass treats them as one intentional
    # group, not three overlapping shapes.
    group_name = f"kpi_card@{int(left)},{int(top)}"
    for shape in (card, value_box, label_box, delta_box):
        if shape is None:
            continue
        try:
            shape.lint_group = group_name
        except (AttributeError, NotImplementedError):
            pass

    return KpiCard(
        card=card, value_box=value_box, label_box=label_box, delta_box=delta_box
    )


def add_progress_bar(
    slide: "Slide",
    *,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    fraction: float,
    tokens: Optional[DesignTokens] = None,
    fill_color: Any = None,
    track_color: Any = None,
) -> ProgressBar:
    """Add a horizontal progress bar (track + fill) to *slide*.

    `fraction` is clamped to ``[0.0, 1.0]``. The fill shape's width is
    ``round(fraction * width)``; when the fraction is zero the fill is
    still emitted (with zero width) so callers can mutate it later
    (e.g. animate the fill on click).

    Colours fall back to the design tokens' ``primary`` / ``surface``
    palette slots when ``fill_color`` / ``track_color`` are ``None``.
    Pass any colour-like (``RGBColor``, hex string, ``(r, g, b)``)
    to override.
    """
    if not 0.0 <= float(fraction) <= 1.0:
        # Clamp rather than raise — values from live data sources are
        # often 99.x or 100.1 due to rounding; raising on those is
        # hostile.
        fraction = max(0.0, min(1.0, float(fraction)))

    resolved_track = _coerce_or_token(track_color, tokens, ("surface", "lt2"))
    resolved_fill = _coerce_or_token(fill_color, tokens, ("primary", "accent", "neutral"))

    track = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    if resolved_track is not None:
        track.fill.solid()
        track.fill.fore_color.rgb = resolved_track
    else:
        track.fill.background()
    track.line.fill.background()
    track.text_frame.text = ""

    fill_w = Length(int(round(int(width) * fraction)))
    fill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, fill_w, height)
    if resolved_fill is not None:
        fill.fill.solid()
        fill.fill.fore_color.rgb = resolved_fill
    else:
        fill.fill.background()
    fill.line.fill.background()
    fill.text_frame.text = ""

    group_name = f"progress_bar@{int(left)},{int(top)}"
    for shape in (track, fill):
        try:
            shape.lint_group = group_name
        except (AttributeError, NotImplementedError):
            pass

    return ProgressBar(track=track, fill=fill)


def _coerce_or_token(value, tokens, fallback_keys):
    """Return an RGBColor for ``value``, or read from ``tokens`` palette."""
    if value is None:
        return _palette(tokens, fallback_keys)
    from power_pptx._color import coerce_color

    return coerce_color(value)
