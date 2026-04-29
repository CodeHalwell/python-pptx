"""Opinionated parameterized slide recipes.

Each recipe is a small callable that produces a fully-styled slide using a
shared :class:`~power_pptx.design.tokens.DesignTokens` palette/typography set.
Recipes are deliberately additive: they sit on top of the low-level shape
APIs and the :class:`~power_pptx.design.style.ShapeStyle` facade, and never invent
new OOXML semantics.

The five recipes cover the slide types that account for most of a
typical pitch deck:

* :func:`title_slide`        — title + subtitle hero slide.
* :func:`bullet_slide`       — title + bulleted body.
* :func:`kpi_slide`          — title + KPI card row.
* :func:`quote_slide`        — large pull-quote with attribution.
* :func:`image_hero_slide`   — full-bleed image with overlay caption.

All recipes accept an optional :class:`DesignTokens` argument; when omitted,
they fall back to PowerPoint's defaults (black text, no fill).  Each
recipe also accepts an optional ``transition`` keyword that names a
:class:`~power_pptx.enum.presentation.MSO_TRANSITION_TYPE` member by lowercase
name (``"fade"``, ``"morph"``, ...).

Example::

    from power_pptx import Presentation
    from power_pptx.design.tokens import DesignTokens
    from power_pptx.design.recipes import title_slide, bullet_slide, kpi_slide

    prs = Presentation()
    tokens = DesignTokens.from_dict({
        "palette": {"primary": "#3C2F80", "neutral": "#222222",
                     "accent":  "#FF6600", "muted":   "#777777"},
        "typography": {
            "heading": {"family": "Inter", "size": 36.0, "bold": True},
            "body":    {"family": "Inter", "size": 16.0},
        },
    })

    title_slide(prs, title="Q4 Review", subtitle="April 2026",
                 tokens=tokens, transition="morph")
    bullet_slide(prs, title="Highlights",
                  bullets=["Two flagships shipped.", "NPS +8 QoQ."],
                  tokens=tokens)
    kpi_slide(prs, title="Run-rate metrics",
               kpis=[{"label": "ARR", "value": "$182M", "delta": +0.27},
                     {"label": "NDR", "value": "131%",  "delta": +0.03}],
               tokens=tokens)

    prs.save("review.pptx")
"""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Any, Mapping, Optional, Sequence, Union

from power_pptx.design.tokens import DesignTokens, TypographyToken
from power_pptx.dml.color import RGBColor
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from power_pptx.util import Emu, Inches, Length, Pt

if TYPE_CHECKING:
    from power_pptx.presentation import Presentation
    from power_pptx.slide import Slide

__all__ = (
    "title_slide",
    "bullet_slide",
    "kpi_slide",
    "quote_slide",
    "image_hero_slide",
)


# ---------------------------------------------------------------------------
# Public recipes
# ---------------------------------------------------------------------------


def title_slide(
    prs: "Presentation",
    *,
    title: str,
    subtitle: Optional[str] = None,
    tokens: Optional[DesignTokens] = None,
    transition: Optional[str] = None,
) -> "Slide":
    """Append a title-hero slide to *prs* and return it.

    Uses the ``Blank`` layout so the recipe owns the geometry and styling
    decisions end-to-end.  Title text is centered horizontally and sits in
    the top half of the slide; the subtitle, if provided, sits just below.

    Tokens consumed:

    * **palette** — title color from ``primary`` (fallback ``neutral``);
      subtitle color from ``muted`` (fallback ``neutral``).
    * **typography** — ``heading`` for the title, ``body`` for the
      subtitle.  Both fall back to Calibri at sensible sizes when the
      token isn't set.
    """
    slide = _add_blank(prs)
    slide_w, slide_h = _slide_dims(prs)

    margin = Inches(1.0)
    title_top = Length(int(slide_h * 0.38))
    title_h = Inches(1.6)
    title_box = slide.shapes.add_textbox(
        margin, title_top, Length(slide_w - 2 * margin), title_h
    )
    _fill_text_frame(
        title_box.text_frame,
        title,
        token=_typography(tokens, "heading", default_size=Pt(44), default_bold=True),
        color=_palette(tokens, ("primary", "neutral")),
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    if subtitle:
        sub_top = Length(title_top + title_h + Inches(0.1))
        sub_h = Inches(0.8)
        sub_box = slide.shapes.add_textbox(
            margin, sub_top, Length(slide_w - 2 * margin), sub_h
        )
        _fill_text_frame(
            sub_box.text_frame,
            subtitle,
            token=_typography(tokens, "body", default_size=Pt(20)),
            color=_palette(tokens, ("muted", "neutral")),
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )

    _apply_transition(slide, transition)
    return slide


def bullet_slide(
    prs: "Presentation",
    *,
    title: str,
    bullets: Sequence[str],
    tokens: Optional[DesignTokens] = None,
    transition: Optional[str] = None,
) -> "Slide":
    """Append a title + bulleted-content slide and return it.

    Tokens consumed:

    * **palette** — title from ``primary`` (fallback ``neutral``);
      bullet text from ``neutral``.
    * **typography** — ``heading`` for the title, ``body`` for the bullet
      lines.
    """
    slide = _add_blank(prs)
    slide_w, slide_h = _slide_dims(prs)

    margin = Inches(0.6)
    title_top = Inches(0.5)
    title_h = Inches(1.0)
    title_box = slide.shapes.add_textbox(
        margin, title_top, Length(slide_w - 2 * margin), title_h
    )
    _fill_text_frame(
        title_box.text_frame,
        title,
        token=_typography(tokens, "heading", default_size=Pt(32), default_bold=True),
        color=_palette(tokens, ("primary", "neutral")),
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    )

    body_top = Length(title_top + title_h + Inches(0.2))
    body_h = Length(slide_h - body_top - Inches(0.5))
    body_box = slide.shapes.add_textbox(
        margin, body_top, Length(slide_w - 2 * margin), body_h
    )
    body_token = _typography(tokens, "body", default_size=Pt(18))
    body_color = _palette(tokens, ("neutral",))
    tf = body_box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = f"•  {bullet}"
        _apply_typography(run.font, body_token)
        if body_color is not None:
            run.font.color.rgb = body_color
        para.space_after = Pt(6)

    _apply_transition(slide, transition)
    return slide


def kpi_slide(
    prs: "Presentation",
    *,
    title: str,
    kpis: Sequence[Mapping[str, Any]],
    tokens: Optional[DesignTokens] = None,
    transition: Optional[str] = None,
) -> "Slide":
    """Append a title + KPI card-row slide and return it.

    Each KPI dict accepts:

    * ``label`` — the small caption under the value.
    * ``value`` — the headline number / string.
    * ``delta`` *(optional)* — numeric change.  Magnitude with absolute
      value at most ``1.0`` is treated as a fraction (``0.27`` → ``+27%``);
      anything larger is rendered as-is with one decimal
      (``14.0`` → ``+14.0``).  Pass a *string* to opt out of the
      auto-format and render the delta verbatim (``"+14 pts"``,
      ``"−$2.3M"``).  Positive deltas are tinted with the palette's
      ``positive`` slot, negative with ``negative`` (falling back to
      green / red).
    * ``delta_text`` *(optional)* — explicit string passthrough; takes
      precedence over ``delta`` when both are set.

    Tokens consumed:

    * **palette** — title from ``primary`` (fallback ``neutral``); card
      fill from ``surface`` (fallback ``lt2``); card border from
      ``muted`` (fallback ``lt1``); value from ``primary`` (fallback
      ``neutral``); label from ``muted``; delta from ``positive`` /
      ``success`` (positive) and ``negative`` / ``danger`` (negative).
    * **typography** — ``heading`` for the title and the KPI value;
      ``body`` for the label and the delta line.
    * **shadows** — ``card`` (optional) is applied as a soft shadow on
      each KPI card.
    """
    slide = _add_blank(prs)
    slide_w, _slide_h = _slide_dims(prs)

    margin = Inches(0.6)
    title_top = Inches(0.5)
    title_h = Inches(0.9)
    title_box = slide.shapes.add_textbox(
        margin, title_top, Length(slide_w - 2 * margin), title_h
    )
    _fill_text_frame(
        title_box.text_frame,
        title,
        token=_typography(tokens, "heading", default_size=Pt(28), default_bold=True),
        color=_palette(tokens, ("primary", "neutral")),
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    )

    n = len(kpis)
    if n == 0:
        _apply_transition(slide, transition)
        return slide

    card_h = Inches(1.9)
    gap = Inches(0.25)
    available = slide_w - 2 * margin - (n - 1) * gap
    card_w = Length(int(available // n))
    top = Length(title_top + title_h + Inches(0.4))

    fill_color = _palette(tokens, ("surface", "lt2"))
    border_color = _palette(tokens, ("muted", "lt1"))
    value_color = _palette(tokens, ("primary", "neutral"))
    label_color = _palette(tokens, ("muted",))

    value_token = _typography(tokens, "heading", default_size=Pt(30), default_bold=True)
    label_token = _typography(tokens, "body", default_size=Pt(12))
    delta_token = _typography(tokens, "body", default_size=Pt(11), default_bold=True)

    for i, kpi in enumerate(kpis):
        left = Length(margin + i * (card_w + gap))

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h
        )
        if fill_color is not None:
            card.fill.solid()
            card.fill.fore_color.rgb = fill_color
        else:
            card.fill.background()
        if border_color is not None:
            card.line.color.rgb = border_color
            card.line.width = Pt(0.75)
        # Drop a soft card shadow if the tokens declared one.
        if tokens is not None:
            card_shadow = tokens.shadows.get("card")
            if card_shadow is not None:
                card.style.shadow = card_shadow
        # Suppress the default text on the autoshape so it doesn't peek
        # through behind our textboxes.
        card.text_frame.text = ""

        # Value
        v_box = slide.shapes.add_textbox(
            left, Length(top + Inches(0.25)), card_w, Inches(0.85)
        )
        _fill_text_frame(
            v_box.text_frame,
            str(kpi.get("value", "")),
            token=value_token,
            color=value_color,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # Label
        l_box = slide.shapes.add_textbox(
            left, Length(top + Inches(1.10)), card_w, Inches(0.4)
        )
        _fill_text_frame(
            l_box.text_frame,
            str(kpi.get("label", "")),
            token=label_token,
            color=label_color,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )

        d_text, d_sign = _resolve_delta(kpi)
        if d_text is not None:
            d_color = _delta_color(tokens, d_sign)
            d_box = slide.shapes.add_textbox(
                left, Length(top + Inches(1.50)), card_w, Inches(0.35)
            )
            _fill_text_frame(
                d_box.text_frame,
                d_text,
                token=delta_token,
                color=d_color,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.TOP,
            )

    _apply_transition(slide, transition)
    return slide


def quote_slide(
    prs: "Presentation",
    *,
    quote: str,
    attribution: Optional[str] = None,
    tokens: Optional[DesignTokens] = None,
    transition: Optional[str] = None,
) -> "Slide":
    """Append a centered pull-quote slide with optional attribution.

    *attribution*, when supplied, is rendered with an em-dash prefix
    (``— Person``).  A leading dash / em-dash / en-dash on the input
    is stripped so callers can pass either ``"Person"`` or ``"— Person"``
    without doubling the dash.

    Tokens consumed:

    * **palette** — quote text from ``primary`` (fallback ``neutral``);
      attribution from ``muted`` (fallback ``neutral``).
    * **typography** — ``heading`` for the quote (italic by default);
      ``body`` for the attribution.
    """
    slide = _add_blank(prs)
    slide_w, slide_h = _slide_dims(prs)

    margin = Inches(1.2)
    quote_h = Inches(3.5)
    quote_top = Length((slide_h - quote_h) // 2)
    box = slide.shapes.add_textbox(
        margin, quote_top, Length(slide_w - 2 * margin), quote_h
    )
    _fill_text_frame(
        box.text_frame,
        f"“{quote}”",
        token=_typography(tokens, "heading", default_size=Pt(32), default_italic=True),
        color=_palette(tokens, ("primary", "neutral")),
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        word_wrap=True,
    )

    if attribution:
        att_top = Length(quote_top + quote_h + Inches(0.1))
        att_box = slide.shapes.add_textbox(
            margin, att_top, Length(slide_w - 2 * margin), Inches(0.6)
        )
        # Strip any leading dash variants the caller may have already
        # written (``-``, ``–`` en-dash, ``—`` em-dash) so the recipe's
        # em-dash isn't doubled — the silent-doubling failure mode is
        # easy to miss in PR review.
        att_clean = _strip_attribution_dash(attribution)
        _fill_text_frame(
            att_box.text_frame,
            f"— {att_clean}",
            token=_typography(tokens, "body", default_size=Pt(16)),
            color=_palette(tokens, ("muted", "neutral")),
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )

    _apply_transition(slide, transition)
    return slide


def image_hero_slide(
    prs: "Presentation",
    *,
    title: str,
    image: Union[str, IO[bytes]],
    caption: Optional[str] = None,
    tokens: Optional[DesignTokens] = None,
    transition: Optional[str] = None,
) -> "Slide":
    """Append a full-bleed image slide with an overlaid title (and caption).

    The image is added at the slide origin and stretched to the slide's
    full extent.  The title sits in a tinted band across the bottom third
    so it remains readable regardless of the underlying image.

    Tokens consumed:

    * **palette** — band fill from ``primary`` (fallback ``neutral``,
      then black); title and caption text from ``on_primary`` (falling
      back to white / near-white).
    * **typography** — ``heading`` for the title, ``body`` for the
      caption.
    """
    slide = _add_blank(prs)
    slide_w, slide_h = _slide_dims(prs)

    slide.shapes.add_picture(image, Emu(0), Emu(0), slide_w, slide_h)

    band_h = Inches(1.6 if caption else 1.2)
    band_top = Length(slide_h - band_h)
    band = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), band_top, slide_w, band_h
    )
    band.line.fill.background()
    band_color = _palette(tokens, ("primary", "neutral")) or RGBColor(0, 0, 0)
    band.fill.solid()
    band.fill.fore_color.rgb = band_color
    band.fill.fore_color.alpha = 0.55
    band.text_frame.text = ""

    margin = Inches(0.6)
    title_box = slide.shapes.add_textbox(
        margin, Length(band_top + Inches(0.2)),
        Length(slide_w - 2 * margin), Inches(0.8),
    )
    _fill_text_frame(
        title_box.text_frame,
        title,
        token=_typography(tokens, "heading", default_size=Pt(28), default_bold=True),
        color=_palette(tokens, ("on_primary",)) or RGBColor(0xFF, 0xFF, 0xFF),
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    )

    if caption:
        cap_box = slide.shapes.add_textbox(
            margin, Length(band_top + Inches(1.0)),
            Length(slide_w - 2 * margin), Inches(0.5),
        )
        _fill_text_frame(
            cap_box.text_frame,
            caption,
            token=_typography(tokens, "body", default_size=Pt(14)),
            color=_palette(tokens, ("on_primary",)) or RGBColor(0xEE, 0xEE, 0xEE),
            align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.TOP,
        )

    _apply_transition(slide, transition)
    return slide


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _add_blank(prs: "Presentation") -> "Slide":
    layouts = prs.slide_layouts
    blank = layouts.get_by_name("Blank")
    if blank is None:
        blank = layouts[-1]
    return prs.slides.add_slide(blank)


def _slide_dims(prs: "Presentation") -> tuple[Length, Length]:
    width = prs.slide_width or Inches(13.333)
    height = prs.slide_height or Inches(7.5)
    return width, height


def _palette(
    tokens: Optional[DesignTokens], names: Sequence[str]
) -> Optional[RGBColor]:
    if tokens is None:
        return None
    for name in names:
        rgb = tokens.palette.get(name)
        if rgb is not None:
            return rgb
    return None


def _typography(
    tokens: Optional[DesignTokens],
    name: str,
    *,
    default_size: Length,
    default_bold: Optional[bool] = None,
    default_italic: Optional[bool] = None,
) -> TypographyToken:
    """Return the named token from *tokens* or a sensible default."""
    if tokens is not None:
        existing = tokens.typography.get(name)
        if existing is not None:
            # Backfill any unset fields from the recipe defaults so callers
            # don't have to repeat them.
            return TypographyToken(
                family=existing.family,
                size=existing.size if existing.size is not None else default_size,
                bold=existing.bold if existing.bold is not None else default_bold,
                italic=existing.italic if existing.italic is not None else default_italic,
                color=existing.color,
            )
    return TypographyToken(
        family="Calibri",
        size=default_size,
        bold=default_bold,
        italic=default_italic,
    )


def _apply_typography(font: Any, token: TypographyToken) -> None:
    font.name = token.family
    if token.size is not None:
        font.size = token.size
    if token.bold is not None:
        font.bold = token.bold
    if token.italic is not None:
        font.italic = token.italic
    if token.color is not None:
        font.color.rgb = token.color


def _fill_text_frame(
    text_frame: Any,
    text: str,
    *,
    token: TypographyToken,
    color: Optional[RGBColor],
    align: PP_ALIGN,
    anchor: MSO_ANCHOR,
    word_wrap: bool = True,
) -> None:
    text_frame.word_wrap = word_wrap
    text_frame.vertical_anchor = anchor
    para = text_frame.paragraphs[0]
    para.alignment = align
    # Clear any default text and add a fresh run we can style.
    para.text = ""
    run = para.add_run()
    run.text = text
    _apply_typography(run.font, token)
    if color is not None:
        run.font.color.rgb = color


def _apply_transition(slide: "Slide", transition: Optional[str]) -> None:
    if not transition:
        return
    from power_pptx.enum.presentation import MSO_TRANSITION_TYPE

    key = transition.upper().replace("-", "_")
    member = getattr(MSO_TRANSITION_TYPE, key, None)
    if member is None:
        raise ValueError(
            f"Unknown transition {transition!r}; "
            f"see power_pptx.enum.presentation.MSO_TRANSITION_TYPE for valid values"
        )
    slide.transition.kind = member


def _delta_color(tokens: Optional[DesignTokens], sign: int) -> RGBColor:
    """Return the palette color tinting a delta with ``sign`` (+1 / -1 / 0).

    ``sign == 0`` and ``sign > 0`` both use the positive slot; only
    strictly-negative deltas use the negative slot.
    """
    if sign >= 0:
        return _palette(tokens, ("positive", "success")) or RGBColor(0x00, 0x8A, 0x3C)
    return _palette(tokens, ("negative", "danger")) or RGBColor(0xCC, 0x00, 0x00)


# Magnitude at or below which a numeric delta is treated as a fraction
# (``0.27`` → ``+27%``).  Outside this range we render the raw number
# with one decimal so callers who already pass percentages — ``14.0`` —
# don't get them silently multiplied by 100.
_DELTA_FRACTION_LIMIT = 1.0


def _resolve_delta(kpi: Mapping[str, Any]) -> tuple[Optional[str], int]:
    """Resolve the formatted delta string and sign from a KPI dict.

    Returns ``(text, sign)`` where *text* is ``None`` when no delta was
    supplied.  The auto-detect rules:

    * ``delta_text="…"`` — explicit string passthrough wins outright.
    * ``delta="…"`` (string) — used verbatim; sign is inferred from a
      leading ``-`` / ``−`` if present.
    * ``delta`` (numeric) with ``|delta| <= 1.0`` — formatted as a
      signed percentage.
    * ``delta`` (numeric) with ``|delta| > 1.0`` — formatted as a
      signed number with one decimal place.
    """
    explicit = kpi.get("delta_text")
    if explicit is not None:
        s = str(explicit)
        sign = -1 if s.lstrip().startswith(("-", "−", "–")) else 1
        return s, sign

    delta = kpi.get("delta")
    if delta is None:
        return None, 0
    if isinstance(delta, str):
        sign = -1 if delta.lstrip().startswith(("-", "−", "–")) else 1
        return delta, sign

    d_val = float(delta)
    sign = -1 if d_val < 0 else 1
    sign_glyph = "+" if d_val >= 0 else "−"  # proper minus glyph
    if abs(d_val) <= _DELTA_FRACTION_LIMIT:
        text = f"{sign_glyph}{abs(d_val):.0%}"
    else:
        text = f"{sign_glyph}{abs(d_val):.1f}"
    return text, sign


def _strip_attribution_dash(attribution: str) -> str:
    """Remove a leading dash variant + whitespace from *attribution*.

    Handles ``-``, ``–`` (en-dash), and ``—`` (em-dash) so callers can
    pass either ``"Person"`` or ``"— Person"`` without producing
    ``"— — Person"`` in the rendered slide.
    """
    s = attribution.lstrip()
    while s and s[0] in ("-", "–", "—"):
        s = s[1:].lstrip()
    return s
