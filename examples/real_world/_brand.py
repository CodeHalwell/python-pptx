"""Shared brand identities for the real-world example decks.

Each Fortune-500-style deck picks a distinct identity so the suite
shows how a token spec drives every recipe and chart palette across
very different verticals.

Tokens stay tight and opinionated: a primary, an accent, neutral
text, surface, semantic positive/negative, and matched typography.
"""

from __future__ import annotations

from power_pptx.design.tokens import DesignTokens


def _make(palette: dict, heading_font: str = "Inter", body_font: str = "Inter") -> DesignTokens:
    return DesignTokens.from_dict(
        {
            "palette": palette,
            "typography": {
                "heading": {"family": heading_font, "size": 40.0, "bold": True},
                "body":    {"family": body_font,    "size": 18.0},
                "caption": {"family": body_font,    "size": 11.0, "italic": True},
            },
            "shadows": {
                "card": {"blur": 18.0, "distance": 4.0, "alpha": 0.16},
                "soft": {"blur": 32.0, "distance": 8.0, "alpha": 0.08},
            },
            "radii":    {"card": 12.0, "button": 6.0},
            "spacings": {"sm": 8.0, "md": 16.0, "lg": 32.0, "xl": 48.0},
        }
    )


# 01 — Q4 Earnings Review (financial / serious navy + gold)
EARNINGS = _make({
    "primary":    "#0B2447",   # deep navy
    "accent":     "#C9A227",   # muted gold
    "neutral":    "#0F172A",
    "muted":      "#475569",
    "surface":    "#F4F5F7",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#1B7F3F",
    "negative":   "#B91C1C",
})
EARNINGS_PALETTE = ["#0B2447", "#C9A227", "#1B7F3F", "#3B82F6", "#94A3B8", "#B91C1C"]

# 02 — Annual Strategic Plan (executive blue)
STRATEGY = _make({
    "primary":    "#1E3A8A",
    "accent":     "#0EA5E9",
    "neutral":    "#0F172A",
    "muted":      "#64748B",
    "surface":    "#F1F5F9",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#10B981",
    "negative":   "#EF4444",
})
STRATEGY_PALETTE = ["#1E3A8A", "#0EA5E9", "#10B981", "#F59E0B", "#8B5CF6", "#EF4444"]

# 03 — Product Launch (modern tech, vibrant)
PRODUCT = _make({
    "primary":    "#7C3AED",
    "accent":     "#22D3EE",
    "neutral":    "#0B1020",
    "muted":      "#6B7280",
    "surface":    "#F5F3FF",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#10B981",
    "negative":   "#EF4444",
})
PRODUCT_PALETTE = ["#7C3AED", "#22D3EE", "#F472B6", "#10B981", "#F59E0B", "#3B82F6"]

# 04 — Investor Pitch (startup-confident, indigo + emerald)
INVESTOR = _make({
    "primary":    "#4338CA",
    "accent":     "#10B981",
    "neutral":    "#111827",
    "muted":      "#6B7280",
    "surface":    "#F8FAFC",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#10B981",
    "negative":   "#EF4444",
})
INVESTOR_PALETTE = ["#4338CA", "#10B981", "#F59E0B", "#06B6D4", "#EC4899", "#EF4444"]

# 05 — Cybersecurity Board Briefing (graphite + amber)
SECURITY = _make({
    "primary":    "#1F2937",
    "accent":     "#F59E0B",
    "neutral":    "#0F172A",
    "muted":      "#64748B",
    "surface":    "#F3F4F6",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#16A34A",
    "negative":   "#DC2626",
})
SECURITY_PALETTE = ["#1F2937", "#F59E0B", "#DC2626", "#16A34A", "#3B82F6", "#8B5CF6"]

# 06 — Sales QBR (energetic blue + green growth)
SALES = _make({
    "primary":    "#0369A1",
    "accent":     "#16A34A",
    "neutral":    "#0F172A",
    "muted":      "#64748B",
    "surface":    "#F0F9FF",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#16A34A",
    "negative":   "#DC2626",
})
SALES_PALETTE = ["#0369A1", "#16A34A", "#F59E0B", "#7C3AED", "#06B6D4", "#DC2626"]

# 07 — M&A Acquisition Proposal (corporate, conservative)
MERGER = _make({
    "primary":    "#0F2D6B",
    "accent":     "#9CA3AF",
    "neutral":    "#111827",
    "muted":      "#6B7280",
    "surface":    "#F9FAFB",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#15803D",
    "negative":   "#B91C1C",
})
MERGER_PALETTE = ["#0F2D6B", "#475569", "#15803D", "#C9A227", "#B91C1C", "#3B82F6"]

# 08 — Operational Excellence (industrial teal)
OPS = _make({
    "primary":    "#0F766E",
    "accent":     "#F97316",
    "neutral":    "#0F172A",
    "muted":      "#64748B",
    "surface":    "#F0FDFA",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#16A34A",
    "negative":   "#DC2626",
})
OPS_PALETTE = ["#0F766E", "#F97316", "#16A34A", "#0EA5E9", "#A855F7", "#DC2626"]

# 09 — Talent & Workforce (warm humanist)
PEOPLE = _make({
    "primary":    "#9D174D",
    "accent":     "#F59E0B",
    "neutral":    "#1F2937",
    "muted":      "#6B7280",
    "surface":    "#FDF2F8",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#16A34A",
    "negative":   "#DC2626",
})
PEOPLE_PALETTE = ["#9D174D", "#F59E0B", "#0EA5E9", "#16A34A", "#7C3AED", "#DC2626"]

# 10 — Marketing Campaign (creative, magenta + cyan)
MARKETING = _make({
    "primary":    "#DB2777",
    "accent":     "#06B6D4",
    "neutral":    "#0F172A",
    "muted":      "#6B7280",
    "surface":    "#FFF1F2",
    "background": "#FFFFFF",
    "on_primary": "#FFFFFF",
    "positive":   "#16A34A",
    "negative":   "#DC2626",
})
MARKETING_PALETTE = ["#DB2777", "#06B6D4", "#F59E0B", "#7C3AED", "#10B981", "#3B82F6"]
