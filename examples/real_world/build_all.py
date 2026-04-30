"""Build every real-world example deck into ``_out/``.

Usage::

    python examples/real_world/build_all.py
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "_out"

# Make sibling helpers (`_brand`, `_common`) importable from each script.
sys.path.insert(0, str(HERE))


SCRIPTS = [
    "01_q4_earnings_review",
    "02_annual_strategic_plan",
    "03_product_launch",
    "04_investor_pitch",
    "05_cybersecurity_briefing",
    "06_sales_qbr",
    "07_acquisition_proposal",
    "08_operational_excellence",
    "09_talent_strategy",
    "10_marketing_campaign",
]


def _load(name: str):
    path = HERE / f"{name}.py"
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"could not import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    OUT.mkdir(exist_ok=True)
    for name in SCRIPTS:
        out_path = OUT / f"{name}.pptx"
        module = _load(name)
        module.build(out_path)
    print(f"\n{len(SCRIPTS)} decks written to {OUT}/")


if __name__ == "__main__":
    main()
