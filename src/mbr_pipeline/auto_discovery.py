"""
Auto-discover everything build_mbr.py needs from the contents of a
sample_data/-style folder, so the pipeline can be re-run with zero
arguments every time new files are dropped in.

WHAT GETS AUTO-DETECTED
-------------------------
- **Reporting month**: the month name that appears most often across
  the team deck filenames in ppt-data/ (e.g. every
  "V2Solutions_..._March_..." file votes for "March"). Ties are broken
  by preferring whichever month is closer to today's date.
- **Base template**: the single .pptx file in final-data/. If more than
  one exists, the most recently modified one is used (on the assumption
  that's the latest finished deck to build on top of), and a warning is
  printed since this is the one case auto-detection could silently pick
  the "wrong" file.
- **Graph workbook folder / ppt-data folder**: fixed conventional
  sub-folder names (Graph/, ppt-data/) under the given root, since
  there's no content-based way to tell these apart from each other if
  they were named differently.

WHY THIS EXISTS
----------------
Without this, every monthly run required typing out 5 CLI flags by hand,
including the exact month name -- easy to typo, and not meaningfully
different information from what's already implied by which files you
dropped in the folder. This module makes `python3 build_mbr.py` (no
arguments) the normal way to run a monthly build; explicit flags remain
available to override any individual choice.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path

_MONTH_NAMES = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
]
_MONTH_ABBREV = {m[:3]: m for m in _MONTH_NAMES}
_MONTH_NUM = {m: i + 1 for i, m in enumerate(_MONTH_NAMES)}


class DiscoveryError(RuntimeError):
    pass


@dataclass
class DiscoveredInputs:
    root: Path
    template_path: Path
    graph_dir: Path
    ppt_data_dir: Path
    month: str
    month_vote_detail: dict[str, int]


def _tokenize_month_votes(filenames: list[str]) -> Counter:
    votes = Counter()
    for name in filenames:
        for token in re.findall(r"[A-Za-z]+", name):
            tl = token.lower()
            if tl in _MONTH_NAMES:
                votes[tl] += 1
            elif tl in _MONTH_ABBREV:
                votes[_MONTH_ABBREV[tl]] += 1
    return votes


def detect_month(ppt_data_dir: Path, today: "date | None" = None) -> tuple[str, dict[str, int]]:
    """Pick the reporting month by majority vote across team deck
    filenames in `ppt_data_dir`. Ties are broken toward whichever
    candidate month is closer to `today` (defaults to the real today),
    looking back up to 11 months (a report is almost always for a
    recently-elapsed month, never one more than a year old)."""
    today = today or date.today()
    filenames = [p.name for p in ppt_data_dir.iterdir() if p.is_file() and not p.name.startswith("~$")]
    votes = _tokenize_month_votes(filenames)

    if not votes:
        raise DiscoveryError(
            f"Could not detect a reporting month from any filename in {ppt_data_dir}. "
            "Pass --month explicitly."
        )

    top_count = max(votes.values())
    candidates = [m for m, c in votes.items() if c == top_count]
    if len(candidates) == 1:
        return candidates[0].capitalize(), dict(votes)

    # Break ties by recency: walk back from this month, pick the first
    # candidate matched.
    for back in range(12):
        probe_month_num = ((today.month - 1 - back) % 12) + 1
        probe_name = _MONTH_NAMES[probe_month_num - 1]
        if probe_name in candidates:
            return probe_name.capitalize(), dict(votes)

    return sorted(candidates)[0].capitalize(), dict(votes)


def detect_template(final_data_dir: Path) -> Path:
    """Pick the base template: the only .pptx in `final_data_dir`, or the
    most-recently-modified one if there's more than one (with a caller-
    visible warning expected from the caller in that case)."""
    candidates = sorted(final_data_dir.glob("*.pptx"))
    if not candidates:
        raise DiscoveryError(f"No .pptx template found in {final_data_dir}")
    if len(candidates) == 1:
        return candidates[0]
    return max(candidates, key=lambda p: p.stat().st_mtime)


def discover_inputs(root: "str | Path" = "sample_data") -> DiscoveredInputs:
    root = Path(root)
    graph_dir = root / "Graph"
    ppt_data_dir = root / "ppt-data"
    final_data_dir = root / "final-data"

    for label, d in (("Graph", graph_dir), ("ppt-data", ppt_data_dir), ("final-data", final_data_dir)):
        if not d.is_dir():
            raise DiscoveryError(f"Expected a '{label}' folder at {d}, but it doesn't exist.")

    template_path = detect_template(final_data_dir)
    month, vote_detail = detect_month(ppt_data_dir)

    return DiscoveredInputs(
        root=root,
        template_path=template_path,
        graph_dir=graph_dir,
        ppt_data_dir=ppt_data_dir,
        month=month,
        month_vote_detail=vote_detail,
    )


def default_output_path(month: str, report_date: "date | None" = None) -> Path:
    report_date = report_date or date.today()
    return Path("output") / f"MBR_{month}_{report_date.year}.pptx"


if __name__ == "__main__":
    import sys

    root = sys.argv[1] if len(sys.argv) > 1 else "sample_data"
    discovered = discover_inputs(root)
    print(f"Template:    {discovered.template_path}")
    print(f"Graph dir:   {discovered.graph_dir}")
    print(f"ppt-data:    {discovered.ppt_data_dir}")
    print(f"Month:       {discovered.month}  (votes: {discovered.month_vote_detail})")
    print(f"Default out: {default_output_path(discovered.month)}")
