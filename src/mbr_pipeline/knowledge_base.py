"""
Build one unified, queryable snapshot of every project's MBR data,
fresh from sample_data/ every time it's called.

WHY THIS EXISTS
----------------
build_mbr.py already extracts everything needed to answer the chatbot's
example questions (managers, achievements, in-progress tickets) while
building the PowerPoint deck -- but that data lives scattered across
per-project local variables inside one big function, thrown away once
the deck is saved. This module reruns the same underlying extractors
(graph_data_extractor, key_highlights_extractor, key_achievements_finder,
a new roadmap extractor below) and assembles the results into one
queryable object per project, with nothing cached -- every call re-reads
the actual files in sample_data/, so a chatbot answer is never stale
relative to the last edit on disk.

DESIGN NOTE
------------
This module deliberately does not import from build_mbr.py (and
build_mbr.py does not import from this module) to avoid coupling the
chatbot's read path to the deck-building write path. Both call the same
lower-level extractors independently. If you change how a project's data
is extracted, update the relevant extractor module directly (e.g.
graph_data_extractor.py) -- both build_mbr.py and this module pick up
the change automatically.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from .auto_discovery import discover_inputs, DiscoveryError
from .graph_data_extractor import read_graph_data_for_sheet, ProjectGraphData
from .project_registry import PROJECT_REGISTRY, ProjectEntry
from .goals_consolidator import find_deck_for_entry
from .odp_converter import ensure_pptx, ConversionError
from .key_highlights_extractor import extract_completed_and_in_progress, ExtractiveSummary
from .key_achievements_finder import find_key_achievements_slide
from .roadmap_extractor import find_roadmap_summary


@dataclass
class ProjectKnowledge:
    entry: ProjectEntry
    graph_data: "ProjectGraphData | None" = None
    deck_path: "Path | None" = None
    deck_read_error: "str | None" = None
    completed: "ExtractiveSummary | None" = None
    in_progress: "ExtractiveSummary | None" = None
    key_achievements_bullets: "list[str] | None" = None
    roadmap_bullets: "list[str] | None" = None

    @property
    def display_name(self) -> str:
        if self.graph_data and self.graph_data.sub_project:
            return self.graph_data.sub_project
        return self.entry.label()

    @property
    def project_manager(self) -> str:
        return self.graph_data.project_manager if self.graph_data else "TBD"

    @property
    def lt_manager(self) -> str:
        if self.graph_data and self.graph_data.lt_manager:
            return self.graph_data.lt_manager
        return self.entry.lt_manager or "TBD"


@dataclass
class KnowledgeBase:
    month: str
    template_path: Path
    graph_dir: Path
    ppt_data_dir: Path
    projects: list[ProjectKnowledge] = field(default_factory=list)

    def find_project(self, query: str) -> "ProjectKnowledge | None":
        """Find a project by fuzzy name match against its label, sheet
        name, slide title hint, live Sub Project value, or any
        parenthesized abbreviation within those names (e.g. "WPM" from
        "WordPress Modernization (WPM)") -- the same multi-candidate
        matching style used elsewhere in this pipeline (e.g.
        key_highlights_table.py), since no single name field is
        reliable across every naming system in this data, and several
        projects' only short/abbreviated form lives in parentheses
        rather than in any dedicated field."""
        query_norm = _normalize(query)
        if not query_norm:
            return None
        for p in self.projects:
            full = [p.entry.label(), p.entry.sheet_name, p.entry.slide_title_hint, p.display_name]
            abbrevs = []
            for name in full:
                abbrevs.extend(re.findall(r"\(([A-Za-z0-9]+)\)", name))
            for c in full + abbrevs:
                c_norm = _normalize(c)
                if c_norm and (c_norm in query_norm or query_norm in c_norm):
                    return p
        return None


def _normalize(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _graph_workbook_for_entry(graph_dir: Path, entry: ProjectEntry) -> "Path | None":
    """Mirrors build_mbr.py's graph_workbook_for_entry: prefer the file
    whose name matches this project, since Graph workbooks can drift out
    of sync with each other once any one of them is edited (see
    MBR_PIPELINE.md)."""
    candidates = sorted(graph_dir.glob("LT-SimplifiedMetrics-*.xlsx"))
    if not candidates:
        return None
    target_tokens = {_normalize(entry.sheet_name), _normalize(entry.slide_title_hint)}
    for path in candidates:
        slug = path.stem.replace("LT-SimplifiedMetrics-", "")
        slug_norm = _normalize(slug)
        if slug_norm and any(slug_norm in t or t in slug_norm for t in target_tokens if t):
            return path
    return candidates[0]


def build_knowledge_base(root: "str | Path" = "sample_data", odp_cache_dir: "str | Path" = "output/_odp_cache") -> KnowledgeBase:
    """Re-read everything in `root` fresh and assemble a KnowledgeBase.
    No caching -- call this again after any file in `root` changes and
    you'll get the current state, not a stale snapshot."""
    discovered = discover_inputs(root)
    kb = KnowledgeBase(
        month=discovered.month,
        template_path=discovered.template_path,
        graph_dir=discovered.graph_dir,
        ppt_data_dir=discovered.ppt_data_dir,
    )

    for entry in PROJECT_REGISTRY:
        pk = ProjectKnowledge(entry=entry)

        workbook_path = _graph_workbook_for_entry(discovered.graph_dir, entry)
        if workbook_path is not None:
            pk.graph_data = read_graph_data_for_sheet(workbook_path, entry.sheet_name, display_name=entry.label())

        deck_path = find_deck_for_entry(discovered.ppt_data_dir, entry.deck_glob)
        if deck_path is not None:
            pk.deck_path = deck_path
            try:
                pptx_path = ensure_pptx(deck_path, odp_cache_dir)
                src_prs = Presentation(pptx_path)
            except (ConversionError, Exception) as exc:
                pk.deck_read_error = str(exc)
            else:
                completed, in_progress = extract_completed_and_in_progress(src_prs)
                pk.completed = completed
                pk.in_progress = in_progress

                ka_info = find_key_achievements_slide(src_prs)
                if ka_info is not None:
                    pk.key_achievements_bullets = ka_info.bullets

                roadmap = find_roadmap_summary(src_prs)
                if roadmap is not None:
                    pk.roadmap_bullets = roadmap.bullets

        kb.projects.append(pk)

    return kb


if __name__ == "__main__":
    import sys

    root = sys.argv[1] if len(sys.argv) > 1 else "sample_data"
    kb = build_knowledge_base(root)
    print(f"Month: {kb.month}")
    for p in kb.projects:
        print(f"\n=== {p.display_name} ===")
        print(f"  Project Manager/Lead: {p.project_manager}")
        print(f"  LT Manager: {p.lt_manager}")
        print(f"  Deck: {p.deck_path.name if p.deck_path else '(none)'}")
        if p.key_achievements_bullets:
            print(f"  Key Achievements: {p.key_achievements_bullets[:2]}")
        if p.in_progress:
            print(f"  In Progress: {p.in_progress.bullets[:2]}")
        if p.roadmap_bullets:
            print(f"  Roadmap: {p.roadmap_bullets[:2]}")
