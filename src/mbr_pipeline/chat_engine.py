"""
Rule-based question answering over the MBR KnowledgeBase.

No external AI API is used. Every answer is retrieved directly from
data already extracted from sample_data/ by knowledge_base.py -- this
module's only job is figuring out (a) which project a question is about,
if any, and (b) which field of that project's data answers it, then
formatting that data as a readable response.

This intentionally does NOT try to handle arbitrary free-form questions
via a language model. It matches against a fixed set of intents (see
INTENT_PATTERNS) that cover the example questions in the spec and
reasonable variations of them. Anything that doesn't match a known
intent gets a clear "I don't understand that yet" response listing what
it *can* answer, rather than guessing.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .knowledge_base import KnowledgeBase, ProjectKnowledge


@dataclass
class ChatAnswer:
    text: str
    matched_intent: "str | None" = None
    matched_project: "str | None" = None


# Each intent: (name, regex patterns to try, handler function name below).
# Patterns are checked in order; first match wins. Project name extraction
# happens separately (see _extract_project_query) since "for SEOD" /
# "in the FLEX project" / "SEOD's" can appear anywhere in the sentence.
INTENT_PATTERNS = [
    ("key_achievements", [r"key achievement", r"what.*achiev", r"accomplish"]),
    ("roadmap", [r"road\s*map", r"what.*plan", r"upcoming\s+work"]),
    ("project_manager", [r"project manager", r"\bpm\b.*who", r"who.*\bpm\b", r"\blead\b.*who", r"who.*\blead\b"]),
    ("lt_manager", [r"lt manager", r"lendingtree manager", r"account manager", r"client poc"]),
    ("in_progress", [r"in[\s-]?progress", r"open ticket", r"current(ly)?\s+work"]),
    ("completed", [r"\bcompleted\b", r"\bdone\b", r"finished", r"goals achieved"]),
    ("metrics", [r"throughput", r"rework", r"carryover", r"\bmetric", r"how many (story )?points", r"velocity"]),
    ("list_projects", [r"list.*project", r"what projects", r"which projects", r"all projects"]),
    ("reporting_month", [r"what month", r"reporting (period|month)", r"which month"]),
    ("generate_report", [r"generate.*report", r"build.*(report|deck|ppt|presentation)", r"create.*(report|deck)"]),
    ("help", [r"^help$", r"what can you", r"what.*ask"]),
]

_COMPILED_INTENTS = [(name, [re.compile(p, re.IGNORECASE) for p in patterns]) for name, patterns in INTENT_PATTERNS]


def classify_intent(question: str) -> "str | None":
    for name, patterns in _COMPILED_INTENTS:
        if any(p.search(question) for p in patterns):
            return name
    return None


def _candidate_names(p: ProjectKnowledge) -> list[str]:
    """All name forms worth matching against a question for this
    project: full names (label, sheet name, slide title hint, live
    display name) plus any short abbreviation written in parentheses
    within those names (e.g. "WPM" from "WordPress Modernization
    (WPM)", "CET" from "Consumer Experience Team (CET)"). The
    abbreviation is often the *only* candidate short enough to match a
    casually-phrased question like "tickets for WPM" -- none of the
    full name fields reduce to that bare 3-letter token on their own.
    """
    full = [p.entry.label(), p.entry.sheet_name, p.entry.slide_title_hint, p.display_name]
    abbrevs = []
    for name in full:
        abbrevs.extend(re.findall(r"\(([A-Za-z0-9]+)\)", name))
    return full + abbrevs


def _extract_project_query(question: str, kb: KnowledgeBase) -> "ProjectKnowledge | None":
    """Try every known project's name candidates (see _candidate_names)
    against the question text, checking both directions: a long
    candidate appearing inside the question ("who is the project
    manager for WordPress Modernization") and the question containing a
    candidate as a substring even when the candidate is short (e.g. the
    extracted abbreviation "WPM" inside "tickets for WPM"). Prefers the
    longest matching candidate so a more specific name wins over a
    shorter coincidental match."""
    q_norm = re.sub(r"[^a-z0-9]", "", question.lower())
    best = None
    best_len = 0
    for p in kb.projects:
        for c in _candidate_names(p):
            c_norm = re.sub(r"[^a-z0-9]", "", c.lower())
            if not c_norm or len(c_norm) < 3:
                continue
            if c_norm in q_norm and len(c_norm) > best_len:
                best = p
                best_len = len(c_norm)
    return best


def _format_bullets(bullets: "list[str] | None", empty_msg: str) -> str:
    if not bullets:
        return empty_msg
    return "\n".join(f"- {b}" for b in bullets)


def _project_not_found_msg(kb: KnowledgeBase) -> str:
    names = ", ".join(p.display_name for p in kb.projects)
    return f"I couldn't tell which project you meant. Known projects: {names}."


def answer_question(question: str, kb: KnowledgeBase) -> ChatAnswer:
    question = question.strip()
    if not question:
        return ChatAnswer(text="Ask me something about this month's MBR data -- e.g. \"What are the key achievements for SEOD?\"")

    intent = classify_intent(question)

    if intent == "help" or intent is None and not question:
        return ChatAnswer(text=_help_text(kb), matched_intent="help")

    if intent == "list_projects":
        names = "\n".join(f"- {p.display_name}" for p in kb.projects)
        return ChatAnswer(text=f"This month's ({kb.month}) projects:\n{names}", matched_intent=intent)

    if intent == "reporting_month":
        return ChatAnswer(text=f"This report covers {kb.month}.", matched_intent=intent)

    if intent == "generate_report":
        return ChatAnswer(
            text="To regenerate the MBR PowerPoint, Excel report, and dashboard, run:\n\n"
                 "    python build_mbr.py\n"
                 "    python main.py\n\n"
                 "or both together with:\n\n"
                 "    python run_all.py\n\n"
                 "This always re-reads sample_data/ fresh -- no need to delete previous output first.",
            matched_intent=intent,
        )

    project = _extract_project_query(question, kb)

    if intent is None:
        if project is not None:
            return ChatAnswer(
                text=f"I found the project ({project.display_name}) but I'm not sure what you're asking about it. "
                     f"{_help_text(kb)}",
                matched_project=project.display_name,
            )
        return ChatAnswer(text=f"I don't understand that yet.\n\n{_help_text(kb)}")

    # For these intents, a question with no project name ("What are the
    # key achievements this month?") reasonably means "across every
    # project", not "which project did you mean?" -- summarize all of
    # them rather than asking for clarification.
    ALL_PROJECTS_INTENTS = {"key_achievements", "roadmap", "in_progress", "completed"}
    if project is None and intent in ALL_PROJECTS_INTENTS:
        return _answer_for_all_projects(intent, kb)

    if project is None:
        return ChatAnswer(text=_project_not_found_msg(kb), matched_intent=intent)

    matched_project = project.display_name

    if intent == "project_manager":
        return ChatAnswer(
            text=f"The Project Manager/Lead for {matched_project} is {project.project_manager}.",
            matched_intent=intent, matched_project=matched_project,
        )

    if intent == "lt_manager":
        return ChatAnswer(
            text=f"The LT Manager for {matched_project} is {project.lt_manager}.",
            matched_intent=intent, matched_project=matched_project,
        )

    if intent == "key_achievements":
        bullets = project.key_achievements_bullets
        body = _format_bullets(bullets, f"No Key Achievements slide was found in {matched_project}'s source deck this month.")
        return ChatAnswer(text=f"Key Achievements for {matched_project}:\n{body}", matched_intent=intent, matched_project=matched_project)

    if intent == "roadmap":
        bullets = project.roadmap_bullets
        body = _format_bullets(bullets, f"No Project Roadmap slide was found in {matched_project}'s source deck this month.")
        return ChatAnswer(text=f"Project Roadmap for {matched_project}:\n{body}", matched_intent=intent, matched_project=matched_project)

    if intent == "in_progress":
        bullets = project.in_progress.bullets if project.in_progress else None
        body = _format_bullets(bullets, f"No in-progress items were found for {matched_project} this month.")
        return ChatAnswer(text=f"In-progress for {matched_project}:\n{body}", matched_intent=intent, matched_project=matched_project)

    if intent == "completed":
        bullets = project.completed.bullets if project.completed else None
        body = _format_bullets(bullets, f"No completed items were found for {matched_project} this month.")
        return ChatAnswer(text=f"Completed for {matched_project}:\n{body}", matched_intent=intent, matched_project=matched_project)

    if intent == "metrics":
        gd = project.graph_data
        if gd is None or gd.error or not gd.weeks:
            return ChatAnswer(text=f"No metrics data was found for {matched_project} this month.", matched_intent=intent, matched_project=matched_project)
        week = gd.latest_week()
        text = (
            f"Latest reported week for {matched_project} ({week.week_label.strip()}):\n"
            f"- Planned SP: {week.planned_sp}\n"
            f"- Actual Delivered: {week.actual_delivered}\n"
            f"- V2: {week.v2}, LT: {week.lt}\n"
            f"- Rework: {week.rework}\n"
            f"- V2 Carryover: {week.v2_carryover}, LT Carryover: {week.lt_carryover}"
        )
        return ChatAnswer(text=text, matched_intent=intent, matched_project=matched_project)

    return ChatAnswer(text=_help_text(kb), matched_intent=intent, matched_project=matched_project)


def _answer_for_all_projects(intent: str, kb: KnowledgeBase) -> ChatAnswer:
    field_map = {
        "key_achievements": ("key_achievements_bullets", "Key Achievements"),
        "roadmap": ("roadmap_bullets", "Project Roadmap"),
        "in_progress": ("in_progress", "In-progress"),
        "completed": ("completed", "Completed"),
    }
    attr_name, label = field_map[intent]
    sections = []
    for p in kb.projects:
        value = getattr(p, attr_name)
        bullets = value.bullets if hasattr(value, "bullets") else value
        if bullets:
            sections.append(f"{p.display_name}:\n" + "\n".join(f"  - {b}" for b in bullets))
    if not sections:
        return ChatAnswer(
            text=f"No {label.lower()} were found for any project this month.",
            matched_intent=intent,
        )
    return ChatAnswer(
        text=f"{label} across all projects this month ({kb.month}):\n\n" + "\n\n".join(sections),
        matched_intent=intent,
    )


def _help_text(kb: KnowledgeBase) -> str:
    return (
        "I can answer questions like:\n"
        "- What are the key achievements for <project>?\n"
        "- What is the project roadmap for <project>?\n"
        "- Who is the Project Manager for <project>?\n"
        "- Who is the LT Manager for <project>?\n"
        "- What are the open/in-progress tickets for <project>?\n"
        "- What's completed for <project>?\n"
        "- What are the metrics for <project>?\n"
        "- List all projects\n"
        "- What month is this report for?\n"
        "- Generate the latest MBR report\n\n"
        f"Known projects this month: {', '.join(p.display_name for p in kb.projects)}"
    )
