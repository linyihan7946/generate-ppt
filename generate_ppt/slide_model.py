from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class SlideKind(str, Enum):
    COVER = "cover"
    SECTION = "section"
    BULLETS = "bullets"
    FLOW = "flow"
    COMPARE = "compare"
    PYRAMID = "pyramid"
    RELATION = "relation"
    PROBLEM_CHAIN = "problem_chain"
    SCREENSHOT = "screenshot"
    SUMMARY = "summary"


@dataclass
class SlideSpec:
    kind: SlideKind
    title: str
    subtitle: str = ""
    items: list[str] = field(default_factory=list)
    accent: str = "blue"
    highlight: str = ""
    screenshot_hint: str = ""


@dataclass
class DeckSpec:
    title: str
    subtitle: str
    template_id: str
    slides: list[SlideSpec]
