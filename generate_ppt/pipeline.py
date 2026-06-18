from __future__ import annotations

import re
import uuid
from pathlib import Path

from generate_ppt.animation import inject_fade_animations
from generate_ppt.document_loader import load_document
from generate_ppt.ppt_renderer import PptRenderer
from generate_ppt.templates import get_template


def generate_ppt(input_path: Path, template_id: str, output_dir: Path) -> Path:
    document = load_document(input_path)
    template_definition = get_template(template_id)
    deck = template_definition.template.build(document)
    safe_stem = _safe_name(document.title) or "technical_deck"
    output_path = output_dir / f"{safe_stem}_{uuid.uuid4().hex[:8]}.pptx"
    PptRenderer().render(deck, output_path)
    inject_fade_animations(output_path)
    return output_path


def _safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" ._")
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_")
    return value[:40]
