from __future__ import annotations

from dataclasses import dataclass

from generate_ppt.templates.technical_no_image import TechnicalNoImageTemplate


@dataclass(frozen=True)
class TemplateDefinition:
    id: str
    name: str
    description: str
    template: object


_TEMPLATES = {
    "technical-no-image": TemplateDefinition(
        id="technical-no-image",
        name="技术类ppt（不生成图片）",
        description="适合技术文档、流程说明、架构讲解；只使用文字、图形、箭头和截图占位。",
        template=TechnicalNoImageTemplate(),
    )
}


def list_templates() -> list[TemplateDefinition]:
    return list(_TEMPLATES.values())


def get_template(template_id: str) -> TemplateDefinition:
    if template_id not in _TEMPLATES:
        raise KeyError(f"未知模板：{template_id}")
    return _TEMPLATES[template_id]
