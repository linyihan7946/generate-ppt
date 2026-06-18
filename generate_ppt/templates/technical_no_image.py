from __future__ import annotations

import re

from generate_ppt.document_loader import SourceDocument
from generate_ppt.slide_model import DeckSpec, SlideKind, SlideSpec


FLOW_WORDS = ("流程", "步骤", "阶段", "链路", "调用", "直到", "循环", "先", "再", "然后")
ARCH_WORDS = ("架构", "模块", "系统", "组件", "服务", "上下游", "依赖", "数据流", "调用链")
COMPARE_WORDS = ("对比", "区别", "优缺点", "问题", "方案", "提升", "减少", "但是", "而是")
SCREENSHOT_WORDS = ("界面", "页面", "配置", "操作", "点击", "输入", "上传", "截图", "控制台", "终端", "IDE")
PROBLEM_WORDS = ("问题", "原因", "方案", "结果", "为什么", "怎么解决")


class TechnicalNoImageTemplate:
    id = "technical-no-image"
    name = "技术类ppt（不生成图片）"

    def build(self, document: SourceDocument) -> DeckSpec:
        units = _split_units(document.text)
        slides: list[SlideSpec] = [
            SlideSpec(
                kind=SlideKind.COVER,
                title=document.title,
                subtitle="技术文档讲解",
                items=["背景", "原理", "落地"],
                accent="blue",
            )
        ]

        overview = _extract_overview(units)
        if overview:
            slides.append(SlideSpec(kind=SlideKind.FLOW, title="先看整体路径", subtitle="原文逻辑", items=overview, accent="blue"))

        for idx, unit in enumerate(units):
            slide = self._unit_to_slide(unit, idx)
            if slide:
                slides.append(slide)

        if len(slides) > 2:
            slides.insert(
                min(2, len(slides)),
                SlideSpec(
                    kind=SlideKind.PROBLEM_CHAIN,
                    title="这页讲什么问题？",
                    subtitle="问题 → 原因 → 方案 → 结果",
                    items=_deck_problem_items(document.text),
                    accent="blue",
                ),
            )
            for offset, insight in enumerate(_deck_insight_slides(document.text), start=3):
                slides.insert(min(offset, len(slides)), insight)

        summary_items = _summary_items(slides)
        slides.append(
            SlideSpec(
                kind=SlideKind.SUMMARY,
                title="完整闭环",
                items=summary_items,
                highlight="讲清楚，才算讲完",
                accent="yellow",
            )
        )
        return DeckSpec(title=document.title, subtitle="技术文档讲解", template_id=self.id, slides=_limit_slides(slides))

    def _unit_to_slide(self, unit: str, index: int) -> SlideSpec | None:
        title, body = _title_and_body(unit)
        if not title:
            return None
        items = _keywords(body or title, max_items=6)
        text = f"{title}\n{body}"

        if _has_any(text, SCREENSHOT_WORDS):
            return SlideSpec(
                kind=SlideKind.SCREENSHOT,
                title=title,
                subtitle=f"第 {index + 1} 段",
                items=items[:2] or ["操作位置", "关键状态"],
                screenshot_hint=_screenshot_hint(text),
                accent="cyan",
            )
        if _has_problem_chain(text):
            return SlideSpec(kind=SlideKind.PROBLEM_CHAIN, title=title, subtitle="问题 → 原因 → 方案 → 结果", items=_problem_items(text), accent="blue")
        if _has_any(text, FLOW_WORDS):
            return SlideSpec(kind=SlideKind.FLOW, title=title, subtitle=f"第 {index + 1} 段", items=items or _split_short(body), accent="green")
        if _has_any(text, ARCH_WORDS):
            relation_items = [title] + (items[:4] or ["输入", "输出", "数据流", "影响范围"])
            return SlideSpec(kind=SlideKind.RELATION, title=title, subtitle="关系图", items=relation_items, accent="orange")
        if _has_any(text, COMPARE_WORDS) and len(items) >= 2:
            return SlideSpec(kind=SlideKind.COMPARE, title=title, subtitle="对比 / 取舍", items=items[:4], highlight=_conclusion(body), accent="yellow")
        if len(items) <= 2:
            return SlideSpec(kind=SlideKind.SECTION, title=title, subtitle="重点", items=items, highlight=_conclusion(body), accent="yellow")
        return SlideSpec(kind=SlideKind.BULLETS, title=title, subtitle=f"第 {index + 1} 段", items=items, accent="blue")


def _split_units(text: str) -> list[str]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    units: list[str] = []
    current: list[str] = []

    for index, line in enumerate(lines):
        is_heading = _looks_like_heading(line, index)
        if is_heading and current and len(current) >= 2:
            units.append("\n".join(current))
            current = [line]
        else:
            current.append(line)
            if len(current) >= 5:
                units.append("\n".join(current))
                current = []
    if current:
        units.append("\n".join(current))
    return [unit for unit in units if len(unit.strip()) > 3]


def _looks_like_heading(line: str, index: int = 0) -> bool:
    if index == 0:
        return True
    if re.match(r"^#{1,4}\s+", line):
        return True
    if re.match(r"^(\d+[\.\、]|第.+[章节步])", line):
        return True
    if len(line) <= 18 and _has_any(line, FLOW_WORDS + ARCH_WORDS + PROBLEM_WORDS):
        return True
    return False


def _title_and_body(unit: str) -> tuple[str, str]:
    lines = [line.strip(" #\t") for line in unit.splitlines() if line.strip()]
    if not lines:
        return "", ""
    title = _shorten(lines[0], 18)
    body = "\n".join(lines[1:]) if len(lines) > 1 else lines[0]
    return title, body


def _extract_overview(units: list[str]) -> list[str]:
    candidates = []
    for unit in units[:8]:
        title, body = _title_and_body(unit)
        if title:
            candidates.append(title)
        if len(candidates) >= 7:
            break
    return candidates


def _keywords(text: str, max_items: int = 6) -> list[str]:
    raw = re.split(r"[\n。；;，,、:：\|]+", text)
    items = []
    seen = set()
    for item in raw:
        clean = _shorten(item.strip(" -•\t"), 16)
        if len(clean) < 2 or clean in seen:
            continue
        seen.add(clean)
        items.append(clean)
        if len(items) >= max_items:
            break
    return items


def _split_short(text: str) -> list[str]:
    return _keywords(text, 6) or [_shorten(text, 16)]


def _has_any(text: str, words: tuple[str, ...]) -> bool:
    return any(word.lower() in text.lower() for word in words)


def _has_problem_chain(text: str) -> bool:
    return sum(1 for word in PROBLEM_WORDS if word in text) >= 2


def _problem_items(text: str) -> list[str]:
    items = _keywords(text, 8)
    defaults = ["现象不清楚", "缺少上下文", "结构化拆解", "验证通过"]
    return (items + defaults)[:4]


def _screenshot_hint(text: str) -> str:
    if "配置" in text:
        return "预留：配置页面截图"
    if "终端" in text or "控制台" in text:
        return "预留：终端 / 控制台截图"
    if "IDE" in text or "代码" in text:
        return "预留：IDE / 代码界面截图"
    return "预留：软件界面 / 操作步骤截图"


def _conclusion(text: str) -> str:
    items = _keywords(text, 4)
    return items[-1] if items else ""


def _summary_items(slides: list[SlideSpec]) -> list[str]:
    result = []
    for slide in slides[1:8]:
        if slide.title not in result:
            result.append(_shorten(slide.title, 8))
    return result[:7] or ["输入", "理解", "拆解", "生成", "验证"]


def _deck_problem_items(text: str) -> list[str]:
    items = _keywords(text, 12)
    problem = _pick_first(items, ("问题", "需求", "目标", "痛点")) or "信息太散"
    reason = _pick_first(items, ("原因", "上下文", "依赖", "调用链", "数据流")) or "缺少结构化理解"
    solution = _pick_first(items, ("方案", "流程", "步骤", "计划", "架构")) or "按流程拆解"
    result = _pick_first(items, ("结果", "成功", "通过", "效果", "提升")) or "验证后交付"
    return [problem, reason, solution, result]


def _deck_insight_slides(text: str) -> list[SlideSpec]:
    items = _keywords(text, 12)
    reason = _pick_first(items, ("原因", "上下文", "依赖", "调用链", "数据流")) or "先看清关系"
    solution = _pick_first(items, ("方案", "流程", "步骤", "计划", "架构")) or "按步骤推进"
    result = _pick_first(items, ("结果", "成功", "通过", "效果", "提升", "减少")) or "验证通过"
    return [
        SlideSpec(kind=SlideKind.SECTION, title="为什么这样？", subtitle="原理", highlight=reason, items=[reason], accent="blue"),
        SlideSpec(kind=SlideKind.SECTION, title="怎么解决？", subtitle="方案", highlight=solution, items=[solution], accent="green"),
        SlideSpec(kind=SlideKind.SECTION, title="最终效果？", subtitle="结论", highlight=result, items=[result], accent="yellow"),
    ]


def _pick_first(items: list[str], words: tuple[str, ...]) -> str:
    for item in items:
        if _has_any(item, words):
            return item
    return ""


def _limit_slides(slides: list[SlideSpec]) -> list[SlideSpec]:
    if len(slides) <= 28:
        return slides
    return slides[:27] + [slides[-1]]


def _shorten(text: str, max_len: int) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= max_len else text[: max_len - 1] + "…"
