from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
import os
from pathlib import Path
from xml.etree import ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


def inject_fade_animations(pptx_path: Path) -> tuple[int, int]:
    if os.environ.get("PPT_ANIMATION_ENGINE", "").lower() == "com":
        try:
            return _inject_with_powerpoint(pptx_path)
        except Exception:
            return _inject_with_xml(pptx_path)
    return _inject_with_xml(pptx_path)


def _inject_with_powerpoint(pptx_path: Path) -> tuple[int, int]:
    import win32com.client

    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    ppt.Visible = True
    presentation = ppt.Presentations.Open(str(pptx_path.resolve()), False, False, False)
    total_objects = 0
    total_slides = 0
    try:
        for slide in presentation.Slides:
            sequence = slide.TimeLine.MainSequence
            for index in range(sequence.Count, 0, -1):
                sequence.Item(index).Delete()

            shapes = [shape for shape in slide.Shapes if str(shape.Name).startswith("anim_")]
            shapes.sort(key=lambda shape: str(shape.Name))
            if shapes:
                total_slides += 1
            for shape in shapes:
                effect = sequence.AddEffect(shape, 10, 0, 1)
                effect.Timing.Duration = 0.28
                total_objects += 1
        presentation.Save()
    finally:
        presentation.Close()
        ppt.Quit()
    return total_slides, total_objects


def _inject_with_xml(pptx_path: Path) -> tuple[int, int]:
    slide_re = re.compile(r"ppt/slides/slide\d+\.xml$")
    total_objects = 0
    total_slides = 0

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp_path = Path(tmp.name)

    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if slide_re.match(item.filename):
                data, count = _inject_slide(data)
                if count:
                    total_slides += 1
                    total_objects += count
            zout.writestr(item, data)

    shutil.move(str(tmp_path), pptx_path)
    return total_slides, total_objects


def _q(tag: str) -> str:
    return f"{{{P_NS}}}{tag}"


def _elm(tag: str, attrs: dict[str, str] | None = None, text: str | None = None) -> ET.Element:
    element = ET.Element(_q(tag), attrs or {})
    if text is not None:
        element.text = text
    return element


def _add(parent: ET.Element, tag: str, attrs: dict[str, str] | None = None, text: str | None = None) -> ET.Element:
    child = _elm(tag, attrs, text)
    parent.append(child)
    return child


def _inject_slide(xml_bytes: bytes) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    anim_shapes = _collect_anim_shapes(root)
    if not anim_shapes:
        return xml_bytes, 0

    for old_timing in list(root.findall(_q("timing"))):
        root.remove(old_timing)

    timing = _make_timing(anim_shapes)
    ext_lst = root.find(_q("extLst"))
    if ext_lst is not None:
        root.insert(list(root).index(ext_lst), timing)
    else:
        root.append(timing)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True, short_empty_elements=True), len(anim_shapes)


def _collect_anim_shapes(root: ET.Element) -> list[str]:
    pairs = []
    for c_nv_pr in root.iter(_q("cNvPr")):
        name = c_nv_pr.attrib.get("name", "")
        shape_id = c_nv_pr.attrib.get("id")
        if name.startswith("anim_") and shape_id:
            pairs.append((name, shape_id))
    return [shape_id for _, shape_id in sorted(pairs, key=lambda item: item[0])]


def _make_timing(shape_ids: list[str]) -> ET.Element:
    timing = _elm("timing")
    tn_lst = _add(timing, "tnLst")
    root_par = _add(tn_lst, "par")
    root_ctn = _add(root_par, "cTn", {"id": "1", "dur": "indefinite", "restart": "never", "nodeType": "tmRoot"})
    root_children = _add(root_ctn, "childTnLst")

    seq = _add(root_children, "seq", {"concurrent": "1", "nextAc": "seek"})
    seq_ctn = _add(seq, "cTn", {"id": "2", "dur": "indefinite", "nodeType": "mainSeq"})
    seq_children = _add(seq_ctn, "childTnLst")

    next_id = 3
    for shape_id in shape_ids:
        effect_par = _add(seq_children, "par")
        effect_ctn = _add(effect_par, "cTn", {"id": str(next_id), "fill": "hold"})
        next_id += 1
        st = _add(effect_ctn, "stCondLst")
        _add(st, "cond", {"delay": "indefinite"})
        child = _add(effect_ctn, "childTnLst")

        click_par = _add(child, "par")
        click_ctn = _add(
            click_par,
            "cTn",
            {
                "id": str(next_id),
                "presetID": "10",
                "presetClass": "entr",
                "presetSubtype": "0",
                "fill": "hold",
                "nodeType": "clickEffect",
            },
        )
        next_id += 1
        st2 = _add(click_ctn, "stCondLst")
        _add(st2, "cond", {"delay": "0"})
        click_children = _add(click_ctn, "childTnLst")

        set_el = _add(click_children, "set")
        c_bhvr = _add(set_el, "cBhvr")
        _add(c_bhvr, "cTn", {"id": str(next_id), "dur": "1", "fill": "hold"})
        next_id += 1
        tgt_el = _add(c_bhvr, "tgtEl")
        _add(tgt_el, "spTgt", {"spid": shape_id})
        attr_lst = _add(c_bhvr, "attrNameLst")
        _add(attr_lst, "attrName", text="style.visibility")
        to = _add(set_el, "to")
        _add(to, "strVal", {"val": "visible"})

        anim = _add(click_children, "animEffect", {"transition": "in", "filter": "fade"})
        c_bhvr2 = _add(anim, "cBhvr")
        _add(c_bhvr2, "cTn", {"id": str(next_id), "dur": "280"})
        next_id += 1
        tgt_el2 = _add(c_bhvr2, "tgtEl")
        _add(tgt_el2, "spTgt", {"spid": shape_id})

    for cond_name in ("prevCondLst", "nextCondLst"):
        cond_lst = _add(seq, cond_name)
        event = "onPrev" if cond_name == "prevCondLst" else "onNext"
        cond = _add(cond_lst, "cond", {"evt": event, "delay": "0"})
        tgt = _add(cond, "tgtEl")
        _add(tgt, "sldTgt")

    bld_lst = _add(timing, "bldLst")
    for shape_id in shape_ids:
        _add(bld_lst, "bldP", {"spid": shape_id, "grpId": "0"})
    return timing
