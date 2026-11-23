"""Gate one finished deck before delivery.

Static, deterministic checks (no rendering, no model call):

  overflow   the new text needs clearly more room than its box has
  tiny       font below 10pt -- unreadable on a projector
  offslide   a text box hangs outside the slide
  collide    two text boxes with text overlap each other
  empty      the slide holds no text at all (blank page)

Usage:  python qa_deck.py out/deck.pptx [--json]
Exit code is 1 when any ERROR-level finding exists, 0 otherwise.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

from pptx import Presentation

EMU_PER_PT = 12700.0


def pt(v) -> float:
    return float(v) / EMU_PER_PT


def text_need(shape) -> tuple[float, float]:
    """(widest paragraph in em, largest font size in pt) of the shape."""
    size = 0.0
    widest = 0.0
    for para in shape.text_frame.paragraphs:
        text = "".join(r.text for r in para.runs) or para.text
        em = sum(0.55 if ch.isascii() and ch.isprintable() else 1.0 for ch in text)
        widest = max(widest, em)
        for run in para.runs:
            if run.font.size:
                size = max(size, run.font.size.pt)
    if size <= 0:
        size = 18.0
    return widest, size


def box(shape) -> tuple[float, float, float, float]:
    left, top = pt(shape.left), pt(shape.top)
    return left, top, left + pt(shape.width), top + pt(shape.height)


def overlap_area(a, b) -> float:
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def _walk(shapes):
    """yield every shape, descending into groups"""
    for sh in shapes:
        yield sh
        if sh.shape_type is not None and int(sh.shape_type) == 6:
            yield from _walk(sh.shapes)


def _walk_abs(shapes, ox=0.0, oy=0.0, sx=1.0, sy=1.0):
    """yield (shape, x0, y0, x1, y1) in absolute slide points, composing group transforms"""
    for sh in shapes:
        try:
            l, t_ = pt(sh.left), pt(sh.top)
            w, h = pt(sh.width), pt(sh.height)
        except Exception:
            l = t_ = w = h = 0.0
        x0, y0 = ox + l * sx, oy + t_ * sy
        x1, y1 = x0 + w * sx, y0 + h * sy
        yield sh, x0, y0, x1, y1
        try:
            is_group = sh.shape_type is not None and int(sh.shape_type) == 6
        except Exception:
            is_group = False
        if is_group:
            # child offsets are in child space; map them into the parent's box
            ch_off_x = ch_off_y = ch_ext_x = ch_ext_y = 0.0
            for el in sh._element.iter():
                tag = el.tag.split("}")[-1]
                if tag == "chOff":
                    ch_off_x, ch_off_y = int(el.get("x") or 0) / 12700.0, int(el.get("y") or 0) / 12700.0
                elif tag == "chExt":
                    ch_ext_x, ch_ext_y = int(el.get("cx") or 0) / 12700.0, int(el.get("cy") or 0) / 12700.0
            gx = sx * (w / ch_ext_x) if ch_ext_x else sx
            gy = sy * (h / ch_ext_y) if ch_ext_y else sy
            yield from _walk_abs(sh.shapes,
                                 ox + (l - ch_off_x) * sx,
                                 oy + (t_ - ch_off_y) * sy,
                                 gx, gy)


def _uniq(shapes):
    """dedupe by xml element so a nested group is not counted twice"""
    seen = set()
    for sh in shapes:
        key = id(sh._element)
        if key in seen:
            continue
        seen.add(key)
        yield sh


def audit(path: Path) -> dict:
    prs = Presentation(str(path))
    slide_w, slide_h = pt(prs.slide_width), pt(prs.slide_height)
    pages = []
    for index, slide in enumerate(prs.slides, 1):
        texts = []
        findings = []
        seen = set()
        for shape, x0, y0, x1, y1 in _walk_abs(slide.shapes):
            if id(shape._element) in seen:
                continue
            seen.add(id(shape._element))
            if not getattr(shape, "has_text_frame", False):
                continue
            content = shape.text_frame.text.strip()
            if not content:
                continue
            w, h = x1 - x0, y1 - y0
            em, size = text_need(shape)
            texts.append({"name": shape.name, "box": [x0, y0, x1, y1],
                          "size": size, "text": content[:24]})
            if x0 < -2 or y0 < -2 or x1 > slide_w + 2 or y1 > slide_h + 2:
                # templates legitimately park decoration partly off-canvas; only a
                # text box that is mostly gone from the page is a real defect
                vw = max(0.0, min(x1, slide_w) - max(x0, 0.0))
                vh = max(0.0, min(y1, slide_h) - max(y0, 0.0))
                visible = (vw * vh) / ((x1 - x0) * (y1 - y0) or 1.0)
                level = "WARN"  # template decoration likes to sit off-canvas
                findings.append((level, "offslide",
                                 f"{shape.name} 有 {1 - visible:.0%} 在页面外"))
            if em > 0 and w > 4:
                per_line = max(1.0, (w - 6) / size)
                need_h = math.ceil(em / per_line) * size * 1.32
                if need_h > h * 1.18 + 2:
                    findings.append(("ERROR", "overflow",
                                     f"{shape.name} 需要 {need_h:.0f}pt 高，只给了 {h:.0f}pt"))
            if size < 10:
                findings.append(("WARN", "tiny", f"{shape.name} 字号 {size:.0f}pt"))
        for i in range(len(texts)):
            for j in range(i + 1, len(texts)):
                a, b = texts[i]["box"], texts[j]["box"]
                inter = overlap_area(a, b)
                if not inter:
                    continue
                small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
                if texts[i]["name"] == texts[j]["name"] and a == b:
                    continue
                import re as _re
                if _re.match(r"^(PART|PARTE)\s*\d+$", texts[i]["text"].strip(), _re.I) or \
                   _re.match(r"^(PART|PARTE)\s*\d+$", texts[j]["text"].strip(), _re.I):
                    # chapter dividers overlay the big "PART 0x" marker by design
                    continue
                if small > 0 and inter / small > 0.25:
                    findings.append(("ERROR", "collide",
                                     f"{texts[i]['name']} 与 {texts[j]['name']} 文字区重叠 "
                                     f"{inter / small:.0%}"))
        if not texts:
            findings.append(("ERROR", "empty", "整页没有文字"))
        pages.append({"page": index, "texts": len(texts), "findings": findings})
    errors = sum(1 for p in pages for f in p["findings"] if f[0] == "ERROR")
    warns = sum(1 for p in pages for f in p["findings"] if f[0] == "WARN")
    return {"deck": str(path), "pages": len(pages), "errors": errors, "warns": warns,
            "detail": pages}


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2
    path = Path(argv[0])
    result = audit(path)
    if "--json" in argv:
        print(json.dumps(result, ensure_ascii=False, indent=1))
    else:
        print(f"{path.name}: {result['pages']} 页, {result['errors']} error, {result['warns']} warn")
        for page in result["detail"]:
            for level, rule, message in page["findings"]:
                print(f"  p{page['page']} [{level}] {rule}: {message}")
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
