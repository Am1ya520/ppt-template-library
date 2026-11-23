#!/usr/bin/env python3
"""pptl.py - native PPTX page-repository toolkit.

Build a library from the user's own .pptx decks (index them, classify pages,
extract slot geometry) and assemble new decks by reusing those native pages,
so nothing is re-authored and the result looks like the source template.

Subcommands
  index   <source.pptx> [-o <index_dir>]      write <id>.json page inventory
  fill    <source.pptx> <plan.json> -o <out.pptx>
  slots   <source.pptx> [--slide N]           print page/slot inventory
"""
from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A_NS}

TOC_RE = re.compile(r"目\s*录|contents|content|目录页|agenda", re.I)
ENDING_RE = re.compile(r"谢谢|感谢|thank|q\s*&\s*a|结束|the end", re.I)
CHAPTER_RE = re.compile(r"^\s*(part|chapter|第[一二三四五六七八九十0-9]+[章节部分]|0\d)\b", re.I)

YEAR_RE = re.compile(r"^\s*(19|20)\d{2}\s*$|^\s*(19|20)\d[Xx]\s*$")
NUM_RE = re.compile(r"^\s*\d{1,3}\s*$")
DESIGN_RE = re.compile(r"^\s*(contents?|agenda|part\s*[\divx]+|section|chapter|\.|-)\s*$", re.I)
LOGO_RE = re.compile(r"^\s*(logo|公司标志|企业标识|公司logo)\s*$", re.I)
FURNITURE_RE = re.compile(
    r"^\s*(目\s*录|contents?|agenda|谢谢|感谢|thank\s*you|many\s*thanks|the\s*end|\.\.\.)\s*$",
    re.I,
)
FILLER_RE = re.compile(
    r"此处|点击(此处|输入|添加)|输入(文本|标题|文字)|添加(文本|标题|文字|内容)|"
    r"标题暨内容|添加标题|在此|键入|请输入|示例(文字|文本)|占位|"
    r"placeholder|lorem|ipsum|单击此处|标题文字",
    re.I,
)
MARKETING_RE = re.compile(
    r"模板|实战|适用于|公众号|素材|请替换|示例文字|点击此处|此处输入|placeholder|lorem|"
    r"公司标志|标志替代|LOGO|企业标识",
    re.I,
)
EN_FILLER_RE = re.compile(
    r"content\s*title|your\s*(title|text)\s*here|insert\s*(title|text)|click\s*(to|here)|sample\s*text|text\s*title|placeholder|\bLOREM\b", re.I)


def is_filler(text: str) -> bool:
    return bool(FILLER_RE.search(text) or EN_FILLER_RE.search(text))


DARK_TEXT = "1F1F1F"
WHITE_TEXT = "FFFFFF"


# ---------------------------------------------------------------- colours

def resolve_theme(prs) -> dict[str, str]:
    """dk1/lt1/dk2/lt2/accent1..6 of the deck's theme, as RRGGBB strings."""
    out = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6"}
    for i in range(1, 7):
        out[f"accent{i}"] = "4472C4"
    try:
        part = prs.slide_masters[0].part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        root = etree.fromstring(part.blob)
    except Exception:
        return out
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return out
    for el in scheme:
        tag = etree.QName(el).localname
        if tag not in out:
            continue
        child = el[0] if len(el) else None
        if child is None:
            continue
        val = child.get("val") or child.get("lastClr")
        if val and re.fullmatch(r"[0-9A-Fa-f]{6}", val):
            out[tag] = val.upper()
    return out


def _theme_name(font_color) -> str | None:
    try:
        tc = font_color.theme_color
    except Exception:
        return None
    if tc is None:
        return None
    name = str(tc).split(".")[-1].split(" ")[0]
    return name.lower()


def run_color(run) -> str:
    """'rgb:RRGGBB' | 'theme:name' | '' (inherit)."""
    font = run.font
    col = font.color
    try:
        if col is None or col.type is None:
            return ""
    except Exception:
        return ""
    try:
        if col.type == 1:  # RGB
            return "rgb:" + str(col.rgb).upper()
    except Exception:
        pass
    name = _theme_name(col)
    return f"theme:{name}" if name else ""


def color_hex(spec: str, theme: dict[str, str]) -> str | None:
    if not spec:
        return None
    if spec.startswith("rgb:"):
        return spec[4:]
    if spec.startswith("theme:"):
        return theme.get(spec[6:])
    return None


def luma(hex_rgb: str | None) -> float | None:
    if not hex_rgb or not re.fullmatch(r"[0-9A-Fa-f]{6}", hex_rgb):
        return None
    r, g, b = (int(hex_rgb[i : i + 2], 16) / 255.0 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


# ---------------------------------------------------------------- inventory

def emu2pt(v) -> float:
    return round(Emu(v).pt, 1) if v is not None else 0.0


@dataclass
class ShapeInfo:
    shape_id: int
    name: str
    kind: str
    placeholder: str | None
    left: float
    top: float
    width: float
    height: float
    max_pt: float
    align: str | None
    paragraphs: list[str] = field(default_factory=list)
    colors: list[str] = field(default_factory=list)
    rot: float = 0.0

    @property
    def area(self) -> float:
        return max(self.width, 0) * max(self.height, 0)

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height

    @property
    def text(self) -> str:
        return " ".join(p for p in self.paragraphs if p.strip())

    def overlaps(self, other: "ShapeInfo", pad: float = 0.0) -> bool:
        return not (
            self.right + pad < other.left
            or other.right + pad < self.left
            or self.bottom + pad < other.top
            or other.bottom + pad < self.top
        )

    def to_dict(self) -> dict:
        return {
            "id": self.shape_id,
            "name": self.name,
            "kind": self.kind,
            "placeholder": self.placeholder,
            "box": [self.left, self.top, self.width, self.height],
            "max_pt": self.max_pt,
            "align": self.align,
            "paragraphs": self.paragraphs,
            "colors": self.colors,
            "rot": self.rot,
        }


def _safe_shape_type(shape):
    """shape_type raises on some legacy/converted shapes; treat them as opaque."""
    try:
        return shape.shape_type
    except Exception:
        return None


def _iter_text_shapes(shapes):
    for shape in shapes:
        if _safe_shape_type(shape) == 6:  # group
            yield from _iter_text_shapes(shape.shapes)
            continue
        if not shape.has_text_frame:
            continue
        if shape.width is None or shape.height is None:
            continue
        yield shape


def read_shapes(slide) -> list[ShapeInfo]:
    out: list[ShapeInfo] = []
    for shape in _iter_text_shapes(slide.shapes):
        paras, max_pt, align, colors = [], 0.0, None, []
        for para in shape.text_frame.paragraphs:
            text = "".join(run.text for run in para.runs) or para.text
            paras.append(text.strip())
            for run in para.runs:
                if run.font.size:
                    max_pt = max(max_pt, run.font.size.pt)
                spec = run_color(run)
                if spec and spec not in colors:
                    colors.append(spec)
            if align is None and para.alignment is not None:
                align = str(para.alignment).split(" ")[0].split(".")[-1]
        ph = shape.placeholder_format.type if shape.is_placeholder else None
        out.append(
            ShapeInfo(
                shape_id=shape.shape_id,
                name=shape.name,
                kind=str(_safe_shape_type(shape)).split(" ")[0] if _safe_shape_type(shape) else "TEXT",
                placeholder=str(ph).split(" ")[0].split(".")[-1] if ph else None,
                left=emu2pt(shape.left),
                top=emu2pt(shape.top),
                width=emu2pt(shape.width),
                height=emu2pt(shape.height),
                max_pt=round(max_pt, 1),
                align=align,
                paragraphs=[p for p in paras],
                colors=colors,
                rot=round(float(getattr(shape, "rotation", 0.0) or 0.0), 1),
            )
        )
    return out


def page_type(idx: int, total: int, slide, shapes: list[ShapeInfo]) -> str:
    texts = [p for s in shapes for p in s.paragraphs if p.strip()]
    joined = " | ".join(texts[:6])
    biggest = max((s.max_pt for s in shapes), default=0)
    if idx == 0:
        return "cover"
    if idx == total - 1:
        return "ending"
    if ENDING_RE.search(joined) and len(texts) <= 6:
        return "ending"
    if TOC_RE.search(joined) and len(texts) <= 12:
        return "toc"
    if len(texts) <= 4 and biggest >= 30:
        return "chapter"
    if texts and CHAPTER_RE.search(texts[0]):
        return "chapter"
    n_with_text = len([s for s in shapes if any(p.strip() for p in s.paragraphs)])
    if n_with_text >= 6:
        return "dense"
    return "content"


def deck_meta(prs) -> dict:
    w, h = prs.slide_width, prs.slide_height
    return {
        "slides": len(prs.slides),
        "width_pt": emu2pt(w),
        "height_pt": emu2pt(h),
        "ratio": round(w / h, 3) if h else 0,
    }


# ---------------------------------------------------------------- classification

def classify(slot: ShapeInfo) -> str:
    """headline | text | year | num | design | furniture | filler | empty | rotated."""
    text = slot.text.strip()
    if not text:
        return "empty"
    if abs(slot.rot) > 2:
        return "rotated"
    if FURNITURE_RE.match(text):
        return "furniture"
    if YEAR_RE.match(text):
        return "year"
    if NUM_RE.match(text) or DESIGN_RE.match(text):
        return "num"
    if re.fullmatch(r"(logo|LOGO|公司标志|企业标识|公司logo)", text, re.I):
        return "design"
    if (is_filler(text) or MARKETING_RE.search(text)) and len(text) <= 60:
        return "filler"
    if len(text) <= 3:
        return "num"
    return "text"


def role_map(slots: list[ShapeInfo]) -> dict[str, list[ShapeInfo]]:
    roles: dict[str, list[ShapeInfo]] = {}
    for s in slots:
        roles.setdefault(classify(s), []).append(s)
    return roles


def pick_headline(
    slots: list[ShapeInfo],
    prefer_top: float | None = None,
    center: float | None = None,
    allow_year: bool = False,
) -> ShapeInfo | None:
    def is_cand(s: ShapeInfo) -> bool:
        c = classify(s)
        if allow_year and c == "year":
            return True
        return c == "text" or (c == "filler" and s.width >= 120 and s.max_pt >= 14)

    def score(s: ShapeInfo) -> float:
        base = min(s.max_pt, 72.0) + min(s.area / 4000.0, 30.0) + min(len(s.text), 24) * 0.4
        if center is not None:
            span = max(center, 1.0)
            drift = abs((s.left + s.width / 2) - center) / span
            base += 12.0 * max(0.0, 1.0 - drift)
        return base

    cands = [s for s in slots if is_cand(s)]
    if not cands:
        return None
    # a title is set in (near) the largest type on the page; body rows with a
    # similar score but half the size are not titles
    top_pt = max(s.max_pt for s in cands)
    if top_pt > 0:
        cands = [s for s in cands if s.max_pt >= 0.75 * top_pt]
    best = max(cands, key=score)
    if prefer_top is not None:
        band = [s for s in cands if s.top <= prefer_top]
        if band:
            inband = max(band, key=score)
            # a small caption sitting high on the page is not the page title
            if inband.max_pt >= 0.6 * best.max_pt:
                best = inband
    return best


def title_capacity(slot: ShapeInfo, floor: float = 0.62) -> int:
    """How many CJK characters fit in a slot while staying above the font floor."""
    if slot.max_pt <= 0:
        return 0
    size = slot.max_pt * floor
    lines = max(1, int(max(slot.height - 2, 14) / (size * 1.32)))
    per_line = max(1, int(max(slot.width - 6, 24) / size))
    return lines * per_line


def read_order(s: ShapeInfo) -> tuple[float, float]:
    return (round(s.top / 10.0), round(s.left / 6.0))


# ---------------------------------------------------------------- rows

def band_rows(slots: list[ShapeInfo], tol: float = 16.0) -> list[list[ShapeInfo]]:
    """Cluster text slots into horizontal rows (one band per visual row)."""
    items = sorted([s for s in slots if s.text.strip()], key=lambda s: (s.top, s.left))
    rows: list[list[ShapeInfo]] = []
    for s in items:
        center = s.top + s.height / 2
        if rows:
            band = rows[-1]
            band_center = sum(m.top + m.height / 2 for m in band) / len(band)
            if abs(center - band_center) <= tol:
                band.append(s)
                continue
        rows.append([s])
    return rows


def row_label(row: list[ShapeInfo]) -> ShapeInfo | None:
    cands = [s for s in row if classify(s) == "text"]
    if not cands:
        return None
    return max(cands, key=lambda s: (s.width, len(s.text)))


# ---------------------------------------------------------------- text writing

def _fill_paragraphs(tf, values: list[str]) -> None:
    if not values:
        return
    existing = list(tf.paragraphs)
    first = existing[0]
    template_run = first.runs[0] if first.runs else None
    for i, value in enumerate(values):
        if i < len(existing):
            para = existing[i]
            runs = list(para.runs)
            if runs:
                runs[0].text = value
                for extra in runs[1:]:
                    extra.text = ""
            else:
                para.add_run().text = value
        else:
            para = tf.add_paragraph()
            run = para.add_run()
            run.text = value
            if template_run is not None:
                if template_run.font.size is not None:
                    run.font.size = template_run.font.size
                run.font.bold = template_run.font.bold
                run.font.name = template_run.font.name
    for para in list(tf.paragraphs)[len(values):]:
        para._p.getparent().remove(para._p)


def _text_width_em(text: str) -> float:
    return sum(1.08 if ord(c) > 0x2E80 else 0.56 for c in text)


def fit_font(tf, slot: ShapeInfo, floor: float = 0.5) -> float:
    """Shrink runs until the text fits; returns the scale actually used."""
    if slot.max_pt <= 0:
        return 1.0
    lines = [p.text for p in tf.paragraphs]
    if not lines:
        return 1.0
    usable_w = max(slot.width - 6, 24)
    orig_lines = max(1, len([p for p in slot.paragraphs if p.strip()]))
    # never let text grow past the footprint the original text block occupied:
    # template boxes are often much taller than the text they actually hold
    usable_h = min(max(slot.height - 2, 14), orig_lines * slot.max_pt * 1.35)
    best = slot.max_pt
    size = slot.max_pt
    for _ in range(18):
        need = 0.0
        for line in lines:
            em = max(_text_width_em(line), 1.0)
            cpl = max(usable_w / size, 1.0)
            need += max(1, math.ceil(em / cpl)) * size * 1.32
        if need <= usable_h + 1.0:
            break
        if best <= slot.max_pt * floor:
            break
        best *= 0.94
        size = best
    ratio = best / slot.max_pt
    if ratio >= 0.995:
        # short text in a big frame looks lost on airy minimal templates: grow it
        # (never past the frame, and never past a sane reading size)
        live = [p.text for p in tf.paragraphs if p.text.strip()]
        if live and sum(_text_width_em(x) for x in live) <= 72:
            widest = max(_text_width_em(x) for x in live)
            by_w = (usable_w / widest) if widest else 0.0
            by_h = (max(slot.height - 4, 12) / (len(live) * 1.45)) if live else 0.0
            target = min(slot.max_pt * 1.22, by_w / 1.15, by_h, 40.0)
            # growing must not turn one line into two: the neighbouring rows in
            # these templates sit close together, and a wrap lands on top of them
            if widest and target * widest * 1.15 > usable_w:
                target = usable_w / (widest * 1.15)
            if len(live) > 1 and target * len(live) * 1.45 > slot.height:
                target = slot.height / (len(live) * 1.45)
            if target > slot.max_pt * 1.15:
                for para in tf.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(round(target, 2))
                return target / slot.max_pt
        return 1.0
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(round(best, 1))
    return ratio


def widen_headline(shape, slot: ShapeInfo, slide_w: float, right_margin: float = 40.0) -> None:
    """A title box narrower than its type deserves is widened into the free band."""
    if slot.width >= 300 or not shape.has_text_frame:
        return
    room = slide_w - slot.left - right_margin
    target = min(max(room, slot.width), 620.0)
    if target > slot.width + 24:
        shape.width = Emu(int(target * 12700))
        slot.width = target


def normalize_alignment(shape) -> None:
    """Undo template justification/spacing that stretches short text across the box.

    Templates often ship titles with algn="dist" (or large character spacing) so
    their own long sentence fills the frame. Replace with our shorter text and the
    letters fly apart, which reads as a layout bug.
    """
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for para in shape.text_frame.paragraphs:
        pPr = para._p.find(ns_a + "pPr")
        if pPr is not None:
            algn = (pPr.get("algn") or "").lower()
            if algn in ("dist", "just", "justlow", "thaiDist", "thaidist"):
                # "distributed" text is meant to span the frame; our shorter line
                # should sit centred in it rather than jammed to the left edge
                pPr.set("algn", "ctr")
        for run in para.runs:
            rPr = run._r.find(ns_a + "rPr")
            if rPr is None:
                continue
            try:
                spc = int(rPr.get("spc") or 0)
            except ValueError:
                spc = 0
            if abs(spc) > 60:  # ~0.6pt per char adds up fast
                rPr.set("spc", "0")



def stabilize_text(shape) -> None:
    """Keep our text inside the frame the template drew.

    Templates rely on "shrink on overflow" and wrap settings; after a shorter or
    longer line is poured in, PowerPoint re-lays the box and the text drifts out
    of its plate. Pinning wrap on and autofit off keeps it where the designer put
    it.
    """
    try:
        tf = shape.text_frame
    except Exception:
        return
    try:
        tf.word_wrap = True
    except Exception:
        pass
    try:
        from pptx.enum.text import MSO_AUTO_SIZE
        tf.auto_size = MSO_AUTO_SIZE.NONE
    except Exception:
        pass



def fit_to_plate(shape, slot: ShapeInfo, slide) -> None:
    """Grow a narrow text frame to the plate drawn behind it.

    Templates often put the text box inside a much wider bar or card. Our line is
    then centred in the small box, wraps early and looks shifted off the bar.
    """
    best = None
    for sh in slide.shapes:
        if sh.shape_id == slot.shape_id:
            continue
        try:
            left, top = emu2pt(sh.left), emu2pt(sh.top)
            w, h = emu2pt(sh.width), emu2pt(sh.height)
        except Exception:
            continue
        if w < slot.width * 1.3 or h < slot.height * 0.8:
            continue
        # only single-line bands, not cards that hold several text lines
        if h > slot.height * 2.4:
            continue
        # the slot must sit inside the plate
        if slot.left < left - 2 or slot.right > left + w + 2:
            continue
        if slot.top < top - 2 or slot.bottom > top + h + 2:
            continue
        kind, _ = shape_paint(sh)
        if kind in ("solid", "gradient", "picture"):
            if best is None or w * h > best[2] * best[3]:
                best = (left, top, w, h)
    if best is None:
        return
    left, top, w, h = best
    inset_x = min(18.0, w * 0.06)
    inset_y = min(6.0, h * 0.12)
    shape.left = Pt(left + inset_x)
    shape.width = Pt(w - 2 * inset_x)
    shape.top = Pt(top + inset_y)
    shape.height = Pt(h - 2 * inset_y)
    slot.left = left + inset_x
    slot.width = w - 2 * inset_x
    slot.top = top + inset_y
    slot.height = h - 2 * inset_y


def paint_text(tf, slot: ShapeInfo, values: list[str]) -> float:
    _fill_paragraphs(tf, values)
    return fit_font(tf, slot)


TRANSITION_TAGS = (
    "transition",
    "timing",
)


def strip_effects(prs) -> int:
    """Remove slide transitions, animation timing and embedded audio/video.

    Every deck this engine emits is meant for live presenting without the
    template's decorative transitions or soundtrack.
    """
    removed = 0
    ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    for slide in prs.slides:
        xml = slide._element
        for tag in TRANSITION_TAGS:
            for node in list(xml.iter(ns + tag)):
                node.getparent().remove(node)
                removed += 1
        for node in list(xml.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}video")):
            node.getparent().remove(node)
            removed += 1
        for node in list(xml.iter(ns + "audio")):
            node.getparent().remove(node)
            removed += 1
    # drop media parts and their relationships
    for slide in prs.slides:
        for rel in list(slide.part.rels.values()):
            if rel.reltype.endswith(("/video", "/audio", "/media")):
                slide.part.drop_rel(rel.rId)
                removed += 1
    return removed


def blank(shape) -> None:
    if not shape.has_text_frame:
        return
    try:
        shape.text_frame.clear()
    except Exception:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                run.text = ""


NUMBER_PREFIX_RE = re.compile(r"^\s*[0-9０-９]{1,2}\s*[.、,，:：)）\-—]?\s+")


def strip_number(text: str) -> str:
    return NUMBER_PREFIX_RE.sub("", text, count=1).strip()


def drop_row(slide, shapes: dict, slots: list[ShapeInfo], row: list[ShapeInfo], used: set[int], report: dict) -> None:
    """Delete one visual row (its text slots, numbers and local decorations)."""
    top = min(s.top for s in row)
    bottom = max(s.bottom for s in row)
    left = min(s.left for s in row)
    right = max(s.right for s in row)
    center = (top + bottom) / 2

    def in_band(t: float, b: float, l: float) -> bool:
        return (top - 12) <= (t + b) / 2 <= (bottom + 12) and l >= left - 60

    for slot in list(slots):
        if slot.shape_id in used:
            continue
        if in_band(slot.top, slot.bottom, slot.left):
            shape = shapes.get(slot.shape_id)
            if shape is not None:
                remove_shape(shape)
                report["removed"] += 1
    for sh in list(slide.shapes):
        try:
            t, b = emu2pt(sh.top), emu2pt(sh.top + sh.height)
            l, r = emu2pt(sh.left), emu2pt(sh.left + sh.width)
        except Exception:
            continue
        if r < left - 40:
            continue
        if in_band(t, b, l):
            if sh.shape_id in used:
                continue
            remove_shape(sh)
            report["removed"] += 1


def remove_shape(shape) -> None:
    el = shape._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


# ---------------------------------------------------------------- contrast

def shape_paint(shape) -> tuple[str, str | None]:
    """('picture'|'solid'|'none', hex) for a shape's own fill."""
    try:
        fill = shape.fill
        ftype = fill.type
    except Exception:
        return ("none", None)
    if ftype is None:
        return ("none", None)
    try:
        if ftype == 6:  # picture
            return ("picture", None)
    except Exception:
        pass
    try:
        t = int(ftype)
    except Exception:
        return ("none", None)
    if t == 1:  # solid
        try:
            return ("solid", str(fill.fore_color.rgb).upper())
        except Exception:
            return ("none", None)
    if t == 3:  # gradient: judge by the mean of its stops
        try:
            cols = []
            for stop in fill.gradient_stops:
                try:
                    cols.append(str(stop.color.rgb).upper())
                except Exception:
                    continue
            if cols:
                return ("solid", _mean_hex(cols))
        except Exception:
            return ("none", None)
    return ("none", None)


def _mean_hex(cols: list[str]) -> str:
    n = len(cols)
    r = sum(int(c[0:2], 16) for c in cols) // n
    g = sum(int(c[2:4], 16) for c in cols) // n
    b = sum(int(c[4:6], 16) for c in cols) // n
    return f"{r:02X}{g:02X}{b:02X}"


def paint_behind(slide, slot: ShapeInfo, theme: dict[str, str]) -> tuple[str, float | None]:
    """What is painted behind a text slot, judged by area coverage.

    Returns ('photo'|'light'|'dark', luma).  A picture counts as 'photo' because
    its brightness cannot be known without rendering.
    """
    shapes = list(slide.shapes)
    top_idx = len(shapes)
    for i, sh in enumerate(shapes):
        if sh.shape_id == slot.shape_id:
            top_idx = i
            break
    slot_area = max(slot.area, 1.0)
    best_frac, best_kind, best_hex = 0.0, "", None
    # the shape's own fill is the closest thing behind its text
    own = next((sh for sh in shapes if sh.shape_id == slot.shape_id), None)
    if own is not None:
        kind, hexc = shape_paint(own)
        if kind in ("solid", "picture"):
            best_frac, best_kind, best_hex = 1.0, kind, hexc
    for sh in reversed(shapes[:top_idx]):
        try:
            box = ShapeInfo(
                sh.shape_id, sh.name, "", None,
                emu2pt(sh.left), emu2pt(sh.top), emu2pt(sh.width), emu2pt(sh.height), 0.0, None,
            )
        except Exception:
            continue
        inter_w = min(box.right, slot.right) - max(box.left, slot.left)
        inter_h = min(box.bottom, slot.bottom) - max(box.top, slot.top)
        if inter_w <= 0 or inter_h <= 0:
            continue
        frac = (inter_w * inter_h) / slot_area
        kind, hexc = shape_paint(sh)
        if kind == "none":
            continue
        if frac > best_frac + 1e-6:
            best_frac, best_kind, best_hex = frac, kind, hexc
    if best_frac >= 0.45:
        if best_kind == "picture":
            return ("photo", None)
        if best_kind == "solid" and best_hex:
            lv = luma(best_hex)
            return ("dark" if (lv or 0) < 0.45 else "light", lv)
    try:
        bg = slide.background.fill
        if bg.type == 1:
            hexc = str(bg.fore_color.rgb).upper()
            lv = luma(hexc)
            return ("dark" if (lv or 0) < 0.45 else "light", lv)
    except Exception:
        pass
    # nothing opaque sits behind the text: fall back to the slide background
    try:
        fill = slide.background.fill
        if fill.type is not None and fill.type == 1:  # MSO_FILL.SOLID
            hexc = str(fill.fore_color.rgb)
            lv = luma(hexc)
            return ("dark" if (lv or 0) < 0.45 else "light", lv)
    except Exception:
        pass
    return ("light", 0.95)


def contrast_ratio(a: str, b: str) -> float:
    def rel(hex_rgb: str) -> float:
        def channel(v: int) -> float:
            c = v / 255.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

        r, g, bl = (int(hex_rgb[i : i + 2], 16) for i in (0, 2, 4))
        return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(bl)

    la, lb = rel(a), rel(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)



def _behind_coverage(slide, slot: ShapeInfo) -> float:
    """Largest opaque overlap fraction behind a slot (0..1)."""
    best = 0.0
    area = max(slot.area, 1.0)
    for sh in slide.shapes:
        if sh.shape_id == slot.shape_id:
            continue
        try:
            left, top = emu2pt(sh.left), emu2pt(sh.top)
            w, h = emu2pt(sh.width), emu2pt(sh.height)
        except Exception:
            continue
        iw = min(left + w, slot.right) - max(left, slot.left)
        ih = min(top + h, slot.bottom) - max(top, slot.top)
        if iw <= 0 or ih <= 0:
            continue
        best = max(best, iw * ih / area)
    return best


def ensure_readable(tf, slot: ShapeInfo, slide, theme: dict[str, str]) -> bool:
    """Make the text visible against the paint behind it.

    Explicit accent colours are kept when they still contrast enough; inherited
    (unstyled) text and clashing colours are flipped to dark or white.
    """
    behind, bg_luma = paint_behind(slide, slot, theme)
    # only trust the probe for a shape that really covers the text; a partial
    # overlap usually means a decorative panel, not the actual background
    if behind in ("light", "dark") and _behind_coverage(slide, slot) < 0.85:
        probe = "F2F2F2"
        try:
            if slide.background.fill.type == 1:
                probe = str(slide.background.fill.fore_color.rgb)
        except Exception:
            pass
        explicit0 = [h for h in (color_hex(c, theme) for c in slot.colors) if h]
        if explicit0 and all(contrast_ratio(h, probe) >= 3.0 for h in explicit0):
            return False
    if behind == "photo":
        # dark lettering over a photo nearly always disappears; lift it to white
        dark = [h for h in (color_hex(c, theme) for c in slot.colors) if h and (luma(h) or 1) < 0.45]
        if not dark:
            return False
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor.from_string(WHITE_TEXT)
        return True
    # unknown (picture / inherited) background: a template that already set light
    # lettering knows its own plate; do not flip it to dark
    try:
        bg_solid = slide.background.fill.type == 1
    except Exception:
        bg_solid = False
    if not bg_solid:
        light_text = [h for h in (color_hex(c, theme) for c in slot.colors) if h and contrast_ratio(h, "000000") >= 3.0]
        if light_text:
            return False

    bg_hex = "F2F2F2" if behind == "light" else "1A1A1A"
    want = DARK_TEXT if behind == "light" else WHITE_TEXT
    # a dark deck must never receive dark lettering, whatever the local probe says
    try:
        bg_fill = slide.background.fill
        if bg_fill.type == 1:
            bl = luma(str(bg_fill.fore_color.rgb)) or 0.5
            want = WHITE_TEXT if bl < 0.45 else DARK_TEXT
    except Exception:
        pass
    explicit = [h for h in (color_hex(c, theme) for c in slot.colors) if h]
    if explicit and all(contrast_ratio(h, bg_hex) >= 3.0 for h in explicit):
        return False
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor.from_string(want)
    return True


# ---------------------------------------------------------------- page filling

def toc_quality(slots: list[ShapeInfo], need: int) -> int:
    """Row-list TOCs score high; diagram TOCs (scattered circles) score low."""
    pool = [s for s in slots if classify(s) in ("text", "num") and s.text.strip()]
    rows = [r for r in band_rows(pool) if row_label(r) is not None]
    if len(rows) < 2:
        return 0
    lefts = {round(row_label(r).left / 10) for r in rows}
    aligned = len(lefts) <= 2
    score = 20 if aligned else 0
    if len(rows) >= need:
        score += 40 - min(abs(len(rows) - need) * 8, 30)
    else:
        score -= (need - len(rows)) * 10
    return score


def slot_usable(slide, slot: ShapeInfo, theme: dict[str, str]) -> bool:
    """A body slot we are willing to pour a sentence into."""
    if abs(slot.rot) > 2:
        return False
    if slot.area < 2200 or slot.height < 20 or slot.width < 120:
        return False
    if slot.max_pt and slot.max_pt < 10.5:
        # 6-9pt template labels are captions under icons; body text there reads
        # as a defect no matter how well it fits
        return False
    if title_capacity(slot, 0.8) < 12:
        return False
    kind, _ = paint_behind(slide, slot, theme)
    if kind == "photo":
        return False
    return True


def page_score(slide, slots: list[ShapeInfo], theme: dict[str, str], ptype: str) -> dict:
    """How well this source page can host generated text."""
    protected = {s.shape_id for s in slots if classify(s) == "furniture"}
    headline = pick_headline([s for s in slots if s.shape_id not in protected])
    usable = [
        s
        for s in slots
        if s.shape_id not in protected
        and classify(s) in ("text", "filler", "empty")
        and slot_usable(slide, s, theme)
        and (headline is None or not s.overlaps(headline, pad=-4))
    ]
    rotated = len([s for s in slots if abs(s.rot) > 2])
    complexity = page_complexity(slide)
    return {
        "headline": headline,
        "body_slots": len(usable),
        "rotated": rotated,
        "complexity": complexity,
        "score": len(usable) * 2 + (3 if headline is not None else 0) - rotated * 2
        + min(complexity, 14),
    }


def page_complexity(slide) -> float:
    """Rough visual richness of a page: pictures, groups, freeforms, autoshapes.

    A page whose graphic is a real composition reads as designed; a page with one
    thin rule and a title box reads as hand-made by an AI.
    """
    score = 0.0
    for sh in _iter_text_shapes(slide.shapes) or []:
        pass
    try:
        for sh in slide.shapes:
            kind = _safe_shape_type(sh)
            if kind == 13:            # PICTURE
                score += 3.0
            elif kind == 6:           # GROUP
                score += 1.5
            elif kind == 5:           # FREEFORM
                score += 1.2
            elif kind == 1:           # AUTO_SHAPE
                score += 0.6
    except Exception:
        pass
    return round(score, 1)


def fill_page(slide, slots: list[ShapeInfo], page: dict, ptype: str, meta: dict, theme: dict) -> dict:
    shapes = {sh.shape_id: sh for sh in _iter_text_shapes(slide.shapes)}
    roles = role_map(slots)
    theme = dict(theme)
    report = {"page_type": ptype, "headline": None, "filled": 0, "blanked": 0, "removed": 0, "recolored": 0}
    used: set[int] = set()

    explicit = {int(k): v for k, v in (page.get("slots") or {}).items()}
    for sid, value in explicit.items():
        shape = shapes.get(sid)
        if shape is None:
            continue
        slot = next((s for s in slots if s.shape_id == sid), None)
        if slot is None:
            continue
        if ptype == "cover" and slot.max_pt >= 40 and slot.width < 420:
            widen_headline(shape, slot, emu2pt(slide.part.package.presentation_part.presentation.slide_width))
            slot.width = emu2pt(shape.width)
            slot.left = emu2pt(shape.left)
        paint_text(shape.text_frame, slot, value if isinstance(value, list) else [value])
        if ptype not in ("cover", "ending") and ensure_readable(shape.text_frame, slot, slide, theme):
            report["recolored"] += 1
        used.add(sid)

    protected = {s.shape_id for s in roles.get("furniture", [])}
    has_furniture = bool(protected)
    pres = slide.part.package.presentation_part.presentation
    slide_h = emu2pt(pres.slide_height)
    slide_w = emu2pt(pres.slide_width)
    prefer = 0.3 * slide_h if ptype in ("content", "dense", "toc") else None
    title = page.get("title")
    centered = ptype in ("cover", "ending")
    headline = pick_headline(
        [s for s in slots if s.shape_id not in used and s.shape_id not in protected],
        prefer_top=prefer,
        center=slide_w / 2 if centered else None,
        allow_year=centered,
    )
    # rotated / decorative placeholders must not survive into the output
    for slot in roles.get("rotated", []) + roles.get("design", []):
        if slot.shape_id not in shapes:
            continue
        if is_filler(slot.text or "") or LOGO_RE.match(slot.text or ""):
            blank(shapes[slot.shape_id])
            report["blanked"] += 1

    if title and headline is not None and not has_furniture:
        if ptype in ("cover", "content", "dense", "ending"):
            widen_headline(shapes[headline.shape_id], headline, slide_w)
        paint_text(shapes[headline.shape_id].text_frame, headline, [title])
        normalize_alignment(shapes[headline.shape_id])
        if ptype not in ("cover", "ending") and ensure_readable(
            shapes[headline.shape_id].text_frame, headline, slide, theme
        ):
            report["recolored"] += 1
        used.add(headline.shape_id)
        report["headline"] = f"#{headline.shape_id} {headline.name} {headline.max_pt}pt"
    elif title and headline is not None and has_furniture:
        report["headline"] = f"kept furniture {[s.text for s in roles.get('furniture', [])][:2]}"

    body = list(page.get("body") or [])
    below_headline: list[ShapeInfo] = []
    if centered and headline is not None:
        below_headline = [
            s
            for s in slots
            if s.shape_id not in used
            and s.shape_id not in protected
            and s.shape_id in shapes
            and classify(s) in ("text", "filler", "empty")
            and slot_usable(slide, s, theme)
            and s.top >= headline.bottom - 4
        ]
        below_headline.sort(key=lambda s: (s.top, abs(s.left + s.width / 2 - slide_w / 2)))
    body_pool = below_headline or [
        s
        for s in slots
        if s.shape_id not in used
        and s.shape_id not in protected
        and s.shape_id in shapes
        and classify(s) in ("text", "filler", "empty")
        and slot_usable(slide, s, theme)
        and (headline is None or not s.overlaps(headline, pad=-4))
    ]
    pre_rows: list[list[ShapeInfo]] = []
    rows_aligned = False
    if ptype in ("dense", "content") and body:
        rp = [
            s
            for s in slots
            if s.shape_id not in protected and classify(s) in ("text", "num") and s.text.strip()
        ]
        pre_rows = [r for r in band_rows(rp) if row_label(r) is not None]
        rows_aligned = len({round(row_label(r).left / 12) for r in pre_rows}) <= 2

    if body and ptype == "toc":
        body = [strip_number(v) for v in body]
        row_pool = [
            s
            for s in slots
            if s.shape_id not in protected and classify(s) in ("text", "num") and s.text.strip()
        ]
        if not has_furniture:
            # no 目录/CONTENTS marker: the top band is the page heading, not a row
            heading_band = 0.22 * slide_h
            row_pool = [s for s in row_pool if s.top > heading_band] or row_pool
        rows = [r for r in band_rows(row_pool) if row_label(r) is not None]
        if len(rows) >= 2:
            labels_left = {round(row_label(r).left / 10) for r in rows}
            consistent = len(labels_left) <= 2
            for i, value in enumerate(body):
                if i >= len(rows):
                    break
                label = row_label(rows[i])
                paint_text(shapes[label.shape_id].text_frame, label, [value])
                if ensure_readable(shapes[label.shape_id].text_frame, label, slide, theme):
                    report["recolored"] += 1
                used.add(label.shape_id)
                report["filled"] += 1
            extra = rows[len(body):]
            if consistent and extra and len(extra) <= 4:
                for row in extra:
                    drop_row(slide, shapes, slots, row, used, report)
                extra = []
            if extra:
                for row in extra:
                    label = row_label(row)
                    if label is not None and label.shape_id in shapes:
                        blank(shapes[label.shape_id])
                        report["blanked"] += 1
            body = body[len(rows):]

    if body:
        pool = [
            s
            for s in body_pool
            if s.shape_id not in used and classify(s) in ("text", "filler", "empty")
        ]
        if headline is not None and ptype in ("cover", "chapter", "ending"):
            below = [
                s
                for s in pool
                if s.top >= headline.top + headline.height * 0.55 and s.bottom > headline.top
            ]
            pool = below or pool
            pool.sort(key=lambda s: (abs(s.left - headline.left), read_order(s)))
        elif ptype == "dense":
            pool.sort(key=lambda s: (-s.area, s.top))
        else:
            pool.sort(key=read_order)
        i = 0
        for slot in pool:
            if i >= len(body):
                break
            shape = shapes[slot.shape_id]
            ratio = paint_text(shape.text_frame, slot, [body[i]])
            eff_pt = slot.max_pt * ratio
            # a one-line caption in a short box is fine; a squeezed paragraph is not
            per_line = max((slot.width - 6) / max(eff_pt, 1.0), 1.0)
            need_lines = max(1, int(len(body[i]) / per_line) + 1)
            cramped = need_lines > 1 and not centered and (
                (ratio < 0.6 and (slot.width < 240 or slot.height <= 26))
                or (eff_pt < 12.0 and slot.height < 40)
            )
            if cramped:
                blank(shape)
                report["blanked"] += 1
                continue
            if ensure_readable(shape.text_frame, slot, slide, theme):
                report["recolored"] += 1
            used.add(slot.shape_id)
            report["filled"] += 1
            i += 1
        if report["filled"] == 0 and i < len(body) and not centered:
            # every candidate looked cramped -> a tight page beats an empty one
            for slot in pool:
                if i >= len(body):
                    break
                if slot.shape_id in used:
                    continue
                shape = shapes[slot.shape_id]
                paint_text(shape.text_frame, slot, [body[i]])
                if ensure_readable(shape.text_frame, slot, slide, theme):
                    report["recolored"] += 1
                used.add(slot.shape_id)
                report["filled"] += 1
                i += 1
        rest = body[i:]
        if rest:
            hosts = [s for s in pool if s.shape_id in used and s.height >= 60 and s.width >= 200]
            if hosts:
                host = max(hosts, key=lambda s: s.area)
                paint_text(shapes[host.shape_id].text_frame, host, list(rest))
                fit_font(shapes[host.shape_id].text_frame, host, floor=0.5)
                report["filled"] += len(rest)

    for sid in list(used):
        if sid in shapes:
            stabilize_text(shapes[sid])

    # a numbered/bulleted list that got fewer items than it has rows: drop the
    # spare rows instead of leaving "04 / 05" hanging with no text
    if ptype in ("dense", "content") and body and page.get("type") != "toc":
        extras = pre_rows[len(body):] if len(pre_rows) > len(body) else []
        if rows_aligned and extras and len(extras) <= 6 and report["filled"] >= len(body):
            for row in extras:
                drop_row(slide, shapes, slots, row, used, report)

    year = None if ptype == "ending" else meta.get("year")
    for slot in roles.get("year", []):
        if slot.shape_id in used or slot.shape_id not in shapes:
            continue
        if headline is not None and slot.overlaps(headline, pad=-6):
            blank(shapes[slot.shape_id])
            report["blanked"] += 1
            continue
        if year and title_capacity(slot, 0.8) >= 4:
            paint_text(shapes[slot.shape_id].text_frame, slot, [str(year)])
            if ensure_readable(shapes[slot.shape_id].text_frame, slot, slide, theme):
                report["recolored"] += 1
        else:
            blank(shapes[slot.shape_id])
            report["blanked"] += 1
        used.add(slot.shape_id)

    for slot in roles.get("filler", []):
        if slot.shape_id in used or slot.shape_id not in shapes:
            continue
        blank(shapes[slot.shape_id])
        report["blanked"] += 1

    for slot in roles.get("text", []):
        if slot.shape_id in used or slot.shape_id not in shapes:
            continue
        same_row = any(
            abs(slot.max_pt - t.max_pt) <= 2.0 and abs(slot.left - t.left) <= 10
            for t in slots
            if t.shape_id in used
        )
        if same_row and ptype in ("toc", "content"):
            remove_shape(shapes[slot.shape_id])
            report["removed"] += 1
        else:
            blank(shapes[slot.shape_id])
            report["blanked"] += 1
    return report


def cmd_fill(args) -> int:
    src = Path(args.source).resolve()
    plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
    out = Path(args.out).resolve() if args.out else src.with_name(src.stem + "-built.pptx")
    prs = Presentation(str(src))
    theme = resolve_theme(prs)
    total = len(prs.slides)
    inventory: dict[int, tuple[str, list[ShapeInfo]]] = {}
    for i, slide in enumerate(prs.slides):
        slots = read_shapes(slide)
        inventory[i + 1] = (page_type(i, total, slide, slots), slots)

    pages = plan["pages"]
    by_type: dict[str, list[int]] = {}
    for idx, (ptype, _slots) in inventory.items():
        by_type.setdefault(ptype, []).append(idx)
    fallback = [i for i in inventory if inventory[i][0] in ("content", "dense")]
    meta = {"year": plan.get("year")}

    chosen: list[tuple[int, dict]] = []
    used: set[int] = set()
    slides = list(prs.slides)

    def rank(pool: list[int], page: dict) -> list[int]:
        req = len(page.get("body") or [])
        want_title = bool(page.get("title"))
        want = page.get("type", "content")
        scored = []
        for i in pool:
            ptype_i, slots_i = inventory[i]
            sc = page_score(slides[i - 1], slots_i, theme, ptype_i)
            bonus = toc_quality(slots_i, req) if want == "toc" else 0
            fits = sc["body_slots"] >= req and (sc["headline"] is not None or not want_title)
            scored.append((1 if fits else 0, bonus, sc["score"], -i, i))
        scored.sort(reverse=True)
        return [t[4] for t in scored]

    for page in pages:
        if page.get("source_slide"):
            pick = int(page["source_slide"])
            if not 1 <= pick <= len(slides):
                print(
                    f"!! page {page.get('type')!r}: source_slide {pick} is out of range "
                    f"(deck has {len(slides)} slides) - picking one automatically"
                )
                want = page.get("type", "content")
                pool = [i for i in by_type.get(want, []) if i not in used] or [
                    i for i in fallback if i not in used
                ]
                if not pool:
                    raise SystemExit(f"no usable source slide for page {page!r}")
                pick = rank(pool, page)[0]
        else:
            want = page.get("type", "content")
            pool = [i for i in by_type.get(want, []) if i not in used]
            if not pool and want in ("content", "dense"):
                pool = [i for i in fallback if i not in used]
            if not pool:
                pool = [i for i in by_type.get(want, [])] or fallback
            if not pool:
                raise SystemExit(f"no source slide for page type {want!r}")
            pick = rank(pool, page)[0]
        used.add(pick)
        chosen.append((pick, page))

    for n, (src_idx, page) in enumerate(chosen, 1):
        slide = slides[src_idx - 1]
        ptype, slots = inventory[src_idx]
        report = fill_page(slide, slots, page, ptype, meta, theme)
        print(
            f"p{n:>2} <- slide {src_idx:>2} [{ptype:<7}] filled={report['filled']} "
            f"blanked={report['blanked']} removed={report['removed']} recy={report['recolored']} "
            f"headline={report['headline']}"
        )

    order = [src_idx for src_idx, _ in chosen]
    sldIdLst = prs.slides._sldIdLst
    id_by_index = {i + 1: sld for i, sld in enumerate(list(sldIdLst))}
    keep = [id_by_index[i] for i in order]
    for sld in list(sldIdLst):
        if sld not in keep:
            rId = sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            prs.part.drop_rel(rId)
            sldIdLst.remove(sld)
    for sld in keep:
        sldIdLst.remove(sld)
    for sld in keep:
        sldIdLst.append(sld)

    out.parent.mkdir(parents=True, exist_ok=True)
    removed = strip_effects(prs)
    if removed:
        print(f"stripped {removed} transition/animation/media nodes")
    prs.save(str(out))
    print(f"built {len(chosen)} pages -> {out}")
    return 0


def cmd_index(args) -> int:
    src = Path(args.source).resolve()
    out_dir = Path(args.out).resolve() if args.out else src.parent / "index"
    prs = Presentation(str(src))
    theme = resolve_theme(prs)
    pages = []
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides):
        shapes = read_shapes(slide)
        layout = slide.slide_layout
        pages.append(
            {
                "index": i + 1,
                "layout": layout.name if layout is not None else "",
                "page_type": page_type(i, total, slide, shapes),
                "shape_count": len(slide.shapes),
                "slots": [s.to_dict() for s in shapes],
            }
        )
    data = {
        "source": str(src),
        "name": src.stem,
        "theme": theme,
        **deck_meta(prs),
        "pages": pages,
    }
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / f"{src.stem}.json"
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"indexed {total} pages -> {dst}")
    return 0


def cmd_slots(args) -> int:
    prs = Presentation(str(Path(args.source).resolve()))
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides):
        if args.slide and i + 1 != args.slide:
            continue
        shapes = read_shapes(slide)
        print(f"\n== page {i + 1}/{total}  layout={slide.slide_layout.name}  type={page_type(i, total, slide, shapes)}")
        for s in sorted(shapes, key=lambda x: (-x.max_pt, -x.width * x.height)):
            print(
                f"  #{s.shape_id:<4} {classify(s):<9} {s.max_pt:>5}pt "
                f"box=({s.left:.0f},{s.top:.0f},{s.width:.0f}x{s.height:.0f})  {s.text[:60]}"
            )
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="pptl")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("index")
    p.add_argument("source")
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_index)

    p = sub.add_parser("fill")
    p.add_argument("source")
    p.add_argument("plan")
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser("slots")
    p.add_argument("source")
    p.add_argument("--slide", type=int)
    p.set_defaults(func=cmd_slots)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
