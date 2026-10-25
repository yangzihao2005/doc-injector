"""Run-level style cloning for injected body text.

Word resolves fonts through three independent slots (``ascii``, ``hAnsi`` and
``eastAsia``).  Setting only ``font.name`` writes the first two, so Chinese text
silently falls back to 等线 on machines where the template normally uses 宋体.
This module captures the style of the anchor's neighbouring body paragraph and
writes all relevant slots when applying it to new runs.
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass
from typing import Any

from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Length, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from src import scanner

DEFAULT_CJK_FONT = "宋体"
DEFAULT_LATIN_FONT = "Times New Roman"
DEFAULT_SIZE_PT = 12.0


@dataclass(frozen=True)
class RunStyleSnapshot:
    """Character-level properties captured from one run."""

    font_name: str | None = None
    east_asia_font: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    color_hex: str | None = None
    underline: bool | None = None


@dataclass(frozen=True)
class ParagraphStyleSnapshot:
    """Everything needed to reproduce a body paragraph's appearance."""

    run: RunStyleSnapshot
    alignment: WD_ALIGN_PARAGRAPH | None = None
    line_spacing: float | int | None = None
    space_before: Length | None = None
    space_after: Length | None = None
    first_line_indent: Length | None = None


def _safe_size_pt(font: Any) -> float | None:
    size = getattr(font, "size", None)
    if size is None:
        return None
    try:
        return float(size.pt)
    except (AttributeError, ValueError):
        return None


def _safe_color_hex(font: Any) -> str | None:
    color = getattr(font, "color", None)
    if color is None:
        return None
    try:
        rgb = color.rgb
    except (AttributeError, ValueError):
        return None
    return str(rgb) if rgb is not None else None


def _east_asia_font(run: Run) -> str | None:
    r_pr = run._r.rPr
    if r_pr is None:
        return None
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        return None
    value = r_fonts.get(qn("w:eastAsia"))
    return value or None


def _style_east_asia_font(paragraph: Paragraph) -> str | None:
    style = paragraph.style
    if style is None:
        return None
    try:
        r_pr = style.element.rPr
    except AttributeError:
        return None
    if r_pr is None or r_pr.rFonts is None:
        return None
    value = r_pr.rFonts.get(qn("w:eastAsia"))
    return value or None


def _capture_run(run: Run, paragraph: Paragraph) -> RunStyleSnapshot:
    font = run.font
    underline = font.underline
    return RunStyleSnapshot(
        font_name=font.name,
        east_asia_font=_east_asia_font(run) or _style_east_asia_font(paragraph),
        size_pt=_safe_size_pt(font),
        bold=font.bold,
        italic=font.italic,
        color_hex=_safe_color_hex(font),
        underline=None if underline is None else bool(underline),
    )


def _snapshot_from_paragraph(paragraph: Paragraph) -> ParagraphStyleSnapshot:
    font = paragraph.style.font if paragraph.style is not None else None
    run_snapshot = RunStyleSnapshot(
        font_name=getattr(font, "name", None),
        east_asia_font=_style_east_asia_font(paragraph),
        size_pt=_safe_size_pt(font),
        bold=getattr(font, "bold", None),
        italic=getattr(font, "italic", None),
        color_hex=_safe_color_hex(font),
        underline=None,
    )
    fmt = paragraph.paragraph_format
    return ParagraphStyleSnapshot(
        run=run_snapshot,
        alignment=fmt.alignment,
        line_spacing=fmt.line_spacing,
        space_before=fmt.space_before,
        space_after=fmt.space_after,
        first_line_indent=fmt.first_line_indent,
    )


def capture_style(paragraph: Paragraph) -> ParagraphStyleSnapshot:
    """Capture paragraph geometry plus the first run's character style."""

    if paragraph.runs:
        run = paragraph.runs[0]
        fmt = paragraph.paragraph_format
        return ParagraphStyleSnapshot(
            run=_capture_run(run, paragraph),
            alignment=fmt.alignment,
            line_spacing=fmt.line_spacing,
            space_before=fmt.space_before,
            space_after=fmt.space_after,
            first_line_indent=fmt.first_line_indent,
        )
    return _snapshot_from_paragraph(paragraph)


def default_style(doc: DocumentObject) -> ParagraphStyleSnapshot:
    """A conservative Chinese body-text fallback that never raises."""

    font_name = DEFAULT_CJK_FONT
    size_pt = DEFAULT_SIZE_PT
    east_asia = DEFAULT_CJK_FONT
    try:
        normal = doc.styles["Normal"]
        font_name = normal.font.name or DEFAULT_CJK_FONT
        size_pt = _safe_size_pt(normal.font) or DEFAULT_SIZE_PT
        r_pr = normal.element.rPr
        if r_pr is not None and r_pr.rFonts is not None:
            east_asia = r_pr.rFonts.get(qn("w:eastAsia")) or DEFAULT_CJK_FONT
    except (KeyError, AttributeError):
        pass

    return ParagraphStyleSnapshot(
        run=RunStyleSnapshot(
            font_name=font_name,
            east_asia_font=east_asia,
            size_pt=size_pt,
            bold=False,
            italic=False,
            color_hex=None,
            underline=False,
        ),
        alignment=None,
        line_spacing=1.5,
        space_before=None,
        space_after=None,
        first_line_indent=None,
    )


def _first_body_after(doc: DocumentObject, anchor: Paragraph) -> Paragraph | None:
    """Nearest body paragraph after ``anchor``, even across a heading boundary."""

    seen = False
    for paragraph in scanner.iter_paragraphs(doc):
        if paragraph._p is anchor._p:
            seen = True
            continue
        if not seen:
            continue
        if not paragraph.text.strip() or scanner.is_heading(paragraph):
            continue
        if scanner.is_caption(paragraph):
            continue
        if paragraph._p.find(f".//{qn('w:drawing')}") is not None:
            continue
        return paragraph
    return None


def capture_body_style(doc: DocumentObject, anchor: Paragraph) -> ParagraphStyleSnapshot:
    """Style source: section body, then nearest later body, then document defaults."""

    sample = scanner.find_body_sample(doc, anchor)
    if sample is None:
        sample = _first_body_after(doc, anchor)
    if sample is None:
        sample = scanner.first_body_paragraph(doc)
    if sample is None:
        return default_style(doc)
    return capture_style(sample)


def set_run_fonts(run: Run, font_name: str, east_asia_font: str | None = None) -> None:
    """Write ascii/hAnsi/cs and eastAsia explicitly (the Chinese inheritance fix)."""

    east_asia = east_asia_font or font_name
    run.font.name = font_name
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:cs"), font_name)
    r_fonts.set(qn("w:eastAsia"), east_asia)


def apply_paragraph_style(paragraph: Paragraph, snapshot: ParagraphStyleSnapshot) -> None:
    """Copy paragraph-level geometry without replacing its named style."""

    fmt = paragraph.paragraph_format
    if snapshot.alignment is not None:
        fmt.alignment = snapshot.alignment
    if snapshot.line_spacing is not None:
        fmt.line_spacing = snapshot.line_spacing  # type: ignore[assignment]
    if snapshot.space_before is not None:
        fmt.space_before = snapshot.space_before
    if snapshot.space_after is not None:
        fmt.space_after = snapshot.space_after
    if snapshot.first_line_indent is not None:
        fmt.first_line_indent = snapshot.first_line_indent


def add_styled_run(
    paragraph: Paragraph,
    text: str,
    snapshot: ParagraphStyleSnapshot,
    *,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> Run:
    """Append a run that reproduces ``snapshot`` with optional overrides."""

    run = paragraph.add_run(text)
    captured = snapshot.run

    latin_font = font or captured.font_name or DEFAULT_LATIN_FONT
    east_asia = font or captured.east_asia_font or DEFAULT_CJK_FONT
    set_run_fonts(run, latin_font, east_asia)

    resolved_size = size_pt if size_pt is not None else captured.size_pt
    if resolved_size is not None:
        run.font.size = Pt(resolved_size)

    resolved_bold = bold if bold is not None else captured.bold
    if resolved_bold is not None:
        run.font.bold = resolved_bold

    resolved_italic = italic if italic is not None else captured.italic
    if resolved_italic is not None:
        run.font.italic = resolved_italic

    if captured.underline is not None:
        run.font.underline = captured.underline

    if captured.color_hex:
        with contextlib.suppress(ValueError):
            run.font.color.rgb = RGBColor.from_string(captured.color_hex)
    return run


def write_styled_paragraph(
    paragraph: Paragraph,
    text: str,
    snapshot: ParagraphStyleSnapshot,
    *,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    alignment: WD_ALIGN_PARAGRAPH | None = None,
) -> Run:
    """Replace a paragraph's runs with one styled run."""

    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)
    apply_paragraph_style(paragraph, snapshot)
    if alignment is not None:
        paragraph.paragraph_format.alignment = alignment
    return add_styled_run(
        paragraph,
        text,
        snapshot,
        font=font,
        size_pt=size_pt,
        bold=bold,
    )


__all__ = [
    "DEFAULT_CJK_FONT",
    "DEFAULT_LATIN_FONT",
    "ParagraphStyleSnapshot",
    "RunStyleSnapshot",
    "add_styled_run",
    "apply_paragraph_style",
    "capture_body_style",
    "capture_style",
    "default_style",
    "set_run_fonts",
    "write_styled_paragraph",
]
