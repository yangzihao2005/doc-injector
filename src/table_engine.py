"""Academic table renderer (three-line and full-grid) with captions above.

All borders are written as explicit ``w:tcBorders`` / ``w:tblBorders`` XML so
the result does not depend on the template's table style or on Word's
locale-specific defaults.  Every row is protected with ``w:cantSplit`` and the
header row repeats with ``w:tblHeader``.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any, Literal

from docx.document import Document as DocumentObject
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.shared import Pt
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from src import scanner
from src.schema import TableModel, TableType
from src.style_cloner import (
    DEFAULT_CJK_FONT,
    ParagraphStyleSnapshot,
    add_styled_run,
    default_style,
    write_styled_paragraph,
)

BorderEdge = Literal["top", "left", "bottom", "right", "insideH", "insideV"]

_EDGE_ORDER: tuple[BorderEdge, ...] = ("top", "left", "bottom", "right", "insideH", "insideV")
_CAPTION_SIZE_PT = 10.5
_THIN = 4  # 0.5 pt
_MEDIUM = 6  # 0.75 pt
_THICK = 12  # 1.5 pt
_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
_TC_BORDERS_SUCCESSORS = (
    "w:shd",
    "w:noWrap",
    "w:tcMar",
    "w:textDirection",
    "w:tcFitText",
    "w:vAlign",
    "w:hideMark",
    "w:headers",
    "w:cellIns",
    "w:cellDel",
    "w:cellMerge",
    "w:tcPrChange",
)
_TBL_PR_BORDERS_SUCCESSORS = (
    "w:shd",
    "w:tblLayout",
    "w:tblCellMar",
    "w:tblLook",
    "w:tblCaption",
    "w:tblDescription",
    "w:tblPrChange",
)
_TR_PR_TAIL = (
    "w:trHeight",
    "w:tblHeader",
    "w:tblCellSpacing",
    "w:jc",
    "w:hidden",
    "w:ins",
    "w:del",
    "w:trPrChange",
)


@dataclass
class TableRenderResult:
    """Handles to the inserted caption and table, plus the next cursor."""

    caption: Paragraph
    table: Table
    last_element: Table


def _get_or_add_tc_borders(cell: _Cell) -> BaseOxmlElement:
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcBorders"))
    if existing is not None:
        tc_pr.remove(existing)
    borders = OxmlElement("w:tcBorders")
    tc_pr.insert_element_before(borders, *_TC_BORDERS_SUCCESSORS)
    return borders


def _get_or_add_tbl_borders(table: Table) -> BaseOxmlElement:
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    tbl_pr.insert_element_before(borders, *_TBL_PR_BORDERS_SUCCESSORS)
    return borders


def _set_edge(
    borders: BaseOxmlElement,
    edge: BorderEdge,
    *,
    val: str,
    sz: int = _THIN,
    color: str = "auto",
) -> None:
    element = borders.find(qn(f"w:{edge}"))
    if element is None:
        element = OxmlElement(f"w:{edge}")
        borders.append(element)
    element.set(qn("w:val"), val)
    if val == "nil":
        return
    element.set(qn("w:sz"), str(sz))
    element.set(qn("w:space"), "0")
    element.set(qn("w:color"), color)


def _clear_table_borders(table: Table) -> None:
    borders = _get_or_add_tbl_borders(table)
    for edge in _EDGE_ORDER:
        _set_edge(borders, edge, val="nil")


def set_cell_border(cell: _Cell, **edges: dict[str, Any]) -> None:
    """Set arbitrary borders on one cell.

    Each keyword argument is an edge name (``top``, ``left``, ``bottom``,
    ``right``) mapping to ``{"val": "single", "sz": 12, "color": "auto"}``.
    """

    borders = _get_or_add_tc_borders(cell)
    for edge, attributes in edges.items():
        if edge not in _EDGE_ORDER:
            raise ValueError(f"unsupported border edge: {edge}")
        _set_edge(
            borders,
            edge,  # type: ignore[arg-type]
            val=str(attributes.get("val", "single")),
            sz=int(attributes.get("sz", _THIN)),
            color=str(attributes.get("color", "auto")),
        )


def apply_three_line_borders(table: Table) -> None:
    """Top/bottom 1.5 pt, header underline 0.75 pt, no vertical rules."""

    _clear_table_borders(table)
    rows = table.rows
    if not rows:
        return
    last_row = len(rows) - 1
    for row_index, row in enumerate(rows):
        for cell in row.cells:
            edges: dict[str, dict[str, Any]] = {
                "top": {"val": "nil"},
                "left": {"val": "nil"},
                "bottom": {"val": "nil"},
                "right": {"val": "nil"},
            }
            if row_index == 0:
                edges["top"] = {"val": "single", "sz": _THICK}
                edges["bottom"] = {"val": "single", "sz": _MEDIUM}
            if row_index == last_row:
                edges["bottom"] = {"val": "single", "sz": _THICK}
            set_cell_border(cell, **edges)


def apply_grid_borders(table: Table) -> None:
    """Every edge of every cell at 0.5 pt (traditional engineering grid)."""

    borders = _get_or_add_tbl_borders(table)
    for edge in _EDGE_ORDER:
        _set_edge(borders, edge, val="single", sz=_THIN)
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(
                cell,
                top={"val": "single", "sz": _THIN},
                left={"val": "single", "sz": _THIN},
                bottom={"val": "single", "sz": _THIN},
                right={"val": "single", "sz": _THIN},
            )


def prevent_row_split(table: Table) -> None:
    """Add ``w:cantSplit`` to every row so rows never break across pages."""

    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is not None:
            continue
        cant_split = OxmlElement("w:cantSplit")
        cant_split.set(qn("w:val"), "1")
        tr_pr.insert_element_before(cant_split, *_TR_PR_TAIL)


def mark_header_row(table: Table) -> None:
    """Mark row 0 as a repeating header row."""

    if not table.rows:
        return
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    existing = tr_pr.find(qn("w:tblHeader"))
    if existing is not None:
        existing.set(qn("w:val"), "1")
        return
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "1")
    tr_pr.insert_element_before(
        header,
        "w:tblCellSpacing",
        "w:jc",
        "w:hidden",
        "w:ins",
        "w:del",
        "w:trPrChange",
    )


def _caption_snapshot(snapshot: ParagraphStyleSnapshot) -> ParagraphStyleSnapshot:
    run = replace(
        snapshot.run,
        font_name=DEFAULT_CJK_FONT,
        east_asia_font=DEFAULT_CJK_FONT,
        size_pt=_CAPTION_SIZE_PT,
        bold=False,
    )
    return replace(snapshot, run=run, line_spacing=1.0)


def render_table(
    doc: DocumentObject,
    model: TableModel,
    *,
    container: Paragraph | Table,
    snapshot: ParagraphStyleSnapshot | None = None,
) -> TableRenderResult:
    """Insert caption + table immediately after ``container``."""

    body_snapshot = snapshot or default_style(doc)
    caption = scanner.insert_paragraph_after(doc, container)
    write_styled_paragraph(
        caption,
        model.title,
        _caption_snapshot(body_snapshot),
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    caption_format = caption.paragraph_format
    caption_format.space_before = Pt(6)
    caption_format.space_after = Pt(3)
    caption_format.keep_with_next = True

    table = scanner.insert_table_after(
        doc,
        caption,
        rows=len(model.rows) + 1,
        cols=len(model.headers),
    )
    table.style = None
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    cell_alignment = _ALIGNMENTS[model.alignment]
    for row_index, row in enumerate(table.rows):
        is_header = row_index == 0
        values = model.headers if is_header else [str(cell) for cell in model.rows[row_index - 1]]
        for column_index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.alignment = cell_alignment
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            add_styled_run(
                paragraph,
                values[column_index],
                body_snapshot,
                bold=is_header,
            )

    if model.table_type is TableType.THREE_LINE:
        apply_three_line_borders(table)
    else:
        apply_grid_borders(table)

    prevent_row_split(table)
    mark_header_row(table)
    return TableRenderResult(caption=caption, table=table, last_element=table)


__all__ = [
    "TableRenderResult",
    "apply_grid_borders",
    "apply_three_line_borders",
    "mark_header_row",
    "prevent_row_split",
    "render_table",
    "set_cell_border",
]
