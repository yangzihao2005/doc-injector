"""Pydantic contracts for the doc-injector payload.

The models are intentionally strict: an invalid payload fails before the
template document is touched, which keeps the surgical guarantee meaningful.
"""

from __future__ import annotations

import json
from enum import Enum
from pathlib import Path
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

Scalar = str | float | int


class TableType(str, Enum):  # noqa: UP042 - spec contract keeps str+Enum
    """Supported academic table layouts."""

    THREE_LINE = "three_line"
    GRID = "grid"


class TableModel(BaseModel):
    """A captioned table. The caption is always rendered above the table."""

    model_config = ConfigDict(extra="forbid")

    title: str = Field(
        ...,
        min_length=1,
        description="表题，必须位于表格上方，例如：表1-1 采样数据",
    )
    table_type: TableType = TableType.THREE_LINE
    headers: list[str] = Field(..., min_length=1)
    rows: list[list[Scalar]] = Field(default_factory=list)
    alignment: str = Field(default="center", description="单元格水平对齐：left/center/right")

    @field_validator("title")
    @classmethod
    def _title_not_blank(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("table.title must not be blank")
        return value

    @field_validator("headers")
    @classmethod
    def _headers_not_blank(cls, value: list[str]) -> list[str]:
        cleaned = [str(item).strip() for item in value]
        if not any(cleaned):
            raise ValueError("table.headers must contain at least one non-empty cell")
        return cleaned

    @field_validator("alignment")
    @classmethod
    def _alignment_supported(cls, value: str) -> str:
        normalized = value.strip().lower()
        if normalized not in {"left", "center", "right"}:
            raise ValueError("alignment must be one of: left, center, right")
        return normalized

    @field_validator("rows")
    @classmethod
    def _cells_are_scalars(cls, value: list[list[Scalar]]) -> list[list[Scalar]]:
        for row in value:
            if not row:
                raise ValueError("table rows must not be empty")
            for cell in row:
                if cell is None:  # type: ignore[comparison-overlap]
                    raise ValueError("table cells must not be None")
        return value

    @model_validator(mode="after")
    def _rows_match_headers(self) -> TableModel:
        width = len(self.headers)
        for row_index, row in enumerate(self.rows, start=1):
            if len(row) != width:
                raise ValueError(
                    f"row {row_index} has {len(row)} cells, expected {width} (len(headers))"
                )
        return self


class ImageModel(BaseModel):
    """A captioned image. The caption is always rendered below the image."""

    model_config = ConfigDict(extra="forbid")

    title: str = Field(
        ...,
        min_length=1,
        description="图题，必须位于图片下方，例如：图2-1 时域波形",
    )
    path: str = Field(..., min_length=1)
    max_width_ratio: float = Field(default=0.8, gt=0, le=1, description="占版心最大宽度比例")

    @field_validator("title", "path")
    @classmethod
    def _not_blank(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("value must not be blank")
        return value


class ContentBlock(BaseModel):
    """One anchor-driven injection unit."""

    model_config = ConfigDict(extra="forbid")

    anchor_keyword: str = Field(..., min_length=1, description="目标锚点关键词，例如：实验目的")
    paragraphs: list[str] | None = None
    custom_font: str | None = None
    custom_size_pt: float | None = Field(default=None, gt=0, le=200)
    table: TableModel | None = None
    images: list[ImageModel] | None = None

    @field_validator("anchor_keyword")
    @classmethod
    def _anchor_not_blank(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("anchor_keyword must not be blank")
        return value

    @field_validator("paragraphs")
    @classmethod
    def _paragraphs_not_blank(cls, value: list[str] | None) -> list[str] | None:
        if value is None:
            return None
        if not value:
            raise ValueError("paragraphs must not be an empty list; omit it instead")
        return value

    @field_validator("images")
    @classmethod
    def _images_not_empty(cls, value: list[ImageModel] | None) -> list[ImageModel] | None:
        if value is not None and not value:
            raise ValueError("images must not be an empty list; omit it instead")
        return value

    @model_validator(mode="after")
    def _has_content(self) -> ContentBlock:
        if not self.paragraphs and self.table is None and not self.images:
            raise ValueError("content block needs at least one of: paragraphs, table, images")
        return self


class ReportInjectionPayload(BaseModel):
    """Root payload accepted by ``doc-injector inject``."""

    model_config = ConfigDict(extra="forbid")

    blocks: list[ContentBlock] = Field(..., min_length=1)


def load_payload(path: str | Path) -> ReportInjectionPayload:
    """Load and validate a JSON payload from disk."""

    raw: Any = json.loads(Path(path).read_text(encoding="utf-8"))
    return ReportInjectionPayload.model_validate(raw)


def payload_json_schema() -> dict[str, Any]:
    """Return the JSON schema for payload authoring / editor autocomplete."""

    return ReportInjectionPayload.model_json_schema()
