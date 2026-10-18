"""Kiểm thử gán trang vào GCN đã biết của đơn."""

from __future__ import annotations

import json
from types import SimpleNamespace
from typing import Any

import pytest

from tach_gcn.gemini import ASSIGN_OTHER, GeminiAnalyzer


class FakeModels:
    """SDK giả: ghi lại prompt và trả JSON dựng sẵn."""

    def __init__(self, pages: list[dict[str, Any]]) -> None:
        self.pages = pages
        self.prompts: list[str] = []

    def generate_content(self, model: str, contents: list[Any], config: Any) -> Any:
        self.prompts.append(contents[0])
        return SimpleNamespace(
            text=json.dumps({"pages": self.pages}),
            usage_metadata=SimpleNamespace(prompt_token_count=10, total_token_count=12),
        )


def _analyzer(pages: list[dict[str, Any]]) -> tuple[GeminiAnalyzer, FakeModels]:
    models = FakeModels(pages)
    return GeminiAnalyzer("key", client=SimpleNamespace(models=models), rpm=6000), models


def test_assign_normalizes_serial_and_marks_unknown() -> None:
    """Số viết khác khoảng trắng/hoa thường vẫn khớp; số lạ thành KHAC."""

    analyzer, models = _analyzer(
        [
            {"i": 1, "gcn": "dh770214", "serial_seen": "DH 770214", "starts_copy": True, "reason": "bìa"},
            {"i": 2, "gcn": "", "serial_seen": "", "attached_to": "DH770214", "reason": "trích lục kèm GCN"},
            {"i": 3, "gcn": "ZZ 999999", "serial_seen": "ZZ 999999", "reason": "GCN cũ"},
        ]
    )
    result, usage = analyzer.assign_certificates(
        [(1, b"a"), (2, b"b"), (3, b"c")],
        [("DH 770214", "CS10707")],
        "trang ngay trước lô này (i=0) không phải trang GCN.",
    )
    assert result[1].gcn == "DH 770214"
    assert result[2].gcn == ""
    assert result[3].gcn == ASSIGN_OTHER
    assert result[1].starts_copy is True and result[2].starts_copy is False
    assert result[2].attached_to == "DH 770214" and result[1].attached_to == ""
    assert usage.input_tokens == 10
    assert "CS10707" in models.prompts[0] and "Ngữ cảnh" in models.prompts[0]


def test_assign_rejects_missing_pages() -> None:
    """Thiếu trang trong JSON trả về phải báo lỗi, không lặng lẽ bỏ qua."""

    analyzer, _ = _analyzer([{"i": 1, "gcn": "", "serial_seen": "", "reason": ""}])
    with pytest.raises(ValueError):
        analyzer.assign_certificates([(1, b"a"), (2, b"b")], [("DH 770214", "")])
