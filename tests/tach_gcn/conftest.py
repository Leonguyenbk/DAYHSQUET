"""Fixture pytest cho tach-gcn."""

from __future__ import annotations

from pathlib import Path

import pytest

from fixtures import create_sample_pdf


@pytest.fixture
def sample_pdf(tmp_path: Path) -> Path:
    """Cung cấp PDF bốn trang được tạo động."""

    path = tmp_path / "nguon" / "cap-1" / "ho-so.pdf"
    create_sample_pdf(path)
    return path
