"""Kiểm thử phần chuyển đổi dữ liệu của GUI Flet."""

from pathlib import Path

import pytest

from tach_gcn.gui import GuiSettings, build_arguments


def test_gui_builds_cli_arguments_without_exposing_key() -> None:
    """GUI truyền key trong bộ nhớ và chuyển đủ các tuỳ chọn sang lõi CLI."""

    settings = GuiSettings(
        input_dir="  C:/PDF  ",
        output_dir="C:/KetQua",
        api_key="  secret-key  ",
        keywords="Xác nhận hoàn tất;CHỨNG CHỈ ĐÀO TẠO\nMẪU RIÊNG",
        mode="each",
        batch="8",
        rpm="12.5",
        workers="2",
        include_rest=True,
        gemini_all=True,
        low_res=True,
    )
    arguments = build_arguments(settings)
    assert arguments.thu_muc == Path("C:/PDF")
    assert arguments.output == Path("C:/KetQua")
    assert arguments.api_key == "secret-key"
    assert arguments.keyword == [
        "Xác nhận hoàn tất",
        "CHỨNG CHỈ ĐÀO TẠO",
        "MẪU RIÊNG",
    ]
    assert (arguments.mode, arguments.batch, arguments.rpm, arguments.workers) == (
        "each",
        8,
        12.5,
        2,
    )
    assert "secret-key" not in repr(settings)


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("input_dir", "", "chọn thư mục"),
        ("batch", "0", "batch"),
        ("rpm", "abc", "RPM"),
        ("workers", "-1", "worker"),
    ],
)
def test_gui_validates_required_numbers(field: str, value: str, message: str) -> None:
    """Thông báo lỗi biểu mẫu rõ ràng trước khi chạy worker."""

    settings = GuiSettings(input_dir="C:/PDF")
    setattr(settings, field, value)
    with pytest.raises(ValueError, match=message):
        build_arguments(settings)
