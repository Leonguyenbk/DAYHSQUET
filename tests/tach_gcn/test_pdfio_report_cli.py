"""Kiểm thử PDF, cấu trúc đầu ra, báo cáo và khả năng chịu lỗi."""

from __future__ import annotations

import csv
from io import StringIO
from pathlib import Path

import pymupdf
from rich.console import Console

from tach_gcn.cli import _enrich_certificate_groups, build_parser, run
from tach_gcn.detect import detect_page, is_certificate_support_page
from tach_gcn.gemini import GeminiCertificateInfo, Usage
from tach_gcn.pdfio import (
    InspectedFile,
    InspectedPage,
    build_certificate_groups,
    inspect_pdf,
    write_outputs,
    write_repaired_pdf,
)

from fixtures import create_broken_page_tree_pdf, create_password_pdf


def test_inspect_sample_pages(sample_pdf: Path) -> None:
    """PDF mẫu có đúng hai chứng nhận và một ảnh scan chưa phân tích."""

    inspected = inspect_pdf(
        str(sample_pdf), "cap-1/ho-so.pdf", (), False, False
    )
    assert not inspected.error
    assert inspected.page_count == 4
    assert [page.result.is_certificate for page in inspected.pages] == [True, True, False, False]
    assert inspected.pages[2].result.score in range(1, 5)
    assert inspected.pages[3].result.method == "ảnh scan"
    assert inspected.pages[3].result.note == "chưa phân tích"


def test_password_and_corrupt_files_continue(tmp_path: Path) -> None:
    """File có mật khẩu và file hỏng được chuyển thành lỗi báo cáo."""

    password = tmp_path / "khoa.pdf"
    corrupt = tmp_path / "hong.pdf"
    create_password_pdf(password)
    corrupt.write_bytes(b"not a pdf")
    locked = inspect_pdf(str(password), "khoa.pdf", (), False, False)
    broken = inspect_pdf(str(corrupt), "hong.pdf", (), False, False)
    assert "mật khẩu" in locked.error
    assert "Không đọc được PDF" in broken.error


def test_broken_page_tree_is_recovered_and_can_be_split(tmp_path: Path) -> None:
    """PDF ghép sai cây trang vẫn quét và tách được các trang /Page hợp lệ."""

    source = tmp_path / "loi-cay-trang.pdf"
    output = tmp_path / "ket-qua"
    create_broken_page_tree_pdf(source)
    inspected = inspect_pdf(str(source), source.name, (), False, False)
    assert not inspected.error
    assert inspected.page_count == 2
    assert len(inspected.recovered_page_xrefs) == 2
    assert "khôi phục được 2/3 trang" in inspected.warning
    assert inspected.pages[0].result.is_certificate

    repaired = write_repaired_pdf(inspected, output)
    assert repaired == output / "loi-cay-trang-DA-SUA.pdf"
    with pymupdf.open(repaired) as repaired_document:
        assert repaired_document.page_count == 2

    created = write_outputs(inspected, output, mode="perfile", include_rest=True)
    assert len(created) == 2
    with pymupdf.open(output / "loi-cay-trang-GCN.pdf") as certificate:
        assert certificate.page_count == 1


def test_cli_output_structure_and_csv(sample_pdf: Path, tmp_path: Path) -> None:
    """CLI giữ cây thư mục, tách đúng trang và ghi CSV có BOM."""

    input_dir = sample_pdf.parents[1]
    output_dir = tmp_path / "ket-qua"
    arguments = build_parser().parse_args(
        [str(input_dir), "-o", str(output_dir), "--no-gemini", "--rest", "--workers", "1"]
    )
    stream = StringIO()
    assert run(arguments, Console(file=stream, force_terminal=False, color_system=None)) == 0

    certificate = output_dir / "cap-1" / "BB 654321-GCN.pdf"
    rest = output_dir / "cap-1" / "ho-so-GT.pdf"
    assert certificate.is_file()
    assert rest.is_file()
    with pymupdf.open(certificate) as document:
        assert document.page_count == 2
    with pymupdf.open(rest) as document:
        assert document.page_count == 2

    report = output_dir / "bao-cao.csv"
    assert report.read_bytes().startswith(b"\xef\xbb\xbf")
    with report.open(encoding="utf-8-sig", newline="") as stream_csv:
        rows = list(csv.DictReader(stream_csv))
    assert len(rows) == 4
    assert rows[0]["File"] == "cap-1/ho-so.pdf"
    assert rows[0]["Giấy chứng nhận"] == "Có"
    assert rows[3]["Cách nhận diện"] == "ảnh scan"
    assert rows[3]["Ghi chú hoặc tiêu đề và lý do của Gemini"] == "chưa phân tích"

    info_report = output_dir / "thong-tin-gcn.csv"
    with info_report.open(encoding="utf-8-sig", newline="") as stream_info:
        info_rows = list(csv.DictReader(stream_info))
    assert len(info_rows) == 1
    assert info_rows[0]["Số giấy chứng nhận (mã phôi)"] == "BB 654321"
    assert info_rows[0]["Số vào sổ cấp GCN"] == "CS 012345"
    assert info_rows[0]["Ngày cấp GCN"] == "01/02/2020"
    assert info_rows[0]["Tên chủ sử dụng hiện tại"] == "Nguyễn Văn An"


def test_front_back_and_map_pages_are_one_certificate() -> None:
    """Hai mặt cách nhau bởi sơ đồ được gom, mã trong sơ đồ không chia file."""

    texts = [
        "GIẤY CHỨNG NHẬN\nSố giấy chứng nhận: BB 654321\n"
        "Số vào sổ cấp GCN: CS 012345\n"
        "Ngày cấp GCN: 01/02/2020\nChủ sử dụng: Nguyễn Văn An",
        "SƠ ĐỒ THỬA ĐẤT\nAB987654\nSố thửa 123, tờ bản đồ 45",
        "Ghi chú kỹ thuật thửa đất",
        "NHỮNG THAY ĐỔI SAU KHI CẤP GCN\n"
        "Chuyển nhượng cho: Trần Thị B",
    ]
    pages = [
        InspectedPage(
            detect_page(text, index + 1, gemini_enabled=False),
            text=text,
            is_certificate_support=is_certificate_support_page(text),
        )
        for index, text in enumerate(texts)
    ]
    inspected = InspectedFile("mau.pdf", "mau.pdf", page_count=4, pages=pages)
    groups = build_certificate_groups(inspected)
    assert len(groups) == 1
    assert groups[0].page_indexes == [0, 1, 2, 3]
    assert groups[0].certificate_serial == "BB 654321"
    assert groups[0].registration_number == "CS 012345"
    assert groups[0].owner_name == "Trần Thị B"


def test_gemini_cannot_replace_serial_with_registration_number(
    sample_pdf: Path,
) -> None:
    """Nếu Gemini chép số vào sổ sang mã phôi thì giữ mã phôi đọc từ mặt GCN."""

    class ConfusedAnalyzer:
        def extract_certificate(self, pages: object) -> tuple[GeminiCertificateInfo, Usage]:
            return (
                GeminiCertificateInfo(
                    certificate_serial="CS 012345",
                    registration_number="CS 012345",
                    issue_date="01/02/2020",
                    owner_name="Nguyễn Văn An",
                    reason="Kết quả thử",
                ),
                Usage(),
            )

    inspected = inspect_pdf(str(sample_pdf), "cap-1/ho-so.pdf", (), False, False)
    _enrich_certificate_groups(
        inspected,
        analyzer=ConfusedAnalyzer(),  # type: ignore[arg-type]
        model="model-test",
        cache={},
        use_cache=False,
        usage=Usage(),
    )
    group = inspected.certificate_groups[0]
    assert group.certificate_serial == "BB 654321"
    assert group.registration_number == "CS 012345"
    assert "trùng số vào sổ" in group.note


def test_dry_run_does_not_write_split_pdf(sample_pdf: Path, tmp_path: Path) -> None:
    """Dry-run chỉ tạo báo cáo và cache."""

    input_dir = sample_pdf.parents[1]
    output_dir = tmp_path / "dry"
    arguments = build_parser().parse_args(
        [str(input_dir), "-o", str(output_dir), "--no-gemini", "--dry-run", "--workers", "1"]
    )
    assert run(arguments, Console(file=StringIO(), force_terminal=False)) == 0
    assert (output_dir / "bao-cao.csv").is_file()
    assert not list(output_dir.rglob("*-GCN.pdf"))
