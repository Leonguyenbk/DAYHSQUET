"""Hàm tạo PDF mẫu ngay trong quá trình kiểm thử."""

from __future__ import annotations

from pathlib import Path
import re

import pymupdf


def unicode_font() -> Path:
    """Tìm font TrueType Unicode phổ biến trên Windows, macOS và Linux."""

    candidates = (
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/tahoma.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise RuntimeError("Không tìm thấy font Unicode để tạo PDF kiểm thử")


def _insert_unicode(page: pymupdf.Page, text: str, font: Path, y: float = 100) -> None:
    """Chèn văn bản Unicode bằng font nhúng."""

    page.insert_font(fontname="UnicodeTest", fontfile=str(font))
    page.insert_text(
        pymupdf.Point(72, y),
        text,
        fontname="UnicodeTest",
        fontsize=20,
    )


def create_sample_pdf(path: Path) -> None:
    """Tạo PDF gồm hai chứng nhận, một quyết định và một trang scan."""

    font = unicode_font()
    path.parent.mkdir(parents=True, exist_ok=True)
    document = pymupdf.open()

    page = document.new_page()
    _insert_unicode(page, "GIẤY CHỨNG NHẬN", font)
    _insert_unicode(page, "Cấp cho: Nguyễn Văn An", font, 150)
    _insert_unicode(page, "Số giấy chứng nhận: BB 654321", font, 200)
    _insert_unicode(page, "Số vào sổ cấp GCN: CS 012345", font, 250)
    _insert_unicode(page, "Ngày cấp GCN: 01/02/2020", font, 300)
    _insert_unicode(page, "Chủ sử dụng: Nguyễn Văn An", font, 350)

    page = document.new_page()
    _insert_unicode(page, "CERTIFICATE OF COMPLETION", font)
    _insert_unicode(page, "AWARDED TO: TRAN AN", font, 150)

    page = document.new_page()
    _insert_unicode(page, "CHỨNG NHẬN là căn cứ trong QUYẾT ĐỊNH này", font)
    _insert_unicode(page, "Điều 1. Ban hành danh sách kèm theo.", font, 150)

    scan_source = pymupdf.open()
    scan_page = scan_source.new_page()
    _insert_unicode(scan_page, "GIẤY CHỨNG NHẬN BẢN QUÉT", font)
    image = scan_page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), alpha=False).tobytes("png")
    scan_source.close()
    page = document.new_page()
    page.insert_image(page.rect, stream=image)

    document.save(path)
    document.close()


def create_password_pdf(path: Path) -> None:
    """Tạo PDF yêu cầu mật khẩu người dùng."""

    path.parent.mkdir(parents=True, exist_ok=True)
    document = pymupdf.open()
    document.new_page().insert_text((72, 72), "SECRET")
    document.save(
        path,
        encryption=pymupdf.PDF_ENCRYPT_AES_256,
        owner_pw="owner-test",
        user_pw="user-test",
    )
    document.close()


def create_broken_page_tree_pdf(path: Path) -> None:
    """Tạo PDF có một mục cây trang giả, giống lỗi ghép trang của ``12.pdf``."""

    font = unicode_font()
    path.parent.mkdir(parents=True, exist_ok=True)
    document = pymupdf.open()
    _insert_unicode(document.new_page(), "GIẤY CHỨNG NHẬN", font)
    _insert_unicode(document.new_page(), "TÀI LIỆU KHÁC", font)
    document.save(path)
    document.close()

    document = pymupdf.open(path)
    root = int(document.xref_get_key(document.pdf_catalog(), "Pages")[1].split()[0])
    kids = document.xref_get_key(root, "Kids")[1]
    pages = [int(value) for value in re.findall(r"(\d+)\s+\d+\s+R", kids)]
    invalid = document.get_new_xref()
    document.update_object(invalid, "<< /ProcSet [/PDF] >>")
    document.xref_set_key(
        root,
        "Kids",
        f"[{pages[0]} 0 R {invalid} 0 R {pages[1]} 0 R]",
    )
    document.xref_set_key(root, "Count", "3")
    document.saveIncr()
    document.close()


def create_handoff_fixtures(directory: Path) -> None:
    """Tạo bộ fixture dùng cho lệnh chạy thử lúc bàn giao."""

    create_sample_pdf(directory / "lop-a" / "chung-chi-mau.pdf")
    create_password_pdf(directory / "co-mat-khau.pdf")
    (directory / "file-hong.pdf").write_bytes(b"day khong phai pdf")


if __name__ == "__main__":
    create_handoff_fixtures(Path(__file__).parents[1] / "fixtures")
