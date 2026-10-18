"""Đọc, render và tách trang PDF bằng PyMuPDF."""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import os
from pathlib import Path
import re
from typing import Iterable, Sequence

import pymupdf

from .detect import (
    DetectionResult,
    detect_page,
    extract_certificate_serial,
    extract_issue_date,
    extract_owner_name,
    extract_registration_number,
    is_certificate_support_page,
)


@dataclass(slots=True)
class InspectedPage:
    """Kết quả đọc một trang, kèm ảnh khi cần Gemini."""

    result: DetectionResult
    jpeg: bytes | None = None
    text: str = ""
    is_certificate_support: bool = False


@dataclass(slots=True)
class CertificateGroup:
    """Các trang thuộc cùng một GCN và thông tin được trích xuất."""

    page_indexes: list[int]
    certificate_serial: str = ""
    registration_number: str = ""
    issue_date: str = ""
    owner_name: str = ""
    note: str = ""
    output_file: str = ""


@dataclass(slots=True)
class InspectedFile:
    """Kết quả kiểm tra một file PDF trong worker."""

    path: str
    relative_path: str
    content_hash: str = ""
    page_count: int = 0
    pages: list[InspectedPage] = field(default_factory=list)
    certificate_groups: list[CertificateGroup] = field(default_factory=list)
    recovered_page_xrefs: list[int] = field(default_factory=list)
    warning: str = ""
    error: str = ""


def hash_file(path: Path) -> str:
    """Tính SHA-256 nội dung file theo từng khối."""

    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


_PDF_REFERENCE = re.compile(r"(\d+)\s+\d+\s+R")
_PAGE_TYPE = re.compile(r"/Type\s*/Page(?!s)\b")
_PAGES_TYPE = re.compile(r"/Type\s*/Pages\b")


def _page_tree_root(document: pymupdf.Document) -> int:
    """Lấy xref nút gốc của cây trang PDF."""

    kind, value = document.xref_get_key(document.pdf_catalog(), "Pages")
    if kind != "xref":
        raise ValueError("PDF không có cây trang hợp lệ")
    match = _PDF_REFERENCE.search(value)
    if match is None:
        raise ValueError("Không đọc được tham chiếu cây trang PDF")
    return int(match.group(1))


def _object_text(path: Path, xref: int) -> str:
    """Đọc một object bằng tài liệu riêng để object hỏng không làm hỏng phiên chính."""

    probe = pymupdf.open(path)
    try:
        return probe.xref_object(xref, compressed=False)
    except Exception:
        return ""
    finally:
        probe.close()


def _recover_page_xrefs(path: Path) -> tuple[list[int], int]:
    """Lần theo cây trang lỗi và giữ các object thực sự có kiểu ``/Page``."""

    document = pymupdf.open(path)
    try:
        root = _page_tree_root(document)
        declared_count = int(document.xref_get_key(root, "Count")[1] or 0)
    finally:
        document.close()

    recovered: list[int] = []
    visiting: set[int] = set()

    def visit(xref: int) -> None:
        if xref in visiting:
            return
        visiting.add(xref)
        value = _object_text(path, xref)
        if _PAGE_TYPE.search(value):
            recovered.append(xref)
        elif _PAGES_TYPE.search(value):
            kids_match = re.search(r"/Kids\s*\[(.*?)\]", value, re.DOTALL)
            if kids_match:
                for child in _PDF_REFERENCE.findall(kids_match.group(1)):
                    visit(int(child))

    old_errors = bool(pymupdf.TOOLS.mupdf_display_errors())
    old_warnings = bool(pymupdf.TOOLS.mupdf_display_warnings())
    pymupdf.TOOLS.mupdf_display_errors(False)
    pymupdf.TOOLS.mupdf_display_warnings(False)
    try:
        visit(root)
        pymupdf.TOOLS.mupdf_warnings(reset=True)
    finally:
        pymupdf.TOOLS.mupdf_display_errors(old_errors)
        pymupdf.TOOLS.mupdf_display_warnings(old_warnings)
    return recovered, declared_count


def open_pdf_document(
    path: Path,
    recovered_page_xrefs: Sequence[int] = (),
) -> pymupdf.Document:
    """Mở PDF và áp cây trang đã phục hồi nếu file nguồn bị ghép sai."""

    document = pymupdf.open(path)
    if not recovered_page_xrefs:
        return document
    root = _page_tree_root(document)
    references = " ".join(f"{xref} 0 R" for xref in recovered_page_xrefs)
    document.xref_set_key(root, "Kids", f"[{references}]")
    document.xref_set_key(root, "Count", str(len(recovered_page_xrefs)))
    for xref in recovered_page_xrefs:
        document.xref_set_key(xref, "Parent", f"{root} 0 R")
    return document


def _open_readable_pdf(path: Path) -> tuple[pymupdf.Document, list[int], str]:
    """Mở PDF; nếu cây trang hỏng thì phục hồi các trang còn nguyên vẹn."""

    document = pymupdf.open(path)
    if document.needs_pass:
        return document, [], ""
    old_errors = bool(pymupdf.TOOLS.mupdf_display_errors())
    old_warnings = bool(pymupdf.TOOLS.mupdf_display_warnings())
    pymupdf.TOOLS.mupdf_display_errors(False)
    pymupdf.TOOLS.mupdf_display_warnings(False)
    try:
        for index in range(document.page_count):
            page = document.load_page(index)
            if document.xref_get_key(page.xref, "Type") != ("name", "/Page"):
                raise ValueError(f"Mục cây trang {index + 1} không phải /Page")
        return document, [], ""
    except Exception as original_error:
        document.close()
        pymupdf.TOOLS.mupdf_warnings(reset=True)
        pymupdf.TOOLS.mupdf_display_errors(old_errors)
        pymupdf.TOOLS.mupdf_display_warnings(old_warnings)
        recovered, declared_count = _recover_page_xrefs(path)
        if not recovered:
            raise original_error
        document = open_pdf_document(path, recovered)
        # Buộc MuPDF kiểm tra cây trang vừa thay trước khi xử lý nội dung.
        for index in range(document.page_count):
            document.load_page(index)
        warning = (
            "PDF lỗi cây trang; khôi phục được "
            f"{len(recovered)}/{declared_count or len(recovered)} trang hợp lệ"
        )
        return document, recovered, warning
    finally:
        pymupdf.TOOLS.mupdf_display_errors(old_errors)
        pymupdf.TOOLS.mupdf_display_warnings(old_warnings)


def render_page_jpeg(page: pymupdf.Page, long_edge: int = 1400, quality: int = 75) -> bytes:
    """Render trang thành JPEG với cạnh dài xấp xỉ kích thước yêu cầu."""

    longest = max(float(page.rect.width), float(page.rect.height), 1.0)
    zoom = long_edge / longest
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    return pixmap.tobytes("jpeg", jpg_quality=quality)


def render_pdf_pages(
    path: Path,
    page_indexes: Sequence[int],
    recovered_page_xrefs: Sequence[int] = (),
) -> list[tuple[int, bytes]]:
    """Render một nhóm trang, trả số trang 1-based và dữ liệu JPEG."""

    images: list[tuple[int, bytes]] = []
    document = open_pdf_document(path, recovered_page_xrefs)
    try:
        for index in page_indexes:
            images.append((index + 1, render_page_jpeg(document.load_page(index))))
    finally:
        document.close()
    return images


def build_certificate_groups(inspected: InspectedFile) -> list[CertificateGroup]:
    """Gom mặt trước/sau và các trang nằm giữa thành một GCN duy nhất."""

    anchors = [
        page
        for page in inspected.pages
        if page.result.is_certificate or page.is_certificate_support
    ]
    anchors.sort(key=lambda item: item.result.page)
    raw_groups: list[tuple[list[int], str, str]] = []
    current_anchors: list[int] = []
    current_registration = ""
    current_serial = ""

    def finish() -> None:
        if current_anchors:
            raw_groups.append(
                (current_anchors.copy(), current_registration, current_serial)
            )

    for page in anchors:
        index = page.result.page - 1
        page_text = "\n".join((page.text, page.result.title, page.result.reason))
        registration = extract_registration_number(page_text)
        serial = (
            extract_certificate_serial(page_text)
            if page.result.is_certificate and not page.is_certificate_support
            else ""
        )
        starts_new_number = bool(
            current_anchors
            and (
                registration
                and current_registration
                and registration != current_registration
                or serial
                and current_serial
                and serial != current_serial
            )
        )
        too_far = bool(current_anchors and index - current_anchors[-1] > 3)
        if starts_new_number or too_far:
            finish()
            current_anchors = []
            current_registration = ""
            current_serial = ""
        current_anchors.append(index)
        if registration and not current_registration:
            current_registration = registration
        if serial and not current_serial:
            current_serial = serial
    finish()

    results: list[CertificateGroup] = []
    certificate_indexes = {
        page.result.page - 1 for page in inspected.pages if page.result.is_certificate
    }
    for group_anchors, registration, serial in raw_groups:
        if not any(index in certificate_indexes for index in group_anchors):
            continue
        first, last = group_anchors[0], group_anchors[-1]
        indexes = list(range(first, last + 1))
        group_text = "\n".join(inspected.pages[index].text for index in indexes)
        certificate_text = "\n".join(
            inspected.pages[index].text
            for index in indexes
            if inspected.pages[index].result.is_certificate
            and not inspected.pages[index].is_certificate_support
        )
        results.append(
            CertificateGroup(
                page_indexes=indexes,
                certificate_serial=serial
                or extract_certificate_serial(certificate_text),
                registration_number=registration
                or extract_registration_number(group_text),
                issue_date=extract_issue_date(group_text),
                owner_name=extract_owner_name(group_text),
                note=inspected.warning,
            )
        )
    inspected.certificate_groups = results
    return results


def inspect_pdf(
    path_text: str,
    relative_path: str,
    custom_keywords: Sequence[str],
    gemini_all: bool,
    gemini_enabled: bool,
) -> InspectedFile:
    """Đọc chữ và chấm điểm một file; hàm này chạy được trong process worker."""

    path = Path(path_text)
    inspected = InspectedFile(path_text, relative_path)
    try:
        inspected.content_hash = hash_file(path)
        document, recovered_xrefs, warning = _open_readable_pdf(path)
        try:
            if document.needs_pass:
                inspected.error = "File PDF có mật khẩu"
                return inspected
            inspected.recovered_page_xrefs = recovered_xrefs
            inspected.warning = warning
            inspected.page_count = document.page_count
            for index in range(document.page_count):
                page = document.load_page(index)
                text = page.get_text("text")
                result = detect_page(
                    text,
                    page=index + 1,
                    custom_keywords=custom_keywords,
                    gemini_all=gemini_all,
                    gemini_enabled=gemini_enabled,
                )
                jpeg = render_page_jpeg(page) if result.needs_gemini else None
                inspected.pages.append(
                    InspectedPage(
                        result=result,
                        jpeg=jpeg,
                        text=text,
                        is_certificate_support=is_certificate_support_page(text),
                    )
                )
            if inspected.warning and inspected.pages:
                inspected.pages[0].result.note = inspected.warning
        finally:
            document.close()
    except Exception as error:
        inspected.error = f"Không đọc được PDF: {error}"
        inspected.pages.clear()
        inspected.page_count = 0
    return inspected


def discover_pdfs(input_dir: Path, output_dir: Path) -> list[Path]:
    """Tìm PDF đệ quy và loại toàn bộ cây thư mục đầu ra."""

    input_resolved = input_dir.resolve()
    output_resolved = output_dir.resolve()
    output_is_inside_input = (
        output_resolved == input_resolved or input_resolved in output_resolved.parents
    )
    files: list[Path] = []
    for current, directories, names in os.walk(input_resolved, followlinks=False):
        current_path = Path(current)
        kept_directories: list[str] = []
        for directory in directories:
            candidate = (current_path / directory).resolve()
            if output_is_inside_input and (
                candidate == output_resolved or output_resolved in candidate.parents
            ):
                continue
            kept_directories.append(directory)
        directories[:] = kept_directories
        for name in names:
            if name.lower().endswith(".pdf"):
                files.append(current_path / name)
    return sorted(files, key=lambda item: str(item).casefold())


def _save_selected_pages(
    source: Path,
    page_indexes: Iterable[int],
    destination: Path,
    recovered_page_xrefs: Sequence[int] = (),
) -> None:
    """Lưu các chỉ số trang 0-based vào file đích theo cách thay thế nguyên tử."""

    indexes = list(page_indexes)
    if not indexes:
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f".{destination.stem}.{os.getpid()}.tmp.pdf")
    source_document = open_pdf_document(source, recovered_page_xrefs)
    output_document = pymupdf.open()
    try:
        for index in indexes:
            output_document.insert_pdf(source_document, from_page=index, to_page=index)
        output_document.save(temporary, garbage=4, deflate=True)
        os.replace(temporary, destination)
    finally:
        output_document.close()
        source_document.close()
        if temporary.exists():
            temporary.unlink()


def write_repaired_pdf(inspected: InspectedFile, output_dir: Path) -> Path | None:
    """Ghi một bản PDF đã dựng lại cây trang khi file nguồn bị ghép hỏng."""

    if not inspected.recovered_page_xrefs or inspected.page_count < 1:
        return None
    source = Path(inspected.path)
    relative = Path(inspected.relative_path)
    destination = output_dir / relative.parent / f"{relative.stem}-DA-SUA.pdf"
    _save_selected_pages(
        source,
        range(inspected.page_count),
        destination,
        inspected.recovered_page_xrefs,
    )
    return destination


def write_outputs(
    inspected: InspectedFile,
    output_dir: Path,
    *,
    mode: str,
    include_rest: bool,
) -> list[Path]:
    """Tách các trang đã phân loại và giữ nguyên thư mục tương đối."""

    source = Path(inspected.path)
    relative = Path(inspected.relative_path)
    target_directory = output_dir / relative.parent
    certificate_pages = [page for page in inspected.pages if page.result.is_certificate]
    certificate_indexes = [page.result.page - 1 for page in certificate_pages]
    groups = inspected.certificate_groups or build_certificate_groups(inspected)
    grouped_indexes = {index for group in groups for index in group.page_indexes}
    rest_indexes = [index for index in range(inspected.page_count) if index not in grouped_indexes]
    created: list[Path] = []

    if mode == "perfile" and certificate_indexes:
        used_names: set[str] = set()
        for position, group in enumerate(groups, start=1):
            base_name = group.certificate_serial or relative.stem
            file_name = f"{base_name}-GCN.pdf"
            if file_name.casefold() in used_names:
                file_name = f"{base_name}-GCN-{position}.pdf"
            used_names.add(file_name.casefold())
            target = target_directory / file_name
            _save_selected_pages(
                source,
                group.page_indexes,
                target,
                inspected.recovered_page_xrefs,
            )
            group.output_file = target.relative_to(output_dir).as_posix()
            created.append(target)
            if not group.certificate_serial:
                group.note = "; ".join(
                    part
                    for part in (
                        group.note,
                        "Không tìm thấy số giấy chứng nhận/mã phôi; dùng tên file gốc",
                    )
                    if part
                )
    elif mode == "each":
        used_names: set[str] = set()
        for page in certificate_pages:
            index = page.result.page - 1
            number = extract_certificate_serial(
                "\n".join((page.text, page.result.title, page.result.reason))
            )
            base_name = f"{number}-GCN" if number else f"{relative.stem}-GCN"
            file_name = f"{base_name}.pdf"
            if file_name.casefold() in used_names:
                file_name = f"{base_name}-trang-{index + 1:03d}.pdf"
            used_names.add(file_name.casefold())
            target = target_directory / file_name
            _save_selected_pages(
                source,
                [index],
                target,
                inspected.recovered_page_xrefs,
            )
            created.append(target)
    if include_rest and rest_indexes:
        target = target_directory / f"{relative.stem}-GT.pdf"
        _save_selected_pages(
            source,
            rest_indexes,
            target,
            inspected.recovered_page_xrefs,
        )
        created.append(target)
    return created
