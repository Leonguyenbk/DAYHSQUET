"""Xử lý một PDF/ảnh theo thứ tự tên file, text layer và OCR."""

from __future__ import annotations

import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from gcn_file_finder.core.cache_manager import CacheManager
from gcn_file_finder.core.gcn_normalizer import (
    extract_controlled_ocr_matches,
    extract_gcn_candidates,
    format_gcn_display,
    match_exact_candidates,
)
from gcn_file_finder.models import FileScanResult, GCNMatch

if TYPE_CHECKING:
    from gcn_file_finder.core.ocr_engine import OCREngine, OCRLine


@dataclass(frozen=True)
class ProcessingOptions:
    """Tùy chọn ảnh hưởng đến kết quả nhận dạng một tài liệu."""

    dpi: int = 150
    verify_content_after_filename: bool = False
    include_red: bool = False
    find_all: bool = True
    thorough_ocr: bool = False

    def cache_key(self, valid_keys: set[str]) -> str:
        return (
            f"v4|{self.dpi}|{int(self.verify_content_after_filename)}|"
            f"{int(self.include_red)}|{int(self.find_all)}|{int(self.thorough_ocr)}|"
            f"{'/'.join(sorted(valid_keys))}"
        )


def _matches_from_text(text: str, valid_keys: set[str], method: str, page: int | None) -> list[GCNMatch]:
    return [
        GCNMatch(item.key, item.display, method, page=page, raw_text=item.raw, confidence=None)
        for item in match_exact_candidates(text, valid_keys)
    ]


def _matches_from_ocr_lines(
    lines: list[OCRLine], valid_keys: set[str], method: str, page: int | None
) -> list[GCNMatch]:
    matches: list[GCNMatch] = []
    for line in lines:
        exact = match_exact_candidates(line.text, valid_keys)
        for item in exact:
            matches.append(GCNMatch(item.key, item.display, method, page, line.text, line.confidence))
        if exact:
            continue
        for corrected in extract_controlled_ocr_matches(line.text, valid_keys):
            if corrected.key:
                matches.append(
                    GCNMatch(
                        corrected.key,
                        format_gcn_display(corrected.key),
                        "OCR SỬA KÝ TỰ CÓ KIỂM SOÁT",
                        page,
                        line.text,
                        line.confidence,
                    )
                )
            elif corrected.needs_review:
                matches.append(
                    GCNMatch(
                        None,
                        None,
                        "OCR SỬA KÝ TỰ CÓ KIỂM SOÁT",
                        page,
                        line.text,
                        line.confidence,
                        needs_review=True,
                        alternatives=corrected.alternatives,
                    )
                )
    # Paddle có thể tách "AM" và "143443" thành hai dòng OCR. Ghép các dòng
    # để regex khoảng trắng vẫn nhận được mã, nhưng báo cáo chỉ giữ đúng token.
    if len(lines) > 1:
        combined = "\n".join(line.text for line in lines)
        confidence = min(line.confidence for line in lines)
        existing = {match.key for match in matches if match.key}
        for item in match_exact_candidates(combined, valid_keys):
            if item.key not in existing:
                matches.append(GCNMatch(item.key, item.display, method, page, item.raw, confidence))
                existing.add(item.key)
        for corrected in extract_controlled_ocr_matches(combined, valid_keys):
            if corrected.key and corrected.key not in existing:
                matches.append(
                    GCNMatch(
                        corrected.key, format_gcn_display(corrected.key),
                        "OCR SỬA KÝ TỰ CÓ KIỂM SOÁT", page, corrected.raw, confidence,
                    )
                )
                existing.add(corrected.key)
            elif corrected.needs_review:
                matches.append(
                    GCNMatch(
                        None, None, "OCR SỬA KÝ TỰ CÓ KIỂM SOÁT", page,
                        corrected.raw, confidence, True, corrected.alternatives,
                    )
                )
    return matches


def _deduplicate(matches: list[GCNMatch]) -> list[GCNMatch]:
    """Bỏ lặp cùng mã/trang, ưu tiên lần phát hiện đầu tiên."""

    output: list[GCNMatch] = []
    seen: set[tuple[object, ...]] = set()
    for match in matches:
        marker = (match.key, match.page, match.needs_review, match.alternatives)
        if marker not in seen:
            output.append(match)
            seen.add(marker)
    return output


def _ocr_image(
    image,
    engine: OCREngine,
    valid_keys: set[str],
    include_red: bool,
    page: int | None,
    find_all: bool,
    thorough_ocr: bool = False,
) -> list[GCNMatch]:
    from gcn_file_finder.core.image_processor import (
        enhanced_gray,
        limit_image_size,
        orientation_variants,
        preprocessing_variants,
    )

    review_matches: list[GCNMatch] = []
    # Normalize once for all rotations/filters. Otherwise PaddleX resizes and
    # prints the same max-side warning for every OCR variant.
    image = limit_image_size(image)
    if not thorough_ocr:
        # Chế độ mặc định ưu tiên tốc độ: ảnh gốc trước, tương tự quy trình
        # đơn giản của locfilescan_2.py. Chỉ khi ảnh gốc không khớp mới thử
        # thêm một lượt ảnh xám tăng tương phản (rẻ, thường đủ sửa các ca chữ
        # đen trên nền không đồng màu mà ảnh màu bỏ sót). Tách màu đỏ chỉ thêm
        # khi người dùng chủ động bật tùy chọn đó.
        variants = [("OCR ẢNH GỐC", image), ("OCR ẢNH XÁM", enhanced_gray(image))]
        if include_red:
            variants.extend(
                (method, variant)
                for method, variant in preprocessing_variants(image, include_red=True)
                if method in ("OCR TÁCH MÀU ĐỎ", "OCR NỀN ĐỎ THÀNH TRẮNG")
            )
        for method, variant in variants:
            matches = _matches_from_ocr_lines(engine.recognize(variant), valid_keys, method, page)
            exact = [match for match in matches if match.key]
            review_matches.extend(match for match in matches if match.needs_review)
            if exact:
                selected = exact if find_all else exact[:1]
                return _deduplicate(selected + review_matches)
        return _deduplicate(review_matches)

    for oriented in orientation_variants(image):
        for method, variant in preprocessing_variants(oriented, include_red):
            matches = _matches_from_ocr_lines(engine.recognize(variant), valid_keys, method, page)
            exact = [match for match in matches if match.key]
            review_matches.extend(match for match in matches if match.needs_review)
            if exact:
                selected = exact if find_all else exact[:1]
                return _deduplicate(selected + review_matches)
    return _deduplicate(review_matches)


def serialize_result(result: FileScanResult) -> dict:
    return {
        "file_type": result.file_type,
        "page_count": result.page_count,
        "has_text_layer": result.has_text_layer,
        "matches": [asdict(match) for match in result.matches],
        "status": result.status,
        "error": result.error,
    }


def deserialize_result(path: Path, payload: dict) -> FileScanResult:
    match_payloads = []
    for item in payload.get("matches", []):
        converted = dict(item)
        converted["alternatives"] = tuple(converted.get("alternatives", ()))
        match_payloads.append(converted)
    return FileScanResult(
        source_path=path,
        file_type=payload["file_type"],
        page_count=int(payload.get("page_count", 1)),
        has_text_layer=bool(payload.get("has_text_layer")),
        matches=[GCNMatch(**item) for item in match_payloads],
        status=payload.get("status", "KHÔNG TÌM THẤY"),
        error=payload.get("error", ""),
        from_cache=True,
    )


def process_document(
    path: str | Path,
    valid_keys: set[str],
    engine: OCREngine,
    options: ProcessingOptions,
    cache: CacheManager | None = None,
) -> FileScanResult:
    """Xử lý độc lập một tài liệu; lỗi file không lan sang toàn bộ phiên chạy."""

    source = Path(path)
    started = time.perf_counter()
    settings_key = options.cache_key(valid_keys)
    if cache:
        cached = cache.get(source, settings_key)
        # A model/IO failure may be temporary, so never reuse an error result.
        if cached and cached.get("status") != "LỖI XỬ LÝ":
            return deserialize_result(source, cached)
    result = FileScanResult(source_path=source, file_type=source.suffix.lower().lstrip(".").upper())
    try:
        filename_candidates = extract_gcn_candidates(source.stem)
        filename_matches = _matches_from_text(source.stem, valid_keys, "TÊN FILE", None)
        if filename_candidates and not options.verify_content_after_filename:
            # Giữ hành vi nhanh của locfilescan_2.py: tên đã có dạng GCN thì
            # không mở/render/OCR file. Mã thuộc Excel vẫn được ghi nhận/copy.
            result.matches = _deduplicate(filename_matches if options.find_all else filename_matches[:1])
        elif source.suffix.lower() == ".pdf":
            from gcn_file_finder.core.pdf_reader import extract_page_texts, open_pdf, render_page_bgr

            with open_pdf(source) as document:
                result.page_count = document.page_count
                text_matches: list[GCNMatch] = []
                for page, text in extract_page_texts(document):
                    if text.strip():
                        result.has_text_layer = True
                    text_matches.extend(_matches_from_text(text, valid_keys, "PDF TEXT", page))
                if text_matches:
                    deduplicated = _deduplicate(text_matches)
                    result.matches = deduplicated if options.find_all else deduplicated[:1]
                else:
                    ocr_matches: list[GCNMatch] = []
                    for page_index in range(document.page_count):
                        image = render_page_bgr(document, page_index, options.dpi)
                        found = _ocr_image(
                            image, engine, valid_keys, options.include_red, page_index + 1,
                            options.find_all, options.thorough_ocr,
                        )
                        ocr_matches.extend(found)
                        del image
                        if any(item.key for item in found) and not options.find_all:
                            break
                    result.matches = _deduplicate(ocr_matches)
        else:
            from gcn_file_finder.core.image_processor import load_image

            image = load_image(source)
            result.matches = _ocr_image(
                image, engine, valid_keys, options.include_red, None,
                options.find_all, options.thorough_ocr,
            )
            del image

        if any(match.needs_review for match in result.matches):
            result.status = "CẦN KIỂM TRA"
        elif any(match.key for match in result.matches):
            result.status = "ĐÃ TÌM THẤY"
        else:
            result.status = "KHÔNG TÌM THẤY"
    except MemoryError:
        result.status = "LỖI XỬ LÝ"
        result.error = "Không đủ bộ nhớ để xử lý tài liệu. Hãy giảm DPI hoặc số luồng."
    except Exception as exc:
        result.status = "LỖI XỬ LÝ"
        result.error = str(exc)
    result.elapsed_seconds = time.perf_counter() - started
    # Caching errors makes later runs immediately return LỖI XỬ LÝ without
    # retrying OCR after the environment or input has been fixed.
    if cache and result.status != "LỖI XỬ LÝ":
        cache.put(source, serialize_result(result), settings_key)
    return result
