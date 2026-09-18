"""Công cụ debug độc lập: đọc PDF/ảnh và in ra chính xác nội dung máy đọc được.

Không copy file, không cần Excel. Dùng để xem text layer PDF đọc được gì,
OCR từng trang đọc được gì (kèm độ tin cậy), và mã GCN nào được trích ra từ
những nội dung đó — để chẩn đoán vì sao một file cụ thể không được nhận diện.

Ví dụ chạy:
    .venv\\Scripts\\python.exe debug_read_pdf.py "D:\\path\\to\\file.pdf"
    .venv\\Scripts\\python.exe debug_read_pdf.py "D:\\thu_muc" --dpi 300 --thorough
    .venv\\Scripts\\python.exe debug_read_pdf.py "D:\\file.pdf" --excel ds.xlsx --sheet "Sheet1"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from gcn_file_finder.core.config import configured_model_dir, load_config
from gcn_file_finder.core.document_processor import _ocr_image
from gcn_file_finder.core.file_scanner import SUPPORTED_EXTENSIONS
from gcn_file_finder.core.gcn_normalizer import extract_controlled_ocr_matches, extract_gcn_candidates
from gcn_file_finder.core.image_processor import limit_image_size, load_image
from gcn_file_finder.core.ocr_engine import OCREngine
from gcn_file_finder.core.pdf_reader import extract_page_texts, open_pdf, render_page_bgr


def _print_candidates(label: str, text: str, valid_keys: set[str] | None, indent: str = "    ") -> None:
    candidates = extract_gcn_candidates(text)
    if not candidates:
        print(f"{indent}{label}: không tìm thấy mã nào khớp định dạng GCN")
        return
    for item in candidates:
        note = ""
        if valid_keys is not None:
            note = "  [CÓ TRONG EXCEL]" if item.key in valid_keys else "  [KHÔNG CÓ TRONG EXCEL]"
        print(f"{indent}{label}: '{item.raw}' -> {item.display} (khóa {item.key}){note}")


def _print_ocr_lines(lines, valid_keys: set[str] | None) -> None:
    if not lines:
        print("    (OCR không đọc được dòng nào)")
        return
    for line in lines:
        print(f"    OCR: '{line.text}'  conf={line.confidence:.3f}")
        _print_candidates("-> mã trích được", line.text, valid_keys, indent="      ")
        for corrected in extract_controlled_ocr_matches(line.text, valid_keys or set()):
            if corrected.key:
                print(f"      -> sửa lỗi OCR khớp Excel: {corrected.key}")
            elif corrected.needs_review:
                print(f"      -> CẦN KIỂM TRA, nhiều phương án: {corrected.alternatives}")


def process_pdf(path: Path, engine: OCREngine, args, valid_keys: set[str] | None) -> None:
    with open_pdf(path) as document:
        print(f"  Số trang: {document.page_count}")
        page_texts = list(extract_page_texts(document))
        any_text_layer = False
        text_matches = []
        for page_number, text in page_texts:
            stripped = text.strip()
            print(f"  --- Trang {page_number} (text layer: {'CÓ' if stripped else 'KHÔNG'}) ---")
            if stripped:
                any_text_layer = True
                preview = stripped if len(stripped) <= 2000 else stripped[:2000] + " ...(cắt bớt)"
                print(f"    Nội dung text layer:\n{preview}")
                _print_candidates("Text layer", stripped, valid_keys)
                text_matches.extend(c for c in extract_gcn_candidates(stripped) if c.key in (valid_keys or set()))

        # Giống hệt document_processor.process_document: quyết định OCR theo
        # TOÀN BỘ tài liệu, không theo từng trang. Chỉ bỏ qua OCR khi text
        # layer GỘP các trang đã khớp Excel; nếu không, OCR TẤT CẢ các trang,
        # kể cả trang có text layer nhưng không chứa mã (ví dụ trang bị chèn
        # chữ ký số/chứng thực không liên quan đến GCN).
        if text_matches and not args.force_ocr:
            print("  (Text layer đã khớp Excel ở trên nên app thật sẽ KHÔNG OCR trang nào.)")
            return

        for page_number, _ in page_texts:
            print(f"  --- OCR trang {page_number} ---")
            image = render_page_bgr(document, page_number - 1, args.dpi)
            matches = _ocr_image(
                image, engine, valid_keys or set(), args.include_red, page_number,
                True, args.thorough,
            )
            lines = engine.recognize(limit_image_size(image))
            _print_ocr_lines(lines, valid_keys)
            if matches:
                for match in matches:
                    tag = "CẦN KIỂM TRA" if match.needs_review else match.key
                    print(f"    => Kết quả pipeline thật (_ocr_image): {tag} (phương thức {match.method})")
            else:
                print("    => Kết quả pipeline thật (_ocr_image): không tìm thấy mã nào")
        if not any_text_layer:
            print("  (Không trang nào có text layer; toàn bộ dựa vào OCR ở trên)")


def process_image(path: Path, engine: OCREngine, args, valid_keys: set[str] | None) -> None:
    image = load_image(path)
    print(f"  Kích thước ảnh: {image.shape[1]}x{image.shape[0]}")
    matches = _ocr_image(image.copy(), engine, valid_keys or set(), args.include_red, None, True, args.thorough)
    lines = engine.recognize(limit_image_size(image))
    _print_ocr_lines(lines, valid_keys)
    if matches:
        for match in matches:
            tag = "CẦN KIỂM TRA" if match.needs_review else match.key
            print(f"  => Kết quả pipeline thật (_ocr_image): {tag} (phương thức {match.method})")
    else:
        print("  => Kết quả pipeline thật (_ocr_image): không tìm thấy mã nào")


def collect_files(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in inputs:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(p for p in path.rglob("*") if p.suffix.lower() in SUPPORTED_EXTENSIONS))
        elif path.is_file():
            files.append(path)
        else:
            print(f"Bỏ qua, không tồn tại: {path}")
    return files


def load_valid_keys(args) -> set[str] | None:
    if not args.excel:
        return None
    from gcn_file_finder.core.excel_reader import read_gcn_excel

    data = read_gcn_excel(args.excel, args.sheet, args.column)
    print(f"Đã nạp {len(data.valid_gcn_keys)} mã GCN hợp lệ từ Excel.")
    return data.valid_gcn_keys


def main() -> None:
    parser = argparse.ArgumentParser(description="Đọc thử PDF/ảnh và in ra nội dung/mã GCN nhận được.")
    parser.add_argument("paths", nargs="+", help="File PDF/ảnh hoặc thư mục cần đọc thử.")
    parser.add_argument("--dpi", type=int, default=150, help="DPI render PDF khi cần OCR (mặc định 150).")
    parser.add_argument("--thorough", action="store_true", help="Thử đủ các biến thể ảnh + 4 hướng xoay (chậm).")
    parser.add_argument("--include-red", action="store_true", help="Thêm biến thể tách màu đỏ.")
    parser.add_argument("--force-ocr", action="store_true", help="OCR cả trang PDF đã có text layer, để so sánh.")
    parser.add_argument("--gpu", action="store_true", help="Dùng GPU nếu khả dụng.")
    parser.add_argument("--model-dir", default=None, help="Đường dẫn model PaddleOCR (mặc định lấy theo config.json).")
    parser.add_argument("--excel", default=None, help="File Excel danh sách GCN, để đối chiếu CÓ/KHÔNG trong Excel.")
    parser.add_argument("--sheet", default=None, help="Tên sheet trong Excel (bắt buộc nếu dùng --excel).")
    parser.add_argument("--column", default="GCN", help="Tên cột GCN trong Excel (mặc định 'GCN').")
    args = parser.parse_args()

    if args.excel and not args.sheet:
        parser.error("--excel yêu cầu phải có --sheet")

    files = collect_files(args.paths)
    if not files:
        print("Không có file nào để đọc.")
        return

    valid_keys = load_valid_keys(args)
    model_dir = Path(args.model_dir) if args.model_dir else configured_model_dir(load_config())
    print(f"Đang nạp PaddleOCR (model: {model_dir})...")
    engine = OCREngine(use_gpu=args.gpu, model_dir=model_dir, cpu_threads=1)

    for path in files:
        print("=" * 100)
        print(f"FILE: {path}")
        try:
            if path.suffix.lower() == ".pdf":
                process_pdf(path, engine, args, valid_keys)
            else:
                process_image(path, engine, args, valid_keys)
        except Exception as exc:  # noqa: BLE001 - công cụ debug, muốn thấy mọi lỗi
            print(f"  LỖI khi đọc file này: {exc}")
    print("=" * 100)
    print(f"Đã đọc thử {len(files)} file.")


if __name__ == "__main__":
    main()
