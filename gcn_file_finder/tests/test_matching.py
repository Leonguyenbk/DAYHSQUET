from dataclasses import dataclass

from gcn_file_finder.core.document_processor import (
    ProcessingOptions,
    _deduplicate,
    _matches_from_ocr_lines,
    process_document,
)
from gcn_file_finder.models import GCNMatch


class NeverCalledEngine:
    def recognize(self, _image):
        raise AssertionError("Tên file đáng tin cậy thì không được OCR")


@dataclass
class Line:
    text: str
    confidence: float


def test_filename_match_must_exist_in_excel(tmp_path):
    source = tmp_path / "AM 143443 CS 999999.pdf"
    source.write_bytes(b"not a real pdf; filename path short-circuits")
    result = process_document(source, {"AM143443"}, NeverCalledEngine(), ProcessingOptions())
    assert [match.key for match in result.matches] == ["AM143443"]


def test_filename_unknown_to_excel_is_not_confirmed(tmp_path):
    source = tmp_path / "CS 999999.pdf"
    source.write_bytes(b"not a real pdf")
    result = process_document(source, {"AM143443"}, NeverCalledEngine(), ProcessingOptions())
    assert result.matches == []
    assert result.status == "KHÔNG TÌM THẤY"


def test_filename_with_gcn_is_opened_when_content_verification_is_enabled(tmp_path):
    source = tmp_path / "CS 999999.pdf"
    source.write_bytes(b"not a real pdf")
    options = ProcessingOptions(verify_content_after_filename=True)

    result = process_document(source, {"AM143443"}, NeverCalledEngine(), options)

    assert result.status == "LỖI XỬ LÝ"


def test_same_gcn_on_many_pages_is_one_copy_key():
    matches = _deduplicate([
        GCNMatch("AM143443", "AM 143443", "PDF TEXT", page=1),
        GCNMatch("AM143443", "AM 143443", "PDF TEXT", page=1),
        GCNMatch("AM143443", "AM 143443", "PDF TEXT", page=3),
    ])
    assert len(matches) == 2
    assert list(dict.fromkeys(match.key for match in matches if match.key)) == ["AM143443"]


def test_ocr_code_split_across_lines_is_found():
    matches = _matches_from_ocr_lines(
        [Line("Số phát hành AM", 0.92), Line("143443", 0.88)], {"AM143443"}, "OCR ẢNH GỐC", 2
    )
    assert [(match.key, match.page) for match in matches] == [("AM143443", 2)]
