from gcn_file_finder.core.gcn_normalizer import (
    controlled_ocr_match,
    extract_gcn_candidates,
    extract_controlled_ocr_matches,
    match_exact_candidates,
)


def keys(text):
    return [item.key for item in extract_gcn_candidates(text)]


def test_extract_supported_separators_and_spaced_letters():
    assert keys("Số AM:143443, CĐ.012345 và A M 143443") == ["AM143443", "CĐ012345", "AM143443"]


def test_does_not_match_partial_long_or_three_letters():
    assert keys("ABC123456 AM1234567 012345678901") == []


def test_does_not_extract_bc_from_abc():
    assert "BC123456" not in keys("ABC123456")


def test_extract_finds_code_after_so_label_glued_by_ocr():
    # "Số" mất dấu do OCR, dính liền "S" vào mã, không còn khoảng trắng: "SAA346849".
    assert keys("SAA346849") == ["AA346849"]
    assert keys("Khi bi mat SAA346849 hu hong") == ["AA346849"]


def test_extract_does_not_duplicate_when_so_label_has_space():
    # Có khoảng trắng ("Số AM143443") thì pattern chuẩn đã khớp đúng rồi,
    # không được khớp thêm lần nữa qua pattern "dính liền".
    assert keys("So AM143443") == ["AM143443"]
    assert keys("So D 1234567") == ["D1234567"]
    assert keys("So AA 12345678") == ["AA12345678"]


def test_extract_d_shape_and_aa_shape():
    assert keys("So D 1234567 va AA 12345678 trong ho so") == ["D1234567", "AA12345678"]


def test_d_shape_does_not_match_other_letters():
    assert keys("A 1234567 B 1234567") == []


def test_aa_shape_does_not_match_other_letter_pairs():
    assert keys("BC 12345678 AB 12345678") == []


def test_only_excel_keys_match():
    matches = match_exact_candidates("AM 143443 CS 012345", {"AM143443"})
    assert [item.key for item in matches] == ["AM143443"]


def test_controlled_digit_correction():
    result = controlled_ocr_match("AM I43443", {"AM143443"})
    assert result.key == "AM143443"
    assert not result.needs_review


def test_d_is_not_changed_without_excel_match():
    result = controlled_ocr_match("DA123456", {"CĐ123456"})
    assert result.key is None
    assert result.alternatives == ()


def test_d_to_d_stroke_only_through_excel_comparison():
    result = controlled_ocr_match("DA123456", {"ĐA123456"})
    assert result.key == "ĐA123456"


def test_multiple_controlled_matches_require_review():
    result = controlled_ocr_match("DD001234", {"DĐ001234", "ĐD001234"})
    assert result.key is None
    assert result.needs_review
    assert result.alternatives == ("DĐ001234", "ĐD001234")
    extracted = extract_controlled_ocr_matches("OCR: DD 001234", {"DĐ001234", "ĐD001234"})
    assert extracted[0].needs_review


def test_existing_d_code_is_not_turned_into_d_stroke():
    result = controlled_ocr_match("DA123456", {"DA123456", "ĐA123456"})
    assert result.key == "DA123456"
    assert not result.needs_review


def test_controlled_digit_correction_d_shape():
    result = controlled_ocr_match("D I234567", {"D1234567"})
    assert result.key == "D1234567"


def test_controlled_digit_correction_aa_shape():
    result = controlled_ocr_match("AA I2345678", {"AA12345678"})
    assert result.key == "AA12345678"


def test_d_shape_not_confused_with_standard_when_both_valid():
    # "DA123456" hợp lệ dạng chuẩn (2 chữ+6 số); không được tự nhận thành
    # dạng D+7 số vì chữ thứ hai không phải số.
    result = controlled_ocr_match("DA123456", {"DA123456"})
    assert result.key == "DA123456"
