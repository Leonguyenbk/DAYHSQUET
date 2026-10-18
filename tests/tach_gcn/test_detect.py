"""Kiểm thử chuẩn hoá và chấm điểm."""

from tach_gcn.detect import (
    detect_page,
    extract_certificate_serial,
    extract_registration_number,
    normalize_text,
    score_text,
)


def test_normalize_vietnamese() -> None:
    """Dấu tiếng Việt và dấu câu được loại bỏ ổn định."""

    assert normalize_text("  GIẤY—chứng nhận: Đỗ!  ") == "GIAY CHUNG NHAN DO"


def test_certificate_scores() -> None:
    """Hai tiêu đề chứng nhận vượt ngưỡng."""

    vietnamese = detect_page("GIẤY CHỨNG NHẬN\nCấp cho Nguyễn Văn A", 1)
    english = detect_page("CERTIFICATE OF COMPLETION", 2)
    assert vietnamese.is_certificate
    assert english.is_certificate
    assert vietnamese.score >= 5
    assert english.score >= 5


def test_decision_is_ambiguous_not_accepted() -> None:
    """Quyết định chỉ nhắc chứng nhận phải chờ Gemini."""

    result = detect_page("CHỨNG NHẬN là căn cứ trong QUYẾT ĐỊNH này", 1)
    assert 0 < result.score < 5
    assert not result.is_certificate
    assert result.needs_gemini


def test_custom_keyword_adds_six_points() -> None:
    """Từ khoá người dùng được chuẩn hoá và cộng đúng sáu điểm."""

    score, found, _ = score_text("Xác nhận hoàn tất khóa học", ["xác nhận hoàn tất"])
    assert score == 7  # +6 từ khoá và +1 vì trang ngắn.
    assert "XAC NHAN HOAN TAT" in found


def test_extract_two_distinct_certificate_numbers() -> None:
    """Mã phôi và số vào sổ là hai trường độc lập, không lấy mã trên sơ đồ."""

    text = (
        "GIẤY CHỨNG NHẬN\nSố giấy chứng nhận: BB 654321\n"
        "Số vào sổ cấp GCN: CS 012345"
    )
    assert extract_certificate_serial(text) == "BB 654321"
    assert extract_registration_number(text) == "CS 012345"
    assert extract_certificate_serial("GIẤY CHỨNG NHẬN\nSố vào sổ cấp GCN: AA 123456") == ""
    assert extract_certificate_serial("Sơ đồ thửa đất\nAB987654") == ""
    assert extract_registration_number("Sơ đồ thửa đất\nAB987654") == ""
