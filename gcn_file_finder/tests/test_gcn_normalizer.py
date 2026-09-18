import pytest

from gcn_file_finder.core.gcn_normalizer import format_gcn_display, normalize_gcn_key


@pytest.mark.parametrize(
    ("value", "key", "display"),
    [
        ("AM143443", "AM143443", "AM 143443"),
        ("AM 143443", "AM143443", "AM 143443"),
        ("AM-143443", "AM143443", "AM 143443"),
        ("AM.143443", "AM143443", "AM 143443"),
        ("am 143443", "AM143443", "AM 143443"),
        ("AM 012345", "AM012345", "AM 012345"),
        ("CĐ 123456", "CĐ123456", "CĐ 123456"),
        ("ĐA 123456", "ĐA123456", "ĐA 123456"),
        ("ĐĐ 001234", "ĐĐ001234", "ĐĐ 001234"),
        ("A M 143443", "AM143443", "AM 143443"),
        ("CÐ.012345", "CĐ012345", "CĐ 012345"),
        ("D 1234567", "D1234567", "D 1234567"),
        ("D-1234567", "D1234567", "D 1234567"),
        ("AA 12345678", "AA12345678", "AA 12345678"),
        ("AA-12345678", "AA12345678", "AA 12345678"),
    ],
)
def test_valid_normalization(value, key, display):
    assert normalize_gcn_key(value) == key
    assert format_gcn_display(key) == display


@pytest.mark.parametrize(
    "value",
    [
        "Đ 123456", "A 123456", "ABC 123456", "ABC123456", "AM 12345", "AM 1234567", "123456",
        # Chỉ riêng D mới được 7 số; chữ khác vẫn phải đúng 6 số.
        "A 1234567", "B 1234567",
        # Chỉ riêng cặp AA mới được 8 số; cặp khác vẫn phải đúng 6 số.
        "BC 12345678", "AB 12345678",
    ],
)
def test_invalid_values(value):
    assert normalize_gcn_key(value) is None


def test_leading_zero_is_preserved():
    assert normalize_gcn_key("ĐĐ 001234") == "ĐĐ001234"
