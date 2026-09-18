"""Chuẩn hóa, trích xuất và sửa lỗi OCR số phát hành GCN."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from itertools import product
from typing import Iterable

GCN_KEY_RE = re.compile(r"^(?:[A-ZĐ]{2}\d{6}|D\d{7}|AA\d{8})$")
GCN_JOINED_RE = re.compile(r"(?<![A-ZĐ0-9])[A-ZĐ]{2}\d{6}(?!\d)")
GCN_RAW_RE = re.compile(
    r"(?<![A-ZĐ0-9])([A-ZĐ])\s*([A-ZĐ])[\s.\-:]*(\d(?:[\s.\-]*\d){5})(?!\d)"
)
# Hai biến thể riêng theo yêu cầu nghiệp vụ: chỉ đúng chữ D có 7 số, và chỉ
# đúng cặp AA có 8 số. Không áp dụng cho chữ/cặp chữ khác.
GCN_RAW_RE_D7 = re.compile(r"(?<![A-ZĐ0-9])(D)[\s.\-:]*(\d(?:[\s.\-]*\d){6})(?!\d)")
GCN_RAW_RE_AA8 = re.compile(r"(?<![A-ZĐ0-9])(A)\s*(A)[\s.\-:]*(\d(?:[\s.\-]*\d){7})(?!\d)")

# OCR thường đọc rụng dấu "Số" (nhãn đứng trước mã, vd "Số phát hành: AA
# 346849") thành chữ "S" dính LIỀN ngay vào mã, không còn khoảng trắng hay
# dấu phân cách nào (vd "SoAA 346849" -> "SAA346849"). Quy tắc chống nhầm
# chữ dài (không lấy "BC" từ "ABC123456") vô tình chặn luôn trường hợp này vì
# "S" đứng sát ngay trước. Ba pattern dưới đây CHỈ áp dụng khi dính liền
# hoàn toàn (không cho phép dấu cách/phân cách sau "S") — nếu có khoảng
# trắng như "Số AM143443" thì 3 pattern chuẩn ở trên đã khớp đúng rồi, thêm
# pattern này vào sẽ trùng lặp.
_LABEL_GLUED = r"S[ỐOÓ0]?"
GCN_RAW_RE_LABELED = re.compile(
    rf"(?<![A-ZĐ0-9]){_LABEL_GLUED}([A-ZĐ])\s*([A-ZĐ])[\s.\-:]*(\d(?:[\s.\-]*\d){{5}})(?!\d)"
)
GCN_RAW_RE_D7_LABELED = re.compile(rf"(?<![A-ZĐ0-9]){_LABEL_GLUED}(D)[\s.\-:]*(\d(?:[\s.\-]*\d){{6}})(?!\d)")
GCN_RAW_RE_AA8_LABELED = re.compile(
    rf"(?<![A-ZĐ0-9]){_LABEL_GLUED}(A)\s*(A)[\s.\-:]*(\d(?:[\s.\-]*\d){{7}})(?!\d)"
)


@dataclass(frozen=True)
class GCNCandidate:
    """Ứng viên GCN trích xuất từ một chuỗi."""

    raw: str
    key: str
    display: str
    start: int
    end: int


@dataclass(frozen=True)
class ControlledMatch:
    """Kết quả sửa OCR có kiểm soát."""

    key: str | None
    alternatives: tuple[str, ...]
    raw: str

    @property
    def needs_review(self) -> bool:
        return len(self.alternatives) > 1


def _clean_unicode(value: object) -> str:
    return unicodedata.normalize("NFC", str(value)).upper().replace("Ð", "Đ")


def normalize_gcn_key(value: object) -> str | None:
    """Trả về khóa 8 ký tự hoặc ``None`` nếu giá trị không đúng quy tắc."""

    if value is None:
        return None
    text = _clean_unicode(value).strip()
    key = re.sub(r"[\s.\-:]", "", text)
    return key if GCN_KEY_RE.fullmatch(key) else None


def _letter_count(key: str) -> int:
    """Số ký tự chữ đứng đầu khóa, tùy theo dạng khóa đã khớp."""

    if re.fullmatch(r"D\d{7}", key):
        return 1
    return 2


def format_gcn_display(gcn_key: str) -> str:
    """Định dạng khóa hợp lệ thành ``AA 123456``, ``D 1234567`` hoặc ``AA 12345678``."""

    key = normalize_gcn_key(gcn_key)
    if key is None:
        raise ValueError(f"Khóa GCN không hợp lệ: {gcn_key!r}")
    letters = _letter_count(key)
    return f"{key[:letters]} {key[letters:]}"


def extract_gcn_candidates(text: object) -> list[GCNCandidate]:
    """Trích mọi mã hợp lệ (hai chữ+sáu số, D+bảy số, AA+tám số), không khớp chuỗi con sai biên."""

    if text is None:
        return []
    normalized = _clean_unicode(text)
    found: list[GCNCandidate] = []
    seen: set[tuple[str, int, int]] = set()

    def _add(match: re.Match[str], key: str) -> None:
        marker = (key, match.start(), match.end())
        if marker not in seen and GCN_KEY_RE.fullmatch(key):
            found.append(GCNCandidate(match.group(0), key, format_gcn_display(key), match.start(), match.end()))
            seen.add(marker)

    for match in GCN_RAW_RE.finditer(normalized):
        _add(match, match.group(1) + match.group(2) + re.sub(r"[\s.\-]", "", match.group(3)))
    for match in GCN_RAW_RE_D7.finditer(normalized):
        _add(match, match.group(1) + re.sub(r"[\s.\-]", "", match.group(2)))
    for match in GCN_RAW_RE_AA8.finditer(normalized):
        _add(match, match.group(1) + match.group(2) + re.sub(r"[\s.\-]", "", match.group(3)))
    # "Số" đứng ngay trước mã, OCR rụng dấu dính liền thành "S" (vd
    # "SAA346849"). Chỉ khớp thêm khi 3 pattern chuẩn ở trên bị chặn bởi tiền
    # tố "S" này, không nới lỏng cho chữ khác.
    for match in GCN_RAW_RE_LABELED.finditer(normalized):
        _add(match, match.group(1) + match.group(2) + re.sub(r"[\s.\-]", "", match.group(3)))
    for match in GCN_RAW_RE_D7_LABELED.finditer(normalized):
        _add(match, match.group(1) + re.sub(r"[\s.\-]", "", match.group(2)))
    for match in GCN_RAW_RE_AA8_LABELED.finditer(normalized):
        _add(match, match.group(1) + match.group(2) + re.sub(r"[\s.\-]", "", match.group(3)))
    found.sort(key=lambda candidate: candidate.start)
    return found


_TO_DIGIT: dict[str, tuple[str, ...]] = {
    "O": ("0",), "Q": ("0",), "I": ("1",), "L": ("1",),
    "Z": ("2",), "S": ("5",), "B": ("8",), "G": ("6",),
}
_TO_LETTER: dict[str, tuple[str, ...]] = {
    "0": ("O",), "1": ("I",), "5": ("S",), "8": ("B",),
}


def _position_options(char: str, letter_position: bool) -> tuple[str, ...]:
    if letter_position:
        options: list[str] = [char] if re.fullmatch(r"[A-ZĐ]", char) else []
        options.extend(_TO_LETTER.get(char, ()))
        # D chỉ được thử thành Đ trong bước đối chiếu này, không sửa toàn cục.
        if char == "D":
            options.append("Đ")
    else:
        options = [char] if char.isascii() and char.isdigit() else []
        options.extend(_TO_DIGIT.get(char, ()))
    return tuple(dict.fromkeys(options))


def _standard_shape_options(cleaned: str) -> list[tuple[str, ...]]:
    """8 ký tự: hai chữ bất kỳ + sáu số (dạng phổ biến)."""

    return [_position_options(ch, index < 2) for index, ch in enumerate(cleaned)]


def _d_shape_options(cleaned: str) -> list[tuple[str, ...]] | None:
    """8 ký tự: đúng chữ D + bảy số. Không áp dụng cho chữ khác."""

    if cleaned[0] != "D":
        return None
    return [("D",)] + [_position_options(ch, False) for ch in cleaned[1:]]


def _aa_shape_options(cleaned: str) -> list[tuple[str, ...]] | None:
    """10 ký tự: đúng cặp AA + tám số. Không áp dụng cho cặp chữ khác."""

    if cleaned[0] != "A" or cleaned[1] != "A":
        return None
    return [("A",), ("A",)] + [_position_options(ch, False) for ch in cleaned[2:]]


def controlled_ocr_match(raw_text: str, valid_gcn_keys: set[str]) -> ControlledMatch:
    """Sinh sửa lỗi theo vị trí và chỉ trả về mã nằm chính xác trong Excel.

    Một kết quả duy nhất được xác nhận. Từ hai kết quả trở lên được giữ ở trạng
    thái cần kiểm tra, tuyệt đối không tự chọn kết quả gần nhất. Thử cả ba
    dạng khóa (hai chữ+sáu số, D+bảy số, AA+tám số) theo độ dài chuỗi sạch.
    """

    cleaned = re.sub(r"[\s.\-:]", "", _clean_unicode(raw_text))
    # Không thử D -> Đ (hay sửa khác) khi chuỗi OCR ban đầu đã khớp Excel.
    if GCN_KEY_RE.fullmatch(cleaned) and cleaned in valid_gcn_keys:
        return ControlledMatch(cleaned, (cleaned,), raw_text)

    candidate_option_sets: list[list[tuple[str, ...]]] = []
    if len(cleaned) == 8:
        candidate_option_sets.append(_standard_shape_options(cleaned))
        d_options = _d_shape_options(cleaned)
        if d_options is not None:
            candidate_option_sets.append(d_options)
    elif len(cleaned) == 10:
        aa_options = _aa_shape_options(cleaned)
        if aa_options is not None:
            candidate_option_sets.append(aa_options)
    else:
        return ControlledMatch(None, (), raw_text)

    matches: set[str] = set()
    for option_sets in candidate_option_sets:
        if any(not options for options in option_sets):
            continue
        matches.update({"".join(chars) for chars in product(*option_sets)} & valid_gcn_keys)
    matches_sorted = sorted(matches)
    if len(matches_sorted) == 1:
        return ControlledMatch(matches_sorted[0], tuple(matches_sorted), raw_text)
    return ControlledMatch(None, tuple(matches_sorted), raw_text)


def extract_controlled_ocr_matches(text: str, valid_gcn_keys: set[str]) -> list[ControlledMatch]:
    """Tìm các cửa sổ OCR có 8 hoặc 10 ký tự và áp dụng sửa lỗi có kiểm soát."""

    normalized = _clean_unicode(text)
    # Giữ các nhóm có thể là mã; không cho phép lấy đuôi từ một nhóm dài hơn.
    token_patterns = (
        r"(?<![A-ZĐ0-9])[A-ZĐ0-9](?:[\s.\-:]*[A-ZĐ0-9]){7}(?![A-ZĐ0-9])",
        r"(?<![A-ZĐ0-9])[A-ZĐ0-9](?:[\s.\-:]*[A-ZĐ0-9]){9}(?![A-ZĐ0-9])",
    )
    tokens = list(re.finditer(token_patterns[0], normalized)) + list(re.finditer(token_patterns[1], normalized))
    results: list[ControlledMatch] = []
    seen: set[tuple[str | None, tuple[str, ...], str]] = set()
    for token in tokens:
        result = controlled_ocr_match(token.group(0), valid_gcn_keys)
        marker = (result.key, result.alternatives, result.raw)
        if (result.key or result.needs_review) and marker not in seen:
            results.append(result)
            seen.add(marker)
    return results


def match_exact_candidates(text: str, valid_gcn_keys: set[str]) -> list[GCNCandidate]:
    """Lọc ứng viên theo tập khóa Excel."""

    return [item for item in extract_gcn_candidates(text) if item.key in valid_gcn_keys]


def unique_keys(candidates: Iterable[GCNCandidate]) -> list[str]:
    """Trả về khóa duy nhất, giữ thứ tự xuất hiện."""

    return list(dict.fromkeys(candidate.key for candidate in candidates))
