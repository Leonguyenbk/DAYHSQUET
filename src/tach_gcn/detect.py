"""Chuẩn hoá văn bản và chấm điểm trang giấy chứng nhận."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import re
import unicodedata
from typing import Any, Iterable


CERTIFICATE_THRESHOLD = 5
SCAN_TEXT_LIMIT = 20

_REGISTRATION_NUMBER_LABELS = (
    re.compile(
        r"\bSO\s+VAO\s+SO\s+(?:CAP\s+)?GCN(?:\s+QSDD)?\s+"
        r"([A-Z]{1,4})\s*(\d{4,10})\b"
    ),
    re.compile(
        r"\bSO\s+VAO\s+SO\s+(?:CAP\s+)?GIAY\s+CHUNG\s+NHAN\s+"
        r"([A-Z]{1,4})\s*(\d{4,10})\b"
    ),
)

_CERTIFICATE_SERIAL_LABELS = (
    re.compile(
        r"\bSO\s+(?:PHAT\s+HANH\s+)?(?:GIAY\s+CHUNG\s+NHAN|GCN)\s+"
        r"([A-Z]{2})\s*(\d{8}|\d{6})\b"
    ),
    re.compile(
        r"\b(?:SO\s+SERI|SERIAL|MA\s+PHOI)\s+([A-Z]{2})\s*(\d{8}|\d{6})\b"
    ),
)
_CERTIFICATE_SERIAL = re.compile(r"\b([A-Z]{2})\s*(\d{8}|\d{6})\b")


@dataclass(slots=True)
class DetectionResult:
    """Kết quả nhận diện của một trang PDF."""

    page: int
    is_certificate: bool
    method: str
    score: int
    keywords: list[str] = field(default_factory=list)
    note: str = ""
    normalized_length: int = 0
    needs_gemini: bool = False
    title: str = ""
    reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        """Chuyển kết quả thành dữ liệu tuần tự hoá được."""

        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "DetectionResult":
        """Khôi phục kết quả từ cache JSON."""

        return cls(**value)


def normalize_text(text: str) -> str:
    """Bỏ dấu tiếng Việt, dấu câu và chuẩn hoá khoảng trắng."""

    text = text.replace("đ", "d").replace("Đ", "D")
    decomposed = unicodedata.normalize("NFD", text)
    without_marks = "".join(
        character
        for character in decomposed
        if unicodedata.category(character) != "Mn"
    )
    cleaned = "".join(
        character if (character.isalnum() or character.isspace()) else " "
        for character in without_marks.upper()
    )
    return re.sub(r"\s+", " ", cleaned).strip()


def extract_registration_number(text: str) -> str:
    """Trích *số vào sổ cấp GCN* khi có đúng nhãn của trường này."""

    normalized = normalize_text(text)
    for pattern in _REGISTRATION_NUMBER_LABELS:
        match = pattern.search(normalized)
        if match:
            return f"{match.group(1)} {match.group(2)}"

    return ""


def extract_certificate_serial(text: str) -> str:
    """Trích số phát hành/mã phôi GCN, tuyệt đối không dùng số vào sổ thay thế."""

    normalized = normalize_text(text)
    registration_number = extract_registration_number(text)
    for pattern in _CERTIFICATE_SERIAL_LABELS:
        match = pattern.search(normalized)
        if match:
            candidate = f"{match.group(1)} {match.group(2)}"
            if candidate != registration_number:
                return candidate

    # Mã phôi thường được in riêng dạng hai chữ cái và sáu chữ số (mẫu A4 mới: AA + tám chữ số). Chỉ cho
    # phép đoán khi chính trang có ngữ cảnh GCN; trang sơ đồ đứng riêng không đủ.
    if "GIAY CHUNG NHAN" not in normalized and "CERTIFICATE" not in normalized:
        return ""
    for match in _CERTIFICATE_SERIAL.finditer(normalized):
        candidate = f"{match.group(1)} {match.group(2)}"
        if candidate != registration_number:
            return candidate
    return ""


def extract_certificate_number(text: str) -> str:
    """Tên tương thích cho số giấy chứng nhận/số phát hành (không phải số vào sổ)."""

    return extract_certificate_serial(text)


def extract_issue_date(text: str) -> str:
    """Lấy ngày cấp khi văn bản có nhãn rõ ràng, tránh ngày chuyển nhượng."""

    patterns = (
        re.compile(
            r"NG[AÀA]Y\s+C[AẤA]P(?:\s+GCN)?\s*[:\-]?\s*"
            r"(\d{1,2})[\s/.-]+(\d{1,2})[\s/.-]+(\d{4})",
            re.IGNORECASE,
        ),
        re.compile(
            r"NG[AÀA]Y\s+C[AẤA]P(?:\s+GCN)?\s*[:\-]?\s*NG[AÀA]Y\s+"
            r"(\d{1,2})\s+TH[AÁA]NG\s+(\d{1,2})\s+N[AĂA]M\s+(\d{4})",
            re.IGNORECASE,
        ),
    )
    for pattern in patterns:
        match = pattern.search(text)
        if match:
            day, month, year = (int(value) for value in match.groups())
            if 1 <= day <= 31 and 1 <= month <= 12:
                return f"{day:02d}/{month:02d}/{year:04d}"

    normalized = normalize_text(text)
    match = re.search(
        r"\bNGAY CAP(?: GCN)?(?: NGAY)? (\d{1,2}) THANG (\d{1,2}) NAM (\d{4})\b",
        normalized,
    )
    if match:
        day, month, year = (int(value) for value in match.groups())
        if 1 <= day <= 31 and 1 <= month <= 12:
            return f"{day:02d}/{month:02d}/{year:04d}"
    return ""


def extract_owner_name(text: str) -> str:
    """Lấy tên chủ; ưu tiên chuyển nhượng/đính chính xuất hiện sau cùng."""

    def clean(value: str) -> str:
        return " ".join(value.strip(" .,:-").split())

    transfer_patterns = (
        re.compile(
            r"(?:CHUYỂN\s+NHƯỢNG|TẶNG\s+CHO|THỪA\s+KẾ)\s+"
            r"(?:CHO\s*[:\-]?\s*)?"
            r"(?:ÔNG|BÀ|ÔNG\s+VÀ\s+BÀ)?\s*[:\-]?\s*([^\n;]{3,100})",
            re.IGNORECASE,
        ),
        re.compile(
            r"ĐÍNH\s+CHÍNH[^\n]{0,80}?(?:TÊN|CHỦ\s+SỬ\s+DỤNG)[^\n]{0,40}?"
            r"(?:THÀNH|LÀ)\s*[:\-]?\s*([^\n;]{3,100})",
            re.IGNORECASE,
        ),
    )
    transferred: list[str] = []
    for pattern in transfer_patterns:
        transferred.extend(clean(match.group(1)) for match in pattern.finditer(text))
    if transferred:
        return transferred[-1]

    initial_patterns = (
        re.compile(
            r"(?:TÊN\s+CHỦ\s+SỬ\s+DỤNG|CHỦ\s+SỬ\s+DỤNG|NGƯỜI\s+SỬ\s+DỤNG\s+ĐẤT)"
            r"\s*[:\-]\s*([^\n;]{3,100})",
            re.IGNORECASE,
        ),
        re.compile(r"^(?:ÔNG|BÀ)\s*[:\-]\s*([^\n;]{3,100})", re.IGNORECASE | re.MULTILINE),
    )
    for pattern in initial_patterns:
        match = pattern.search(text)
        if match:
            return clean(match.group(1))
    return ""


def is_certificate_support_page(text: str) -> bool:
    """Nhận ra trang 2–4 thuộc GCN dù trang đó không lặp tiêu đề GCN."""

    normalized = normalize_text(text)
    markers = (
        "SO DO THUA DAT",
        "THUA DAT NHA O VA TAI SAN KHAC GAN LIEN VOI DAT",
        "NHUNG THAY DOI SAU KHI CAP GIAY CHUNG NHAN",
        "NHUNG THAY DOI SAU KHI CAP GCN",
        "NOI DUNG THAY DOI VA CO SO PHAP LY",
    )
    return any(marker in normalized for marker in markers)


def _add_if_present(
    text: str,
    phrase: str,
    weight: int,
    found: list[str],
) -> int:
    """Cộng trọng số một lần nếu cụm từ xuất hiện."""

    if phrase in text:
        found.append(phrase)
        return weight
    return 0


def score_text(text: str, custom_keywords: Iterable[str] = ()) -> tuple[int, list[str], str]:
    """Chấm điểm văn bản theo bộ quy tắc đã định nghĩa."""

    normalized = normalize_text(text)
    first_part = normalized[: max(1, int(len(normalized) * 0.35))]
    found: list[str] = []
    score = 0

    score += _add_if_present(normalized, "GIAY CHUNG NHAN", 6, found)
    score += _add_if_present(normalized, "CERTIFICATE", 6, found)
    score += _add_if_present(normalized, "CHUNG CHI", 4, found)
    score += _add_if_present(normalized, "THIS IS TO CERTIFY", 3, found)
    score += _add_if_present(normalized, "CERTIFIES THAT", 3, found)

    if "GIAY CHUNG NHAN" not in normalized and "CHUNG NHAN" in normalized:
        found.append("CHUNG NHAN")
        score += 2
        if "CHUNG NHAN" in first_part:
            score += 1

    for phrase in ("AWARDED TO", "PRESENTED TO", "DA HOAN THANH", "CAP CHO"):
        score += _add_if_present(normalized, phrase, 1, found)

    for raw_keyword in custom_keywords:
        keyword = normalize_text(raw_keyword)
        if keyword and keyword in normalized:
            found.append(keyword)
            score += 6

    if score > 0 and len(normalized) < 1800:
        score += 1

    # Không lặp cụm từ trong cột báo cáo nếu từ khoá người dùng trùng từ khoá có sẵn.
    found = list(dict.fromkeys(found))
    return score, found, normalized


def detect_page(
    text: str,
    page: int,
    custom_keywords: Iterable[str] = (),
    gemini_all: bool = False,
    gemini_enabled: bool = True,
) -> DetectionResult:
    """Nhận diện sơ bộ một trang và quyết định có cần gửi Gemini không."""

    score, keywords, normalized = score_text(text, custom_keywords)
    is_scan = len(normalized) < SCAN_TEXT_LIMIT
    is_ambiguous = 0 < score < CERTIFICATE_THRESHOLD
    needs_gemini = gemini_enabled and (is_scan or is_ambiguous or gemini_all)
    method = "ảnh scan" if is_scan else "từ khoá"
    note = "chưa phân tích" if is_scan and not gemini_enabled else ""
    return DetectionResult(
        page=page,
        is_certificate=score >= CERTIFICATE_THRESHOLD,
        method=method,
        score=score,
        keywords=keywords,
        note=note,
        normalized_length=len(normalized),
        needs_gemini=needs_gemini,
    )


def apply_gemini_result(
    result: DetectionResult,
    is_certificate: bool,
    title: str,
    reason: str,
) -> DetectionResult:
    """Ghi đè kết luận từ khoá bằng kết luận có cấu trúc của Gemini."""

    result.is_certificate = is_certificate
    result.method = "Gemini"
    result.title = title.strip()
    result.reason = reason.strip()
    result.note = " — ".join(part for part in (result.title, result.reason) if part)
    result.needs_gemini = False
    return result
