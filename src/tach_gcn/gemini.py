"""Gọi Gemini để phân loại ảnh trang PDF theo lô."""

from __future__ import annotations

from dataclasses import dataclass
import json
import random
import re
import threading
import time
from typing import Any, Callable, Mapping, Sequence


# Tra ngày 18/10/2026 tại:
# https://ai.google.dev/gemini-api/docs/models/gemini-3.5-flash-lite
# https://ai.google.dev/gemini-api/docs/deprecations
# Đây là Flash-Lite ổn định mới nhất chưa bị deprecate tại thời điểm tra cứu.
DEFAULT_MODEL = "gemini-3.5-flash-lite"

# Bảng giá tra ngày 18/10/2026 tại https://ai.google.dev/gemini-api/docs/pricing.
# Cần cập nhật các hằng số dưới đây khi Google đổi giá.
INPUT_USD_PER_MILLION_TOKENS = 0.30
OUTPUT_USD_PER_MILLION_TOKENS = 2.50

RESPONSE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "pages": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "i": {"type": "integer"},
                    "is_certificate": {"type": "boolean"},
                    "title": {"type": "string"},
                    "reason": {"type": "string"},
                },
                "required": ["i", "is_certificate", "title", "reason"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["pages"],
    "additionalProperties": False,
}

PROMPT = """Bạn đang phân loại từng ảnh trang PDF.

GIẤY CHỨNG NHẬN / CHỨNG CHỈ / CERTIFICATE là văn bản do một tổ chức cấp để xác
nhận một người hoặc đơn vị đã đạt được, hoàn thành hoặc sở hữu điều gì đó. Trang
thường có tiêu đề lớn, bố cục trang trọng, khung viền, chữ ký hoặc con dấu.

ĐỐI VỚI giấy chứng nhận quyền sử dụng đất: tính CÓ cả bìa/quyển màu đỏ hoặc
hồng có Quốc huy và tiêu đề "GIẤY CHỨNG NHẬN", lẫn các mặt trong như trang
thông tin chủ sử dụng, sơ đồ thửa đất và trang ghi biến động khi ảnh cho thấy
chúng thực sự là các mặt của quyển GCN.

KHÔNG tính: quyết định, công văn, hợp đồng, hoá đơn, báo cáo, đơn từ, danh sách,
bìa hồ sơ chung, phụ lục, hoặc trang chỉ nhắc đến giấy chứng nhận/chứng chỉ mà
bản thân không phải giấy chứng nhận. "Bìa hồ sơ" ở đây không bao gồm bìa đỏ/hồng
của chính quyển GCN quyền sử dụng đất.

Hãy trả đúng một phần tử cho mỗi ảnh, giữ nguyên số i được ghi ngay trước ảnh.
title là tiêu đề ngắn đọc được (để trống nếu không rõ). Không đưa số thửa, số tờ
bản đồ, số vào sổ hoặc mã phôi vào title. reason là lý do ngắn gọn bằng tiếng
Việt. Chỉ xuất JSON theo schema được yêu cầu.
"""

CERTIFICATE_INFO_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "certificate_serial": {"type": "string"},
        "registration_number": {"type": "string"},
        "issue_date": {"type": "string"},
        "owner_name": {"type": "string"},
        "reason": {"type": "string"},
    },
    "required": [
        "certificate_serial",
        "registration_number",
        "issue_date",
        "owner_name",
        "reason",
    ],
    "additionalProperties": False,
}

CERTIFICATE_INFO_PROMPT = """Các ảnh sau là những trang/mặt của CÙNG MỘT giấy
chứng nhận quyền sử dụng đất. Hãy đọc toàn bộ các trang và trích:

- certificate_serial: SỐ GIẤY CHỨNG NHẬN / số phát hành / mã phôi in trên
  GCN, thường có dạng hai chữ cái và sáu chữ số, ví dụ AA 123456. Đây là số
  dùng đặt tên file. Không lấy "Số vào sổ cấp GCN", số thửa, số tờ bản đồ,
  mã vạch hoặc số in ở trang sơ đồ.
- registration_number: chỉ lấy nội dung tại nhãn "Số vào sổ cấp GCN". Tuyệt
  đối không lấy số giấy chứng nhận/số phát hành/mã phôi, số thửa, số tờ bản đồ,
  mã vạch hoặc số in ở sơ đồ. Trường này chỉ dùng trong báo cáo, không đặt tên.
- issue_date: ngày cấp GCN ban đầu, định dạng DD/MM/YYYY; không lấy ngày ký hợp
  đồng, ngày đăng ký biến động, ngày chuyển nhượng hay ngày đính chính.
- owner_name: tên người/chủ sử dụng hiện tại. Phải kiểm tra trang "Những thay
  đổi sau khi cấp giấy chứng nhận" (thường là trang 4); nếu có chuyển nhượng,
  tặng cho, thừa kế hoặc đính chính hợp lệ thì lấy chủ mới nhất thay cho chủ ban
  đầu. Giữ đủ tên đồng sở hữu/vợ chồng nếu có.
- reason: giải thích ngắn nguồn trang và việc có/không có biến động.

Không chắc hoặc không nhìn thấy trường nào thì để chuỗi rỗng, không suy đoán.
certificate_serial và registration_number là hai trường khác nhau; không được
chép cùng một giá trị vào cả hai trường.
Chỉ xuất JSON theo schema.
"""


@dataclass(frozen=True, slots=True)
class GeminiPageResult:
    """Kết luận Gemini cho một trang."""

    i: int
    is_certificate: bool
    title: str
    reason: str


@dataclass(frozen=True, slots=True)
class GeminiCertificateInfo:
    """Thông tin tổng hợp từ toàn bộ các mặt của một GCN."""

    certificate_serial: str
    registration_number: str
    issue_date: str
    owner_name: str
    reason: str


@dataclass(slots=True)
class Usage:
    """Số token dùng để tính chi phí ước tính."""

    input_tokens: int = 0
    output_tokens: int = 0

    def add(self, other: "Usage") -> None:
        """Cộng số token từ một request."""

        self.input_tokens += other.input_tokens
        self.output_tokens += other.output_tokens

    @property
    def estimated_cost_usd(self) -> float:
        """Ước tính chi phí theo hằng số đơn giá hiện tại."""

        return (
            self.input_tokens * INPUT_USD_PER_MILLION_TOKENS
            + self.output_tokens * OUTPUT_USD_PER_MILLION_TOKENS
        ) / 1_000_000


class RateLimiter:
    """Giới hạn tốc độ tuần tự, dùng chung cho mọi lời gọi Gemini."""

    def __init__(
        self,
        rpm: float,
        *,
        clock: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        if rpm <= 0:
            raise ValueError("rpm phải lớn hơn 0")
        self._interval = 60.0 / rpm
        self._clock = clock
        self._sleep = sleep
        self._next_allowed = 0.0
        self._lock = threading.Lock()

    def wait(self) -> None:
        """Chờ đến lượt request tiếp theo."""

        with self._lock:
            now = self._clock()
            delay = self._next_allowed - now
            if delay > 0:
                self._sleep(delay)
                now = self._clock()
            self._next_allowed = max(now, self._next_allowed) + self._interval


def _status_code(error: BaseException) -> int | None:
    """Lấy mã HTTP từ các kiểu ngoại lệ khác nhau của SDK."""

    for name in ("status_code", "code"):
        value = getattr(error, name, None)
        if callable(value):
            try:
                value = value()
            except TypeError:
                value = None
        if isinstance(value, int):
            return value
        enum_value = getattr(value, "value", None)
        if isinstance(enum_value, int):
            return enum_value
    match = re.search(r"\b(429|5\d\d)\b", str(error))
    return int(match.group(1)) if match else None


def is_retryable(error: BaseException) -> bool:
    """Cho biết lỗi có phải 429 hoặc 5xx hay không."""

    code = _status_code(error)
    return code == 429 or (code is not None and 500 <= code <= 599)


def call_with_retry(
    operation: Callable[[], Any],
    *,
    attempts: int = 6,
    sleep: Callable[[float], None] = time.sleep,
    random_value: Callable[[], float] = random.random,
) -> Any:
    """Gọi thao tác với exponential backoff và jitter cho lỗi tạm thời."""

    if attempts < 1:
        raise ValueError("attempts phải ít nhất là 1")
    for attempt in range(attempts):
        try:
            return operation()
        except Exception as error:
            if attempt == attempts - 1 or not is_retryable(error):
                raise
            delay = min(30.0, 2.0**attempt) + random_value()
            sleep(delay)
    raise RuntimeError("Không thể tới nhánh này")


def _extract_usage(response: Any) -> Usage:
    """Đọc usage metadata mà không phụ thuộc phiên bản nhỏ của SDK."""

    metadata = getattr(response, "usage_metadata", None)
    if metadata is None:
        return Usage()
    input_tokens = int(getattr(metadata, "prompt_token_count", 0) or 0)
    total_tokens = int(getattr(metadata, "total_token_count", 0) or 0)
    if total_tokens:
        output_tokens = max(0, total_tokens - input_tokens)
    else:
        output_tokens = int(getattr(metadata, "candidates_token_count", 0) or 0)
        output_tokens += int(getattr(metadata, "thoughts_token_count", 0) or 0)
    return Usage(input_tokens, output_tokens)


def _parse_response(value: str, expected_pages: set[int]) -> dict[int, GeminiPageResult]:
    """Kiểm tra JSON trả về có đủ và đúng kiểu cho mọi trang."""

    payload = json.loads(value)
    if not isinstance(payload, dict) or not isinstance(payload.get("pages"), list):
        raise ValueError("Gemini trả về JSON không đúng cấu trúc pages")
    parsed: dict[int, GeminiPageResult] = {}
    for item in payload["pages"]:
        if not isinstance(item, Mapping):
            raise ValueError("Phần tử pages không phải object")
        index = item.get("i")
        certificate = item.get("is_certificate")
        title = item.get("title")
        reason = item.get("reason")
        if (
            not isinstance(index, int)
            or type(certificate) is not bool
            or not isinstance(title, str)
            or not isinstance(reason, str)
        ):
            raise ValueError("Gemini trả về sai kiểu dữ liệu")
        if index in parsed:
            raise ValueError(f"Gemini trả trùng trang {index}")
        parsed[index] = GeminiPageResult(index, certificate, title, reason)
    if set(parsed) != expected_pages:
        raise ValueError("Gemini không trả đủ đúng các trang trong batch")
    return parsed


class GeminiAnalyzer:
    """Bộ phân tích ảnh trang qua SDK chính thức google-genai."""

    def __init__(
        self,
        api_key: str,
        model: str = DEFAULT_MODEL,
        *,
        low_res: bool = False,
        rpm: float = 30,
        client: Any | None = None,
        sleep: Callable[[float], None] = time.sleep,
        random_value: Callable[[], float] = random.random,
    ) -> None:
        self.model = model
        self.low_res = low_res
        self._sleep = sleep
        self._random_value = random_value
        self._limiter = RateLimiter(rpm, sleep=sleep)
        if client is None:
            from google import genai

            client = genai.Client(api_key=api_key)
        self._client = client

    def analyze(self, pages: Sequence[tuple[int, bytes]]) -> tuple[dict[int, GeminiPageResult], Usage]:
        """Phân loại một lô ảnh JPEG và trả kết quả theo số trang."""

        if not pages:
            return {}, Usage()
        from google.genai import types

        contents: list[Any] = [PROMPT]
        for page_number, image in pages:
            contents.append(f"Ảnh kế tiếp có i={page_number}.")
            contents.append(types.Part.from_bytes(data=image, mime_type="image/jpeg"))

        config_arguments: dict[str, Any] = {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_json_schema": RESPONSE_SCHEMA,
            "thinking_config": types.ThinkingConfig(
                thinking_level=types.ThinkingLevel.MINIMAL
            ),
        }
        if self.low_res:
            config_arguments["media_resolution"] = (
                types.MediaResolution.MEDIA_RESOLUTION_LOW
            )
        config = types.GenerateContentConfig(**config_arguments)

        def request() -> Any:
            self._limiter.wait()
            return self._client.models.generate_content(
                model=self.model,
                contents=contents,
                config=config,
            )

        response = call_with_retry(
            request,
            attempts=6,
            sleep=self._sleep,
            random_value=self._random_value,
        )
        response_text = getattr(response, "text", None)
        if not isinstance(response_text, str) or not response_text.strip():
            raise ValueError("Gemini không trả nội dung JSON")
        expected = {page_number for page_number, _ in pages}
        return _parse_response(response_text, expected), _extract_usage(response)

    def extract_certificate(
        self, pages: Sequence[tuple[int, bytes]]
    ) -> tuple[GeminiCertificateInfo, Usage]:
        """Trích thông tin từ các ảnh thuộc cùng một GCN."""

        if not pages:
            raise ValueError("Nhóm GCN không có trang")
        from google.genai import types

        contents: list[Any] = [CERTIFICATE_INFO_PROMPT]
        for page_number, image in pages:
            contents.append(f"Ảnh kế tiếp là trang PDF số {page_number}.")
            contents.append(types.Part.from_bytes(data=image, mime_type="image/jpeg"))
        config_arguments: dict[str, Any] = {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_json_schema": CERTIFICATE_INFO_SCHEMA,
            "thinking_config": types.ThinkingConfig(
                thinking_level=types.ThinkingLevel.MINIMAL
            ),
        }
        if self.low_res:
            config_arguments["media_resolution"] = (
                types.MediaResolution.MEDIA_RESOLUTION_LOW
            )
        config = types.GenerateContentConfig(**config_arguments)

        def request() -> Any:
            self._limiter.wait()
            return self._client.models.generate_content(
                model=self.model,
                contents=contents,
                config=config,
            )

        response = call_with_retry(
            request,
            attempts=6,
            sleep=self._sleep,
            random_value=self._random_value,
        )
        response_text = getattr(response, "text", None)
        if not isinstance(response_text, str) or not response_text.strip():
            raise ValueError("Gemini không trả thông tin GCN")
        payload = json.loads(response_text)
        if not isinstance(payload, dict):
            raise ValueError("Gemini trả thông tin GCN không đúng cấu trúc")
        values = [payload.get(name) for name in CERTIFICATE_INFO_SCHEMA["required"]]
        if not all(isinstance(value, str) for value in values):
            raise ValueError("Gemini trả sai kiểu thông tin GCN")
        info = GeminiCertificateInfo(
            certificate_serial=str(payload["certificate_serial"]).strip(),
            registration_number=str(payload["registration_number"]).strip(),
            issue_date=str(payload["issue_date"]).strip(),
            owner_name=str(payload["owner_name"]).strip(),
            reason=str(payload["reason"]).strip(),
        )
        return info, _extract_usage(response)
