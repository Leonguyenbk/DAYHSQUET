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
Bản scan có thể TRẢI 2 MẶT trên một ảnh khổ ngang (A3): mặt ngoài gồm trang 4
(những thay đổi) + bìa/trang 1, mặt trong gồm trang 2 + trang 3 (thông tin thửa, sơ
đồ). Ảnh như vậy vẫn là MỘT ảnh của quyển GCN; đọc cả hai nửa.

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
        "owner_id": {"type": "string"},
        "owner_address": {"type": "string"},
        "reason": {"type": "string"},
    },
    "required": [
        "certificate_serial",
        "registration_number",
        "issue_date",
        "owner_name",
        "owner_id",
        "owner_address",
        "reason",
    ],
    "additionalProperties": False,
}

CERTIFICATE_INFO_PROMPT = """Các ảnh sau là những trang/mặt của CÙNG MỘT giấy
chứng nhận quyền sử dụng đất. Bản scan có thể TRẢI 2 MẶT trên một ảnh khổ ngang (A3): mặt ngoài gồm trang 4
(những thay đổi) + bìa/trang 1, mặt trong gồm trang 2 + trang 3 (thông tin thửa, sơ
đồ). Ảnh như vậy vẫn là MỘT ảnh của quyển GCN; đọc cả hai nửa.
Hãy đọc toàn bộ các trang và trích:

- certificate_serial: SỐ GIẤY CHỨNG NHẬN / số phát hành / mã phôi in trên
  GCN, thường có dạng hai chữ cái và sáu chữ số, ví dụ AA 123456 (GCN mẫu A4 mới:
  AA và tám chữ số, ví dụ AA 01234567). Đây là số
  dùng đặt tên file. Không lấy "Số vào sổ cấp GCN", số thửa, số tờ bản đồ,
  mã vạch hoặc số in ở trang sơ đồ.
- registration_number: chỉ lấy nội dung tại nhãn "Số vào sổ cấp GCN" (mẫu cũ:
  "Vào sổ cấp giấy chứng nhận ... số"), in cùng trang với thông tin thửa đất (số
  thửa, tờ bản đồ, diện tích...), thường gần chữ ký/ngày cấp. Tuyệt đối không lấy
  số giấy chứng nhận/số phát hành/mã phôi, số thửa, số tờ bản đồ, mã vạch hoặc số
  in ở sơ đồ. Trường này chỉ dùng trong báo cáo, không đặt tên.
- issue_date: ngày cấp GCN ban đầu, định dạng DD/MM/YYYY; không lấy ngày ký hợp
  đồng, ngày đăng ký biến động, ngày chuyển nhượng hay ngày đính chính.
- owner_name: tên người/chủ sử dụng hiện tại. Phải kiểm tra trang "Những thay
  đổi sau khi cấp giấy chứng nhận" (thường là trang 4); nếu có chuyển nhượng,
  tặng cho, thừa kế hoặc đính chính hợp lệ thì lấy chủ mới nhất thay cho chủ ban
  đầu. Giữ đủ tên đồng sở hữu/vợ chồng nếu có.
- owner_id: số CCCD / CMND / số định danh / hộ chiếu của (các) chủ hiện tại ghi
  trên GCN, nhiều người thì cách nhau "; " theo đúng thứ tự owner_name. Chủ hiện
  tại là chủ mới ở trang biến động mà trang đó không ghi số giấy tờ thì để rỗng,
  không lấy số của chủ cũ.
- owner_address: địa chỉ thường trú của chủ hiện tại ghi trên GCN (cùng nguyên
  tắc như owner_id; không lấy địa chỉ thửa đất).
- reason: giải thích ngắn nguồn trang và việc có/không có biến động.

Không chắc hoặc không nhìn thấy trường nào thì để chuỗi rỗng, không suy đoán.
certificate_serial và registration_number là hai trường khác nhau; không được
chép cùng một giá trị vào cả hai trường.
Chỉ xuất JSON theo schema.
"""


ASSIGN_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "pages": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "i": {"type": "integer"},
                    "gcn": {"type": "string"},
                    "serial_seen": {"type": "string"},
                    "starts_copy": {"type": "boolean"},
                    "attached_to": {"type": "string"},
                    "reason": {"type": "string"},
                },
                "required": ["i", "gcn", "serial_seen", "starts_copy", "attached_to", "reason"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["pages"],
    "additionalProperties": False,
}

# Giá trị "gcn" cho trang thuộc một GCN không có trong danh sách của đơn.
ASSIGN_OTHER = "KHAC"

ASSIGN_PROMPT = """Các ảnh sau là các trang của hồ sơ quét một đơn đăng ký đất đai. Theo hệ
thống, đơn này có các giấy chứng nhận quyền sử dụng đất (GCN) sau:

{certificates}

Bản scan có thể TRẢI 2 MẶT trên một ảnh khổ ngang (A3): mặt ngoài gồm trang 4
(những thay đổi) + bìa/trang 1, mặt trong gồm trang 2 + trang 3 (thông tin thửa, sơ
đồ). Ảnh như vậy vẫn là MỘT ảnh của quyển GCN; đọc cả hai nửa.

Với MỖI ảnh, xác định trang đó có phải là một mặt của quyển GCN không (bìa đỏ/hồng có
Quốc huy và tiêu đề "GIẤY CHỨNG NHẬN", trang thông tin người sử dụng/thửa đất, trang sơ
đồ thửa đất, trang "Những thay đổi sau khi cấp giấy chứng nhận"). Nếu có thì nó thuộc
GCN nào trong danh sách:

- gcn: ghi ĐÚNG NGUYÊN VĂN số phát hành trong danh sách trên nếu trang thuộc GCN đó;
  "{other}" nếu là trang GCN nhưng của một GCN không có trong danh sách (GCN cũ, thửa
  khác...); chuỗi rỗng nếu không phải trang GCN (đơn, hợp đồng, quyết định, trích lục,
  biên bản, bản vẽ, CCCD, giấy tờ khác — kể cả khi có nhắc tới số GCN).
- Số phát hành (mã phôi, dạng 2 chữ cái + 6 chữ số, vd AA 123456) in ở trang đầu của
  quyển GCN. GCN mẫu A4 mới có số AA + 8 chữ số (vd AA 01234567), gồm 1 tờ A4 2 mặt: mặt
  trước có số phát hành, mặt sau (sơ đồ thửa đất / những thay đổi) thuộc cùng GCN đó. Các trang sơ đồ / trang thay đổi thường không có số: chúng thuộc GCN có
  trang đầu đứng ngay trước trong cùng file, trừ khi số vào sổ, số thửa hoặc tên người
  trên trang cho thấy là GCN khác. Đối chiếu thêm số vào sổ trong danh sách.
- Mỗi trang chỉ thuộc tối đa một GCN. Không chắc thì để rỗng, không đoán.
- serial_seen: số phát hành đọc được trên chính trang đó (rỗng nếu không thấy).
- attached_to: nếu ảnh là TRÍCH LỤC BẢN ĐỒ / SƠ ĐỒ THỬA ĐẤT ghi "đính kèm (kèm theo)
  giấy chứng nhận số ..." (thường đi với GCN cấp trước 2003) thì ghi số phát hành GCN mà nó
  kèm theo (đúng nguyên văn số trong danh sách nếu khớp, không khớp thì ghi số đọc được);
  trang như vậy vẫn để gcn rỗng. Các trang khác: chuỗi rỗng.
- starts_copy: true nếu ảnh này mở đầu MỘT BẢN (một lần scan) quyển GCN mới; false nếu
  là mặt tiếp theo của bản đang xem ở ảnh liền trước (kể cả khi lặp lại số phát hành).
  Cùng một GCN được scan 2 lần (vd trước và sau đính chính) thì ảnh đầu của lần scan
  sau là true. Trang không phải GCN: false.
- reason: lý do ngắn bằng tiếng Việt.
{previous}
Trả đúng một phần tử cho mỗi ảnh, giữ nguyên số i ghi trước ảnh. Chỉ xuất JSON theo schema.
"""


APPLICATION_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "pages": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "i": {"type": "integer"},
                    "is_application": {"type": "boolean"},
                    "starts_copy": {"type": "boolean"},
                    "title": {"type": "string"},
                    "reason": {"type": "string"},
                },
                "required": ["i", "is_application", "starts_copy", "title", "reason"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["pages"],
    "additionalProperties": False,
}

APPLICATION_PROMPT = """Các ảnh sau là các trang của hồ sơ quét một đơn đăng ký đất đai.
Với MỖI ảnh, xác định trang đó có thuộc ĐƠN ĐĂNG KÝ của người sử dụng đất không:
"Đơn đăng ký đất đai, tài sản gắn liền với đất", "Đơn đề nghị cấp giấy chứng nhận",
"Đơn đăng ký, cấp giấy chứng nhận...", "Đơn đăng ký biến động..." (mẫu cũ hay mới đều tính).

- is_application: true cho trang đầu của đơn (có tiêu đề "ĐƠN ...") VÀ các trang tiếp theo
  của cùng mẫu đơn (phần kê khai thửa đất, tài sản, giấy tờ, cam đoan, chữ ký người viết đơn,
  phần xác nhận / ý kiến của UBND xã, của cơ quan đăng ký đất đai ghi trên mẫu đơn).
  false cho: GCN, hợp đồng, quyết định, tờ khai thuế, thông báo, biên bản, trích lục, bản vẽ,
  CCCD, sổ hộ khẩu, giấy uỷ quyền, phiếu lấy ý kiến, danh sách, bìa hồ sơ, đơn từ khác không
  phải đơn đăng ký đất đai (đơn xin xác nhận, đơn khiếu nại...).
- starts_copy: true nếu ảnh mở đầu MỘT BẢN đơn mới (trang có tiêu đề đơn); false nếu là trang
  tiếp theo của đơn ở ảnh liền trước, hoặc không phải trang đơn.
- title: tiêu đề ngắn đọc được (rỗng nếu không có).
- reason: lý do ngắn bằng tiếng Việt.
{previous}
Trả đúng một phần tử cho mỗi ảnh, giữ nguyên số i ghi trước ảnh. Chỉ xuất JSON theo schema.
"""


FIELDS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "fields": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "key": {"type": "string"},
                    "value": {"type": "string"},
                    "page": {"type": "integer"},
                    "confidence": {"type": "string", "enum": ["cao", "trung bình", "thấp"]},
                    "evidence": {"type": "string"},
                },
                "required": ["key", "value", "page", "confidence", "evidence"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["fields"],
    "additionalProperties": False,
}

FIELDS_PROMPT = """Các ảnh sau là các trang hồ sơ quét (giấy chứng nhận quyền sử dụng đất,
đơn đăng ký...) của MỘT đơn đăng ký đất đai. Hệ thống đang thiếu các thông tin dưới đây;
hãy đọc chúng trên ảnh:

{fields}
{note}
Quy tắc:
- Mỗi khoá trả đúng một phần tử. value là giá trị đọc được, ghi đúng như trên giấy (ngày
  theo dạng DD/MM/YYYY; diện tích chỉ ghi số m², dấu chấm thập phân; năm sinh 4 chữ số).
- Chỉ lấy thông tin của ĐÚNG đối tượng được mô tả (đúng người, đúng thửa, đúng GCN). Trên GCN
  đã có biến động ở trang 4 thì ưu tiên thông tin mới nhất.
- Không thấy trên ảnh thì value rỗng, page = 0. Không suy đoán, TRỪ các khoá ghi "(suy luận)":
  khi đó được suy ra từ mẫu giấy / tiêu đề / ghi chú trên ảnh và để confidence tối đa "trung bình".
- page: số i của ảnh có thông tin. evidence: trích nguyên văn ngắn đoạn chữ trên ảnh làm căn cứ.
- confidence: "cao" khi đọc rõ, "trung bình" khi chữ mờ/viết tay hoặc suy luận, "thấp" khi không chắc.
Chỉ xuất JSON theo schema.
"""


@dataclass(frozen=True, slots=True)
class GeminiFieldValue:
    """Giá trị một trường đọc được trên hồ sơ quét."""

    value: str
    page: int
    confidence: str
    evidence: str


@dataclass(frozen=True, slots=True)
class GeminiApplicationPage:
    """Trang có thuộc đơn đăng ký đất đai không."""

    i: int
    is_application: bool
    starts_copy: bool
    title: str
    reason: str


@dataclass(frozen=True, slots=True)
class GeminiPageAssignment:
    """Trang thuộc GCN nào của đơn (``gcn`` rỗng: không phải trang GCN)."""

    i: int
    gcn: str
    serial_seen: str
    reason: str
    starts_copy: bool = False
    attached_to: str = ""


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
    owner_id: str = ""
    owner_address: str = ""


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

    def analyze(
        self,
        pages: Sequence[tuple[int, bytes]],
        *,
        low_res: bool | None = None,
    ) -> tuple[dict[int, GeminiPageResult], Usage]:
        """Phân loại một lô ảnh JPEG và trả kết quả theo số trang.

        ``low_res`` ghi đè thiết lập của analyzer cho riêng lô này (None: theo analyzer).
        """

        if low_res is None:
            low_res = self.low_res
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
        if low_res:
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
        # owner_id / owner_address bổ sung sau: thiếu thì coi là rỗng.
        values = [payload.get(name, "" if name.startswith("owner_") and name != "owner_name" else None)
                  for name in CERTIFICATE_INFO_SCHEMA["required"]]
        if not all(isinstance(value, str) for value in values):
            raise ValueError("Gemini trả sai kiểu thông tin GCN")
        info = GeminiCertificateInfo(
            certificate_serial=str(payload["certificate_serial"]).strip(),
            registration_number=str(payload["registration_number"]).strip(),
            issue_date=str(payload["issue_date"]).strip(),
            owner_name=str(payload["owner_name"]).strip(),
            reason=str(payload["reason"]).strip(),
            owner_id=str(payload.get("owner_id") or "").strip(),
            owner_address=str(payload.get("owner_address") or "").strip(),
        )
        return info, _extract_usage(response)

    def assign_certificates(
        self,
        pages: Sequence[tuple[int, bytes]],
        certificates: Sequence[tuple[str, str]],
        previous: str = "",
    ) -> tuple[dict[int, GeminiPageAssignment], Usage]:
        """Gán từng trang vào một GCN đã biết của đơn.

        ``certificates`` là các cặp (số phát hành, số vào sổ) lấy từ hệ thống (rỗng: đơn chưa có
        GCN, mọi trang GCN thành ``ASSIGN_OTHER``);
        ``previous`` mô tả trang ngay trước lô này để nối tiếp GCN giữa các lô.
        """

        if not pages:
            return {}, Usage()
        from google.genai import types

        listing = "\n".join(
            f"- Số phát hành: {serial}" + (f" | Số vào sổ cấp GCN: {book}" if book else "")
            for serial, book in certificates
        ) or (
            f"(Đơn chưa có GCN nào trên hệ thống: mọi trang GCN ghi gcn = \"{ASSIGN_OTHER}\" và đọc "
            "kỹ serial_seen trên trang đầu của từng quyển.)"
        )
        prompt = ASSIGN_PROMPT.format(
            certificates=listing,
            other=ASSIGN_OTHER,
            previous=f"\nNgữ cảnh: {previous}\n" if previous else "",
        )
        contents: list[Any] = [prompt]
        for page_number, image in pages:
            contents.append(f"Ảnh kế tiếp có i={page_number}.")
            contents.append(types.Part.from_bytes(data=image, mime_type="image/jpeg"))
        config_arguments: dict[str, Any] = {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_json_schema": ASSIGN_SCHEMA,
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
            raise ValueError("Gemini không trả kết quả gán GCN")
        return (
            _parse_assignments(
                response_text,
                {page_number for page_number, _ in pages},
                {serial for serial, _ in certificates},
            ),
            _extract_usage(response),
        )


    def find_applications(
        self,
        pages: Sequence[tuple[int, bytes]],
        previous: str = "",
        *,
        low_res: bool | None = None,
    ) -> tuple[dict[int, GeminiApplicationPage], Usage]:
        """Đánh dấu trang thuộc đơn đăng ký đất đai; ``previous`` nối ngữ cảnh lô trước."""

        if low_res is None:
            low_res = self.low_res
        if not pages:
            return {}, Usage()
        from google.genai import types

        prompt = APPLICATION_PROMPT.format(
            previous=f"\nNgữ cảnh: {previous}\n" if previous else ""
        )
        contents: list[Any] = [prompt]
        for page_number, image in pages:
            contents.append(f"Ảnh kế tiếp có i={page_number}.")
            contents.append(types.Part.from_bytes(data=image, mime_type="image/jpeg"))
        config_arguments: dict[str, Any] = {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_json_schema": APPLICATION_SCHEMA,
            "thinking_config": types.ThinkingConfig(
                thinking_level=types.ThinkingLevel.MINIMAL
            ),
        }
        if low_res:
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
            raise ValueError("Gemini không trả kết quả tìm đơn đăng ký")
        return (
            _parse_applications(response_text, {page_number for page_number, _ in pages}),
            _extract_usage(response),
        )


    def extract_fields(
        self,
        pages: Sequence[tuple[int, bytes]],
        fields: Sequence[tuple[str, str]],
        note: str = "",
    ) -> tuple[dict[str, GeminiFieldValue], Usage]:
        """Đọc các trường còn thiếu (``fields``: (khoá, mô tả)) trên ảnh trang hồ sơ."""

        if not pages or not fields:
            return {}, Usage()
        from google.genai import types

        listing = "\n".join(f"- {key}: {desc}" for key, desc in fields)
        prompt = FIELDS_PROMPT.format(
            fields=listing, note=f"\nGhi chú: {note}\n" if note else ""
        )
        contents: list[Any] = [prompt]
        for page_number, image in pages:
            contents.append(f"Ảnh kế tiếp là trang i={page_number}.")
            contents.append(types.Part.from_bytes(data=image, mime_type="image/jpeg"))
        config = types.GenerateContentConfig(
            temperature=0,
            response_mime_type="application/json",
            response_json_schema=FIELDS_SCHEMA,
            thinking_config=types.ThinkingConfig(thinking_level=types.ThinkingLevel.LOW),
        )

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
            raise ValueError("Gemini không trả kết quả đọc trường")
        return _parse_fields(response_text, {key for key, _ in fields}), _extract_usage(response)


def _parse_fields(value: str, keys: set[str]) -> dict[str, GeminiFieldValue]:
    """JSON đọc trường → {khoá: giá trị}; khoá lạ bị bỏ, khoá thiếu coi như không đọc được."""

    payload = json.loads(value)
    if not isinstance(payload, dict) or not isinstance(payload.get("fields"), list):
        raise ValueError("Gemini trả về JSON không đúng cấu trúc fields")
    parsed: dict[str, GeminiFieldValue] = {}
    for item in payload["fields"]:
        if not isinstance(item, Mapping) or item.get("key") not in keys:
            continue
        page = item.get("page")
        parsed[str(item["key"])] = GeminiFieldValue(
            value=str(item.get("value") or "").strip(),
            page=page if isinstance(page, int) else 0,
            confidence=str(item.get("confidence") or "").strip(),
            evidence=str(item.get("evidence") or "").strip(),
        )
    return parsed


def _parse_applications(value: str, expected_pages: set[int]) -> dict[int, GeminiApplicationPage]:
    """Kiểm tra JSON đánh dấu trang đơn đăng ký có đủ và đúng kiểu cho mọi trang."""

    payload = json.loads(value)
    if not isinstance(payload, dict) or not isinstance(payload.get("pages"), list):
        raise ValueError("Gemini trả về JSON không đúng cấu trúc pages")
    parsed: dict[int, GeminiApplicationPage] = {}
    for item in payload["pages"]:
        if not isinstance(item, Mapping):
            raise ValueError("Phần tử pages không phải object")
        index = item.get("i")
        is_app = item.get("is_application")
        title, reason = item.get("title"), item.get("reason")
        if (not isinstance(index, int) or type(is_app) is not bool
                or not isinstance(title, str) or not isinstance(reason, str)):
            raise ValueError("Gemini trả về sai kiểu dữ liệu")
        if index in parsed:
            raise ValueError(f"Gemini trả trùng trang {index}")
        parsed[index] = GeminiApplicationPage(index, is_app, item.get("starts_copy") is True,
                                              title.strip(), reason.strip())
    if set(parsed) != expected_pages:
        raise ValueError("Gemini không trả đủ đúng các trang trong batch")
    return parsed


def _parse_assignments(
    value: str, expected_pages: set[int], serials: set[str]
) -> dict[int, GeminiPageAssignment]:
    """Kiểm tra JSON gán trang; số GCN ngoài danh sách được coi là ``ASSIGN_OTHER``."""

    def key(text: str) -> str:
        return re.sub(r"[^0-9A-Z]", "", text.upper())

    by_key = {key(serial): serial for serial in serials}
    payload = json.loads(value)
    if not isinstance(payload, dict) or not isinstance(payload.get("pages"), list):
        raise ValueError("Gemini trả về JSON không đúng cấu trúc pages")
    parsed: dict[int, GeminiPageAssignment] = {}
    for item in payload["pages"]:
        if not isinstance(item, Mapping):
            raise ValueError("Phần tử pages không phải object")
        index = item.get("i")
        values = [item.get(name) for name in ("gcn", "serial_seen", "reason")]
        if not isinstance(index, int) or not all(isinstance(v, str) for v in values):
            raise ValueError("Gemini trả về sai kiểu dữ liệu")
        if index in parsed:
            raise ValueError(f"Gemini trả trùng trang {index}")
        gcn = values[0].strip()
        if gcn and gcn != ASSIGN_OTHER:
            gcn = by_key.get(key(gcn), ASSIGN_OTHER)
        attached = str(item.get("attached_to") or "").strip()
        if attached:
            attached = by_key.get(key(attached), ASSIGN_OTHER)
        parsed[index] = GeminiPageAssignment(index, gcn, values[1].strip(), values[2].strip(),
                                             item.get("starts_copy") is True, attached)
    if set(parsed) != expected_pages:
        raise ValueError("Gemini không trả đủ đúng các trang trong batch")
    return parsed
