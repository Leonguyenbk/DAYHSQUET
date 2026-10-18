"""Kiểm thử mock Gemini, retry và cache."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

from tach_gcn.cli import _apply_gemini_to_file
from tach_gcn.detect import detect_page
from tach_gcn.gemini import (
    GeminiAnalyzer,
    GeminiPageResult,
    Usage,
    call_with_retry,
)
from tach_gcn.pdfio import InspectedFile, InspectedPage


class Http429(Exception):
    """Ngoại lệ giả lập quota HTTP."""

    status_code = 429


def test_retry_429_with_exponential_backoff() -> None:
    """Lỗi 429 được thử lại, còn lần gọi thành công được trả về."""

    calls = 0
    delays: list[float] = []

    def operation() -> str:
        nonlocal calls
        calls += 1
        if calls < 3:
            raise Http429("429 quota")
        return "xong"

    assert call_with_retry(
        operation,
        sleep=delays.append,
        random_value=lambda: 0.25,
    ) == "xong"
    assert calls == 3
    assert delays == [1.25, 2.25]


class FakeAnalyzer:
    """Gemini giả để đếm số trang thực sự được gửi."""

    def __init__(self) -> None:
        self.calls = 0

    def analyze(self, pages: Any) -> tuple[dict[int, GeminiPageResult], Usage]:
        """Trả một kết luận cố định cho từng trang."""

        self.calls += 1
        return (
            {
                page: GeminiPageResult(page, True, "Giấy chứng nhận", "Có dấu và chữ ký")
                for page, _ in pages
            },
            Usage(100, 20),
        )


def _scan_file() -> InspectedFile:
    """Tạo kết quả worker giả có một trang scan."""

    result = detect_page("", 1, gemini_enabled=True)
    return InspectedFile(
        "mau.pdf",
        "mau.pdf",
        content_hash="abc",
        page_count=1,
        pages=[InspectedPage(result, b"jpeg")],
    )


def test_cache_avoids_second_gemini_call() -> None:
    """Chạy lại cùng hash/trang/model phải dùng cache."""

    cache: dict[str, dict[str, object]] = {}
    analyzer = FakeAnalyzer()
    usage = Usage()
    first = _scan_file()
    sent, cached = _apply_gemini_to_file(
        first,
        analyzer=analyzer,  # type: ignore[arg-type]
        model="model-test",
        batch_size=6,
        cache=cache,
        use_cache=True,
        usage=usage,
    )
    assert (sent, cached, analyzer.calls) == (1, 0, 1)
    assert first.pages[0].result.method == "Gemini"

    second = _scan_file()
    sent, cached = _apply_gemini_to_file(
        second,
        analyzer=analyzer,  # type: ignore[arg-type]
        model="model-test",
        batch_size=6,
        cache=cache,
        use_cache=True,
        usage=usage,
    )
    assert (sent, cached, analyzer.calls) == (0, 1, 1)
    assert second.pages[0].result.is_certificate
    assert second.pages[0].result.note == "Giấy chứng nhận — Có dấu và chữ ký"


def test_gemini_analyzer_uses_structured_response_and_mock_client() -> None:
    """Analyzer đọc structured JSON và không gọi mạng khi dùng client giả."""

    response = SimpleNamespace(
        text='{"pages":[{"i":3,"is_certificate":false,"title":"","reason":"Quyết định"}]}',
        usage_metadata=SimpleNamespace(prompt_token_count=120, total_token_count=150),
    )

    class Models:
        def __init__(self) -> None:
            self.config: Any = None

        def generate_content(self, **kwargs: Any) -> Any:
            self.config = kwargs["config"]
            return response

    models = Models()
    client = SimpleNamespace(models=models)
    analyzer = GeminiAnalyzer(
        "fake",
        client=client,
        low_res=True,
        rpm=1_000_000,
        sleep=lambda _: None,
    )
    results, usage = analyzer.analyze([(3, b"fake-jpeg")])
    assert not results[3].is_certificate
    assert results[3].reason == "Quyết định"
    assert (usage.input_tokens, usage.output_tokens) == (120, 30)
    assert models.config.response_mime_type == "application/json"
    assert models.config.temperature == 0
    assert models.config.response_json_schema["required"] == ["pages"]


def test_gemini_extracts_group_information_with_mock_client() -> None:
    """Một request đọc chung các mặt trả riêng mã phôi và số vào sổ."""

    response = SimpleNamespace(
        text=(
            '{"certificate_serial":"BB 654321",'
            '"registration_number":"CS 012345","issue_date":"01/02/2020",'
            '"owner_name":"Trần Thị B","reason":"Chủ mới ở trang 4"}'
        ),
        usage_metadata=SimpleNamespace(prompt_token_count=200, total_token_count=250),
    )

    class Models:
        def generate_content(self, **kwargs: Any) -> Any:
            assert len(kwargs["contents"]) == 5  # prompt + hai nhãn + hai ảnh
            return response

    analyzer = GeminiAnalyzer(
        "fake",
        client=SimpleNamespace(models=Models()),
        rpm=1_000_000,
        sleep=lambda _: None,
    )
    info, usage = analyzer.extract_certificate([(1, b"front"), (4, b"changes")])
    assert info.certificate_serial == "BB 654321"
    assert info.registration_number == "CS 012345"
    assert info.issue_date == "01/02/2020"
    assert info.owner_name == "Trần Thị B"
    assert info.reason == "Chủ mới ở trang 4"
    assert (usage.input_tokens, usage.output_tokens) == (200, 50)
