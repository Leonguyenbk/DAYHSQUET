"""Giao diện dòng lệnh tiếng Việt cho ``tach-gcn``."""

from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor, as_completed
import os
from pathlib import Path
import sys
from typing import Sequence

from rich.console import Console
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn

from .detect import (
    apply_gemini_result,
    extract_certificate_serial,
    extract_registration_number,
)
from .gemini import DEFAULT_MODEL, GeminiAnalyzer, Usage
from .pdfio import (
    InspectedFile,
    build_certificate_groups,
    discover_pdfs,
    inspect_pdf,
    render_pdf_pages,
    write_outputs,
    write_repaired_pdf,
)
from .report import (
    cache_key,
    cached_gemini_result,
    load_cache,
    rows_for_file,
    save_cache,
    serialize_gemini_result,
    write_certificate_csv,
    write_csv,
)


def _configure_utf8_console() -> None:
    """Ưu tiên UTF-8 để console Windows luôn in được thông báo tiếng Việt."""

    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except (OSError, ValueError):
                pass


def build_parser() -> argparse.ArgumentParser:
    """Tạo bộ phân tích tham số dòng lệnh."""

    parser = argparse.ArgumentParser(
        prog="tach-gcn",
        description="Quét đệ quy PDF và tách các trang giấy chứng nhận/chứng chỉ.",
    )
    parser.add_argument("thu_muc", type=Path, help="Thư mục PDF cần quét")
    parser.add_argument("-o", "--output", type=Path, help="Thư mục kết quả")
    parser.add_argument(
        "-k",
        "--keyword",
        action="append",
        default=[],
        metavar="TỪ_KHOÁ",
        help="Từ khoá bổ sung (+6 điểm), có thể dùng nhiều lần",
    )
    parser.add_argument(
        "--mode",
        choices=("perfile", "each"),
        default="perfile",
        help="Gộp theo file hoặc tạo một file cho mỗi trang (mặc định: perfile)",
    )
    parser.add_argument("--rest", action="store_true", help="Xuất thêm các trang còn lại")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ quét và ghi báo cáo")
    parser.add_argument("--no-gemini", action="store_true", help="Không gọi Gemini")
    parser.add_argument("--gemini-all", action="store_true", help="Gửi mọi trang cho Gemini")
    parser.add_argument(
        "--model",
        default=os.environ.get("GEMINI_MODEL", DEFAULT_MODEL),
        help="Model Gemini (mặc định: GEMINI_MODEL hoặc %(default)s)",
    )
    parser.add_argument("--batch", type=int, default=6, help="Số trang mỗi request Gemini")
    parser.add_argument("--low-res", action="store_true", help="Dùng media resolution thấp")
    parser.add_argument("--rpm", type=float, default=30, help="Giới hạn request Gemini/phút")
    parser.add_argument("--workers", type=int, default=4, help="Số process đọc PDF song song")
    parser.add_argument("--no-cache", action="store_true", help="Bỏ qua cache Gemini")
    return parser


def _positive_arguments(parser: argparse.ArgumentParser, arguments: argparse.Namespace) -> None:
    """Báo lỗi cho các tham số số không hợp lệ."""

    if arguments.batch < 1:
        parser.error("--batch phải lớn hơn 0")
    if arguments.workers < 1:
        parser.error("--workers phải lớn hơn 0")
    if arguments.rpm <= 0:
        parser.error("--rpm phải lớn hơn 0")


def _default_output(input_dir: Path) -> Path:
    """Tạo đường dẫn đầu ra mặc định cạnh thư mục đầu vào."""

    return input_dir.parent / f"{input_dir.name}_giay-chung-nhan"


def _append_note(current: str, message: str) -> str:
    """Nối ghi chú mà không tạo dấu phân cách thừa."""

    return "; ".join(part for part in (current, message) if part)


def _apply_gemini_to_file(
    inspected: InspectedFile,
    *,
    analyzer: GeminiAnalyzer | None,
    model: str,
    batch_size: int,
    cache: dict[str, dict[str, object]],
    use_cache: bool,
    usage: Usage,
) -> tuple[int, int]:
    """Dùng cache rồi gọi Gemini cho các trang còn thiếu của một file."""

    candidates = [page for page in inspected.pages if page.result.needs_gemini]
    to_send = []
    cached_count = 0
    for page in candidates:
        key = cache_key(inspected.content_hash, page.result.page, model)
        cached = cached_gemini_result(cache.get(key, {})) if use_cache else None
        if cached is not None:
            apply_gemini_result(
                page.result,
                cached.is_certificate,
                cached.title,
                cached.reason,
            )
            page.jpeg = None
            cached_count += 1
        else:
            to_send.append(page)

    sent_count = 0
    if analyzer is None:
        return sent_count, cached_count
    for start in range(0, len(to_send), batch_size):
        batch = to_send[start : start + batch_size]
        images = [(page.result.page, page.jpeg or b"") for page in batch]
        sent_count += len(batch)
        try:
            responses, request_usage = analyzer.analyze(images)
            usage.add(request_usage)
            for page in batch:
                response = responses[page.result.page]
                apply_gemini_result(
                    page.result,
                    response.is_certificate,
                    response.title,
                    response.reason,
                )
                if use_cache:
                    key = cache_key(inspected.content_hash, page.result.page, model)
                    cache[key] = serialize_gemini_result(response)
                page.jpeg = None
        except Exception as error:
            message = f"Lỗi Gemini: {error}; giữ kết quả từ khoá"
            for page in batch:
                page.result.note = _append_note(page.result.note, message)
                page.result.needs_gemini = False
                page.jpeg = None
    return sent_count, cached_count


def _enrich_certificate_groups(
    inspected: InspectedFile,
    *,
    analyzer: GeminiAnalyzer | None,
    model: str,
    cache: dict[str, dict[str, object]],
    use_cache: bool,
    usage: Usage,
) -> tuple[int, int]:
    """Gom trang và trích thông tin tổng hợp của từng GCN."""

    groups = build_certificate_groups(inspected)
    if analyzer is None:
        return 0, 0
    sent_pages = 0
    cached_pages = 0
    for group in groups:
        pages_token = ",".join(str(index + 1) for index in group.page_indexes)
        key = f"info:{inspected.content_hash}:{pages_token}:{model}"
        cached = cache.get(key) if use_cache else None
        if isinstance(cached, dict) and all(
            isinstance(cached.get(name), str)
            for name in (
                "certificate_serial",
                "registration_number",
                "issue_date",
                "owner_name",
                "reason",
            )
        ):
            cached_serial = str(cached["certificate_serial"])
            cached_registration = str(cached["registration_number"])
            if cached_serial and cached_serial != cached_registration:
                group.certificate_serial = cached_serial
            group.registration_number = cached_registration
            group.issue_date = str(cached["issue_date"])
            group.owner_name = str(cached["owner_name"])
            group.note = str(cached["reason"])
            if cached_serial and cached_serial == cached_registration:
                group.note = _append_note(
                    group.note,
                    "Bỏ số giấy chứng nhận do trùng số vào sổ cấp GCN",
                )
            cached_pages += len(group.page_indexes)
            continue
        try:
            images = render_pdf_pages(
                Path(inspected.path),
                group.page_indexes,
                inspected.recovered_page_xrefs,
            )
            sent_pages += len(images)
            info, request_usage = analyzer.extract_certificate(images)
            usage.add(request_usage)
            canonical_registration = extract_registration_number(
                f"Số vào sổ cấp GCN: {info.registration_number}"
            )
            canonical_serial = extract_certificate_serial(
                f"Số giấy chứng nhận: {info.certificate_serial}"
            )
            duplicate_numbers = bool(
                canonical_serial
                and canonical_registration
                and canonical_serial == canonical_registration
            )
            if duplicate_numbers:
                canonical_serial = ""
            if canonical_serial:
                group.certificate_serial = canonical_serial
            if canonical_registration:
                group.registration_number = canonical_registration
            if info.issue_date:
                group.issue_date = info.issue_date
            if info.owner_name:
                group.owner_name = info.owner_name
            group.note = _append_note(info.reason, inspected.warning)
            if duplicate_numbers:
                group.note = _append_note(
                    group.note,
                    "Gemini trả số giấy chứng nhận trùng số vào sổ; đã bỏ giá trị sai",
                )
            if use_cache:
                cache[key] = {
                    "certificate_serial": group.certificate_serial,
                    "registration_number": group.registration_number,
                    "issue_date": group.issue_date,
                    "owner_name": group.owner_name,
                    "reason": group.note,
                }
        except Exception as error:
            group.note = _append_note(group.note, f"Lỗi trích thông tin Gemini: {error}")
    return sent_pages, cached_pages


def run(arguments: argparse.Namespace, console: Console | None = None) -> int:
    """Thực thi một lượt quét và trả mã thoát."""

    console = console or Console()
    input_dir = arguments.thu_muc.expanduser().resolve()
    if not input_dir.is_dir():
        console.print(f"[red]Lỗi:[/red] Không tìm thấy thư mục đầu vào: {input_dir}")
        return 2
    output_dir = (arguments.output or _default_output(input_dir)).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # GUI truyền key trực tiếp trong bộ nhớ; CLI tiếp tục đọc từ môi trường.
    supplied_api_key = getattr(arguments, "api_key", None)
    api_key = (
        str(supplied_api_key).strip()
        if supplied_api_key is not None
        else os.environ.get("GEMINI_API_KEY", "").strip()
    )
    gemini_enabled = not arguments.no_gemini and bool(api_key)
    if not arguments.no_gemini and not api_key:
        console.print(
            "[yellow]Không có GEMINI_API_KEY: chỉ nhận diện bằng từ khoá; "
            "trang scan sẽ được ghi 'chưa phân tích'.[/yellow]"
        )

    pdf_files = discover_pdfs(input_dir, output_dir)
    console.print(f"Tìm thấy [bold]{len(pdf_files)}[/bold] file PDF.")

    cache_path = output_dir / ".tach-gcn-cache.json"
    cache = load_cache(cache_path) if not arguments.no_cache else {}
    analyzer = (
        GeminiAnalyzer(
            api_key,
            arguments.model,
            low_res=arguments.low_res,
            rpm=arguments.rpm,
        )
        if gemini_enabled
        else None
    )

    completed_files: list[InspectedFile] = []
    usage = Usage()
    sent_pages = 0
    cached_pages = 0
    certificates = 0
    total_pages = 0

    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TextColumn("• GCN: {task.fields[gcn]} • gửi Gemini: {task.fields[gemini]}"),
        console=console,
    )
    with progress:
        task = progress.add_task("Đang quét", total=len(pdf_files), gcn=0, gemini=0)
        with ProcessPoolExecutor(max_workers=arguments.workers) as executor:
            futures = {
                executor.submit(
                    inspect_pdf,
                    str(path),
                    str(path.relative_to(input_dir)),
                    tuple(arguments.keyword),
                    arguments.gemini_all,
                    gemini_enabled,
                ): path
                for path in pdf_files
            }
            for future in as_completed(futures):
                path = futures[future]
                try:
                    inspected = future.result()
                except Exception as error:
                    inspected = InspectedFile(
                        str(path),
                        str(path.relative_to(input_dir)),
                        error=f"Lỗi worker: {error}",
                    )

                sent, cached = _apply_gemini_to_file(
                    inspected,
                    analyzer=analyzer,
                    model=arguments.model,
                    batch_size=arguments.batch,
                    cache=cache,
                    use_cache=not arguments.no_cache,
                    usage=usage,
                )
                sent_pages += sent
                cached_pages += cached
                info_sent, info_cached = _enrich_certificate_groups(
                    inspected,
                    analyzer=analyzer,
                    model=arguments.model,
                    cache=cache,
                    use_cache=not arguments.no_cache,
                    usage=usage,
                )
                sent_pages += info_sent
                cached_pages += info_cached
                if not arguments.dry_run and not inspected.error:
                    try:
                        repaired = write_repaired_pdf(inspected, output_dir)
                        if repaired is not None:
                            repaired_relative = repaired.relative_to(output_dir).as_posix()
                            inspected.warning = _append_note(
                                inspected.warning,
                                f"Đã tự tạo bản Save As: {repaired_relative}",
                            )
                        write_outputs(
                            inspected,
                            output_dir,
                            mode=arguments.mode,
                            include_rest=arguments.rest,
                        )
                    except Exception as error:
                        message = f"Lỗi ghi PDF đầu ra: {error}"
                        if inspected.pages:
                            inspected.pages[0].result.note = _append_note(
                                inspected.pages[0].result.note, message
                            )
                        else:
                            inspected.error = message
                completed_files.append(inspected)
                total_pages += inspected.page_count
                certificates += sum(
                    page.result.is_certificate for page in inspected.pages
                )
                progress.update(
                    task,
                    advance=1,
                    gcn=certificates,
                    gemini=sent_pages,
                )

    completed_files.sort(key=lambda item: item.relative_path.casefold())
    rows = [row for item in completed_files for row in rows_for_file(item)]
    report_path = output_dir / "bao-cao.csv"
    write_csv(report_path, rows)
    certificate_report_path = output_dir / "thong-tin-gcn.csv"
    write_certificate_csv(certificate_report_path, completed_files)
    if not arguments.no_cache:
        save_cache(cache_path, cache)

    console.print("\n[bold green]Hoàn tất.[/bold green]")
    console.print(f"Số file: {len(pdf_files)}")
    console.print(f"Số trang: {total_pages}")
    console.print(f"Trang giấy chứng nhận: {certificates}")
    console.print(f"Trang đã gửi Gemini: {sent_pages}")
    if cached_pages:
        console.print(f"Trang dùng cache Gemini: {cached_pages}")
    console.print(f"Chi phí Gemini ước tính: {usage.estimated_cost_usd:.6f} USD")
    console.print(f"Báo cáo: {report_path}")
    console.print(f"Thông tin GCN: {certificate_report_path}")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """Entry point của lệnh ``tach-gcn``."""

    _configure_utf8_console()
    parser = build_parser()
    arguments = parser.parse_args(argv)
    _positive_arguments(parser, arguments)
    try:
        return run(arguments)
    except KeyboardInterrupt:
        Console().print("\n[yellow]Đã dừng theo yêu cầu người dùng.[/yellow]")
        return 130


if __name__ == "__main__":
    sys.exit(main())
