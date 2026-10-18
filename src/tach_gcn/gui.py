"""Giao diện Flet cho công cụ tách giấy chứng nhận."""

from __future__ import annotations

import argparse
import asyncio
from dataclasses import dataclass, field
from io import StringIO
import multiprocessing
import os
from pathlib import Path
import re
from typing import Any

import flet as ft
from rich.console import Console

from .cli import run
from .gemini import DEFAULT_MODEL


@dataclass(slots=True)
class GuiSettings:
    """Giá trị biểu mẫu GUI, tách khỏi Flet để dễ kiểm thử."""

    input_dir: str
    output_dir: str = ""
    api_key: str = field(default="", repr=False)
    model: str = DEFAULT_MODEL
    keywords: str = ""
    mode: str = "perfile"
    batch: str = "6"
    rpm: str = "30"
    workers: str = "4"
    include_rest: bool = False
    dry_run: bool = False
    no_gemini: bool = False
    gemini_all: bool = False
    low_res: bool = False
    no_cache: bool = False


def _positive_int(value: str, label: str) -> int:
    """Đọc một số nguyên dương từ biểu mẫu."""

    try:
        parsed = int(value)
    except ValueError as error:
        raise ValueError(f"{label} phải là số nguyên") from error
    if parsed < 1:
        raise ValueError(f"{label} phải lớn hơn 0")
    return parsed


def _positive_float(value: str, label: str) -> float:
    """Đọc một số thực dương từ biểu mẫu."""

    try:
        parsed = float(value)
    except ValueError as error:
        raise ValueError(f"{label} phải là số") from error
    if parsed <= 0:
        raise ValueError(f"{label} phải lớn hơn 0")
    return parsed


def build_arguments(settings: GuiSettings) -> argparse.Namespace:
    """Chuyển biểu mẫu thành Namespace tương thích hoàn toàn với lõi CLI."""

    input_text = settings.input_dir.strip()
    if not input_text:
        raise ValueError("Hãy chọn thư mục đầu vào")
    if settings.mode not in {"perfile", "each"}:
        raise ValueError("Chế độ tách không hợp lệ")
    keywords = [
        item.strip()
        for item in re.split(r"[;\r\n]+", settings.keywords)
        if item.strip()
    ]
    return argparse.Namespace(
        thu_muc=Path(input_text),
        output=Path(settings.output_dir.strip()) if settings.output_dir.strip() else None,
        keyword=keywords,
        mode=settings.mode,
        rest=settings.include_rest,
        dry_run=settings.dry_run,
        no_gemini=settings.no_gemini,
        gemini_all=settings.gemini_all,
        model=settings.model.strip() or DEFAULT_MODEL,
        batch=_positive_int(settings.batch.strip(), "Số trang mỗi batch"),
        low_res=settings.low_res,
        rpm=_positive_float(settings.rpm.strip(), "RPM"),
        workers=_positive_int(settings.workers.strip(), "Số worker"),
        no_cache=settings.no_cache,
        # Không đưa key vào argv, biến môi trường, cache hoặc file báo cáo.
        api_key=settings.api_key.strip(),
    )


async def app(page: ft.Page) -> None:
    """Tạo và điều khiển cửa sổ chính của ứng dụng."""

    page.title = "Tách giấy chứng nhận"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 18
    page.scroll = ft.ScrollMode.AUTO
    page.window.width = 980
    page.window.height = 820
    page.window.min_width = 760
    page.window.min_height = 620

    input_picker = ft.FilePicker()
    output_picker = ft.FilePicker()
    page.services.extend([input_picker, output_picker])

    input_field = ft.TextField(
        label="Thư mục PDF đầu vào *",
        hint_text="Chọn thư mục chứa PDF cần quét",
        expand=True,
    )
    output_field = ft.TextField(
        label="Thư mục kết quả",
        hint_text="Để trống để dùng <thư-mục-vào>_giay-chung-nhan",
        expand=True,
    )
    api_key_field = ft.TextField(
        label="Gemini API key",
        value=os.environ.get("GEMINI_API_KEY", ""),
        password=True,
        can_reveal_password=True,
        helper="Chỉ giữ trong bộ nhớ phiên chạy, không ghi vào báo cáo hoặc cache.",
        expand=True,
    )
    model_field = ft.TextField(label="Model Gemini", value=DEFAULT_MODEL, col={"sm": 12, "md": 6})
    keywords_field = ft.TextField(
        label="Từ khoá bổ sung",
        hint_text="Mỗi dòng hoặc dấu ; là một từ khoá",
        multiline=True,
        min_lines=2,
        max_lines=3,
        col=12,
    )
    mode_field = ft.Dropdown(
        label="Kiểu xuất",
        value="perfile",
        options=[
            ft.DropdownOption(key="perfile", text="Gộp theo file gốc"),
            ft.DropdownOption(key="each", text="Mỗi trang một file"),
        ],
        col={"sm": 12, "md": 6},
    )
    batch_field = ft.TextField(label="Trang / batch", value="6", col={"sm": 4, "md": 2})
    rpm_field = ft.TextField(label="Request / phút", value="30", col={"sm": 4, "md": 2})
    workers_field = ft.TextField(label="Số worker", value="4", col={"sm": 4, "md": 2})

    rest_check = ft.Checkbox(label="Xuất phần còn lại thành <tên-gốc>-GT.pdf", value=True)
    dry_run_check = ft.Checkbox(label="Chỉ quét, không tách PDF (--dry-run)")
    no_gemini_check = ft.Checkbox(label="Không dùng Gemini (--no-gemini)")
    gemini_all_check = ft.Checkbox(label="Gửi mọi trang cho Gemini")
    low_res_check = ft.Checkbox(label="Giảm media resolution")
    no_cache_check = ft.Checkbox(label="Bỏ qua cache")

    status = ft.Text("Sẵn sàng.", weight=ft.FontWeight.W_600)
    progress = ft.ProgressBar(value=None, visible=False)
    log_field = ft.TextField(
        label="Nhật ký và tổng kết",
        value="",
        multiline=True,
        min_lines=8,
        max_lines=14,
        read_only=True,
    )

    async def choose_input(_: Any) -> None:
        selected = await input_picker.get_directory_path(
            dialog_title="Chọn thư mục PDF đầu vào"
        )
        if selected:
            input_field.value = selected
            if not output_field.value.strip():
                source = Path(selected)
                output_field.value = str(source.parent / f"{source.name}_giay-chung-nhan")
            page.update()

    async def choose_output(_: Any) -> None:
        selected = await output_picker.get_directory_path(
            dialog_title="Chọn thư mục lưu kết quả"
        )
        if selected:
            output_field.value = selected
            page.update()

    def toggle_gemini(_: Any) -> None:
        disabled = bool(no_gemini_check.value)
        api_key_field.disabled = disabled
        model_field.disabled = disabled
        gemini_all_check.disabled = disabled
        low_res_check.disabled = disabled
        page.update()

    no_gemini_check.on_change = toggle_gemini

    run_button = ft.Button(
        content="Bắt đầu quét",
        icon=ft.Icons.PLAY_ARROW,
        bgcolor=ft.Colors.BLUE_700,
        color=ft.Colors.WHITE,
    )

    async def start_scan(_: Any) -> None:
        try:
            settings = GuiSettings(
                input_dir=input_field.value,
                output_dir=output_field.value,
                api_key=api_key_field.value,
                model=model_field.value,
                keywords=keywords_field.value,
                mode=mode_field.value or "perfile",
                batch=batch_field.value,
                rpm=rpm_field.value,
                workers=workers_field.value,
                include_rest=bool(rest_check.value),
                dry_run=bool(dry_run_check.value),
                no_gemini=bool(no_gemini_check.value),
                gemini_all=bool(gemini_all_check.value),
                low_res=bool(low_res_check.value),
                no_cache=bool(no_cache_check.value),
            )
            arguments = build_arguments(settings)
        except ValueError as error:
            status.value = f"Lỗi nhập liệu: {error}"
            status.color = ft.Colors.RED_700
            page.update()
            return

        run_button.disabled = True
        progress.visible = True
        status.value = "Đang quét PDF, vui lòng chờ…"
        status.color = ft.Colors.BLUE_700
        log_field.value = ""
        page.update()

        output = StringIO()
        console = Console(file=output, force_terminal=False, no_color=True, width=100)
        try:
            exit_code = await asyncio.to_thread(run, arguments, console)
            log_field.value = output.getvalue().strip()
            if exit_code == 0:
                status.value = "Hoàn tất. Xem tổng kết và đường dẫn báo cáo bên dưới."
                status.color = ft.Colors.GREEN_700
            else:
                status.value = f"Không hoàn tất (mã lỗi {exit_code})."
                status.color = ft.Colors.RED_700
        except Exception as error:
            log_field.value = output.getvalue().strip()
            status.value = f"Lỗi không mong đợi: {error}"
            status.color = ft.Colors.RED_700
        finally:
            # Xoá bản sao key nằm trong Namespace ngay khi lượt chạy kết thúc.
            arguments.api_key = ""
            progress.visible = False
            run_button.disabled = False
            page.update()

    run_button.on_click = start_scan

    page.add(
        ft.Text("TÁCH GIẤY CHỨNG NHẬN", size=25, weight=ft.FontWeight.BOLD),
        ft.Text(
            "Quét PDF đệ quy, nhận diện bằng từ khoá và Gemini, giữ nguyên cấu trúc thư mục.",
            color=ft.Colors.GREY_700,
        ),
        ft.Card(
            content=ft.Container(
                padding=18,
                content=ft.Column(
                    [
                        ft.Text("Thư mục", size=18, weight=ft.FontWeight.BOLD),
                        ft.Row(
                            [
                                input_field,
                                ft.Button(
                                    content="Chọn…",
                                    icon=ft.Icons.FOLDER_OPEN,
                                    on_click=choose_input,
                                ),
                            ]
                        ),
                        ft.Row(
                            [
                                output_field,
                                ft.Button(
                                    content="Chọn…",
                                    icon=ft.Icons.FOLDER_OPEN,
                                    on_click=choose_output,
                                ),
                            ]
                        ),
                    ]
                ),
            )
        ),
        ft.Card(
            content=ft.Container(
                padding=18,
                content=ft.Column(
                    [
                        ft.Text("Gemini", size=18, weight=ft.FontWeight.BOLD),
                        api_key_field,
                        ft.ResponsiveRow([model_field, mode_field, batch_field, rpm_field, workers_field]),
                        ft.Text(
                            "Lưu ý quyền riêng tư: các trang cần phân tích sẽ được gửi dưới dạng ảnh lên máy chủ Google.",
                            color=ft.Colors.ORANGE_800,
                            size=12,
                        ),
                    ]
                ),
            )
        ),
        ft.Card(
            content=ft.Container(
                padding=18,
                content=ft.Column(
                    [
                        ft.Text("Tuỳ chọn", size=18, weight=ft.FontWeight.BOLD),
                        keywords_field,
                        ft.ResponsiveRow(
                            [
                                ft.Container(rest_check, col={"sm": 12, "md": 6}),
                                ft.Container(dry_run_check, col={"sm": 12, "md": 6}),
                                ft.Container(no_gemini_check, col={"sm": 12, "md": 6}),
                                ft.Container(gemini_all_check, col={"sm": 12, "md": 6}),
                                ft.Container(low_res_check, col={"sm": 12, "md": 6}),
                                ft.Container(no_cache_check, col={"sm": 12, "md": 6}),
                            ]
                        ),
                    ]
                ),
            )
        ),
        ft.Row([run_button, status], vertical_alignment=ft.CrossAxisAlignment.CENTER),
        progress,
        log_field,
        ft.Text(
            "API key không được lưu. Nếu muốn dùng tự động ở lần sau, hãy đặt biến môi trường GEMINI_API_KEY.",
            size=12,
            color=ft.Colors.GREY_700,
        ),
    )


def main() -> None:
    """Entry point ``tach-gcn-gui``."""

    multiprocessing.freeze_support()
    ft.run(app)


if __name__ == "__main__":
    main()
