# CÔNG CỤ TÌM FILE THEO SỐ GCN

Ứng dụng Windows đọc danh sách số phát hành Giấy chứng nhận quyền sử dụng đất (GCN) từ Excel, tìm mã trong tên file, text layer PDF và OCR ảnh/PDF, rồi sao chép **nguyên file nguồn** sang thư mục kết quả. Dữ liệu được xử lý hoàn toàn trên máy; chương trình không dùng API cloud và không gửi tài liệu ra Internet.

## Quy tắc GCN

GCN hợp lệ có đúng hai chữ cái (`A-Z` hoặc `Đ`) và đúng sáu chữ số. Ứng dụng dùng khóa nội bộ `AM143443`, nhưng mọi nội dung người dùng nhìn thấy và mọi tên file đều dùng dạng `AM 143443`.

Hai ngoại lệ theo yêu cầu nghiệp vụ: **chỉ riêng chữ `D`** + đúng bảy chữ số (`D 1234567`), và **chỉ riêng cặp `AA`** + đúng tám chữ số (`AA 12345678`) cũng được coi là GCN hợp lệ. Chữ/cặp chữ khác vẫn phải đúng sáu chữ số như quy tắc chung.

Các dạng `AM143443`, `AM-143443`, `AM.143443`, `AM:143443`, `A M 143443` đều được chuẩn hóa thành `AM 143443`. Các dạng một/ba chữ cái (ngoài ngoại lệ D và AA ở trên), năm chữ số, số CCCD hoặc mã không có trong Excel không được xác nhận.

Một ô Excel có thể chứa nhiều GCN cách nhau bằng dấu `;` (ví dụ `DH 358799;DH 358800`); mỗi phần được tách và đối chiếu như một dòng riêng.

Khi một GCN khớp nhiều nguồn, tên lần lượt là:

```text
AM 143443.pdf
AM 143443_02.pdf
AM 143443_03.jpg
```

Ứng dụng không ghi đè. Khi chạy lại, file cùng kích thước và cùng SHA-256 được ghi trạng thái `ĐÃ TỒN TẠI`, không tạo bản sao mới.

## Yêu cầu hệ thống

- Windows 11 64-bit.
- Python 3.11 trở lên, bản 64-bit, có `py launcher` là thuận tiện nhất.
- RAM tối thiểu 8 GB; nên có 16 GB khi OCR PDF 350–400 DPI.
- CPU chạy được đầy đủ. GPU là tùy chọn và cần bộ CUDA/PaddlePaddle tương thích.
- Internet chỉ cần cho lần cài thư viện và tải model. Sau đó có thể chạy hoàn toàn offline.
- Không cần cài Tesseract.

Requirements dùng PaddleOCR 3.x và PaddlePaddle 3.3 CPU. Tài liệu chính thức xác nhận PaddleOCR 3.x cần PaddlePaddle 3.0+ để suy luận cục bộ và PaddlePaddle Windows hỗ trợ Python 3.11: [PaddleOCR installation](https://www.paddleocr.ai/main/en/version3.x/installation.html), [PaddlePaddle Windows installation](https://www.paddlepaddle.org.cn/documentation/docs/en/install/pip/windows-pip_en.html).

Khi cài Python từ trang Python chính thức, bật lựa chọn thêm Python vào `PATH` và cài `py launcher`. Kiểm tra:

```powershell
py -3.11 --version
```

## Cài đặt lần đầu và tải model offline

Mở PowerShell tại thư mục này:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\run.ps1 -InstallOnly
.\download_models.ps1
```

Lệnh đầu tạo `.venv`, cài dependencies trong `requirements.txt`; lệnh thứ hai là thao tác mạng **chủ động duy nhất** để tải model tiếng Việt vào:

```text
models\paddleocr\det
models\paddleocr\rec
models\paddleocr\cls
```

Downloader dùng PP-OCRv5 mobile detection, model nhận dạng Latin hỗ trợ tiếng Việt và model xoay dòng chữ; cache tải tạm được xóa sau khi ba model đã được sao chép và khởi tạo kiểm tra thành công. Khi chạy offline, ứng dụng cung cấp cả ba đường dẫn local và đặt `PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK=1` để PaddleX không dò các máy chủ model.

Ứng dụng không âm thầm tải model. Nếu một trong ba thư mục trên thiếu hoặc rỗng, bước `KIỂM TRA DỮ LIỆU` dừng với thông báo rõ ràng. Sau khi tải đủ, có thể ngắt Internet.

Nếu PowerShell chặn script, chỉ dùng `Set-ExecutionPolicy -Scope Process Bypass`; thiết lập này hết hiệu lực khi đóng cửa sổ PowerShell.

## Chạy chương trình

```powershell
.\run.ps1
```

Nếu môi trường đã đầy đủ và không muốn kiểm tra/cài lại dependency:

```powershell
.\run.ps1 -SkipInstall
```

Quy trình sử dụng:

1. Chọn file `.xlsx`/`.xlsm`. Chương trình liệt kê sheet và cột.
2. Chọn sheet. Cột có tiêu đề `GCN` (không phân biệt hoa thường/khoảng trắng) được chọn tự động; có thể chọn cột khác.
3. Chọn thư mục nguồn PDF/ảnh và thư mục kết quả. Vì an toàn, thư mục kết quả không được nằm trong thư mục nguồn.
4. Chọn quét thư mục con, kiểm tra nội dung sau tên file, OCR màu đỏ, tìm tất cả GCN, DPI và số luồng.
5. Nhấn `KIỂM TRA DỮ LIỆU`. Bảng tóm tắt cho biết số dòng, GCN hợp lệ/sai/trùng và tổng file.
6. Nhấn `BẮT ĐẦU TÌM`. Có thể tạm dừng, tiếp tục hoặc dừng an toàn.

Mặc định CPU dùng hai worker. PaddleOCR chỉ có một model trong RAM và việc gọi model được tuần tự hóa; các worker vẫn có thể đọc text PDF, render và chuẩn bị ảnh song song. DPI 350 cân bằng độ rõ/dung lượng; giảm xuống 300 nếu thiếu RAM, tăng 400 cho bản quét chữ quá nhỏ.

### CPU và GPU

`requirements.txt` cài `paddlepaddle` bản CPU. Để dùng GPU:

1. Xác định CUDA/driver của máy.
2. Trong `.venv`, gỡ `paddlepaddle` CPU và cài đúng gói `paddlepaddle-gpu` theo [ma trận cài đặt Windows chính thức](https://www.paddlepaddle.org.cn/documentation/docs/en/install/pip/windows-pip_en.html).
3. Chạy `.\run.ps1 -SkipInstall` để tránh cài lại gói CPU.
4. Bật `Sử dụng GPU nếu khả dụng` trước khi kiểm tra dữ liệu.

Nếu phiên bản CUDA/Paddle không khớp, bỏ chọn GPU để trở về CPU. Không bật checkbox GPU khi mới chỉ cài bản CPU.

## Cách tìm và đối chiếu

Thứ tự xử lý là tên file → PDF text → OCR. PDF chỉ được OCR khi text layer không tìm thấy mã Excel. OCR lần lượt dùng ảnh gốc, ảnh xám/tăng tương phản, adaptive threshold, làm nét và tách chữ đỏ HSV. Ảnh được sửa EXIF và thử các hướng 0/90/180/270 độ. Trang PDF được render trong RAM, không để ảnh tạm.

Các sửa nhầm `O/Q→0`, `I/L→1`, `Z→2`, `S→5`, `B→8`, `G→6`, chiều ngược lại ở hai vị trí chữ, và `D→Đ` chỉ được sinh theo đúng vị trí. Mã chỉ được chấp nhận khi tồn tại chính xác trong tập Excel. Nhiều phương án cùng hợp lệ được ghi `CẦN KIỂM TRA` và không sao chép.

Checkbox `Kiểm tra nội dung dù tên file đã có GCN` có nghĩa tên file không đủ để xác nhận; tài liệu phải khớp qua PDF text/OCR. Checkbox `Tìm tất cả GCN` khiến một nguồn chứa nhiều GCN được sao chép thành một bản cho từng GCN.

## Báo cáo, cache và log

Thư mục kết quả có:

- Các bản sao đặt tên chuẩn; nguồn không bị sửa, xóa hoặc di chuyển.
- `KET_QUA_TIM_GCN.xlsx` gồm `KET_QUA`, `FILE_DA_QUET`, `TONG_HOP`.
- `.gcn_scan_cache.json`, dùng đường dẫn tuyệt đối + kích thước + mtime + cấu hình/tập GCN để tránh OCR lại file không đổi.
- File log `gcn_file_finder.log` ở thư mục làm việc khi khởi chạy.

Báo cáo tạm và cache được lưu sau mỗi 10 file. Khi dừng, chương trình không nhận file mới, chờ task đang chạy xong rồi lưu cache/báo cáo. Không tự sửa `.gcn_scan_cache.json` khi chương trình đang chạy.

## Kiểm thử

```powershell
.\.venv\Scripts\python.exe -m pytest .\tests -q
```

Test bao phủ chuẩn hóa, Unicode `Đ/Ð`, biên chuỗi, CCCD, sửa OCR có kiểm soát, đối chiếu Excel, tên có dấu cách, hậu tố `_02/_03`, không ghi đè, hash chống lặp, nhiều GCN trong một nguồn và nhiều trang của cùng GCN.

### Kiểm tra thủ công với Excel/PDF mẫu

Tạo workbook có sheet `Danh sách`, tiêu đề cột `GCN`, các dòng `AM 143443`, `CS 012345`, `ĐĐ 001234`. Tạo PDF có text `Số phát hành: AM-143443` hoặc ảnh quét cùng nội dung. Sau khi chạy, thư mục kết quả phải có `AM 143443.pdf`; báo cáo ghi đúng dòng Excel, trang, phương thức và đường dẫn nguồn/kết quả. Mã không có trong Excel không được sao chép.

## Build EXE Windows

Phải tải đủ model trước, sau đó:

```powershell
.\build_exe.ps1
```

Script tạo môi trường riêng `.venv-build`, cài dependency, chạy pytest và chỉ build khi test đạt. PyInstaller dùng `--onedir` vì PaddleOCR ổn định hơn dạng one-file, gom CustomTkinter/PaddleOCR/Paddle/OpenCV/PyMuPDF và đóng gói model offline.

EXE sau build:

```text
dist\GCNFileFinder\GCNFileFinder.exe
```

Phải phân phối **toàn bộ** thư mục `GCNFileFinder`, không chỉ riêng file `.exe`. Nếu muốn đổi cấu hình, sao chép `config.example.json` thành `config.json` cạnh EXE rồi chỉnh đường dẫn model tương đối hoặc tuyệt đối.

## Xử lý sự cố

- **Thiếu model:** chạy `.\download_models.ps1` khi có Internet; kiểm tra cả `det`, `rec`, `cls` không rỗng.
- **Không đọc Excel:** đóng Excel nếu file đang khóa, kiểm tra file không hỏng và sheet/cột đúng.
- **PDF có mật khẩu:** tạo bản PDF đã mở khóa hợp pháp rồi quét bản đó; ứng dụng không phá mật khẩu.
- **Hết RAM/PDF rất lớn:** chọn 300 DPI, một worker; chia PDF nếu cần.
- **GPU lỗi:** bỏ chọn GPU hoặc cài lại bản PaddlePaddle GPU đúng CUDA.
- **Đường dẫn quá dài/quyền:** dùng thư mục ngắn hơn và thư mục người dùng có quyền đọc/ghi.
- **Ảnh hỏng:** lỗi được ghi theo file; các file khác vẫn tiếp tục.

## Giới hạn OCR

Độ chính xác phụ thuộc DPI, độ nghiêng, nhòe, nén JPEG, nền hoa văn, dấu mộc che chữ và chất lượng chữ đỏ. Bộ sửa OCR cố ý bảo thủ: nó không fuzzy-match và không đoán mã gần nhất, nên có thể bỏ sót để tránh sao chép nhầm. Trường hợp `CẦN KIỂM TRA` phải được con người đối chiếu với tài liệu gốc và Excel. OCR chữ viết tay, PDF mã hóa đặc biệt và ảnh mất nét nặng không được bảo đảm.
