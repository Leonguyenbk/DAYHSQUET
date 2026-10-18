# tach-gcn

`tach-gcn` là công cụ dòng lệnh Python quét đệ quy một thư mục, nhận diện từng
trang PDF là **GIẤY CHỨNG NHẬN / CHỨNG CHỈ / CERTIFICATE**, rồi tách các trang
đó sang thư mục kết quả mà vẫn giữ nguyên cây thư mục nguồn. Giao diện và báo
cáo đều bằng tiếng Việt; chạy trên Windows, macOS và Linux.

## Yêu cầu và cài đặt

- Python 3.10 trở lên.
- Không cần Tesseract.
- Gemini là tùy chọn; nếu không dùng, công cụ vẫn nhận diện bằng lớp chữ PDF.

Từ thư mục chứa dự án:

```powershell
python -m pip install .
```

Lệnh cài đặt trên cũng cài Flet Desktop để dùng giao diện đồ hoạ.

Để chạy kiểm thử:

```powershell
python -m pip install ".[test]"
pytest
```

## Thiết lập Gemini

Tạo API key tại [Google AI Studio](https://aistudio.google.com/app/apikey). Trên
Windows, lưu key bằng:

```powershell
setx GEMINI_API_KEY "API_KEY_CUA_BAN"
```

Đóng rồi mở lại PowerShell sau khi chạy `setx`. Cho phiên PowerShell hiện tại,
có thể dùng `$env:GEMINI_API_KEY="API_KEY_CUA_BAN"`. Trên macOS/Linux:

```bash
export GEMINI_API_KEY="API_KEY_CUA_BAN"
```

Model mặc định là `gemini-3.5-flash-lite`; có thể đặt `GEMINI_MODEL` hoặc dùng
`--model`. Nếu thiếu API key, chương trình tự chuyển sang chế độ chỉ dùng từ
khóa và đánh dấu trang scan là `chưa phân tích`.

> **Quyền riêng tư:** những trang cần Gemini được render thành ảnh JPEG và gửi
> lên máy chủ Google để phân tích. Khi lấy thông tin GCN, toàn bộ các mặt/trang
> đã gom của GCN được gửi chung để xác định đúng chủ hiện tại. PDF có thể chứa dữ liệu cá nhân hoặc dữ liệu
> nhạy cảm; chỉ bật Gemini khi bạn có quyền gửi dữ liệu đó và chấp nhận điều
> khoản xử lý dữ liệu của Google. Dùng `--no-gemini` để không gửi dữ liệu.

## Ví dụ

### Giao diện Flet

Mở giao diện đồ hoạ:

```powershell
tach-gcn-gui
```

Trong cửa sổ, chọn thư mục đầu vào/đầu ra và nhập Gemini API key vào ô mật
khẩu. Key nhập tại GUI chỉ tồn tại trong bộ nhớ của phiên chạy, không được ghi
vào cache hay `bao-cao.csv`; có thể bấm biểu tượng con mắt để kiểm tra key. Mọi
tuỳ chọn quan trọng của CLI đều có trên GUI, gồm model, batch, RPM, worker,
dry-run, `--rest`, `--gemini-all`, low-res và cache.

### Dòng lệnh

Quét cơ bản, kết quả nằm cạnh thư mục nguồn với hậu tố
`_giay-chung-nhan`:

```powershell
tach-gcn "D:\HoSo\PDF"
```

Chọn thư mục đầu ra, thêm từ khóa, xuất mỗi trang thành một file và giữ cả phần
còn lại:

```powershell
tach-gcn "D:\HoSo\PDF" -o "D:\KetQua" -k "XÁC NHẬN HOÀN THÀNH" --mode each --rest
```

Chỉ xem báo cáo, hoàn toàn không gửi dữ liệu lên Gemini:

```powershell
tach-gcn "D:\HoSo\PDF" --dry-run --no-gemini
```

Gửi mọi trang cho Gemini, sáu trang mỗi request, giới hạn 20 request/phút và
giảm độ phân giải media để tiết kiệm token:

```powershell
tach-gcn "D:\HoSo\PDF" --gemini-all --batch 6 --rpm 20 --low-res
```

Các tùy chọn chính:

- `--mode perfile` (mặc định): gộp các mặt/trang cùng GCN thành
  `<số-giấy-chứng-nhận>-GCN.pdf`, ví dụ `AA 123456-GCN.pdf`. Tên này lấy từ
  số phát hành/mã phôi, không lấy “Số vào sổ cấp GCN”. Nếu không đọc được mã
  phôi, công cụ dùng `<tên-gốc>-GCN.pdf`.
- `--mode each`: tạo một file riêng cho từng trang GCN, ưu tiên mã phôi.
- `--rest`: tạo thêm `<tên-gốc>-GT.pdf` cho các trang còn lại. GUI bật sẵn
  tuỳ chọn này.
- `--workers 4`: số file đọc song song bằng process.
- `--no-cache`: bỏ qua `.tach-gcn-cache.json`.
- `--no-gemini`: chỉ dùng từ khóa.
- `--dry-run`: không tách PDF, nhưng vẫn ghi báo cáo (và vẫn có thể gọi Gemini
  nếu không thêm `--no-gemini`).

Chạy `tach-gcn --help` để xem đầy đủ.

## Nhận diện và đầu ra

Mỗi trang được chuẩn hóa (bỏ dấu tiếng Việt, viết hoa, bỏ dấu câu), chấm điểm
từ khóa, rồi mới gửi Gemini nếu là ảnh scan, có điểm mập mờ từ 1–4, hoặc khi
bật `--gemini-all`. Kết luận Gemini ghi đè kết luận từ khóa. Một batch Gemini
lỗi 429/5xx được thử lại tối đa sáu lần; nếu vẫn lỗi, chương trình giữ kết quả
từ khóa và ghi lỗi vào báo cáo.

`bao-cao.csv` dùng UTF-8 có BOM để Excel đọc đúng tiếng Việt, gồm:

- file và số trang;
- kết luận `Có` nếu là giấy chứng nhận;
- cách nhận diện (`từ khoá`, `Gemini`, `ảnh scan`);
- điểm, từ khóa tìm thấy;
- ghi chú hoặc tiêu đề/lý do của Gemini.

`thong-tin-gcn.csv` có một dòng cho mỗi GCN đã được gom, gồm file nguồn, file
GCN đầu ra, các trang, **số giấy chứng nhận/số phát hành/mã phôi**, **số vào sổ
cấp GCN**, **ngày cấp** và **tên chủ sử dụng hiện tại**. Đây là hai loại số khác
nhau: chỉ mã phôi (thường dạng `AA 123456`) được dùng đặt tên PDF; số vào sổ chỉ
được ghi báo cáo. Công cụ không dùng số thửa, số tờ bản đồ hay mã in ở trang sơ
đồ. Khi có Gemini, toàn bộ các mặt của cùng GCN được đọc chung; thông tin chuyển
nhượng, tặng cho, thừa kế hoặc đính chính mới nhất ở trang 4 được ưu tiên hơn
tên chủ ban đầu.

Các trang mặt trong như “Sơ đồ thửa đất” và “Những thay đổi sau khi cấp GCN”
được ghép với mặt đầu thành một file GCN. Nếu hai mặt được nhận diện cách nhau
tối đa ba trang, những trang nằm giữa cũng được giữ trong cùng file, tránh tách
một GCN thành nhiều PDF.

File hỏng và file có mật khẩu được ghi thành một dòng lỗi, các file khác vẫn
tiếp tục. Thư mục kết quả nằm trong thư mục nguồn sẽ tự được bỏ qua khi quét.
Nếu PDF bị ghép sai cây trang, công cụ cố phục hồi các object `/Page` còn đọc
được, tiếp tục quét/tách chúng, tự tạo `<tên-gốc>-DA-SUA.pdf` trong thư mục kết
quả và ghi tỷ lệ trang phục hồi trong báo cáo. Chế độ `--dry-run` không ghi bản
PDF đã sửa.

## Cache và chi phí

Kết quả Gemini được cache theo SHA-256 nội dung file, số trang và model tại
`.tach-gcn-cache.json`. Vì vậy chạy lại cùng dữ liệu không gửi lại những trang
đã phân tích. `--no-cache` chỉ bỏ qua cache, không xóa file cache hiện có.

Cuối lượt, công cụ in số token thực tế do API báo và chi phí USD ước tính. Đơn
giá là hằng số tham chiếu trong `gemini.py`, tra ngày 18/10/2026; giá Gemini có
thể thay đổi, và tài khoản/free tier của bạn có thể có cách tính khác. Hãy kiểm
tra [bảng giá Gemini API](https://ai.google.dev/gemini-api/docs/pricing) trước
khi xử lý số lượng lớn.
