import json
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

BASE_URL = "https://dla.mplis.gov.vn/dc/"
REFERER_URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
API_URL = "https://dla.mplis.gov.vn/dc/DangKyAjax/AdvancedSearchTinhHinhDangKy"


def lay_token_tu_trang(driver):
    js = """
    return (
        document.querySelector('input[name="__RequestVerificationToken"]')?.value ||
        document.querySelector('input[name="__requestverificationtoken"]')?.value ||
        document.querySelector('meta[name="__RequestVerificationToken"]')?.content ||
        document.querySelector('meta[name="__requestverificationtoken"]')?.content ||
        document.querySelector('meta[name="RequestVerificationToken"]')?.content ||
        ''
    );
    """
    return driver.execute_script(js)


def tao_session_tu_selenium(driver):
    session = requests.Session()

    user_agent = driver.execute_script("return navigator.userAgent;")
    token = lay_token_tu_trang(driver)

    print("TOKEN:", token[:30] + "..." if token else "KHÔNG LẤY ĐƯỢC TOKEN")

    session.headers.update({
        "User-Agent": user_agent,
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://dla.mplis.gov.vn",
        "Referer": REFERER_URL,
        "__requestverificationtoken": token,
    })

    for c in driver.get_cookies():
        session.cookies.set(
            name=c["name"],
            value=c["value"],
            domain=c.get("domain"),
            path=c.get("path", "/")
        )

    return session


def tao_payload(xa_id, so_to, so_thua):
    return {
        "draw": "2",

        "columns[0][data]": "",
        "columns[0][name]": "",
        "columns[0][searchable]": "true",
        "columns[0][orderable]": "false",
        "columns[0][search][value]": "",
        "columns[0][search][regex]": "false",

        "columns[1][data]": "tinhHinhDangKyId",
        "columns[1][name]": "tinhHinhDangKyId",
        "columns[1][searchable]": "true",
        "columns[1][orderable]": "true",
        "columns[1][search][value]": "",
        "columns[1][search][regex]": "false",

        "columns[2][data]": "maDon",
        "columns[2][name]": "maDon",
        "columns[2][searchable]": "true",
        "columns[2][orderable]": "true",
        "columns[2][search][value]": "",
        "columns[2][search][regex]": "false",

        "columns[3][data]": "soThuTu",
        "columns[3][name]": "soThuTu",
        "columns[3][searchable]": "true",
        "columns[3][orderable]": "true",
        "columns[3][search][value]": "",
        "columns[3][search][regex]": "false",

        "columns[4][data]": "DaiDienKhaiTrinh",
        "columns[4][name]": "DaiDienKhaiTrinh",
        "columns[4][searchable]": "true",
        "columns[4][orderable]": "false",
        "columns[4][search][value]": "",
        "columns[4][search][regex]": "false",

        "columns[5][data]": "ngayTiepNhan",
        "columns[5][name]": "ngayTiepNhan",
        "columns[5][searchable]": "true",
        "columns[5][orderable]": "true",
        "columns[5][search][value]": "",
        "columns[5][search][regex]": "false",

        "columns[6][data]": "thoiDiemDangKy",
        "columns[6][name]": "thoiDiemDangKy",
        "columns[6][searchable]": "true",
        "columns[6][orderable]": "true",
        "columns[6][search][value]": "",
        "columns[6][search][regex]": "false",

        "order[0][column]": "5",
        "order[0][dir]": "desc",
        "start": "0",
        "length": "10",
        "search[value]": "",
        "search[regex]": "false",

        "model[xaId]": str(xa_id),
        "model[huyenId]": "",
        "model[tinhHinhDangKyId]": "",
        "model[maDon]": "",
        "model[soThuTu]": "",
        "model[ngayTiepNhan]": "",
        "model[thoiDiemDangKy]": "",
        "model[loaiGiayChungNhanId]": "",
        "model[soPhatHanh]": "",
        "model[maVach]": "",
        "model[soVaoSo]": "",
        "model[soVaoSoCu]": "",
        "model[ngayVaoSo]": "",
        "model[soHoSoGoc]": "",
        "model[soHoSoGocCu]": "",
        "model[hoTen]": "",
        "model[soGiayTo]": "",
        "model[namSinh]": "",
        "model[soThuTuThua]": str(so_thua),
        "model[soHieuToBanDo]": str(so_to),
        "model[soThuTuThuaCu]": "",
        "model[soHieuToBanDoCu]": "",
        "model[soNha]": "",
        "model[diaChiChiTiet]": "",
        "model[dieuKienCapGiay]": "",
        "model[phucHoiDuLieu]": "false",
    }


def tim_to_thua(session, xa_id, so_to, so_thua):
    payload = tao_payload(xa_id, so_to, so_thua)
    res = session.post(API_URL, data=payload, timeout=30)

    print("Status:", res.status_code)
    print("Content-Type:", res.headers.get("content-type"))

    try:
        return res.json()
    except Exception:
        print("Không phải JSON, response:")
        print(res.text[:1000])
        return None


options = Options()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)
driver.get(REFERER_URL)

input("Đăng nhập MPLIS xong, mở đúng màn hình kê khai đăng ký rồi nhấn ENTER...")

session = tao_session_tu_selenium(driver)

result = tim_to_thua(
    session=session,
    xa_id=24133,
    so_to=364,
    so_thua=158
)

print(json.dumps(result, ensure_ascii=False, indent=2))

if result and result.get("success") is False:
    print("Lỗi server:", result.get("Error"))

if result and result.get("recordsTotal", 0) > 0:
    item = result["data"][0]
    print("Tìm thấy ID:", item.get("tinhHinhDangKyId"))
else:
    print("Không tìm thấy bản ghi.")