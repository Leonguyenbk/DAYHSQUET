import requests

BASE_URL = "https://dla.mplis.gov.vn"

token = "KD9tHn0G_ByF9RY5cqtSThhACrla1tM7sIaTyDTGRVBDWeU5c_bsoi3pvKLZTNuFXk44wAhz-vBJzMtSkJK1O_f2dJSo6Byqh30fkb4HVxx9nTQe0"
cookie_raw = "ASP.NET_SessionId=pslpkrsc5xdipzxtvi0mcsxd; __RequestVerificationTokenDC=y4jj0flrwcB_WqUrFFWuIhvZlW92OTHpIcbb3DazLTWvpD0RgY6iz-ie0PZKLduRyc-6wseJPbUG80Hw4aapzM1_X5w1; _Vbdlis.DC.Cookie=I1ndzLo3rmiiIm-2hqav6WEV4DNPFEJPR_VAoiyHLREUl1OJcn0BRs5vC93fHeKj4nM55lkZ16YzC2XOPyT1mbeyODYLHk_ZxSmev_ubH0N8fanqIKY2qdxCfC1UveI-5Y463mNWTA_s9EvZRiQamIiaoioUzuZFxNxCkcDDOhP3YOjM9NezLCsCxSE5ChRjyweok8_BnrTyTwhq-axI2I3F4rnslF1iDTLlh_DC5JH4hxNo-kv0hYzJKVDJ8xD-D62gv1aw9ndTlhHBLlOpgcdJpEiO8sGqv3s8ARsJrEl8nJk5P48uj7J-L5gQKkgu1_bWFahGBmER3D83-x1Qi1z82XOdrwT9It2zUSdxjZXZA8Tr3nEGWP99ZjrTeVa6hOOvCDa25fDeadl3tD2GErwsOHDxpZgRJ4rMDwWScP7I2IxcDdgVd5YkElNw_V6XlQXBvA"

s = requests.Session()

s.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Content-Type": "application/json; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
    "Referer": f"https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2",
    "__requestverificationtoken": token,
    "__RequestVerificationToken": token,
    "RequestVerificationToken": token,
    "Cookie": cookie_raw,
})

payload =  {
        "soThuTuThua": 4000,
        "soHieuToBanDo": 2,
        "dienTich": 621.3,
        "dienTichPhapLy": "621.3",
        "soThuTuThuaCu": "",
        "soHieuToBanDoCu": "",
        "loaiThuaDat": "",
        "quaTrinhSuDung": "",
        "inSoLieuCu": False,
        "lichSuHinhThanh": "",
        "noiDungQuyHoach": "",
        "ghiChuDienTich": "",
        "ListMucDichSuDung": [
            {
            "loaiMucDichSuDungId": "SON",
            "dienTich": 621.3,
            "soThuTu": 1,
            "ngayHinhThanh": None,
            "ngaySuDung": None,
            "loaiMucDichSuDungQuyHoachId": "0",
            "mucDichSuDungChiTiet": "",
            "thoiHanSuDung": "",
            "loaiMucDichSuDungPhuId": "0",
            "ghiChu": "",
            "ListNguonGocSuDungDat": [
                {
                "loaiNguonGocSuDungDatId": "6",
                "loaiNguonGocChuyenQuyenId": "0",
                "dienTich": 621.3,
                "chiTiet": ""
                }
            ],
            "LoaiMucDichSuDung": {
                "loaiMucDichSuDungId": "SON",
                "kyHieuLoaiMucDichSuDung": None,
                "tenLoaiMucDichSuDung": "Đất có mặt nước dạng sông, ngòi, kênh, rạch, suối",
                "moTaLoaiMucDichSuDung": "SON (Đất có mặt nước dạng sông, ngòi, kênh, rạch, suối)",
                "trangThai": True
            },
            "LoaiMucDichSuDungPhu": None
            }
        ],
        "ListDiaChi": [
            {
            "tinhId": "66",
            "huyenId": "0",
            "xaId": "24175",
            "duongId": "",
            "ngoPho": "",
            "soNha": "",
            "toDanPhoId": "",
            "laDiaChiChinh": False,
            "laDiaChiCu": False,
            "diaChiChiTiet": "Xã Hòa Phú, Tỉnh Đắk Lắk"
            }
        ],
        "TaiLieuDoDac": None,
        "diaChi": "Xã Hòa Phú, Tỉnh Đắk Lắk",
        "tinhId": 66,
        "huyenId": 0,
        "xaId": "24175",
        "duongDanSoDo": None,
        "tenFileSoDo": None
        }

r = s.post(
    "https://dla.mplis.gov.vn/dc/TaiSanAjax/AddThuaDat",
    json=payload,
    timeout=120,
    allow_redirects=False
)

print("STATUS CODE:", r.status_code)
print("REASON:", r.reason)
print("URL:", r.url)
print("CONTENT-TYPE:", r.headers.get("Content-Type"))
print("CONTENT-LENGTH:", r.headers.get("Content-Length"))
print("LOCATION:", r.headers.get("Location"))
print("HEADERS:", dict(r.headers))
print("TEXT repr:", repr(r.text))
print("CONTENT bytes:", repr(r.content))