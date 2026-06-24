import requests
from difflib import get_close_matches
import unicodedata

nguongoc ={
  "1": "Nhà nước công nhận quyền sử dụng đất",
  "2": "Công nhận QSDĐ như giao đất có thu tiền sử dụng đất",
  "3": "Công nhận QSDĐ như giao đất không thu tiền sử dụng đất",
  "4": "Nhà nước giao đất có thu tiền sử dụng đất",
  "5": "Nhà nước giao đất không thu tiền sử dụng đất",
  "6": "Nhà nước giao đất để quản lý",
  "7": "Thuê đất của doanh nghiệp đầu tư hạ tầng khu công nghiệp, khu kinh tế, khu công nghệ cao",
  "8": "Thuê đất trả tiền hàng năm của chủ đầu tư xây dựng kinh doanh kết cấu hạ tầng khu công nghiệp, khu kinh tế, khu công nghệ cao",
  "9": "Thuê đất trả tiền một lần của chủ đầu tư hạ tầng khu công nghiệp, khu kinh tế, khu công nghệ cao",
  "10": "Nhà nước cho thuê đất trả tiền hàng năm",
  "11": "Nhà nước cho thuê đất trả tiền một lần",
  "12": "Nhà nước công nhận quyền sử dụng đất như Nhà nước giao đất có thu tiền sử dụng đất",
  "13": "Nhà nước công nhận quyền sử dụng đất như Nhà nước giao đất không thu tiền sử dụng đất",
  "14": "Nhà nước công nhận quyền sử dụng đất như Nhà nước cho thuê đất trả tiền hàng năm",
  "15": "Nhà nước công nhận quyền sử dụng đất như Nhà nước cho thuê đất trả tiền một lần",
  "16": "Sở hữu căn hộ nhà chung cư"
}

muc_dich_su_dung = {
  "BCS": "Đất bằng chưa sử dụng",
  "BHK": "Đất bằng trồng cây hàng năm khác",
  "CAN": "Đất an ninh",
  "CCC": "Đất sử dụng vào mục đích công cộng",
  "CDG": "Đất chuyên dùng",
  "CGT": "Đất do Nhà nước thu hồi theo quy định của pháp luật đất đai chưa giao, chưa cho thuê",
  "CHN": "Đất trồng cây hàng năm",
  "CLN": "Đất trồng cây lâu năm",
  "CNC": "Đất khu công nghệ cao",
  "CNT": "Đất chăn nuôi tập trung",
  "COC": "Đất cỏ dùng vào chăn nuôi",
  "CQA": "Đất quốc phòng, an ninh",
  "CQP": "Đất quốc phòng",
  "CSD": "Đất chưa sử dụng",
  "CSK": "Đất sản xuất, kinh doanh phi nông nghiệp",
  "CTS": "Đất trụ sở cơ quan, công trình sự nghiệp",
  "DBV": "Đất công trình hạ tầng bưu chính, viễn thông, công nghệ thông tin",
  "DCH": "Đất chợ dân sinh, chợ đầu mối",
  "DCK": "Đất công trình công cộng khác",
  "DCS": "Đất đồi núi chưa sử dụng",
  "DCT": "Đất công trình cấp nước, thoát nước",
  "DDD": "Đất có di tích lịch sử - văn hóa, danh lam thắng cảnh, di sản thiên nhiên",
  "DDL": "Đất danh lam thắng cảnh",
  "DDT": "Đất có di tích",
  "DGD": "Đất xây dựng cơ sở giáo dục và đào tạo",
  "DGT": "Đất công trình giao thông",
  "DKH": "Đất xây dựng cơ sở khoa học và công nghệ",
  "DKT": "Đất xây dựng cơ sở khí tượng thủy văn",
  "DKV": "Đất khu vui chơi, giải trí công cộng, sinh hoạt cộng đồng",
  "DMT": "Đất xây dựng cơ sở môi trường",
  "DNG": "Đất xây dựng cơ sở ngoại giao",
  "DNL": "Đất công trình năng lượng, chiếu sáng công cộng",
  "DNT": "Khu dân cư nông thôn",
  "DPC": "Đất công trình phòng, chống thiên tai",
  "DRA": "Đất công trình xử lý chất thải",
  "DSH": "Đất sinh hoạt cộng đồng",
  "DSK": "Đất xây dựng công trình sự nghiệp khác",
  "DSN": "Đất xây dựng công trình sự nghiệp",
  "DTL": "Đất công trình thủy lợi",
  "DTS": "Đất xây dựng trụ sở của tổ chức sự nghiệp",
  "DTT": "Đất xây dựng cơ sở thể dục, thể thao",
  "DVH": "Đất xây dựng cơ sở văn hóa",
  "DXH": "Đất xây dựng cơ sở xã hội",
  "DYT": "Đất xây dựng cơ sở y tế",
  "HNK": "Đất trồng cây hằng năm khác",
  "KBT": "Đất khu bảo tồn",
  "KDT": "Đất đô thị",
  "KKT": "Đất khu kinh tế",
  "KĐD": "Đất cơ sở bảo tồn đa dạng sinh học",
  "LMU": "Đất làm muối",
  "LNC": "Đất trồng cây công nghiệp lâu năm",
  "LNK": "Đất trồng cây lâu năm khác",
  "LNP": "Đất lâm nghiệp",
  "LNQ": "Đất trồng cây ăn quả lâu năm",
  "LUA": "Đất trồng lúa",
  "LUC": "Đất chuyên trồng lúa",
  "LUK": "Đất trồng lúa còn lại",
  "LUN": "Đất trồng lúa nương",
  "MCS": "Đất có mặt nước chưa sử dụng",
  "MNC": "Đất có mặt nước chuyên dùng dạng ao, hồ, đầm, phá",
  "MVB": "Đất có mặt nước ven biển",
  "MVK": "Đất mặt nước ven biển có mục đích khác",
  "MVR": "Đất mặt nước ven biển có rừng ngập mặn",
  "MVT": "Đất mặt nước ven biển nuôi trồng thủy sản",
  "NCS": "Đất núi đá không có rừng cây",
  "NHK": "Đất nương rẫy trồng cây hàng năm khác",
  "NKH": "Đất nông nghiệp khác",
  "NNP": "Đất nông nghiệp",
  "NTD": "Đất nghĩa trang, nhà tang lễ, cơ sở hỏa táng; đất cơ sở lưu trữ tro cốt",
  "NTS": "Đất nuôi trồng thủy sản",
  "ODT": "Đất ở tại đô thị",
  "ONT": "Đất ở tại nông thôn",
  "OTC": "Đất ở",
  "PNK": "Đất phi nông nghiệp khác",
  "PNN": "Đất phi nông nghiệp",
  "RDD": "Đất rừng đặc dụng",
  "RDK": "Đất khoanh nuôi phục hồi rừng đặc dụng",
  "RDM": "Đất trồng rừng đặc dụng",
  "RDN": "Đất có rừng tự nhiên đặc dụng",
  "RDT": "Đất có rừng trồng đặc dụng",
  "RPH": "Đất rừng phòng hộ",
  "RPK": "Đất khoanh nuôi phục hồi rừng phòng hộ",
  "RPM": "Đất trồng rừng phòng hộ",
  "RPN": "Đất có rừng tự nhiên phòng hộ",
  "RPT": "Đất có rừng trồng phòng hộ",
  "RSK": "Đất khoanh nuôi phục hồi rừng sản xuất",
  "RSM": "Đất trồng rừng sản xuất",
  "RSN": "Đất có rừng tự nhiên sản xuất",
  "RST": "Đất có rừng trồng sản xuất",
  "RSX": "Đất rừng sản xuất",
  "SCC": "Đất khu công nghiệp, cụm công nghiệp",
  "SCT": "Đất khu công nghệ thông tin tập trung",
  "SKC": "Đất cơ sở sản xuất phi nông nghiệp",
  "SKK": "Đất khu công nghiệp",
  "SKN": "Đất cụm công nghiệp",
  "SKS": "Đất sử dụng cho hoạt động khoáng sản",
  "SKT": "Đất khu chế xuất",
  "SKX": "Đất sản xuất vật liệu xây dựng, làm đồ gốm",
  "SMN": "Đất sông suối và mặt nước chuyên dùng",
  "SON": "Đất có mặt nước dạng sông, ngòi, kênh, rạch, suối",
  "SXN": "Đất sản xuất nông nghiệp",
  "TIN": "Đất tín ngưỡng",
  "TMD": "Đất thương mại, dịch vụ",
  "TON": "Đất tôn giáo",
  "TSC": "Đất xây dựng trụ sở cơ quan",
  "TSK": "Đất trụ sở khác",
  "TSL": "Đất nuôi trồng thủy sản nước lợ, mặn",
  "TSN": "Đất nuôi trồng thủy sản nước ngọt",
  "TTN": "Đất tôn giáo, tín ngưỡng",
  "TVC": "Đất có mặt nước chuyên dùng"
}

def normalize_text(s):
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d")
    s = " ".join(s.split())
    return s

def resolve_nguon_goc_id(value):
    key = normalize_text(value)

    # nhập trực tiếp mã ID
    if key.isdigit() and key in NGUON_GOC_SU_DUNG_DAT:
        return key

    # nhóm dễ nhầm nhất: giao đất để quản lý
    if "quan ly" in key:
        if "giao" in key or "nha nuoc" in key:
            return "6"

    # thuê đất
    if "thue" in key:
        if "mot lan" in key or "1 lan" in key:
            return "11"
        if "hang nam" in key or "nam" in key:
            return "10"

    # giao đất
    if "giao" in key:
        if "khong thu" in key:
            return "5"
        if "co thu" in key:
            return "4"
        if "quan ly" in key:
            return "6"

    # công nhận
    if "cong nhan" in key:
        if "khong thu" in key:
            return "3"
        if "co thu" in key:
            return "2"
        if "thue" in key and ("hang nam" in key or "nam" in key):
            return "14"
        if "thue" in key and ("mot lan" in key or "1 lan" in key):
            return "15"
        return "1"

    # chung cư
    if "chung cu" in key or "can ho" in key:
        return "16"

    # fallback so khớp gần đúng
    match = get_close_matches(key, NGUON_GOC_ALIAS.keys(), n=1, cutoff=0.6)
    if match:
        return NGUON_GOC_ALIAS[match[0]]

    raise ValueError(f"Không nhận diện được nguồn gốc sử dụng đất: {value}")

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