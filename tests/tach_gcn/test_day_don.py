import day_don_theo_thu_muc as dd


def _detail(hsq=None, gcn=None):
    return {
        "TinhHinhDangKy": {"tinhHinhDangKyId": 13868776, "coQuyenQuanLy": False, "dieuKienCapGiay": 0,
                           "daiDienKhaiTrinhId": 15424884, "CreatedDate": "/Date(1762708922142)/", "ghiChu": None,
                           "maDotChuyenDoi": None},
        "ChuSoHuu": {"VoChongs": [{"voChongId": 2874082}]},
        "TaiSan": {"ThuaDats": [{"thuaDatId": 2960871}]},
        "ListDangKyQuyen": [{"dangKyQuyenSuDungId": 18892394, "tinhHinhDangKyId": 13868776, "loaiDoiTuong": 1,
                             "chuSuDungId": 2874082, "caNhanId": None, "typeItem": 6, "itemId": 2960871,
                             "subItemId": "7090315", "giayChungNhanId": None}],
        "ListGiayChungNhan": gcn or [],
        "ListHoSoQuet": hsq or [],
    }


def _hsq(*nodes):
    return {"hoSoQuetId": 5, "ListFileHoSoQuet": {"ListFileHoSoQuet": list(nodes)}}


def test_nhan_file_don():
    assert dd.la_file_don("D:/x/01_Hồ sơ_tr2-4-DON.pdf")
    assert dd.la_file_don("D:/x/Đơn đăng ký_signed.pdf")
    assert not dd.la_file_don("D:/x/PC_NGUYEN VAN A.pdf")


def test_payload_theo_mau_responses():
    p = dd.build_payload_ttdk(_detail())
    th = p["TinhHinhDangKy"]
    assert th["coQuyenQuanLy"] is True and th["dieuKienCapGiay"] == "3"
    assert th["daiDienKhaiTrinhId"] == "15424884" and th["ghiChu"] == ""
    assert "maDotChuyenDoi" not in th and th["CreatedDate"].endswith("Z")
    [q] = p["ListDangKyQuyen"]
    assert q["duDieuKienCapGiay"] is False and q["subItemId"] == 7090315 and q["dangKyQuyenSuDungId"] == 0
    assert q["chuSuDungId"] == 2874082 and q["giayChungNhanId"] == 0
    assert set(p) >= {"ChuSoHuu", "TaiSan", "ListGiayChungNhan", "ListHanCheThuaLienKe", "GhiChuDongSuDungTrang1"}


def test_ke_hoach():
    # chưa có HSQ → tạo mới
    assert dd.lap_ke_hoach(_detail(), ["D:/x/a-DON.pdf"])[0]["cach"] == "d"
    # mô tả đúng bằng tên → thay; file tách trang không đè thành phần gốc → thêm mới
    d = _detail([_hsq({"moTa": "Đơn đăng ký", "loaiHoSoQuet": 0, "nodeId": "n1"},
                      {"moTa": "Hồ sơ cấp giấy", "loaiHoSoQuet": 0, "nodeId": "n2"})])
    kh = dd.lap_ke_hoach(d, ["D:/x/01_Đơn đăng ký-DON.pdf", "D:/x/02_Hồ sơ cấp giấy_tr2-4-DON.pdf"])
    assert [k["cach"] for k in kh] == ["b", "c"] and kh[0]["vi_tri"] == 0
    items = dd.items_hsq_co_san(d["ListHoSoQuet"][0], kh)
    assert len(items) == 3 and items[0][0]["loaiHoSoQuet"] == 2 and items[2][0]["moTa"] == "Đơn đăng ký"
    # đã có thành phần loại Đơn đăng ký → không đẩy
    d2 = _detail([_hsq({"moTa": "Đơn", "loaiHoSoQuet": 2, "nodeId": "n1"})])
    assert dd.lap_ke_hoach(d2, ["D:/x/a-DON.pdf"])[0]["cach"] == "co_san"


def test_bo_qua_don_co_gcn():
    class C:
        def lay_chi_tiet_don(self, madon):
            return _detail(gcn=[{"giayChungNhanId": 2112692, "soPhatHanh": None, "isLastest": True}])
    [r] = dd.xu_ly_don(C(), 13868776, ["D:/x/a-DON.pdf"], chay_thu=True)
    assert r["trang_thai"] == "Bỏ qua" and "2112692" in r["ghi_chu"]
