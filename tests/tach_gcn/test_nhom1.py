import cap_nhat_don_hsq as core
import day_gcn_theo_thu_muc as dg
import day_hsq_theo_mota as dm
import nhom1_mplis as nhom1

DETAIL = {"TinhHinhDangKy": {"xaId": 24526, "tinhId": 66},
          "TaiSan": {"ThuaDats": [{"thuaDatId": 1, "soHieuToBanDo": 120, "soThuTuThua": 1807}]},
          "ListGiayChungNhan": [{"giayChungNhanId": 9, "soPhatHanh": "AA 123456", "isLastest": True}],
          "ListHoSoQuet": [{"hoSoQuetId": 5, "ListFileHoSoQuet": {"ListFileHoSoQuet": [{"moTa": "Đơn"}]}}]}


class Resp:
    def __init__(self, data):
        self.data = data


class Client:
    def __init__(self, ban_ghi=None, loi=None):
        self.ban_ghi, self.loi, self.goi = ban_ghi or [], loi, 0

    def lay_chi_tiet_don(self, madon):
        return DETAIL

    def _post_doc(self, url, **kw):
        self.goi += 1
        if self.loi:
            raise self.loi
        return Resp({"data": self.ban_ghi})

    def _response_json(self, resp, name):
        return resp.data


def test_kiem_tra_nhom_1():
    dat = [{"tinhHinhDangKyId": 7, "thongTinDangKyDapUngNhom1": True}]
    assert nhom1.kiem_tra(Client(dat), 7, DETAIL) == (True, dat)
    chua = dat + [{"tinhHinhDangKyIds": [7], "thongTinDangKyDapUngNhom1": False}]
    assert nhom1.kiem_tra(Client(chua), 7, DETAIL)[0] is False
    assert nhom1.kiem_tra(Client([]), 7, DETAIL)[0] is False  # không có bản ghi → chưa đạt
    # bản ghi của đơn khác không tính
    assert nhom1.kiem_tra(Client([{"tinhHinhDangKyId": 8, "thongTinDangKyDapUngNhom1": True}]), 7, DETAIL)[0] is False


def test_tra_loi_van_xu_ly_het_phien_thi_dung():
    assert nhom1.kiem_tra(Client(loi=RuntimeError("500")), 7, DETAIL) == (False, None)
    try:
        nhom1.kiem_tra(Client(loi=core.HetPhien("hết phiên")), 7, DETAIL)
    except core.HetPhien:
        pass
    else:
        raise AssertionError("phải ném HetPhien")


def test_tool_day_bo_qua_don_dat_nhom_1():
    c = Client([{"tinhHinhDangKyId": 7, "thongTinDangKyDapUngNhom1": True}])
    [r] = dg.xu_ly_don(c, 7, ["D:/x/AA 123456.pdf"], chay_thu=True)
    assert r["trang_thai"] == "Bỏ qua" and r["ghi_chu"] == nhom1.LY_DO_DA_DAT
    rows = dm.xu_ly_don(c, 7, ["D:/x/Đơn.pdf"], chay_thu=True)
    assert all(r["ghi_chu"] == nhom1.LY_DO_DA_DAT for r in rows)
    # tắt kiểm tra → không gọi tra cứu làm sạch
    c2 = Client([{"tinhHinhDangKyId": 7, "thongTinDangKyDapUngNhom1": True}])
    dm.xu_ly_don(c2, 7, ["D:/x/Đơn.pdf"], chay_thu=True, bo_qua_nhom_1=False)
    assert c2.goi == 0
