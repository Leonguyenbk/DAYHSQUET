import tai_tach_hsq_don as tt

GCN = {"giayChungNhanId": 2112692, "soPhatHanh": None, "soVaoSo": "--\--", "isLastest": True}


def test_ghi_so_doc_duoc_cho_gcn_chua_so():
    rows = [{"loai": "GCN", "trang_thai": "Không khớp", "so_ai_doc": "AA 01234567",
             "file_nguon": "hsq.pdf", "trang": "3-4"}]
    [dong] = tt._dong_gcn_chua_so(13868776, [GCN], rows)
    assert dong["giaychungnhanid"] == 2112692 and dong["so_ai_doc"] == "AA 01234567"
    assert dong["trang_thai"] == "Chưa có số GCN" and "hsq.pdf tr.3-4" in dong["ghi_chu"]


def test_khong_co_trang_gcn():
    [dong] = tt._dong_gcn_chua_so(13868776, [GCN], [])
    assert dong["so_ai_doc"] == "" and "không có trang GCN" in dong["ghi_chu"]
    assert tt._dong_gcn_chua_so(1, [], []) == []
