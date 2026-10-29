import tach_gcn_theo_don as tg
from tach_gcn.detect import extract_certificate_serial

A4 = (60, 84)


def _gan(n, **trang):
    gan = {i: {"gcn": "", "serial_seen": "", "reason": "", "bat_dau": False} for i in range(n)}
    for i, v in trang.items():
        gan[int(i[1:])].update(v)
    return gan


def test_la_so_a4():
    assert tg.la_so_a4("AA 01234567") and tg.la_so_a4("aa01234567")
    assert not tg.la_so_a4("AA 123456") and not tg.la_so_a4("AB 01234567")


def test_gcn_a4_ghep_mat_sau_trong_file_toan_a4():
    gan = _gan(4, t1={"gcn": "AA 01234567", "serial_seen": "AA 01234567"})
    trang, ghi_chu = tg.bo_sung_mat([1], gan, [A4] * 4, "AA 01234567")
    assert trang == [1, 2] and "A4" in ghi_chu


def test_gcn_mau_cu_file_toan_a4_khong_ghep():
    gan = _gan(4, t1={"gcn": "AA 123456", "serial_seen": "AA 123456"})
    assert tg.bo_sung_mat([1], gan, [A4] * 4, "AA 123456")[0] == [1]


def test_doc_so_phat_hanh_8_chu_so():
    assert extract_certificate_serial("GIẤY CHỨNG NHẬN\nSố phát hành GCN: AA 01234567") == "AA 01234567"
