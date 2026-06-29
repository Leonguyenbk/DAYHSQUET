import lamsach_beta as module


import pytest


@pytest.mark.parametrize(
    ("cccd", "expected"),
    [
        ("066087123456", 1987),
        ("066001123456", 2001),
    ],
)
def test_derive_nam_sinh_from_cccd_uses_vietnam_rule(cccd, expected):
    assert module.derive_nam_sinh_from_cccd(cccd) == expected


def test_update_thua_dat_from_nhom1_falls_back_to_tai_san_thua_dats(monkeypatch, tmp_path):
    raw_nhom1 = {
        "thongTinDangKy": {
            "TinhHinhDangKy": {"tinhHinhDangKyId": 12869380},
        },
        "TaiSan": {
            "ThuaDats": [
                {
                    "thuaDatId": 1360947,
                    "version": 11,
                    "ListMucDichSuDung": [
                        {
                            "mucDichSuDungId": 5657183,
                            "loaiMucDichSuDungId": "ONT",
                            "ListNguonGocSuDungDat": [],
                        }
                    ],
                }
            ]
        },
    }

    captured = []

    def fake_api_get_thong_tin_dang_ky_nhom1(session, tinh_hinh_dang_ky_id):
        return {"ok": True, "raw": raw_nhom1}

    def fake_api_update_thua_dat(session, thua_dat):
        captured.append(thua_dat.get("thuaDatId"))
        return {"ok": True, "raw": {}}

    monkeypatch.setattr(module, "api_get_thong_tin_dang_ky_nhom1", fake_api_get_thong_tin_dang_ky_nhom1)
    monkeypatch.setattr(module, "api_update_thua_dat", fake_api_update_thua_dat)

    notes = []
    result = module.update_thua_dat_from_nhom1(
        session=object(),
        tinh_hinh_dang_ky_id=12869380,
        phan_loai_item={"thuaDatId": 1360947},
        debug_dir=str(tmp_path),
        row_excel=1,
        notes=notes,
    )

    assert result["ok"] is True
    assert captured == [1360947]
    assert notes
