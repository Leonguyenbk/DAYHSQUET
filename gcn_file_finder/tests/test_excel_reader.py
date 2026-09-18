from openpyxl import Workbook

from gcn_file_finder.core.excel_reader import read_gcn_excel


def test_excel_reader_keeps_rows_invalid_duplicates_and_zero(tmp_path):
    path = tmp_path / "input.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Danh sách"
    sheet.append(["STT", " gcn "])
    sheet.append([1, "AM 012345"])
    sheet.append([2, "am-012345"])
    sheet.append([3, "A 123456"])
    sheet.append([4, None])
    workbook.save(path)
    data = read_gcn_excel(path, "Danh sách", "GCN")
    assert data.valid_gcn_keys == {"AM012345"}
    assert data.gcn_rows["AM012345"] == [2, 3]
    assert data.invalid_count == 1
    assert data.duplicate_count == 1
    assert data.total_data_rows == 3


def test_excel_reader_splits_semicolon_separated_values(tmp_path):
    path = tmp_path / "input.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Danh sách"
    sheet.append(["STT", "GCN"])
    sheet.append([1, "DH 358799;DH 358800"])
    workbook.save(path)
    data = read_gcn_excel(path, "Danh sách", "GCN")
    assert data.valid_gcn_keys == {"DH358799", "DH358800"}
    assert data.gcn_rows["DH358799"] == [2]
    assert data.gcn_rows["DH358800"] == [2]
    assert data.total_data_rows == 2
