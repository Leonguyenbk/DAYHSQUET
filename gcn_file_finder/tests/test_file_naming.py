from pathlib import Path

from gcn_file_finder.core.file_copier import copy_for_gcn, is_valid_result_filename


def test_required_filename_shapes():
    assert is_valid_result_filename("AM 143443.pdf", "AM143443")
    assert is_valid_result_filename("AM 143443_02.pdf", "AM143443")
    assert is_valid_result_filename("AM 143443_03.jpg", "AM143443")
    assert not is_valid_result_filename("AM143443.pdf", "AM143443")
    assert not is_valid_result_filename("AM_143443.pdf", "AM143443")
    assert not is_valid_result_filename("AM-143443.pdf", "AM143443")


def test_two_and_three_files_get_stable_suffixes(tmp_path):
    output = tmp_path / "output"
    sources = []
    for index in range(3):
        source = tmp_path / f"source_{index}.pdf"
        source.write_bytes(f"different-{index}".encode())
        sources.append(source)
    results = [copy_for_gcn(source, output, "AM143443") for source in sources]
    assert [result.destination.name for result in results] == [
        "AM 143443.pdf", "AM 143443_02.pdf", "AM 143443_03.pdf"
    ]


def test_existing_same_content_is_not_copied_again(tmp_path):
    source = tmp_path / "source.pdf"
    source.write_bytes(b"same-content")
    output = tmp_path / "output"
    first = copy_for_gcn(source, output, "CĐ123456")
    second = copy_for_gcn(source, output, "CĐ123456")
    assert first.destination == second.destination
    assert second.status == "ĐÃ TỒN TẠI"
    assert len(list(output.iterdir())) == 1


def test_existing_different_content_is_never_overwritten(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    existing = output / "ĐA 001234.jpg"
    existing.write_bytes(b"old")
    source = tmp_path / "new.jpg"
    source.write_bytes(b"new")
    result = copy_for_gcn(source, output, "ĐA001234")
    assert existing.read_bytes() == b"old"
    assert result.destination.name == "ĐA 001234_02.jpg"


def test_one_source_can_be_copied_once_for_each_gcn(tmp_path):
    source = tmp_path / "document.pdf"
    source.write_bytes(b"one document with two GCN")
    output = tmp_path / "output"
    one = copy_for_gcn(source, output, "AM143443")
    two = copy_for_gcn(source, output, "CS012345")
    assert one.destination.name == "AM 143443.pdf"
    assert two.destination.name == "CS 012345.pdf"
