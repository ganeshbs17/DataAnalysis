from pathlib import Path

from notice_generator.core.models import CellType
from notice_generator.importers.cell_reader import number_cell, read_workbook
from notice_generator.importers.prepared_importer import import_prepared


def test_import_maps_all_columns_and_skips_empty(sample_xlsx):
    res = import_prepared(sample_xlsx)
    assert res.sheet == "Sheet1"
    assert res.mapping.header_row == 1
    assert len(res.mapping.field_to_col) == 11
    assert set(res.mapping.how.values()) == {"exact"}
    assert len(res.records) == 6
    assert res.skipped_empty == 1


def test_account_numbers_are_text_never_float(sample_xlsx):
    recs = import_prepared(sample_xlsx).records
    victim = recs[0].raw("BankAccountNumber")
    assert victim.text == "1234567890123400"
    assert victim.ctype == CellType.INT and victim.digits == 16
    # float cell → integer text, no ".0" or exponent
    assert recs[2].raw("SuspectAccountNumber").text == "20100036980152"
    # text cell kept as text
    assert recs[3].raw("SuspectAccountNumber").ctype == CellType.TEXT
    for r in recs:
        t = r.raw("SuspectAccountNumber").text
        assert "e" not in t.lower() and not t.endswith(".0")


def test_number_cell_conversion():
    assert number_cell(7692500102064500).text == "7692500102064500"
    assert number_cell(7692500102064500.0).text == "7692500102064500"
    assert number_cell(2.01e13).text == "20100000000000"
    assert number_cell(12.5).text == "12.5"


def test_dates_and_times_typed(sample_xlsx):
    r = import_prepared(sample_xlsx).records[0]
    assert r.raw("AppearanceDate").value.isoformat() == "2030-12-02"
    assert r.raw("AppearanceTime").value.strftime("%H:%M") == "11:00"


def test_csv_scientific_notation_detected(tmp_path: Path):
    p = tmp_path / "d.csv"
    p.write_text("SuspectName,BankAccountNumber,SuspectAccountNumber\nA,7.6925E+15,0012345678\n",
                 encoding="utf-8")
    rows = read_workbook(p)[0].rows
    assert rows[1][1].scientific
    assert rows[1][2].text == "0012345678"  # leading zeros kept


def test_csv_cp1252(tmp_path: Path):
    p = tmp_path / "d.csv"
    p.write_bytes("SuspectName,Address,AmountDeposited\nJos\xe9,Caf\xe9,100\n".encode("cp1252"))
    assert read_workbook(p)[0].rows[1][0].text == "José"
