import datetime as dt
from decimal import Decimal

from notice_generator.core import formatting as f
from notice_generator.core.normaliser import apply_prefix_mode, name_key, normalise_address, normalise_text


def test_indian_grouping():
    assert f.indian_group(2000000) == "20,00,000"
    assert f.indian_group(13700000) == "1,37,00,000"
    assert f.indian_group(999) == "999"
    assert f.indian_group(1000) == "1,000"


def test_amount_format_has_no_rupee_symbol():
    assert f.format_amount(Decimal("13700000")) == "1,37,00,000/-"
    assert f.format_amount(Decimal("1500.5"), "") == "1,500.50"
    assert f.parse_amount("₹ 1,37,00,000/-") == Decimal("13700000")
    assert f.parse_amount("abc") is None


def test_amount_in_words():
    assert f.amount_in_words(Decimal(2000000)) == "Rupees Twenty Lakh only"
    assert f.amount_in_words(Decimal(13700000)) == "Rupees One Crore Thirty Seven Lakh only"
    assert f.amount_in_words(Decimal(2365000)) == "Rupees Twenty Three Lakh Sixty Five Thousand only"


def test_multiple_transfer_dates():
    text, bad = f.format_transfer_dates("18.02.2025, 20.02.2025,\n21.02.2025,\n25.02.2025 ")
    assert text == "18.02.2025, 20.02.2025, 21.02.2025 and 25.02.2025"
    assert bad == []
    text, bad = f.format_transfer_dates("22.01.2025; sometime")
    assert text == "22.01.2025 and sometime" and bad == ["sometime"]


def test_time_and_date_formats():
    assert f.format_time(dt.time(11, 0)) == "11:00 AM"
    assert f.format_time(dt.time(14, 30)) == "02:30 PM"
    assert dt.date(2026, 12, 2).strftime("%d/%m/%Y") == "02/12/2026"


def test_text_normalisation():
    assert normalise_text("\n\n  A   B \nC  \n\n") == "A B\nC"
    assert normalise_address("X\nAddress:Y\n", remove_label=True) == "X\nY"
    assert normalise_address("X\nAddress:Y\n") == "X\nAddress:Y"


def test_prefixes():
    assert name_key("Mr.MOLUGU Test") == name_key("molugu test")
    assert name_key("Mrinal Sen") == "mrinal sen"
    assert apply_prefix_mode("Mr. A B", "remove") == "A B"
    assert apply_prefix_mode("Smt A B", "standardise") == "Ms. A B"
    assert apply_prefix_mode("Mr.A", "keep") == "Mr.A"
