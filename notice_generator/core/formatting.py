"""Indian amount format, amount in words, date and time formats (FR-05)."""
from __future__ import annotations

import datetime as dt
import re
from decimal import Decimal, InvalidOperation

# ------------------------------------------------------------------ amounts


def parse_amount(text: str) -> Decimal | None:
    t = (text or "").strip()
    t = re.sub(r"(?i)(₹|rs\.?|inr|/-|,|\s)", "", t)
    if not t:
        return None
    try:
        d = Decimal(t)
    except InvalidOperation:
        return None
    if not d.is_finite():
        return None
    return d


def indian_group(n: int) -> str:
    s = str(abs(n))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts) + "," + tail
    return ("-" if n < 0 else "") + s


def format_amount(d: Decimal, suffix: str = "/-") -> str:
    """2000000 → '20,00,000/-'. No currency symbol: the template supplies ₹."""
    q = d.quantize(Decimal("0.01"))
    rupees = int(q)
    paise = int((abs(q) - abs(Decimal(rupees))) * 100)
    out = indian_group(rupees)
    if paise:
        out += f".{paise:02d}"
    return out + suffix


_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
         "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
         "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _two(n: int) -> str:
    if n < 20:
        return _ONES[n]
    return (_TENS[n // 10] + (" " + _ONES[n % 10] if n % 10 else "")).strip()


def _three(n: int) -> str:
    h, r = divmod(n, 100)
    parts = []
    if h:
        parts.append(_ONES[h] + " Hundred")
    if r:
        parts.append(_two(r))
    return " ".join(parts)


def int_to_indian_words(n: int) -> str:
    if n == 0:
        return "Zero"
    parts = []
    crore, n = divmod(n, 10_000_000)
    lakh, n = divmod(n, 100_000)
    thousand, n = divmod(n, 1000)
    if crore:
        parts.append(int_to_indian_words(crore) + " Crore")
    if lakh:
        parts.append(_two(lakh) + " Lakh")
    if thousand:
        parts.append(_two(thousand) + " Thousand")
    if n:
        parts.append(_three(n))
    return " ".join(parts)


def amount_in_words(d: Decimal) -> str:
    q = d.quantize(Decimal("0.01"))
    rupees = int(q)
    paise = int((q - rupees) * 100)
    s = "Rupees " + int_to_indian_words(rupees)
    if paise:
        s += " and " + _two(paise) + " Paise"
    return s + " only"


# ------------------------------------------------------------------ dates

_DATE_FORMATS = ["%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
                 "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%Y-%m-%d %H:%M:%S"]
_SPLIT_RE = re.compile(r"[,;\n]+")


def parse_date(text: str) -> dt.date | None:
    t = (text or "").strip().rstrip(".")
    for f in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(t, f).date()
        except ValueError:
            continue
    return None


def split_dates(text: str) -> list[str]:
    return [p.strip() for p in _SPLIT_RE.split(text or "") if p.strip()]


def join_list(items: list[str]) -> str:
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " and " + items[-1]


def format_transfer_dates(text: str, fmt: str = "%d.%m.%Y", value=None) -> tuple[str, list[str]]:
    """Return (formatted text, unparsed parts). Unparsed parts are printed as raw text."""
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        return value.strftime(fmt), []
    out, bad = [], []
    for part in split_dates(text):
        d = parse_date(part)
        if d is None:
            bad.append(part)
            out.append(part)
        else:
            out.append(d.strftime(fmt))
    return join_list(out), bad


def parse_time(text: str) -> dt.time | None:
    t = (text or "").strip().upper().replace(".", ":")
    for f in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M%p", "%I %p", "%I:%M:%S %p"):
        try:
            return dt.datetime.strptime(t, f).time()
        except ValueError:
            continue
    return None


def format_time(t: dt.time) -> str:
    return t.strftime("%I:%M %p")


def to_date(value, text: str) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return parse_date(text)


def to_time(value, text: str) -> dt.time | None:
    if isinstance(value, dt.datetime):
        return value.time()
    if isinstance(value, dt.time):
        return value
    return parse_time(text)
