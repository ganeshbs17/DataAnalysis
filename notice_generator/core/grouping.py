"""Address-variant groups, shared accounts and exact duplicates (FR-03).

Information only: nothing is merged or removed, and two records are never
concluded to be the same person because a field matches.
"""
from __future__ import annotations

from collections import defaultdict

from .models import Record
from .normaliser import name_key, normalise_account, normalise_address


def address_variant_groups(records: list[Record]) -> list[list[Record]]:
    """Same normalised name AND same suspect account, more than one record."""
    g: dict[tuple[str, str], list[Record]] = defaultdict(list)
    for r in records:
        acc = normalise_account(r.text("SuspectAccountNumber"))
        nk = name_key(r.text("SuspectName"))
        if acc and nk:
            g[(nk, acc)].append(r)
    return [v for v in g.values() if len(v) > 1]


def shared_accounts(records: list[Record]) -> dict[str, list[Record]]:
    """Suspect accounts that appear under more than one normalised name."""
    by_acc: dict[str, list[Record]] = defaultdict(list)
    for r in records:
        acc = normalise_account(r.text("SuspectAccountNumber"))
        if acc:
            by_acc[acc].append(r)
    return {a: rs for a, rs in by_acc.items()
            if len({name_key(r.text("SuspectName")) for r in rs}) > 1}


def exact_duplicates(records: list[Record]) -> list[list[Record]]:
    g: dict[tuple[str, str, str], list[Record]] = defaultdict(list)
    for r in records:
        key = (name_key(r.text("SuspectName")),
               normalise_address(r.text("Address")).lower(),
               normalise_account(r.text("SuspectAccountNumber")))
        g[key].append(r)
    return [v for v in g.values() if len(v) > 1]
