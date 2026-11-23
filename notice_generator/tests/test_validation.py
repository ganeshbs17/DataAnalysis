import datetime as dt
from collections import Counter

from notice_generator.core.grouping import address_variant_groups, shared_accounts
from notice_generator.core.models import Status
from notice_generator.core.validation import validate_batch
from notice_generator.importers.prepared_importer import import_prepared

TODAY = dt.date(2026, 11, 23)


def _load(path, settings):
    recs = import_prepared(path).records
    validate_batch(recs, settings, today=TODAY)
    return recs


def rules(r):
    return {i.rule for i in r.issues}


def test_truncation_blocks_until_confirmed(sample_xlsx, settings):
    recs = _load(sample_xlsx, settings)
    assert all("V-03" in rules(r) for r in recs)  # 16-digit victim account in every row
    r = recs[0]
    assert r.status == Status.BLOCKED
    r.confirmations.add("V-03:BankAccountNumber")
    assert r.status != Status.BLOCKED


def test_edited_account_clears_v03(sample_xlsx, settings):
    recs = import_prepared(sample_xlsx).records
    recs[0].edits["BankAccountNumber"] = "1234567890123401"
    validate_batch(recs, settings, today=TODAY)
    assert "V-03" not in rules(recs[0])


def test_groups_and_shared_accounts(sample_xlsx, settings):
    recs = _load(sample_xlsx, settings)
    groups = address_variant_groups(recs)
    assert len(groups) == 1 and len(groups[0]) == 2
    assert "201030653952" in shared_accounts(recs)
    assert "V-10" in rules(recs[3]) and "V-10" in rules(recs[4])
    assert "V-09" in rules(recs[0])


def test_bank_typo_suggested_not_applied(sample_xlsx, settings):
    recs = _load(sample_xlsx, settings)
    v12 = [i for i in recs[2].issues if i.rule == "V-12"]
    assert v12 and v12[0].suggestion == "City Union Bank"
    assert recs[2].text("SuspectBankName") == "City Unin Bank"


def test_missing_required_blocks_only_that_record(sample_xlsx, settings):
    recs = _load(sample_xlsx, settings)
    for r in recs:
        r.confirmations.update({"V-03:BankAccountNumber", "V-03:SuspectAccountNumber"})
    bad = recs[-1]
    assert "V-01" in rules(bad) and bad.status == Status.BLOCKED
    assert Counter(r.status for r in recs)[Status.BLOCKED] == 1


def test_appearance_date_rules(sample_xlsx, settings):
    settings.appearance_date = dt.date(2026, 1, 1)
    recs = _load(sample_xlsx, settings)
    assert "V-07" in rules(recs[0])
    settings.appearance_date = dt.date(2030, 12, 1)  # Sunday
    recs = _load(sample_xlsx, settings)
    assert "V-08" in rules(recs[0]) and "V-07" not in rules(recs[0])


def test_gui_vs_file_conflict(sample_xlsx, settings):
    settings.officer_name = "Officer B"
    recs = _load(sample_xlsx, settings)
    assert "V-13" in rules(recs[0])
    settings.officer_name = "Officer A"
    recs = _load(sample_xlsx, settings)
    assert "V-13" not in rules(recs[0])


def test_amount_rules(sample_xlsx, settings):
    recs = import_prepared(sample_xlsx).records
    recs[0].edits["AmountDeposited"] = "abc"
    recs[1].edits["AmountDeposited"] = "0"
    validate_batch(recs, settings, today=TODAY)
    assert "V-05" in rules(recs[0]) and "V-05" in rules(recs[1])
