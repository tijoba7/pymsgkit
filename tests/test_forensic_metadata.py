"""
Forensic metadata control: setting original timestamps, explicit Message-ID,
and producing reproducible output.
"""

import struct
from datetime import datetime, timezone

import pytest

from pymsgkit import MSGWriter, create_email
from pymsgkit.properties import PropertyTag, filetime_to_datetime


ORIG = datetime(2019, 3, 14, 9, 26, 53, tzinfo=timezone.utc)


def _msg():
    m = MSGWriter()
    m.set_subject("Case")
    m.set_sender("suspect@corp.com", "Suspect")
    m.add_recipient("victim@corp.com", "Victim")
    m.set_body("body")
    return m


def test_set_sent_time_roundtrips(tmp_path, read_msg):
    m = _msg()
    m.set_sent_time(ORIG)
    path = tmp_path / "m.msg"
    m.save(str(path))

    raw = read_msg(path).fixed_props()[PropertyTag.PR_CLIENT_SUBMIT_TIME]
    dt = filetime_to_datetime(struct.unpack("<Q", raw)[0])
    assert abs((dt - ORIG).total_seconds()) < 1


def test_set_dates_all_fields(tmp_path, read_msg):
    m = _msg()
    m.set_dates(sent=ORIG, received=ORIG, created=ORIG, modified=ORIG)
    path = tmp_path / "m.msg"
    m.save(str(path))

    fp = read_msg(path).fixed_props()
    for tag in (PropertyTag.PR_CLIENT_SUBMIT_TIME,
                PropertyTag.PR_MESSAGE_DELIVERY_TIME,
                PropertyTag.PR_CREATION_TIME,
                PropertyTag.PR_LAST_MODIFICATION_TIME):
        dt = filetime_to_datetime(struct.unpack("<Q", fp[tag])[0])
        assert abs((dt - ORIG).total_seconds()) < 1


def test_explicit_message_id_preserved(tmp_path, read_msg):
    m = _msg()
    m.set_message_id("case-2024-001@evidence.local")
    path = tmp_path / "m.msg"
    m.save(str(path))

    mid = read_msg(path).string8(PropertyTag.PR_INTERNET_MESSAGE_ID)
    assert mid == "<case-2024-001@evidence.local>"


def test_deterministic_output_with_pinned_metadata(tmp_path):
    """Same input + pinned timestamps + Message-ID => byte-identical files."""
    def build(path):
        m = create_email(
            subject="Reproducible",
            body="same bytes every time",
            sender_email="a@b.com",
            sender_name="Alice",
            to_recipients=[("c@d.com", "Carol")],
        )
        m.set_dates(sent=ORIG, received=ORIG, created=ORIG, modified=ORIG)
        m.set_message_id("fixed-id@local")
        m.save(str(path))

    p1 = tmp_path / "a.msg"
    p2 = tmp_path / "b.msg"
    build(p1)
    build(p2)
    assert p1.read_bytes() == p2.read_bytes()


def test_extract_msg_reads_original_date(tmp_path):
    extract_msg = pytest.importorskip("extract_msg")
    m = _msg()
    m.set_dates(sent=ORIG, received=ORIG)
    path = tmp_path / "m.msg"
    m.save(str(path))

    parsed = extract_msg.Message(str(path))
    try:
        # extract-msg exposes the submit time via .date; it must reflect 2019.
        assert "2019" in (parsed.date or "")
    finally:
        parsed.close()
