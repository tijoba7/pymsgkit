"""
Round-trip tests: write a MSG, then read it back with olefile and assert the
decoded MAPI properties match what we put in.
"""

import struct

import pytest

from pymsgkit import MSGWriter, create_email, RecipientType
from pymsgkit.properties import PropertyTag, filetime_to_datetime
from pymsgkit.types import PropertyType


def test_subject_and_body_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("Quarterly Report")
    msg.set_sender("boss@corp.com", "The Boss")
    msg.add_recipient("team@corp.com", "Team")
    msg.set_body("Plain body content")
    msg.save(str(path))

    r = read_msg(path)
    assert r.unicode(PropertyTag.PR_SUBJECT) == "Quarterly Report"
    assert r.unicode(PropertyTag.PR_BODY) == "Plain body content"
    # RE:/FW: stripped conversation topic equals the subject here.
    assert r.unicode(PropertyTag.PR_CONVERSATION_TOPIC) == "Quarterly Report"


def test_conversation_topic_strips_reply_prefix(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("RE: Original Thread")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("x")
    msg.save(str(path))

    r = read_msg(path)
    assert r.unicode(PropertyTag.PR_SUBJECT) == "RE: Original Thread"
    assert r.unicode(PropertyTag.PR_CONVERSATION_TOPIC) == "Original Thread"


def test_html_body_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    html = "<html><body><p>Hello &amp; welcome</p></body></html>"
    msg = MSGWriter()
    msg.set_subject("HTML")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body(html, is_html=True)
    msg.save(str(path))

    r = read_msg(path)
    stored = r.binary(PropertyTag.PR_HTML)
    assert stored == html.encode("utf-8")


def test_sender_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("forensic.sender@archive.com", "Original Sender")
    msg.add_recipient("legal@corp.com")
    msg.set_body("b")
    msg.save(str(path))

    r = read_msg(path)
    assert r.unicode(PropertyTag.PR_SENDER_NAME) == "Original Sender"
    assert r.unicode(PropertyTag.PR_SENDER_EMAIL_ADDRESS) == "forensic.sender@archive.com"
    assert r.unicode(PropertyTag.PR_SENDER_ADDRTYPE) == "SMTP"
    # sent-representing mirrors the sender for a normal email
    assert r.unicode(PropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS) == "forensic.sender@archive.com"


def test_unicode_content_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = create_email(
        subject="会議の予定 café ☕",
        body="Ω emoji 🚀 and accents é ü",
        sender_email="tanaka@example.jp",
        sender_name="田中 太郎",
        to_recipients=[("zoe@example.com", "Zoë")],
    )
    msg.save(str(path))

    r = read_msg(path)
    assert r.unicode(PropertyTag.PR_SUBJECT) == "会議の予定 café ☕"
    assert r.unicode(PropertyTag.PR_BODY) == "Ω emoji 🚀 and accents é ü"
    assert r.unicode(PropertyTag.PR_SENDER_NAME) == "田中 太郎"


def test_display_recipients(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.set_body("b")
    msg.add_recipient("to1@x.com", "Alice", RecipientType.TO)
    msg.add_recipient("to2@x.com", "Bob", RecipientType.TO)
    msg.add_recipient("cc@x.com", "Carol", RecipientType.CC)
    msg.add_recipient("bcc@x.com", "Dave", RecipientType.BCC)
    msg.save(str(path))

    r = read_msg(path)
    assert r.unicode(PropertyTag.PR_DISPLAY_TO) == "Alice; Bob"
    assert r.unicode(PropertyTag.PR_DISPLAY_CC) == "Carol"
    assert r.unicode(PropertyTag.PR_DISPLAY_BCC) == "Dave"


def test_recipient_storages_written(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.set_body("b")
    msg.add_recipient("to1@x.com", "Alice", RecipientType.TO)
    msg.add_recipient("cc@x.com", "Carol", RecipientType.CC)
    msg.save(str(path))

    r = read_msg(path)
    recips = r.storages("__recip_version1.0")
    assert len(recips) == 2

    # First recipient storage decodes to Alice / TO
    s0 = "__recip_version1.0_#00000000"
    assert r.unicode(PropertyTag.PR_DISPLAY_NAME, storage=s0) == "Alice"
    assert r.unicode(PropertyTag.PR_EMAIL_ADDRESS, storage=s0) == "to1@x.com"
    rtype = r.fixed_props(storage=s0)[PropertyTag.PR_RECIPIENT_TYPE]
    assert struct.unpack("<i", rtype[:4])[0] == int(RecipientType.TO)


def test_attachment_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    payload = b"%PDF-1.4 fake pdf bytes" * 10
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("see attached")
    msg.add_attachment("report.pdf", payload, mime_type="application/pdf")
    msg.save(str(path))

    r = read_msg(path)
    attaches = r.storages("__attach_version1.0")
    assert len(attaches) == 1
    s0 = "__attach_version1.0_#00000000"
    assert r.unicode(PropertyTag.PR_ATTACH_LONG_FILENAME, storage=s0) == "report.pdf"
    assert r.unicode(PropertyTag.PR_ATTACH_MIME_TAG, storage=s0) == "application/pdf"
    assert r.binary(PropertyTag.PR_ATTACH_DATA_BIN, storage=s0) == payload
    assert r.unicode(PropertyTag.PR_ATTACH_EXTENSION, storage=s0) == ".pdf"


def test_large_attachment_integrity(tmp_path, read_msg):
    """Data above the 4096-byte mini-stream cutoff uses regular sectors."""
    path = tmp_path / "m.msg"
    payload = bytes(range(256)) * 500  # 128 KB
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("b")
    msg.add_attachment("big.bin", payload)
    msg.save(str(path))

    r = read_msg(path)
    s0 = "__attach_version1.0_#00000000"
    assert r.binary(PropertyTag.PR_ATTACH_DATA_BIN, storage=s0) == payload


def test_inline_attachment_flags(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body('<img src="cid:logo">', is_html=True)
    msg.add_attachment("logo.png", b"PNGDATA", content_id="logo",
                       mime_type="image/png", is_inline=True)
    msg.save(str(path))

    r = read_msg(path)
    s0 = "__attach_version1.0_#00000000"
    assert r.unicode(PropertyTag.PR_ATTACH_CONTENT_ID, storage=s0) == "logo"


def test_hasattach_flag_set(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("b")
    msg.add_attachment("f.txt", b"data")
    msg.save(str(path))

    r = read_msg(path)
    flags = struct.unpack("<i", r.fixed_props()[PropertyTag.PR_MESSAGE_FLAGS][:4])[0]
    assert flags & 0x10  # MSGFLAG_HASATTACH


def test_submit_time_roundtrip(tmp_path, read_msg):
    path = tmp_path / "m.msg"
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("b")
    msg.save(str(path))

    r = read_msg(path)
    raw = r.fixed_props()[PropertyTag.PR_CLIENT_SUBMIT_TIME]
    filetime = struct.unpack("<Q", raw)[0]
    dt = filetime_to_datetime(filetime)
    assert dt.year >= 2020  # sanity: decodes to a real, recent timestamp
