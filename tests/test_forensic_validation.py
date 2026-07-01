"""
Validation against extract-msg, the de-facto forensic MSG parser.

These tests open PyMsgKit output with an *independent* MAPI implementation in
its strict default mode. That catches spec-compliance problems (property-tag
byte order, stream header sizes, string terminators) that a CFB-only check with
olefile cannot see. If extract-msg is not installed the module is skipped.
"""

import pytest

extract_msg = pytest.importorskip("extract_msg")

from pymsgkit import MSGWriter, create_email, RecipientType


def _open(path):
    # Strict default error behaviour: raises on standards violations.
    return extract_msg.Message(str(path))


def test_extract_msg_reads_full_message(tmp_path):
    path = tmp_path / "case.msg"
    msg = create_email(
        subject="Case 2024-001",
        body="Reconstructed evidence body.",
        sender_email="suspect@corp.com",
        sender_name="A Suspect",
        to_recipients=[("victim@corp.com", "The Victim")],
        cc_recipients=[("legal@corp.com", "Legal")],
    )
    msg.add_attachment("evidence.pdf", b"%PDF-1.4 data" * 100, mime_type="application/pdf")
    msg.save(str(path))

    m = _open(path)
    try:
        assert m.subject == "Case 2024-001"
        assert m.sender == "A Suspect"
        assert "victim@corp.com" in m.to
        assert "legal@corp.com" in m.cc
        assert m.body.startswith("Reconstructed evidence body.")
        assert len(m.attachments) == 1
        assert m.attachments[0].longFilename == "evidence.pdf"
    finally:
        m.close()


def test_extract_msg_strings_have_no_trailing_null(tmp_path):
    path = tmp_path / "s.msg"
    msg = MSGWriter()
    msg.set_subject("No Null Please")
    msg.set_sender("a@b.com", "Alice")
    msg.add_recipient("c@d.com", "Carol")
    msg.set_body("clean body")
    msg.save(str(path))

    m = _open(path)
    try:
        assert m.subject == "No Null Please"
        assert "\x00" not in (m.subject or "")
        assert "\x00" not in (m.sender or "")
    finally:
        m.close()


def test_extract_msg_recipient_types(tmp_path):
    path = tmp_path / "r.msg"
    msg = MSGWriter()
    msg.set_subject("types")
    msg.set_sender("a@b.com")
    msg.set_body("b")
    msg.add_recipient("to@x.com", "To Person", RecipientType.TO)
    msg.add_recipient("cc@x.com", "Cc Person", RecipientType.CC)
    msg.save(str(path))

    m = _open(path)
    try:
        types = sorted(r.type.value for r in m.recipients)
        assert types == [int(RecipientType.TO), int(RecipientType.CC)]
    finally:
        m.close()


def test_extract_msg_large_attachment(tmp_path):
    """A >7 MB attachment exercises the CFB DIFAT path end to end."""
    import hashlib

    path = tmp_path / "big.msg"
    payload = bytes((i * 13) % 256 for i in range(9 * 1024 * 1024))  # 9 MB
    msg = MSGWriter()
    msg.set_subject("large evidence")
    msg.set_sender("a@b.com", "A")
    msg.add_recipient("c@d.com", "C")
    msg.set_body("see attachment")
    msg.add_attachment("disk_image.bin", payload)
    msg.save(str(path))

    m = _open(path)
    try:
        assert m.subject == "large evidence"
        got = m.attachments[0].data
        assert hashlib.sha256(got).hexdigest() == hashlib.sha256(payload).hexdigest()
    finally:
        m.close()


def test_extract_msg_unicode(tmp_path):
    path = tmp_path / "u.msg"
    msg = create_email(
        subject="会議 café",
        body="Ω é ü",
        sender_email="tanaka@example.jp",
        sender_name="田中",
        to_recipients=[("zoe@example.com", "Zoë")],
    )
    msg.save(str(path))

    m = _open(path)
    try:
        assert m.subject == "会議 café"
    finally:
        m.close()
