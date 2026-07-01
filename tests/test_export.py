"""
Tests for EML and MBOX export.
"""

import email
from email.policy import default as default_policy
import mailbox

import pytest

from pymsgkit import MSGWriter, MboxWriter, create_email, RecipientType


def _parse(data: bytes):
    """Parse EML bytes with the modern email API (decoded headers, get_body)."""
    return email.message_from_bytes(data, policy=default_policy)


def _sample(subject="Hello", body="Body text"):
    msg = MSGWriter()
    msg.set_subject(subject)
    msg.set_sender("alice@example.com", "Alice")
    msg.add_recipient("bob@example.com", "Bob", RecipientType.TO)
    msg.add_recipient("carol@example.com", "Carol", RecipientType.CC)
    msg.set_body(body)
    return msg


def test_eml_bytes_parse_back():
    msg = _sample()
    parsed = _parse(msg.to_eml_bytes())
    assert parsed["Subject"] == "Hello"
    assert "alice@example.com" in parsed["From"]
    assert "bob@example.com" in parsed["To"]
    assert "carol@example.com" in parsed["Cc"]


def test_save_eml_file(tmp_path):
    msg = _sample(subject="Saved", body="On disk")
    path = tmp_path / "out.eml"
    msg.save_eml(str(path))
    parsed = _parse(path.read_bytes())
    assert parsed["Subject"] == "Saved"
    assert path.read_bytes().count(b"On disk") >= 1


def test_eml_plain_body_content():
    msg = _sample(body="The quick brown fox")
    parsed = _parse(msg.to_eml_bytes())
    body = parsed.get_body(preferencelist=("plain",))
    assert "The quick brown fox" in body.get_content()


def test_eml_html_alternative():
    msg = MSGWriter()
    msg.set_subject("HTML")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("<p>Hi <b>there</b></p>", is_html=True)
    parsed = _parse(msg.to_eml_bytes())
    html_part = parsed.get_body(preferencelist=("html",))
    assert html_part is not None
    assert "<b>there</b>" in html_part.get_content()


def test_eml_attachment_preserved():
    msg = _sample()
    msg.add_attachment("notes.txt", b"attached bytes here", mime_type="text/plain")
    parsed = _parse(msg.to_eml_bytes())
    attachments = [p for p in parsed.iter_attachments()]
    assert len(attachments) == 1
    assert attachments[0].get_filename() == "notes.txt"
    assert attachments[0].get_content().rstrip("\n") == "attached bytes here"


def test_eml_unicode_headers_and_body():
    msg = create_email(
        subject="会議 café",
        body="Ω 🚀 é",
        sender_email="tanaka@example.jp",
        sender_name="田中",
        to_recipients=[("zoe@example.com", "Zoë")],
    )
    parsed = _parse(msg.to_eml_bytes())
    assert parsed["Subject"] == "会議 café"
    body = parsed.get_body(preferencelist=("plain",)).get_content()
    assert "🚀" in body


def test_mbox_multiple_messages(tmp_path):
    box = MboxWriter()
    box.add(_sample(subject="First", body="one"))
    box.add(_sample(subject="Second", body="two"))
    assert len(box) == 2

    path = tmp_path / "archive.mbox"
    box.save(str(path))

    mb = mailbox.mbox(str(path))
    try:
        subjects = sorted(m["Subject"] for m in mb)
    finally:
        mb.close()
    assert subjects == ["First", "Second"]


def test_mbox_save_is_not_appending(tmp_path):
    """Saving twice must not accumulate duplicate messages."""
    box = MboxWriter()
    box.add(_sample(subject="Only"))
    path = tmp_path / "a.mbox"
    box.save(str(path))
    box.save(str(path))

    mb = mailbox.mbox(str(path))
    try:
        count = len(mb)
    finally:
        mb.close()
    assert count == 1


def test_mbox_roundtrip_body(tmp_path):
    box = MboxWriter()
    box.add(_sample(subject="Body check", body="unique-token-xyz"))
    path = tmp_path / "b.mbox"
    box.save(str(path))

    mb = mailbox.mbox(str(path))
    try:
        msg = list(mb)[0]
        payload = msg.get_payload(decode=True)
        if payload is None:  # multipart
            payload = msg.get_payload(0).get_payload(decode=True)
    finally:
        mb.close()
    assert b"unique-token-xyz" in payload
