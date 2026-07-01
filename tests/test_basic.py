"""
Basic smoke tests: files are created and are structurally compound files.
"""

import pytest

from pymsgkit import MSGWriter, create_email, RecipientType
from conftest import CFB_SIGNATURE


def test_create_simple_email(tmp_path):
    filepath = tmp_path / "test.msg"
    msg = create_email(
        subject="Test Subject",
        body="Test Body",
        sender_email="sender@test.com",
        sender_name="Test Sender",
        to_recipients=[("recipient@test.com", "Test Recipient")],
    )
    msg.save(str(filepath))

    assert filepath.exists()
    assert filepath.stat().st_size > 0
    with open(filepath, "rb") as f:
        assert f.read(8) == CFB_SIGNATURE


def test_html_email(tmp_path):
    filepath = tmp_path / "html.msg"
    msg = MSGWriter()
    msg.set_subject("HTML Test")
    msg.set_sender("sender@test.com")
    msg.set_body("<h1>Test</h1>", is_html=True)
    msg.add_recipient("recipient@test.com")
    msg.save(str(filepath))

    assert filepath.exists() and filepath.stat().st_size > 0


def test_attachment(tmp_path):
    filepath = tmp_path / "attachment.msg"
    msg = MSGWriter()
    msg.set_subject("Attachment Test")
    msg.set_sender("sender@test.com")
    msg.set_body("See attachment")
    msg.add_recipient("recipient@test.com")
    msg.add_attachment("test.txt", b"Test content", mime_type="text/plain")
    msg.save(str(filepath))

    assert filepath.exists()


def test_multiple_recipients(tmp_path):
    filepath = tmp_path / "multi_recipient.msg"
    msg = MSGWriter()
    msg.set_subject("Multi Recipient")
    msg.set_sender("sender@test.com")
    msg.set_body("Test")
    msg.add_recipient("to1@test.com", "To 1", RecipientType.TO)
    msg.add_recipient("to2@test.com", "To 2", RecipientType.TO)
    msg.add_recipient("cc@test.com", "CC", RecipientType.CC)
    msg.add_recipient("bcc@test.com", "BCC", RecipientType.BCC)
    msg.save(str(filepath))

    assert filepath.exists()


def test_save_is_idempotent(tmp_path):
    """Calling save() twice must not corrupt state or duplicate streams."""
    msg = MSGWriter()
    msg.set_subject("Idem")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("hi")
    first = tmp_path / "a.msg"
    second = tmp_path / "b.msg"
    msg.save(str(first))
    msg.save(str(second))
    assert first.read_bytes()[:8] == CFB_SIGNATURE
    assert second.read_bytes()[:8] == CFB_SIGNATURE


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
