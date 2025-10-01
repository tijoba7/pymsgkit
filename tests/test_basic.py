"""
Basic functionality tests
"""

import pytest
import os
from pymsgkit import MSGWriter, create_email, RecipientType

def test_create_simple_email(tmp_path):
    """Test creating a simple email"""
    filepath = tmp_path / "test.msg"

    msg = create_email(
        subject="Test Subject",
        body="Test Body",
        sender_email="sender@test.com",
        sender_name="Test Sender",
        to_recipients=[("recipient@test.com", "Test Recipient")]
    )
    msg.save(str(filepath))

    assert filepath.exists()
    assert filepath.stat().st_size > 0

    # Check CFB signature
    with open(filepath, 'rb') as f:
        signature = f.read(8)
        assert signature == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

def test_html_email(tmp_path):
    """Test HTML email creation"""
    filepath = tmp_path / "html.msg"

    msg = MSGWriter()
    msg.set_subject("HTML Test")
    msg.set_sender("sender@test.com")
    msg.set_body("<h1>Test</h1>", is_html=True)
    msg.add_recipient("recipient@test.com")
    msg.save(str(filepath))

    assert filepath.exists()
    assert filepath.stat().st_size > 0

def test_attachment(tmp_path):
    """Test email with attachment"""
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
    """Test email with multiple recipients"""
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

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
