"""
PyMsgKit - Pure Python library for creating Outlook MSG files
"""

from .writer import MSGWriter
from .types import RecipientType, PropertyType, AttachMethod
from .properties import PropertyTag

def create_email(subject, body, sender_email, sender_name="",
                 to_recipients=None, cc_recipients=None, bcc_recipients=None,
                 is_html=False):
    """
    Quick helper function to create a simple email.

    Args:
        subject: Email subject
        body: Email body text
        sender_email: Sender email address
        sender_name: Sender display name (optional)
        to_recipients: List of (email, name) tuples for TO recipients
        cc_recipients: List of (email, name) tuples for CC recipients
        bcc_recipients: List of (email, name) tuples for BCC recipients
        is_html: Whether body is HTML (default: False)

    Returns:
        MSGWriter instance
    """
    msg = MSGWriter()
    msg.set_subject(subject)
    msg.set_body(body, is_html=is_html)
    msg.set_sender(sender_email, sender_name)

    if to_recipients:
        for email, name in to_recipients:
            msg.add_recipient(email, name, RecipientType.TO)

    if cc_recipients:
        for email, name in cc_recipients:
            msg.add_recipient(email, name, RecipientType.CC)

    if bcc_recipients:
        for email, name in bcc_recipients:
            msg.add_recipient(email, name, RecipientType.BCC)

    return msg

__version__ = "1.0.0"
__all__ = ['MSGWriter', 'create_email', 'RecipientType', 'PropertyType', 'AttachMethod', 'PropertyTag']
