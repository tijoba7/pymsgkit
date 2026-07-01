"""
Export helpers: turn the messages you build with :class:`MSGWriter` into the
standard, widely-importable RFC 5322 ``.eml`` format and into ``.mbox``
archives that hold many messages in a single file.

Why not native ``.pst``?
------------------------
A ``.pst`` file is not "a bunch of MSG files in a container" -- it is a full
on-disk database (the MS-PST NDB/LTP layers: CRC-protected pages, node and
block B-trees, heap-on-node, table/property contexts, encoded blocks, a folder
hierarchy, and a named-property map).  Writing one correctly from scratch is an
order of magnitude more work than the entire MSG writer, which is why no
pure-Python PST *writer* exists in the wild.  For the common need behind "make
a PST" -- collecting many messages into one importable archive -- MBOX and EML
do the job and import cleanly into Outlook, Thunderbird, Apple Mail, and most
eDiscovery tools.  See ``docs`` / README for the full rationale.
"""

import mailbox
from email.message import EmailMessage
from email.utils import format_datetime, formataddr
from datetime import datetime, timezone

from .properties import PropertyTag
from .types import RecipientType


def _prop(msg, tag, default=None):
    """Fetch a property value from an MSGWriter, or a default if unset."""
    prop = msg.properties.get(tag)
    return prop.value if prop is not None else default


def msg_to_email_message(msg) -> EmailMessage:
    """Convert an :class:`~pymsgkit.writer.MSGWriter` to an ``EmailMessage``.

    Plain-text and HTML bodies are both preserved (as a multipart/alternative
    when both are present); attachments are carried over, including inline
    attachments with their Content-ID.
    """
    em = EmailMessage()

    subject = _prop(msg, PropertyTag.PR_SUBJECT, "")
    if subject:
        em["Subject"] = subject

    sender_email = _prop(msg, PropertyTag.PR_SENDER_EMAIL_ADDRESS, "")
    sender_name = _prop(msg, PropertyTag.PR_SENDER_NAME, "")
    if sender_email:
        em["From"] = formataddr((sender_name or "", sender_email))

    to_addrs = [formataddr((r["name"], r["email"]))
                for r in msg.recipients if r["type"] == RecipientType.TO]
    cc_addrs = [formataddr((r["name"], r["email"]))
                for r in msg.recipients if r["type"] == RecipientType.CC]
    bcc_addrs = [formataddr((r["name"], r["email"]))
                 for r in msg.recipients if r["type"] == RecipientType.BCC]
    if to_addrs:
        em["To"] = ", ".join(to_addrs)
    if cc_addrs:
        em["Cc"] = ", ".join(cc_addrs)
    if bcc_addrs:
        em["Bcc"] = ", ".join(bcc_addrs)

    date = _prop(msg, PropertyTag.PR_CLIENT_SUBMIT_TIME)
    if isinstance(date, datetime):
        if date.tzinfo is None:
            date = date.replace(tzinfo=timezone.utc)
        em["Date"] = format_datetime(date)

    message_id = _prop(msg, PropertyTag.PR_INTERNET_MESSAGE_ID)
    if message_id:
        em["Message-ID"] = message_id

    # Bodies. PR_BODY is plain text; PR_HTML is HTML stored as UTF-8 bytes.
    plain = _prop(msg, PropertyTag.PR_BODY, "")
    html = _prop(msg, PropertyTag.PR_HTML)
    em.set_content(plain if isinstance(plain, str) else "")
    if html:
        html_str = html.decode("utf-8") if isinstance(html, (bytes, bytearray)) else str(html)
        em.add_alternative(html_str, subtype="html")

    for att in msg.attachments:
        maintype, _, subtype = (att.get("mime_type") or "application/octet-stream").partition("/")
        subtype = subtype or "octet-stream"
        kwargs = dict(maintype=maintype, subtype=subtype, filename=att["filename"])
        if att.get("is_inline"):
            kwargs["disposition"] = "inline"
        if att.get("content_id"):
            kwargs["cid"] = att["content_id"]
        em.add_attachment(att["data"], **kwargs)

    return em


def msg_to_eml_bytes(msg) -> bytes:
    """Serialize an MSGWriter as RFC 5322 ``.eml`` bytes."""
    return msg_to_email_message(msg).as_bytes()


def save_eml(msg, filepath: str):
    """Write an MSGWriter to a ``.eml`` file on disk."""
    with open(filepath, "wb") as f:
        f.write(msg_to_eml_bytes(msg))


class MboxWriter:
    """Collect multiple messages into a single ``.mbox`` archive.

    This is the practical stand-in for a PST when you need many messages in one
    importable file::

        box = MboxWriter()
        box.add(msg1)
        box.add(msg2)
        box.save("archive.mbox")
    """

    def __init__(self):
        self._messages = []

    def add(self, msg):
        """Add an MSGWriter (or a pre-built EmailMessage) to the archive."""
        if isinstance(msg, EmailMessage):
            self._messages.append(msg)
        else:
            self._messages.append(msg_to_email_message(msg))
        return self

    def __len__(self):
        return len(self._messages)

    def save(self, filepath: str):
        """Write all collected messages to an mbox file.

        The file is truncated first so ``save`` produces exactly the messages
        added to this writer rather than appending to a stale archive.
        """
        open(filepath, "wb").close()  # start from an empty archive
        box = mailbox.mbox(filepath)
        box.lock()
        try:
            for em in self._messages:
                box.add(mailbox.mboxMessage(em))
            box.flush()
        finally:
            box.unlock()
            box.close()
