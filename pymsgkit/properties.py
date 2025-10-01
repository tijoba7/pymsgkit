"""
MAPI property tags and property encoding
Based on MS-OXPROPS specification
"""

import struct
from datetime import datetime, timezone
from typing import Any, Union
from .types import PropertyType


class PropertyTag:
    """Common MAPI property tags (PidTag*)"""

    # Message Envelope Properties
    PR_MESSAGE_CLASS = 0x001A
    PR_SUBJECT = 0x0037
    PR_CONVERSATION_TOPIC = 0x0070
    PR_CONVERSATION_INDEX = 0x0071
    PR_IMPORTANCE = 0x0017
    PR_PRIORITY = 0x0026
    PR_SENSITIVITY = 0x0036
    PR_MESSAGE_FLAGS = 0x0E07
    PR_MESSAGE_SIZE = 0x0E08

    # Time Properties
    PR_CLIENT_SUBMIT_TIME = 0x0039
    PR_MESSAGE_DELIVERY_TIME = 0x0E06
    PR_CREATION_TIME = 0x3007
    PR_LAST_MODIFICATION_TIME = 0x3008

    # Body Properties
    PR_BODY = 0x1000
    PR_HTML = 0x1013
    PR_RTF_COMPRESSED = 0x1009
    PR_BODY_CONTENT_LOCATION = 0x1014
    PR_BODY_CONTENT_ID = 0x1015

    # Internet Headers and Message ID
    PR_TRANSPORT_MESSAGE_HEADERS = 0x007D
    PR_INTERNET_MESSAGE_ID = 0x1035
    PR_IN_REPLY_TO_ID = 0x1042
    PR_INTERNET_REFERENCES = 0x1039

    # Sender Properties
    PR_SENDER_NAME = 0x0C1A
    PR_SENDER_EMAIL_ADDRESS = 0x0C1F
    PR_SENDER_ADDRTYPE = 0x0C1E
    PR_SENDER_ENTRYID = 0x0C19
    PR_SENDER_SEARCH_KEY = 0x0C1D

    # Sent Representing Properties (for "on behalf of")
    PR_SENT_REPRESENTING_NAME = 0x0042
    PR_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065
    PR_SENT_REPRESENTING_ADDRTYPE = 0x0064
    PR_SENT_REPRESENTING_ENTRYID = 0x0041
    PR_SENT_REPRESENTING_SEARCH_KEY = 0x003B

    # Recipient Properties (in recipient table)
    PR_RECIPIENT_TYPE = 0x0C15
    PR_DISPLAY_NAME = 0x3001
    PR_EMAIL_ADDRESS = 0x3003
    PR_ADDRTYPE = 0x3002
    PR_ENTRYID = 0x0FFF
    PR_SEARCH_KEY = 0x300B
    PR_SMTP_ADDRESS = 0x39FE
    PR_OBJECT_TYPE = 0x0FFE
    PR_DISPLAY_TYPE = 0x3900

    # Recipient Display Properties (in message)
    PR_DISPLAY_TO = 0x0E04
    PR_DISPLAY_CC = 0x0E03
    PR_DISPLAY_BCC = 0x0E02

    # Attachment Properties
    PR_ATTACH_NUM = 0x0E21
    PR_ATTACH_SIZE = 0x0E20
    PR_ATTACH_FILENAME = 0x3704
    PR_ATTACH_LONG_FILENAME = 0x3707
    PR_ATTACH_EXTENSION = 0x3703
    PR_ATTACH_METHOD = 0x3705
    PR_ATTACH_DATA_BIN = 0x3701
    PR_ATTACH_DATA_OBJ = 0x3701
    PR_ATTACH_MIME_TAG = 0x370E
    PR_ATTACH_CONTENT_ID = 0x3712
    PR_ATTACH_CONTENT_LOCATION = 0x3713
    PR_RENDERING_POSITION = 0x370B
    PR_ATTACHMENT_HIDDEN = 0x7FFE
    PR_ATTACHMENT_FLAGS = 0x3714

    # Named Property Mapping Streams
    PR_MAPPING_SIGNATURE = 0x0FF8
    PR_RECORD_KEY = 0x0FF9
    PR_STORE_RECORD_KEY = 0x0FFA
    PR_STORE_ENTRYID = 0x0FFB
    PR_OBJECT_TYPE_PROP = 0x0FFE

    # Exchange Server Properties
    PR_HASATTACH = 0x0E1B
    PR_MESSAGE_CODEPAGE = 0x3FFD
    PR_INTERNET_CPID = 0x3FDE
    PR_MESSAGE_LOCALE_ID = 0x3FF1
    PR_CREATOR_NAME = 0x3FF8
    PR_CREATOR_ENTRYID = 0x3FF9
    PR_LAST_MODIFIER_NAME = 0x3FFA
    PR_LAST_MODIFIER_ENTRYID = 0x3FFB

    # Additional Message Properties
    PR_READ_RECEIPT_REQUESTED = 0x0029
    PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED = 0x0023
    PR_REPLY_RECIPIENT_ENTRIES = 0x004F
    PR_REPLY_RECIPIENT_NAMES = 0x0050

    # Message Status
    PR_MSG_STATUS = 0x0E17
    PR_MESSAGE_FLAGS_2 = 0x0E17


class Property:
    """Represents a MAPI property with tag, type, and value"""

    def __init__(self, tag: int, prop_type: PropertyType, value: Any):
        self.tag = tag
        self.prop_type = prop_type
        self.value = value

    def get_stream_name(self) -> str:
        """
        Get the property stream name for CFB.
        Format: __substg1.0_TTTTIIII
        where TTTT = property tag (4 hex digits)
              IIII = property type (4 hex digits)
        """
        return f"__substg1.0_{self.tag:04X}{self.prop_type:04X}"

    def encode_value(self) -> bytes:
        """Encode property value according to its type"""
        return encode_property_value(self.value, self.prop_type)

    def is_fixed_length(self) -> bool:
        """Check if property is fixed-length (fits in 8 bytes)"""
        return self.prop_type in (
            PropertyType.PT_SHORT,
            PropertyType.PT_LONG,
            PropertyType.PT_FLOAT,
            PropertyType.PT_DOUBLE,
            PropertyType.PT_BOOLEAN,
            PropertyType.PT_LONGLONG,
            PropertyType.PT_SYSTIME,
            PropertyType.PT_ERROR
        )

    def get_fixed_entry(self) -> bytes:
        """
        Get the 16-byte entry for __properties_version1.0 stream.
        Format: 4 bytes property tag + 4 bytes flags + 8 bytes value/size
        """
        prop_tag_combined = (self.prop_type << 16) | self.tag
        flags = 0  # Typically zero

        if self.is_fixed_length():
            # Value fits directly in 8-byte field
            value_bytes = self.encode_value()
            # Pad to 8 bytes
            value_bytes += b'\x00' * (8 - len(value_bytes))
            value_field = value_bytes[:8]
        else:
            # Variable length - store size in value field
            encoded = self.encode_value()
            size = len(encoded)
            value_field = struct.pack('<Q', size)

        return struct.pack('<II', prop_tag_combined, flags) + value_field


def encode_property_value(value: Any, prop_type: PropertyType) -> bytes:
    """
    Encode a property value according to its MAPI property type.
    Returns bytes suitable for writing to property stream.
    """

    if prop_type == PropertyType.PT_UNICODE:
        # Unicode string: UTF-16LE with null terminator
        if isinstance(value, str):
            return value.encode('utf-16le') + b'\x00\x00'
        return b'\x00\x00'

    elif prop_type == PropertyType.PT_STRING8:
        # ASCII string with null terminator
        if isinstance(value, str):
            return value.encode('cp1252') + b'\x00'
        elif isinstance(value, bytes):
            return value + b'\x00'
        return b'\x00'

    elif prop_type == PropertyType.PT_BINARY:
        # Binary data - pass through as-is
        if isinstance(value, bytes):
            return value
        elif isinstance(value, str):
            return value.encode('utf-8')
        return b''

    elif prop_type == PropertyType.PT_LONG:
        # 32-bit signed integer
        if isinstance(value, int):
            return struct.pack('<i', value)
        return struct.pack('<i', 0)

    elif prop_type == PropertyType.PT_SHORT:
        # 16-bit signed integer
        if isinstance(value, int):
            return struct.pack('<h', value)
        return struct.pack('<h', 0)

    elif prop_type == PropertyType.PT_BOOLEAN:
        # Boolean - 1 or 0 as 16-bit value
        if value:
            return struct.pack('<H', 1)
        return struct.pack('<H', 0)

    elif prop_type == PropertyType.PT_LONGLONG:
        # 64-bit signed integer
        if isinstance(value, int):
            return struct.pack('<q', value)
        return struct.pack('<q', 0)

    elif prop_type == PropertyType.PT_SYSTIME:
        # FILETIME - 64-bit value representing 100-nanosecond intervals since Jan 1, 1601 UTC
        if isinstance(value, datetime):
            return datetime_to_filetime(value)
        elif isinstance(value, int):
            return struct.pack('<Q', value)
        return struct.pack('<Q', 0)

    elif prop_type == PropertyType.PT_FLOAT:
        # 32-bit floating point
        if isinstance(value, (int, float)):
            return struct.pack('<f', float(value))
        return struct.pack('<f', 0.0)

    elif prop_type == PropertyType.PT_DOUBLE:
        # 64-bit floating point
        if isinstance(value, (int, float)):
            return struct.pack('<d', float(value))
        return struct.pack('<d', 0.0)

    elif prop_type == PropertyType.PT_ERROR:
        # Error code - 32-bit
        if isinstance(value, int):
            return struct.pack('<I', value)
        return struct.pack('<I', 0)

    else:
        # Unknown type - return empty
        return b''


def datetime_to_filetime(dt: datetime) -> bytes:
    """
    Convert Python datetime to Windows FILETIME (64-bit).
    FILETIME = number of 100-nanosecond intervals since January 1, 1601 UTC
    """
    # Ensure UTC
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    else:
        dt = dt.astimezone(timezone.utc)

    # Epoch for FILETIME
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)

    # Calculate delta in seconds
    delta = dt - epoch

    # Convert to 100-nanosecond intervals
    # 1 second = 10,000,000 intervals
    filetime = int(delta.total_seconds() * 10000000)

    return struct.pack('<Q', filetime)


def filetime_to_datetime(filetime: int) -> datetime:
    """
    Convert Windows FILETIME (64-bit) to Python datetime.
    """
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    # Convert 100-nanosecond intervals to seconds
    seconds = filetime / 10000000
    return epoch + timezone.timedelta(seconds=seconds)


def create_entryid(email: str, display_name: str, addr_type: str = "SMTP") -> bytes:
    """
    Create a simple EntryID for email address.
    This is a simplified version - full implementation would follow MS-OXCDATA spec.
    """
    # Simplified one-off EntryID structure
    # In practice, this should follow the full specification
    flags = 0x00000000
    provider_uid = b'\x00' * 16  # Simplified
    version = 0
    addr_type_bytes = (addr_type + '\x00').encode('ascii')
    email_bytes = (email + '\x00').encode('ascii')
    display_bytes = (display_name + '\x00').encode('ascii')

    return (struct.pack('<I', flags) + provider_uid +
            struct.pack('<I', version) +
            addr_type_bytes + email_bytes + display_bytes)


def create_search_key(addr_type: str, email: str) -> bytes:
    """
    Create search key for email address.
    Format: ADDRTYPE:EMAIL in uppercase
    """
    search_key_str = f"{addr_type}:{email}".upper()
    return search_key_str.encode('ascii') + b'\x00'


def generate_message_id(domain: str = "pymsgkit.local") -> str:
    """
    Generate a unique RFC 5322 compliant Message-ID.
    Format: <uniquestring@domain>
    """
    import uuid
    import time

    # Create unique ID using timestamp and UUID
    timestamp = int(time.time() * 1000000)
    unique_id = uuid.uuid4().hex[:16]

    return f"<{timestamp}.{unique_id}@{domain}>"


def generate_internet_headers(subject: str, sender_email: str, sender_name: str,
                              to_recipients: list, cc_recipients: list = None,
                              message_id: str = None, date: datetime = None) -> str:
    """
    Generate RFC 5322 compliant internet message headers.
    These headers make the MSG file more compatible with online viewers.
    """
    if message_id is None:
        # Extract domain from sender email
        domain = sender_email.split('@')[1] if '@' in sender_email else 'pymsgkit.local'
        message_id = generate_message_id(domain)

    if date is None:
        date = datetime.now(timezone.utc)

    # Format date as RFC 5322
    # Example: Mon, 2 Oct 2025 14:30:00 +0000
    date_str = date.strftime('%a, %d %b %Y %H:%M:%S %z')
    if not date_str.endswith('+0000') and not date_str.endswith('-'):
        date_str += ' +0000'

    # Build headers
    headers = []
    headers.append(f"Date: {date_str}")

    # From header
    if sender_name:
        headers.append(f"From: \"{sender_name}\" <{sender_email}>")
    else:
        headers.append(f"From: {sender_email}")

    # To header
    to_list = []
    for email, name in to_recipients:
        if name:
            to_list.append(f"\"{name}\" <{email}>")
        else:
            to_list.append(email)
    if to_list:
        headers.append(f"To: {', '.join(to_list)}")

    # CC header
    if cc_recipients:
        cc_list = []
        for email, name in cc_recipients:
            if name:
                cc_list.append(f"\"{name}\" <{email}>")
            else:
                cc_list.append(email)
        headers.append(f"Cc: {', '.join(cc_list)}")

    # Subject
    headers.append(f"Subject: {subject}")

    # Message-ID
    headers.append(f"Message-ID: {message_id}")

    # MIME headers
    headers.append("MIME-Version: 1.0")
    headers.append("Content-Type: text/plain; charset=\"utf-8\"")
    headers.append("Content-Transfer-Encoding: quoted-printable")

    # X-Mailer
    headers.append("X-Mailer: PyMsgKit")

    return '\r\n'.join(headers) + '\r\n'
