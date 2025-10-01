"""
Complete MSG file writer implementation
Based on MS-OXMSG specification for creating Outlook-compatible MSG files
"""

import struct
import os
from datetime import datetime, timezone
from typing import List, Dict, Optional
from .cfb import CFBWriter
from .properties import Property, PropertyTag, encode_property_value, create_entryid, create_search_key, datetime_to_filetime
from .types import RecipientType, PropertyType, AttachMethod


class MSGWriter:
    """
    Main class for creating MSG files with full MAPI property support.
    Creates valid Outlook MSG files with proper CFB structure.
    """

    def __init__(self):
        self.cfb = CFBWriter()
        self.properties: Dict[int, Property] = {}
        self.recipients: List[Dict] = []
        self.attachments: List[Dict] = []

        # Set required default properties
        self._set_default_properties()

    def _set_default_properties(self):
        """Set default required properties for MSG file"""
        # Message class - IPM.Note for standard email
        self.set_property(
            PropertyTag.PR_MESSAGE_CLASS,
            PropertyType.PT_UNICODE,
            "IPM.Note"
        )

        # Message flags
        self.set_property(
            PropertyTag.PR_MESSAGE_FLAGS,
            PropertyType.PT_LONG,
            0  # Will be updated based on content
        )

        # Set timestamps
        now = datetime.now(timezone.utc)
        self.set_property(
            PropertyTag.PR_CLIENT_SUBMIT_TIME,
            PropertyType.PT_SYSTIME,
            now
        )
        self.set_property(
            PropertyTag.PR_MESSAGE_DELIVERY_TIME,
            PropertyType.PT_SYSTIME,
            now
        )
        self.set_property(
            PropertyTag.PR_CREATION_TIME,
            PropertyType.PT_SYSTIME,
            now
        )
        self.set_property(
            PropertyTag.PR_LAST_MODIFICATION_TIME,
            PropertyType.PT_SYSTIME,
            now
        )

        # Priority and importance
        self.set_property(PropertyTag.PR_IMPORTANCE, PropertyType.PT_LONG, 1)  # Normal
        self.set_property(PropertyTag.PR_PRIORITY, PropertyType.PT_LONG, 0)    # Normal
        self.set_property(PropertyTag.PR_SENSITIVITY, PropertyType.PT_LONG, 0)  # None

    def set_property(self, tag: int, prop_type: PropertyType, value):
        """Set a MAPI property"""
        prop = Property(tag, prop_type, value)
        self.properties[tag] = prop

    def set_subject(self, subject: str):
        """Set email subject"""
        self.set_property(PropertyTag.PR_SUBJECT, PropertyType.PT_UNICODE, subject)
        # Also set conversation topic (normalized subject)
        normalized_subject = subject
        # Remove RE:, FW: prefixes for conversation topic
        for prefix in ["RE:", "FW:", "Re:", "Fw:", "RE :", "FW :"]:
            if normalized_subject.startswith(prefix):
                normalized_subject = normalized_subject[len(prefix):].strip()
        self.set_property(PropertyTag.PR_CONVERSATION_TOPIC, PropertyType.PT_UNICODE, normalized_subject)

    def set_body(self, body: str, is_html: bool = False):
        """Set email body (plain text or HTML)"""
        if is_html:
            # HTML body stored as PT_BINARY
            self.set_property(
                PropertyTag.PR_HTML,
                PropertyType.PT_BINARY,
                body.encode('utf-8')
            )
        # Always include plain text body
        self.set_property(PropertyTag.PR_BODY, PropertyType.PT_UNICODE, body)

    def set_sender(self, email: str, name: str = "", addr_type: str = "SMTP"):
        """Set sender information (enables custom sender for eDiscovery)"""
        display_name = name if name else email

        # Sender properties
        self.set_property(PropertyTag.PR_SENDER_NAME, PropertyType.PT_UNICODE, display_name)
        self.set_property(PropertyTag.PR_SENDER_EMAIL_ADDRESS, PropertyType.PT_UNICODE, email)
        self.set_property(PropertyTag.PR_SENDER_ADDRTYPE, PropertyType.PT_UNICODE, addr_type)
        self.set_property(PropertyTag.PR_SENDER_SEARCH_KEY, PropertyType.PT_BINARY, create_search_key(addr_type, email))
        self.set_property(PropertyTag.PR_SENDER_ENTRYID, PropertyType.PT_BINARY, create_entryid(email, display_name, addr_type))

        # Sent representing properties (same as sender for normal emails)
        self.set_property(PropertyTag.PR_SENT_REPRESENTING_NAME, PropertyType.PT_UNICODE, display_name)
        self.set_property(PropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS, PropertyType.PT_UNICODE, email)
        self.set_property(PropertyTag.PR_SENT_REPRESENTING_ADDRTYPE, PropertyType.PT_UNICODE, addr_type)
        self.set_property(PropertyTag.PR_SENT_REPRESENTING_SEARCH_KEY, PropertyType.PT_BINARY, create_search_key(addr_type, email))
        self.set_property(PropertyTag.PR_SENT_REPRESENTING_ENTRYID, PropertyType.PT_BINARY, create_entryid(email, display_name, addr_type))

    def add_recipient(self, email: str, name: str = "", recipient_type: RecipientType = RecipientType.TO):
        """Add a recipient (To, Cc, or Bcc)"""
        display_name = name if name else email

        recipient = {
            'email': email,
            'name': display_name,
            'type': recipient_type,
            'addr_type': 'SMTP'
        }
        self.recipients.append(recipient)

    def add_attachment(self, filename: str, data: bytes, content_id: str = None,
                       mime_type: str = None, is_inline: bool = False):
        """Add an attachment (regular or inline)"""
        attachment = {
            'filename': filename,
            'data': data,
            'content_id': content_id,
            'mime_type': mime_type or 'application/octet-stream',
            'is_inline': is_inline,
            'method': AttachMethod.BY_VALUE
        }
        self.attachments.append(attachment)

    def set_conversation_index(self, parent_index: bytes = None):
        """Set conversation index for email threading"""
        if parent_index is None:
            # Create new conversation index (22 bytes)
            # Byte 0: reserved (0x01)
            # Bytes 1-5: FILETIME compressed (current time)
            # Bytes 6-21: GUID (16 bytes)
            import time

            # Current time as FILETIME
            now = datetime.now(timezone.utc)
            filetime_bytes = datetime_to_filetime(now)
            # Take first 5 bytes (compressed FILETIME)
            compressed_time = filetime_bytes[0:5]

            # Generate GUID
            guid = os.urandom(16)

            # Build conversation index
            index = b'\x01' + compressed_time + guid
        else:
            # Reply - append 5-byte child block to parent index
            import time
            # Time delta as 5 bytes
            time_delta = os.urandom(5)  # Simplified
            index = parent_index + time_delta

        self.set_property(
            PropertyTag.PR_CONVERSATION_INDEX,
            PropertyType.PT_BINARY,
            index
        )

    def save(self, filepath: str):
        """Save MSG file to disk"""
        # Update message flags based on content
        flags = 0
        if self.attachments:
            flags |= 0x00000010  # MSGFLAG_HASATTACH
        flags |= 0x00000001  # MSGFLAG_READ (default to read)
        self.set_property(PropertyTag.PR_MESSAGE_FLAGS, PropertyType.PT_LONG, flags)

        # Update display recipient properties
        self._update_display_recipients()

        # Write all properties to CFB
        self._write_properties()

        # Write recipients
        for idx, recipient in enumerate(self.recipients):
            self._write_recipient(idx, recipient)

        # Write attachments
        for idx, attachment in enumerate(self.attachments):
            self._write_attachment(idx, attachment)

        # Write CFB to file
        self.cfb.write(filepath)

    def _update_display_recipients(self):
        """Update PR_DISPLAY_TO, PR_DISPLAY_CC, PR_DISPLAY_BCC"""
        to_names = []
        cc_names = []
        bcc_names = []

        for recipient in self.recipients:
            if recipient['type'] == RecipientType.TO:
                to_names.append(recipient['name'])
            elif recipient['type'] == RecipientType.CC:
                cc_names.append(recipient['name'])
            elif recipient['type'] == RecipientType.BCC:
                bcc_names.append(recipient['name'])

        if to_names:
            self.set_property(PropertyTag.PR_DISPLAY_TO, PropertyType.PT_UNICODE, '; '.join(to_names))
        if cc_names:
            self.set_property(PropertyTag.PR_DISPLAY_CC, PropertyType.PT_UNICODE, '; '.join(cc_names))
        if bcc_names:
            self.set_property(PropertyTag.PR_DISPLAY_BCC, PropertyType.PT_UNICODE, '; '.join(bcc_names))

    def _write_properties(self):
        """Write all message properties to CFB"""
        # Build __properties_version1.0 stream
        properties_data = bytearray()

        # Reserved header (8 bytes of zeros)
        properties_data.extend(b'\x00' * 8)

        # Add all fixed-length property entries
        for tag, prop in sorted(self.properties.items()):
            if prop.is_fixed_length():
                properties_data.extend(prop.get_fixed_entry())

        # Write __properties_version1.0 stream
        if properties_data:
            self.cfb.add_stream("__properties_version1.0", bytes(properties_data))

        # Write variable-length property streams
        for tag, prop in self.properties.items():
            if not prop.is_fixed_length():
                stream_name = prop.get_stream_name()
                stream_data = prop.encode_value()
                self.cfb.add_stream(stream_name, stream_data)

    def _write_recipient(self, idx: int, recipient: Dict):
        """Write recipient storage to CFB"""
        # Create recipient storage
        storage_name = f"__recip_version1.0_#{idx:08X}"
        storage_did = self.cfb.add_storage(storage_name)

        # Build recipient properties
        recip_props = {}

        # Required recipient properties
        recip_props[PropertyTag.PR_RECIPIENT_TYPE] = Property(
            PropertyTag.PR_RECIPIENT_TYPE, PropertyType.PT_LONG, int(recipient['type'])
        )
        recip_props[PropertyTag.PR_DISPLAY_NAME] = Property(
            PropertyTag.PR_DISPLAY_NAME, PropertyType.PT_UNICODE, recipient['name']
        )
        recip_props[PropertyTag.PR_EMAIL_ADDRESS] = Property(
            PropertyTag.PR_EMAIL_ADDRESS, PropertyType.PT_UNICODE, recipient['email']
        )
        recip_props[PropertyTag.PR_ADDRTYPE] = Property(
            PropertyTag.PR_ADDRTYPE, PropertyType.PT_UNICODE, recipient['addr_type']
        )
        recip_props[PropertyTag.PR_SMTP_ADDRESS] = Property(
            PropertyTag.PR_SMTP_ADDRESS, PropertyType.PT_UNICODE, recipient['email']
        )
        recip_props[PropertyTag.PR_SEARCH_KEY] = Property(
            PropertyTag.PR_SEARCH_KEY, PropertyType.PT_BINARY,
            create_search_key(recipient['addr_type'], recipient['email'])
        )
        recip_props[PropertyTag.PR_ENTRYID] = Property(
            PropertyTag.PR_ENTRYID, PropertyType.PT_BINARY,
            create_entryid(recipient['email'], recipient['name'], recipient['addr_type'])
        )

        # Write recipient __properties_version1.0
        recip_properties_data = bytearray(b'\x00' * 8)  # Reserved header
        for tag, prop in sorted(recip_props.items()):
            if prop.is_fixed_length():
                recip_properties_data.extend(prop.get_fixed_entry())

        self.cfb.add_stream("__properties_version1.0", bytes(recip_properties_data), storage_did)

        # Write variable-length property streams
        for tag, prop in recip_props.items():
            if not prop.is_fixed_length():
                stream_name = prop.get_stream_name()
                stream_data = prop.encode_value()
                self.cfb.add_stream(stream_name, stream_data, storage_did)

    def _write_attachment(self, idx: int, attachment: Dict):
        """Write attachment storage to CFB"""
        # Create attachment storage
        storage_name = f"__attach_version1.0_#{idx:08X}"
        storage_did = self.cfb.add_storage(storage_name)

        # Build attachment properties
        attach_props = {}

        # Attachment method
        attach_props[PropertyTag.PR_ATTACH_METHOD] = Property(
            PropertyTag.PR_ATTACH_METHOD, PropertyType.PT_LONG, int(attachment['method'])
        )

        # Attachment size
        attach_props[PropertyTag.PR_ATTACH_SIZE] = Property(
            PropertyTag.PR_ATTACH_SIZE, PropertyType.PT_LONG, len(attachment['data'])
        )

        # Attachment filename
        attach_props[PropertyTag.PR_ATTACH_LONG_FILENAME] = Property(
            PropertyTag.PR_ATTACH_LONG_FILENAME, PropertyType.PT_UNICODE, attachment['filename']
        )
        attach_props[PropertyTag.PR_ATTACH_FILENAME] = Property(
            PropertyTag.PR_ATTACH_FILENAME, PropertyType.PT_UNICODE, attachment['filename']
        )

        # Extension
        ext = os.path.splitext(attachment['filename'])[1]
        if ext:
            attach_props[PropertyTag.PR_ATTACH_EXTENSION] = Property(
                PropertyTag.PR_ATTACH_EXTENSION, PropertyType.PT_UNICODE, ext
            )

        # MIME type
        attach_props[PropertyTag.PR_ATTACH_MIME_TAG] = Property(
            PropertyTag.PR_ATTACH_MIME_TAG, PropertyType.PT_UNICODE, attachment['mime_type']
        )

        # Content ID (for inline attachments)
        if attachment['content_id']:
            attach_props[PropertyTag.PR_ATTACH_CONTENT_ID] = Property(
                PropertyTag.PR_ATTACH_CONTENT_ID, PropertyType.PT_UNICODE, attachment['content_id']
            )

        # Rendering position (for inline images)
        if attachment['is_inline']:
            attach_props[PropertyTag.PR_RENDERING_POSITION] = Property(
                PropertyTag.PR_RENDERING_POSITION, PropertyType.PT_LONG, -1
            )
            attach_props[PropertyTag.PR_ATTACHMENT_HIDDEN] = Property(
                PropertyTag.PR_ATTACHMENT_HIDDEN, PropertyType.PT_BOOLEAN, True
            )

        # Attachment data
        attach_props[PropertyTag.PR_ATTACH_DATA_BIN] = Property(
            PropertyTag.PR_ATTACH_DATA_BIN, PropertyType.PT_BINARY, attachment['data']
        )

        # Attachment number
        attach_props[PropertyTag.PR_ATTACH_NUM] = Property(
            PropertyTag.PR_ATTACH_NUM, PropertyType.PT_LONG, idx
        )

        # Write attachment __properties_version1.0
        attach_properties_data = bytearray(b'\x00' * 8)  # Reserved header
        for tag, prop in sorted(attach_props.items()):
            if prop.is_fixed_length():
                attach_properties_data.extend(prop.get_fixed_entry())

        self.cfb.add_stream("__properties_version1.0", bytes(attach_properties_data), storage_did)

        # Write variable-length property streams
        for tag, prop in attach_props.items():
            if not prop.is_fixed_length():
                stream_name = prop.get_stream_name()
                stream_data = prop.encode_value()
                self.cfb.add_stream(stream_name, stream_data, storage_did)
