"""
Complete MSG file writer implementation
Based on MS-OXMSG specification for creating Outlook-compatible MSG files
"""

import struct
import os
from datetime import datetime, timezone
from typing import List, Dict, Optional
from .cfb import CFBWriter
from .properties import (Property, PropertyTag, encode_property_value, create_entryid,
                        create_search_key, datetime_to_filetime, generate_message_id,
                        generate_internet_headers)
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

        # Exchange Server properties
        self.set_property(PropertyTag.PR_HASATTACH, PropertyType.PT_BOOLEAN, False)
        self.set_property(PropertyTag.PR_MESSAGE_CODEPAGE, PropertyType.PT_LONG, 65001)  # UTF-8
        self.set_property(PropertyTag.PR_INTERNET_CPID, PropertyType.PT_LONG, 65001)  # UTF-8
        self.set_property(PropertyTag.PR_MESSAGE_LOCALE_ID, PropertyType.PT_LONG, 0x0409)  # en-US
        self.set_property(PropertyTag.PR_STORE_SUPPORT_MASK, PropertyType.PT_LONG, 0x00040000)

        # Additional message properties
        self.set_property(PropertyTag.PR_READ_RECEIPT_REQUESTED, PropertyType.PT_BOOLEAN, False)
        self.set_property(PropertyTag.PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED, PropertyType.PT_BOOLEAN, False)
        self.set_property(PropertyTag.PR_MSG_STATUS, PropertyType.PT_LONG, 0)

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
            self.set_property(PropertyTag.PR_HASATTACH, PropertyType.PT_BOOLEAN, True)
        flags |= 0x00000001  # MSGFLAG_READ (default to read)
        self.set_property(PropertyTag.PR_MESSAGE_FLAGS, PropertyType.PT_LONG, flags)

        # Generate and add internet headers for better compatibility
        self._add_internet_headers()

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

        # Write named properties structure (required by some readers even if empty)
        self._write_named_properties()

        # Write CFB to file
        self.cfb.write(filepath)

    def _add_internet_headers(self):
        """Add internet message headers and Message-ID for compatibility"""
        # Get sender info
        sender_email = ""
        sender_name = ""
        if PropertyTag.PR_SENDER_EMAIL_ADDRESS in self.properties:
            sender_email = self.properties[PropertyTag.PR_SENDER_EMAIL_ADDRESS].value
        if PropertyTag.PR_SENDER_NAME in self.properties:
            sender_name = self.properties[PropertyTag.PR_SENDER_NAME].value

        # Get subject
        subject = ""
        if PropertyTag.PR_SUBJECT in self.properties:
            subject = self.properties[PropertyTag.PR_SUBJECT].value

        # Generate Message-ID
        domain = sender_email.split('@')[1] if '@' in sender_email else 'pymsgkit.local'
        message_id = generate_message_id(domain)
        self.set_property(PropertyTag.PR_INTERNET_MESSAGE_ID, PropertyType.PT_UNICODE, message_id)

        # Collect recipients by type
        to_recips = [(r['email'], r['name']) for r in self.recipients if r['type'] == RecipientType.TO]
        cc_recips = [(r['email'], r['name']) for r in self.recipients if r['type'] == RecipientType.CC]

        # Get timestamp
        send_time = datetime.now(timezone.utc)
        if PropertyTag.PR_CLIENT_SUBMIT_TIME in self.properties:
            send_time = self.properties[PropertyTag.PR_CLIENT_SUBMIT_TIME].value

        # Generate internet headers
        if sender_email and to_recips:
            headers = generate_internet_headers(
                subject=subject,
                sender_email=sender_email,
                sender_name=sender_name,
                to_recipients=to_recips,
                cc_recipients=cc_recips if cc_recips else None,
                message_id=message_id,
                date=send_time
            )
            self.set_property(PropertyTag.PR_TRANSPORT_MESSAGE_HEADERS, PropertyType.PT_UNICODE, headers)

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
        # Build __properties_version1.0 stream with 32-byte header
        properties_data = bytearray()

        recipient_count = len(self.recipients)
        attachment_count = len(self.attachments)

        # 8-byte reserved block followed by counts and next IDs (per MS-OXMSG 2.4.1)
        properties_data.extend(b'\x00' * 8)
        properties_data.extend(struct.pack(
            '<IIII',
            recipient_count,
            attachment_count,
            recipient_count,
            attachment_count
        ))

        variable_streams = []

        # Add all property entries (fixed and variable)
        for tag, prop in sorted(self.properties.items()):
            encoded_value = prop.encode_value()
            properties_data.extend(prop.get_property_entry(encoded_value))

            if not prop.is_fixed_length():
                variable_streams.append((prop.get_stream_name(), encoded_value))

        # Write __properties_version1.0 stream
        self.cfb.add_stream("__properties_version1.0", bytes(properties_data))

        # Write variable-length property streams
        for stream_name, stream_data in variable_streams:
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
        recip_variable_streams = []
        for tag, prop in sorted(recip_props.items()):
            encoded_value = prop.encode_value()
            recip_properties_data.extend(prop.get_property_entry(encoded_value))
            if not prop.is_fixed_length():
                recip_variable_streams.append((prop.get_stream_name(), encoded_value))

        self.cfb.add_stream("__properties_version1.0", bytes(recip_properties_data), storage_did)

        # Write variable-length property streams
        for stream_name, stream_data in recip_variable_streams:
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
        attach_variable_streams = []
        for tag, prop in sorted(attach_props.items()):
            encoded_value = prop.encode_value()
            attach_properties_data.extend(prop.get_property_entry(encoded_value))
            if not prop.is_fixed_length():
                attach_variable_streams.append((prop.get_stream_name(), encoded_value))

        self.cfb.add_stream("__properties_version1.0", bytes(attach_properties_data), storage_did)

        # Write variable-length property streams
        for stream_name, stream_data in attach_variable_streams:
            self.cfb.add_stream(stream_name, stream_data, storage_did)

    def _write_named_properties(self):
        """
        Write __nameid_version1.0 storage with required streams.
        This is required by some MSG readers even if we don't use named properties.
        Creates minimal valid structure.
        """
        # Create __nameid_version1.0 storage
        nameid_storage = self.cfb.add_storage("__nameid_version1.0")

        # GUID stream (__substg1.0_00020102) - stores property set GUIDs (16 bytes each)
        # Add a placeholder GUID (PS_MAPI - all zeros is valid but unused)
        guid_stream = b'\x00' * 16  # One GUID (all zeros = PS_MAPI placeholder)
        self.cfb.add_stream("__substg1.0_00020102", guid_stream, nameid_storage)

        # Entry stream (__substg1.0_00030102) - stores named property entries
        # Format per entry: 4 bytes (name offset/id) + 2 bytes (GUID index) + 2 bytes (property type/kind)
        # Add one placeholder entry
        entry_stream = struct.pack('<I', 0) + struct.pack('<H', 0) + struct.pack('<H', 0)  # 8 bytes
        self.cfb.add_stream("__substg1.0_00030102", entry_stream, nameid_storage)

        # String stream (__substg1.0_00040102) - stores string names (optional)
        # Only needed if we have string-named properties
        # We can skip this for now as it's optional
