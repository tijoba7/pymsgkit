"""
Unit tests for the property encoding layer.
"""

import struct
from datetime import datetime, timezone

import pytest

from pymsgkit.properties import (
    Property,
    PropertyTag,
    encode_property_value,
    datetime_to_filetime,
    filetime_to_datetime,
    create_search_key,
    create_entryid,
    generate_message_id,
    generate_internet_headers,
)
from pymsgkit.types import PropertyType


def test_unicode_stream_has_no_null_terminator():
    # MSG string streams store the raw string; the null terminator is omitted
    # from the stream and only counted in the property-table size.
    out = encode_property_value("Hi", PropertyType.PT_UNICODE)
    assert out == "Hi".encode("utf-16le")
    assert not out.endswith(b"\x00\x00")


def test_unicode_table_size_counts_terminator():
    p = Property(PropertyTag.PR_SUBJECT, PropertyType.PT_UNICODE, "Hi")
    size = struct.unpack("<I", p.get_entry()[8:12])[0]
    # 2 chars * 2 bytes + 2-byte UTF-16 null terminator
    assert size == 6


def test_string8_replaces_unmappable_chars():
    # CJK cannot be represented in cp1252; must not raise and must not append a
    # stream null terminator.
    out = encode_property_value("会", PropertyType.PT_STRING8)
    assert isinstance(out, bytes) and not out.endswith(b"\x00")


def test_long_and_short_encoding():
    assert encode_property_value(-5, PropertyType.PT_LONG) == struct.pack("<i", -5)
    assert encode_property_value(7, PropertyType.PT_SHORT) == struct.pack("<h", 7)


def test_boolean_encoding():
    assert encode_property_value(True, PropertyType.PT_BOOLEAN) == struct.pack("<H", 1)
    assert encode_property_value(False, PropertyType.PT_BOOLEAN) == struct.pack("<H", 0)


def test_binary_passthrough():
    assert encode_property_value(b"\x00\x01\x02", PropertyType.PT_BINARY) == b"\x00\x01\x02"


@pytest.mark.parametrize("dt", [
    datetime(1601, 1, 1, tzinfo=timezone.utc),
    datetime(2000, 1, 1, tzinfo=timezone.utc),
    datetime(2026, 7, 1, 12, 30, 45, tzinfo=timezone.utc),
])
def test_filetime_roundtrip(dt):
    raw = datetime_to_filetime(dt)
    ft = struct.unpack("<Q", raw)[0]
    back = filetime_to_datetime(ft)
    # sub-second precision is lossy; compare to the second.
    assert abs((back - dt).total_seconds()) < 1


def test_naive_datetime_treated_as_utc():
    naive = datetime(2020, 5, 5, 10, 0, 0)
    aware = datetime(2020, 5, 5, 10, 0, 0, tzinfo=timezone.utc)
    assert datetime_to_filetime(naive) == datetime_to_filetime(aware)


def test_search_key_format():
    assert create_search_key("SMTP", "a@b.com") == b"SMTP:A@B.COM\x00"


def test_search_key_non_ascii_does_not_raise():
    key = create_search_key("SMTP", "田中@example.jp")
    assert key.endswith(b"\x00")


def test_entryid_is_one_off_with_unicode_addresses():
    eid = create_entryid("user@x.com", "User Name", "SMTP")
    # One-Off EntryID provider UID (MS-OXCDATA 2.2.5.1).
    assert eid[4:20] == bytes.fromhex('812B1FA4BEA310199D6E00DD010F5402')
    # Strings are UTF-16LE with the Unicode flag set.
    assert "user@x.com".encode("utf-16le") in eid
    assert "SMTP".encode("utf-16le") in eid


def test_entryid_non_ascii_does_not_raise():
    eid = create_entryid("田中@example.jp", "田中", "SMTP")
    assert isinstance(eid, bytes) and len(eid) > 0


def test_message_id_format():
    mid = generate_message_id("example.com")
    assert mid.startswith("<") and mid.endswith("@example.com>")


def test_message_ids_are_unique():
    assert generate_message_id("x.com") != generate_message_id("x.com")


def test_property_entry_is_16_bytes():
    p = Property(PropertyTag.PR_SUBJECT, PropertyType.PT_UNICODE, "hello")
    assert len(p.get_entry()) == 16


def test_fixed_length_classification():
    assert Property(PropertyTag.PR_MESSAGE_FLAGS, PropertyType.PT_LONG, 1).is_fixed_length()
    assert not Property(PropertyTag.PR_SUBJECT, PropertyType.PT_UNICODE, "x").is_fixed_length()


def test_stream_name_format():
    p = Property(PropertyTag.PR_SUBJECT, PropertyType.PT_UNICODE, "x")
    assert p.get_stream_name() == "__substg1.0_0037001F"


def test_internet_headers_contain_core_fields():
    headers = generate_internet_headers(
        subject="Hi",
        sender_email="a@b.com",
        sender_name="Alice",
        to_recipients=[("c@d.com", "Carol")],
        cc_recipients=[("e@f.com", "Eve")],
    )
    assert "From:" in headers and "a@b.com" in headers
    assert "To:" in headers and "c@d.com" in headers
    assert "Cc:" in headers and "e@f.com" in headers
    assert "Subject: Hi" in headers
    assert "Message-ID:" in headers
