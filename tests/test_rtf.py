"""
Tests for the synthesized RTF body (PR_RTF_COMPRESSED, uncompressed form).
"""

import struct

import pytest

from pymsgkit import MSGWriter
from pymsgkit.properties import PropertyTag, build_uncompressed_rtf


def test_uncompressed_rtf_header():
    data = build_uncompressed_rtf("hello")
    comp_size, raw_size, comp_type, crc = struct.unpack("<IIII", data[:16])
    assert comp_type == 0x414C454D          # 'MELA' = uncompressed
    assert crc == 0
    assert raw_size == len(data) - 16
    assert comp_size == raw_size + 12
    assert data[16:].startswith(b"{\\rtf1")


def test_rtf_escapes_special_chars():
    data = build_uncompressed_rtf("a{b}c\\d")
    body = data[16:]
    assert b"\\{" in body and b"\\}" in body and b"\\\\" in body


def test_rtf_newlines_become_par():
    body = build_uncompressed_rtf("one\ntwo")[16:]
    assert b"\\par" in body


def test_rtf_unicode_escaped():
    body = build_uncompressed_rtf("café ☕")[16:]
    # Non-ASCII becomes \uN? control words; the raw bytes stay ASCII.
    assert b"\\u" in body
    body.decode("ascii")  # must not raise


def test_plain_body_gets_rtf(tmp_path, read_msg):
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("plain text body")
    msg.save(str(tmp_path / "m.msg"))

    r = read_msg(tmp_path / "m.msg")
    rtf = r.binary(PropertyTag.PR_RTF_COMPRESSED)
    assert rtf is not None
    assert struct.unpack("<I", rtf[8:12])[0] == 0x414C454D


def test_html_body_has_no_rtf_override(tmp_path, read_msg):
    """HTML messages display from PR_HTML; we must not force a plain RTF body."""
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com")
    msg.add_recipient("c@d.com")
    msg.set_body("<p>hello</p>", is_html=True)
    msg.save(str(tmp_path / "m.msg"))

    r = read_msg(tmp_path / "m.msg")
    assert r.binary(PropertyTag.PR_RTF_COMPRESSED) is None
    assert r.binary(PropertyTag.PR_HTML) is not None


def test_extract_msg_decompresses_rtf(tmp_path):
    extract_msg = pytest.importorskip("extract_msg")
    msg = MSGWriter()
    msg.set_subject("s")
    msg.set_sender("a@b.com", "Alice")
    msg.add_recipient("c@d.com", "Carol")
    msg.set_body("body with {braces} and a backslash \\ end")
    path = tmp_path / "m.msg"
    msg.save(str(path))

    m = extract_msg.Message(str(path))
    try:
        assert b"\\rtf1" in m.rtfBody
    finally:
        m.close()
