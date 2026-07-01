"""
Tests for the Compound File Binary writer, exercised directly and through
the MSG writer. Validation is done with the independent olefile parser.
"""

import struct

import olefile
import pytest

from pymsgkit.cfb import CFBWriter, SectorType
from conftest import CFB_SIGNATURE


def test_empty_compound_file(tmp_path):
    """A file with only the root entry must still be a valid compound file."""
    path = tmp_path / "empty.cfb"
    cfb = CFBWriter()
    cfb.write(str(path))
    assert olefile.isOleFile(str(path))


def test_single_small_stream(tmp_path):
    path = tmp_path / "s.cfb"
    cfb = CFBWriter()
    cfb.add_stream("hello", b"world")
    cfb.write(str(path))

    ole = olefile.OleFileIO(str(path))
    try:
        assert ole.exists("hello")
        assert ole.openstream("hello").read() == b"world"
    finally:
        ole.close()


def test_nested_storage(tmp_path):
    path = tmp_path / "n.cfb"
    cfb = CFBWriter()
    did = cfb.add_storage("folder")
    cfb.add_stream("inner", b"data", parent_did=did)
    cfb.write(str(path))

    ole = olefile.OleFileIO(str(path))
    try:
        assert ole.exists("folder/inner")
        assert ole.openstream("folder/inner").read() == b"data"
    finally:
        ole.close()


def test_many_siblings_all_readable(tmp_path):
    """Sibling chaining must keep every stream reachable."""
    path = tmp_path / "many.cfb"
    cfb = CFBWriter()
    expected = {}
    for i in range(25):
        name = f"stream_{i:02d}"
        data = f"payload-{i}".encode()
        cfb.add_stream(name, data)
        expected[name] = data
    cfb.write(str(path))

    ole = olefile.OleFileIO(str(path))
    try:
        for name, data in expected.items():
            assert ole.exists(name), f"{name} missing"
            assert ole.openstream(name).read() == data
    finally:
        ole.close()


def test_mini_and_regular_streams_coexist(tmp_path):
    """A file mixing sub-cutoff (mini) and above-cutoff (regular) streams."""
    path = tmp_path / "mix.cfb"
    cfb = CFBWriter()
    small = b"tiny"
    large = bytes(range(256)) * 30  # 7680 bytes > 4096 cutoff
    cfb.add_stream("small", small)
    cfb.add_stream("large", large)
    cfb.write(str(path))

    ole = olefile.OleFileIO(str(path))
    try:
        assert ole.openstream("small").read() == small
        assert ole.openstream("large").read() == large
    finally:
        ole.close()


def test_stream_spanning_multiple_regular_sectors(tmp_path):
    path = tmp_path / "big.cfb"
    cfb = CFBWriter()
    data = bytes((i * 7) % 256 for i in range(200000))  # ~200 KB, many sectors
    cfb.add_stream("blob", data)
    cfb.write(str(path))

    ole = olefile.OleFileIO(str(path))
    try:
        assert ole.openstream("blob").read() == data
    finally:
        ole.close()


def test_difat_large_file(tmp_path):
    """A stream past the 109-FAT-sector (~7 MB) limit needs DIFAT sectors.

    Before DIFAT support this produced a silently corrupt, unreadable file.
    """
    path = tmp_path / "difat.cfb"
    cfb = CFBWriter()
    data = bytes((i * 13) % 256 for i in range(9 * 1024 * 1024))  # 9 MB
    cfb.add_stream("blob", data)
    cfb.write(str(path))

    assert olefile.isOleFile(str(path))
    ole = olefile.OleFileIO(str(path))
    try:
        assert ole.openstream("blob").read() == data
    finally:
        ole.close()

    # Header must advertise a DIFAT chain (num DIFAT sectors > 0 at offset 72).
    raw = path.read_bytes()
    assert struct.unpack("<I", raw[72:76])[0] >= 1


def test_header_fields(tmp_path):
    path = tmp_path / "h.cfb"
    cfb = CFBWriter()
    cfb.add_stream("x", b"y")
    cfb.write(str(path))

    raw = path.read_bytes()
    assert raw[:8] == CFB_SIGNATURE
    # byte order marker 0xFFFE at offset 28
    assert raw[28:30] == b"\xFE\xFF"
    # sector shift 0x0009 (512-byte sectors) at offset 30
    assert struct.unpack("<H", raw[30:32])[0] == 0x0009
    # mini sector shift 0x0006 (64-byte) at offset 32
    assert struct.unpack("<H", raw[32:34])[0] == 0x0006
    # mini stream cutoff 4096 at offset 56
    assert struct.unpack("<I", raw[56:60])[0] == 4096
