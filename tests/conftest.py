"""
Shared test helpers.

The whole point of these helpers is to read the MSG files PyMsgKit produces
*back* with an independent parser (olefile) and decode the MAPI properties, so
the tests validate real output rather than just "a file was written".
"""

import struct
import pytest
import olefile

from pymsgkit.types import PropertyType


# CFB signature that every valid MSG/compound file must start with.
CFB_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


def _stream_name(tag: int, prop_type: int) -> str:
    return f"__substg1.0_{tag:04X}{prop_type:04X}"


class MsgReader:
    """Minimal read-back parser for MSG files, built on olefile.

    Only implements what the tests need: decoding variable-length property
    streams (strings/binary) and enumerating recipient/attachment storages.
    """

    def __init__(self, path):
        self.ole = olefile.OleFileIO(str(path))

    def close(self):
        self.ole.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # -- structural --------------------------------------------------------
    def paths(self):
        return ["/".join(p) for p in self.ole.listdir()]

    def exists(self, path):
        return self.ole.exists(path)

    def storages(self, prefix):
        """Return the set of top-level storage names starting with prefix."""
        names = set()
        for parts in self.ole.listdir():
            if parts[0].startswith(prefix):
                names.add(parts[0])
        return names

    # -- property access ---------------------------------------------------
    def _read(self, storage, tag, prop_type):
        name = _stream_name(tag, prop_type)
        path = f"{storage}/{name}" if storage else name
        if not self.ole.exists(path):
            return None
        return self.ole.openstream(path).read()

    def unicode(self, tag, storage=""):
        data = self._read(storage, tag, PropertyType.PT_UNICODE)
        if data is None:
            return None
        return data.decode("utf-16le").rstrip("\x00")

    def string8(self, tag, storage=""):
        data = self._read(storage, tag, PropertyType.PT_STRING8)
        if data is None:
            return None
        return data.decode("cp1252", errors="replace").rstrip("\x00")

    def binary(self, tag, storage=""):
        return self._read(storage, tag, PropertyType.PT_BINARY)

    def fixed_props(self, storage=""):
        """Decode the fixed-length entries from __properties_version1.0.

        Returns {tag: raw_8_byte_value}. Header size differs between the
        top-level stream (8-byte reserved) and sub-storages (8-byte reserved
        too, per the writer); recipients/attachments also use an 8-byte header.
        """
        path = f"{storage}/__properties_version1.0" if storage else "__properties_version1.0"
        if not self.ole.exists(path):
            return {}
        data = self.ole.openstream(path).read()
        # Top-level header is 32 bytes (MS-OXMSG 2.4.1.1); recipient/attachment
        # sub storages use an 8-byte reserved header only.
        header = 32 if storage == "" else 8
        props = {}
        body = data[header:]
        for i in range(0, len(body) - 15, 16):
            entry = body[i:i + 16]
            combined = struct.unpack("<I", entry[0:4])[0]
            # MAPI tag packs PropId in the high word, PropType in the low word.
            tag = (combined >> 16) & 0xFFFF
            props[tag] = entry[8:16]
        return props


@pytest.fixture
def read_msg():
    """Fixture returning a factory that opens a saved MSG for inspection."""
    readers = []

    def _open(path):
        r = MsgReader(path)
        readers.append(r)
        return r

    yield _open

    for r in readers:
        r.close()
