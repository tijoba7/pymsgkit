"""
Complete Compound File Binary (CFB) format implementation
Based on MS-CFB specification for creating valid MSG files
"""

import struct
import io
from datetime import datetime, timezone
from typing import List, Dict, Optional, BinaryIO
from enum import IntEnum


class SectorType(IntEnum):
    """CFB sector type markers"""
    MAXREGSECT = 0xFFFFFFFA
    DIFSECT = 0xFFFFFFFC
    FATSECT = 0xFFFFFFFD
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF


class EntryType(IntEnum):
    """Directory entry types"""
    EMPTY = 0
    STORAGE = 1
    STREAM = 2
    ROOT = 5


class Color(IntEnum):
    """Red-black tree colors (simplified - all black)"""
    RED = 0
    BLACK = 1


class DirectoryEntry:
    """CFB Directory Entry (128 bytes)"""

    NOSTREAM = 0xFFFFFFFF

    def __init__(self, name: str = "", entry_type: EntryType = EntryType.EMPTY):
        self.name = name[:31]  # Max 31 characters
        self.entry_type = entry_type
        self.color = Color.BLACK
        self.left_sibling = DirectoryEntry.NOSTREAM
        self.right_sibling = DirectoryEntry.NOSTREAM
        self.child = DirectoryEntry.NOSTREAM
        self.clsid = b'\x00' * 16
        self.state_bits = 0
        self.creation_time = 0
        self.modified_time = 0
        self.starting_sector = 0
        self.stream_size = 0

    def to_bytes(self) -> bytes:
        """Serialize directory entry to 128 bytes"""
        # Name as UTF-16LE, padded to 64 bytes
        name_bytes = self.name.encode('utf-16le')
        name_len = len(name_bytes) + 2  # Include null terminator
        name_field = name_bytes + b'\x00\x00' + b'\x00' * (64 - len(name_bytes) - 2)

        entry = struct.pack(
            '<64sHBBIII16sIQQII',
            name_field,              # 64 bytes: name in UTF-16LE
            name_len,                # 2 bytes: name length including terminator
            self.entry_type,         # 1 byte: entry type
            self.color,              # 1 byte: color
            self.left_sibling,       # 4 bytes: left sibling DID
            self.right_sibling,      # 4 bytes: right sibling DID
            self.child,              # 4 bytes: child DID
            self.clsid,              # 16 bytes: CLSID
            self.state_bits,         # 4 bytes: state bits
            self.creation_time,      # 8 bytes: creation time
            self.modified_time,      # 8 bytes: modified time
            self.starting_sector,    # 4 bytes: starting sector
            self.stream_size & 0xFFFFFFFF  # 4 bytes: stream size (low 32 bits for v3)
        )
        # Pad to 128 bytes (entry is currently 120 bytes, need 8 more)
        entry += struct.pack('<I', (self.stream_size >> 32) & 0xFFFFFFFF)  # High 32 bits (for v4)
        entry += b'\x00' * 4  # Reserved

        # Ensure exactly 128 bytes
        return entry[:128]


class CFBWriter:
    """
    Complete CFB writer with proper sector management, FAT chains, and directory structure.
    Creates valid Compound File Binary files that Outlook can read.
    """

    SECTOR_SIZE = 512
    MINI_SECTOR_SIZE = 64
    MINI_STREAM_CUTOFF = 4096
    HEADER_SIGNATURE = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

    def __init__(self):
        self.sector_size = self.SECTOR_SIZE
        self.mini_sector_size = self.MINI_SECTOR_SIZE
        self.directory_entries: List[DirectoryEntry] = []
        self.streams: Dict[int, bytes] = {}  # DID -> stream data
        self.fat: List[int] = []  # File Allocation Table
        self.mini_fat: List[int] = []  # Mini FAT for small streams
        self.mini_stream_data = bytearray()  # Mini stream container

        # Create root entry (always at index 0)
        root = DirectoryEntry("Root Entry", EntryType.ROOT)
        # starting_sector will be set later if we have mini streams
        self.directory_entries.append(root)

    def add_storage(self, name: str, parent_did: int = 0) -> int:
        """Add a storage (directory) entry"""
        entry = DirectoryEntry(name, EntryType.STORAGE)
        entry.starting_sector = SectorType.ENDOFCHAIN  # Storages have no data stream
        did = len(self.directory_entries)
        self.directory_entries.append(entry)

        # Link to parent's child chain (simplified - just use child pointer)
        parent = self.directory_entries[parent_did]
        if parent.child == DirectoryEntry.NOSTREAM:
            parent.child = did
        else:
            # Add as right sibling of the first child
            sibling_did = parent.child
            while self.directory_entries[sibling_did].right_sibling != DirectoryEntry.NOSTREAM:
                sibling_did = self.directory_entries[sibling_did].right_sibling
            self.directory_entries[sibling_did].right_sibling = did

        return did

    def add_stream(self, name: str, data: bytes, parent_did: int = 0) -> int:
        """Add a stream (file) entry"""
        entry = DirectoryEntry(name, EntryType.STREAM)
        entry.stream_size = len(data)
        did = len(self.directory_entries)
        self.directory_entries.append(entry)
        self.streams[did] = data

        # Link to parent's child chain (simplified - just use child pointer)
        parent = self.directory_entries[parent_did]
        if parent.child == DirectoryEntry.NOSTREAM:
            parent.child = did
        else:
            # Add as right sibling of the first child
            sibling_did = parent.child
            while self.directory_entries[sibling_did].right_sibling != DirectoryEntry.NOSTREAM:
                sibling_did = self.directory_entries[sibling_did].right_sibling
            self.directory_entries[sibling_did].right_sibling = did

        return did

    def _allocate_sectors_for_data(self, data: bytes) -> List[int]:
        """Allocate sectors for data and build FAT chain"""
        if not data:
            return []

        sectors_needed = (len(data) + self.sector_size - 1) // self.sector_size
        sector_chain = []

        for i in range(sectors_needed):
            sector_id = len(self.fat)
            sector_chain.append(sector_id)

            # Set FAT entry to next sector or end of chain
            if i < sectors_needed - 1:
                self.fat.append(sector_id + 1)
            else:
                self.fat.append(SectorType.ENDOFCHAIN)

        return sector_chain

    def _allocate_mini_sectors_for_data(self, data: bytes) -> List[int]:
        """Allocate mini sectors for small streams"""
        if not data:
            return []

        sectors_needed = (len(data) + self.mini_sector_size - 1) // self.mini_sector_size
        sector_chain = []

        for i in range(sectors_needed):
            sector_id = len(self.mini_fat)
            sector_chain.append(sector_id)

            # Set Mini FAT entry
            if i < sectors_needed - 1:
                self.mini_fat.append(sector_id + 1)
            else:
                self.mini_fat.append(SectorType.ENDOFCHAIN)

            # Append data to mini stream
            start = i * self.mini_sector_size
            end = min(start + self.mini_sector_size, len(data))
            chunk = data[start:end]
            self.mini_stream_data.extend(chunk)
            # Pad to mini sector size
            if len(chunk) < self.mini_sector_size:
                self.mini_stream_data.extend(b'\x00' * (self.mini_sector_size - len(chunk)))

        return sector_chain

    def write(self, file_path: str):
        """Write complete CFB file"""
        with open(file_path, 'wb') as f:
            self._write_to_stream(f)

    def _write_to_stream(self, f: BinaryIO):
        """Write CFB structure to binary stream"""
        # Build sector allocation for all streams
        stream_sectors = {}

        for did, data in self.streams.items():
            entry = self.directory_entries[did]

            if not data:
                entry.starting_sector = SectorType.ENDOFCHAIN
                continue

            # Decide whether to use mini stream or regular sectors
            if len(data) < self.MINI_STREAM_CUTOFF:
                # Use mini stream for small data
                mini_sectors = self._allocate_mini_sectors_for_data(data)
                if mini_sectors:
                    entry.starting_sector = mini_sectors[0]
            else:
                # Use regular sectors for large data
                sectors = self._allocate_sectors_for_data(data)
                if sectors:
                    entry.starting_sector = sectors[0]
                    stream_sectors[did] = (data, sectors)

        # Store mini stream data in root entry if we have any
        if self.mini_stream_data:
            # Mini stream is stored as a regular stream
            mini_stream_sectors = self._allocate_sectors_for_data(bytes(self.mini_stream_data))
            if mini_stream_sectors:
                self.directory_entries[0].starting_sector = mini_stream_sectors[0]
                self.directory_entries[0].stream_size = len(self.mini_stream_data)
                stream_sectors[-1] = (bytes(self.mini_stream_data), mini_stream_sectors)
        else:
            # No mini stream
            self.directory_entries[0].starting_sector = SectorType.ENDOFCHAIN
            self.directory_entries[0].stream_size = 0

        # Allocate sectors for directory entries
        dir_data = b''.join(entry.to_bytes() for entry in self.directory_entries)
        # Pad to sector boundary
        dir_data += b'\xFF' * ((self.sector_size - (len(dir_data) % self.sector_size)) % self.sector_size)
        dir_sectors = self._allocate_sectors_for_data(dir_data)
        stream_sectors[-2] = (dir_data, dir_sectors)

        # Allocate sectors for Mini FAT if we have mini streams
        mini_fat_start_sector = SectorType.ENDOFCHAIN
        num_mini_fat_sectors = 0
        if self.mini_fat:
            mini_fat_data = struct.pack(f'<{len(self.mini_fat)}I', *self.mini_fat)
            # Pad to sector boundary
            padding_needed = (self.sector_size - (len(mini_fat_data) % self.sector_size)) % self.sector_size
            mini_fat_data += struct.pack(f'<{padding_needed // 4}I', *([SectorType.FREESECT] * (padding_needed // 4)))
            mini_fat_sectors = self._allocate_sectors_for_data(mini_fat_data)
            if mini_fat_sectors:
                mini_fat_start_sector = mini_fat_sectors[0]
                num_mini_fat_sectors = len(mini_fat_sectors)
                stream_sectors[-3] = (mini_fat_data, mini_fat_sectors)

        # Allocate sectors for the FAT (and DIFAT, if the file is large enough
        # to need more than 109 FAT sectors). Every FAT and DIFAT sector is
        # itself a sector the FAT must describe, so the counts are interdependent
        # and solved with a small fixpoint loop.
        fat_entries_per_sector = self.sector_size // 4          # 128 for 512-byte sectors
        difat_entries_per_sector = fat_entries_per_sector - 1   # 127; last slot chains
        HEADER_DIFAT_SLOTS = 109

        content_sectors = len(self.fat)  # data / directory / mini-FAT sectors

        num_fat_sectors = 0
        num_difat_sectors = 0
        while True:
            total = content_sectors + num_fat_sectors + num_difat_sectors
            new_fat = (total + fat_entries_per_sector - 1) // fat_entries_per_sector
            overflow = max(0, new_fat - HEADER_DIFAT_SLOTS)
            new_difat = (overflow + difat_entries_per_sector - 1) // difat_entries_per_sector
            if new_fat == num_fat_sectors and new_difat == num_difat_sectors:
                break
            num_fat_sectors, num_difat_sectors = new_fat, new_difat

        # Sector ids: FAT sectors first, then DIFAT sectors.
        fat_sector_ids = list(range(content_sectors, content_sectors + num_fat_sectors))
        difat_sector_ids = list(range(content_sectors + num_fat_sectors,
                                      content_sectors + num_fat_sectors + num_difat_sectors))
        total_sectors = content_sectors + num_fat_sectors + num_difat_sectors

        # Extend the FAT to describe every sector, marking FAT/DIFAT sectors.
        while len(self.fat) < total_sectors:
            self.fat.append(SectorType.FREESECT)
        for sid in fat_sector_ids:
            self.fat[sid] = SectorType.FATSECT
        for sid in difat_sector_ids:
            self.fat[sid] = SectorType.DIFSECT

        # The header holds the first 109 FAT locations; any overflow goes into
        # the DIFAT sector chain.
        header_fat_locs = fat_sector_ids[:HEADER_DIFAT_SLOTS]
        overflow_fat_locs = fat_sector_ids[HEADER_DIFAT_SLOTS:]
        difat_start = difat_sector_ids[0] if difat_sector_ids else SectorType.ENDOFCHAIN

        # Write header
        self._write_header(f, dir_sectors[0], header_fat_locs, num_fat_sectors,
                           mini_fat_start_sector, num_mini_fat_sectors,
                           difat_start, num_difat_sectors)

        # Write all sectors in order
        # Collect all sectors by ID
        sector_data = {}

        # Add stream sectors
        for did, (data, sectors) in stream_sectors.items():
            for i, sector_id in enumerate(sectors):
                start = i * self.sector_size
                end = min(start + self.sector_size, len(data))
                chunk = data[start:end]
                # Pad to sector size
                chunk += b'\x00' * (self.sector_size - len(chunk))
                sector_data[sector_id] = chunk

        # Add FAT sectors
        fat_data = struct.pack(f'<{len(self.fat)}I', *self.fat)
        # Pad to fill all FAT sectors
        fat_data += struct.pack('<I', SectorType.FREESECT) * (num_fat_sectors * fat_entries_per_sector - len(self.fat))

        for i, sector_id in enumerate(fat_sector_ids):
            start = i * self.sector_size
            end = start + self.sector_size
            sector_data[sector_id] = fat_data[start:end]

        # Add DIFAT sectors: 127 FAT locations followed by the next DIFAT
        # sector id (or ENDOFCHAIN for the last one).
        for j, sector_id in enumerate(difat_sector_ids):
            locs = overflow_fat_locs[j * difat_entries_per_sector:(j + 1) * difat_entries_per_sector]
            entries = list(locs)
            entries += [SectorType.FREESECT] * (difat_entries_per_sector - len(entries))
            if j < len(difat_sector_ids) - 1:
                entries.append(difat_sector_ids[j + 1])
            else:
                entries.append(SectorType.ENDOFCHAIN)
            sector_data[sector_id] = struct.pack(f'<{fat_entries_per_sector}I', *entries)

        # Write sectors in order
        for sector_id in sorted(sector_data.keys()):
            f.write(sector_data[sector_id])

    def _write_header(self, f: BinaryIO, dir_start_sector: int, fat_sectors: List[int],
                      num_fat_sectors: int, mini_fat_start: int, num_mini_fat_sectors: int,
                      difat_start: int, num_difat_sectors: int):
        """Write CFB header (512 bytes)"""
        # Signature
        f.write(self.HEADER_SIGNATURE)

        # CLSID (16 bytes of zeros)
        f.write(b'\x00' * 16)

        # Minor version (0x003E)
        f.write(struct.pack('<H', 0x003E))

        # Major version (0x0003 for 512-byte sectors)
        f.write(struct.pack('<H', 0x0003))

        # Byte order (0xFFFE = little-endian)
        f.write(struct.pack('<H', 0xFFFE))

        # Sector size power (0x0009 = 2^9 = 512)
        f.write(struct.pack('<H', 0x0009))

        # Mini sector size power (0x0006 = 2^6 = 64)
        f.write(struct.pack('<H', 0x0006))

        # Reserved (6 bytes)
        f.write(b'\x00' * 6)

        # Number of directory sectors (0 for version 3)
        f.write(struct.pack('<I', 0))

        # Number of FAT sectors (total, including any indexed via DIFAT sectors)
        f.write(struct.pack('<I', num_fat_sectors))

        # First directory sector
        f.write(struct.pack('<I', dir_start_sector))

        # Transaction signature (0)
        f.write(struct.pack('<I', 0))

        # Mini stream cutoff (4096)
        f.write(struct.pack('<I', self.MINI_STREAM_CUTOFF))

        # First mini FAT sector
        f.write(struct.pack('<I', mini_fat_start))

        # Number of mini FAT sectors
        f.write(struct.pack('<I', num_mini_fat_sectors))

        # First DIFAT sector (ENDOFCHAIN when the 109 header slots suffice)
        f.write(struct.pack('<I', difat_start))

        # Number of DIFAT sectors
        f.write(struct.pack('<I', num_difat_sectors))

        # DIFAT array: the first up to 109 FAT sector locations (436 bytes)
        for i in range(109):
            if i < len(fat_sectors):
                f.write(struct.pack('<I', fat_sectors[i]))
            else:
                f.write(struct.pack('<I', SectorType.FREESECT))
