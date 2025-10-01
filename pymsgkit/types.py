"""
Type definitions and enums for MAPI properties
"""

from enum import IntEnum


class RecipientType(IntEnum):
    """Recipient type enumeration"""
    TO = 1      # MAPI_TO
    CC = 2      # MAPI_CC
    BCC = 3     # MAPI_BCC


class PropertyType(IntEnum):
    """MAPI property type enumeration"""
    PT_UNSPECIFIED = 0x0000
    PT_NULL = 0x0001
    PT_SHORT = 0x0002
    PT_LONG = 0x0003
    PT_FLOAT = 0x0004
    PT_DOUBLE = 0x0005
    PT_CURRENCY = 0x0006
    PT_APPTIME = 0x0007
    PT_ERROR = 0x000A
    PT_BOOLEAN = 0x000B
    PT_OBJECT = 0x000D
    PT_LONGLONG = 0x0014
    PT_STRING8 = 0x001E
    PT_UNICODE = 0x001F
    PT_SYSTIME = 0x0040
    PT_CLSID = 0x0048
    PT_BINARY = 0x0102
    PT_MV_SHORT = 0x1002
    PT_MV_LONG = 0x1003
    PT_MV_UNICODE = 0x101F


class AttachMethod(IntEnum):
    """Attachment method enumeration"""
    NO_ATTACHMENT = 0x0000
    BY_VALUE = 0x0001
    BY_REFERENCE = 0x0002
    BY_REF_RESOLVE = 0x0003
    BY_REF_ONLY = 0x0004
    EMBEDDED_MSG = 0x0005
    OLE = 0x0006
