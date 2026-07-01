# PyMsgKit

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Pure Python library for creating Microsoft Outlook MSG files without requiring Outlook or external dependencies. Designed for eDiscovery, forensic email reconstruction, and programmatic email generation with full control over sender properties and MAPI attributes.

## Features

- ✅ **Create MSG files from scratch** - No Outlook installation required
- ✅ **Custom sender control** - Bypass Outlook restrictions for forensic reconstruction
- ✅ **Email threading** - Support for conversation chains and replies
- ✅ **Attachments** - Regular and inline attachments with content IDs
- ✅ **HTML & Plain Text** - Full support for both body formats
- ✅ **Pure Python** - No external dependencies, works on Linux/Mac/Windows
- ✅ **Full MAPI property support** - Complete control over message properties
- ✅ **EML & MBOX export** - Emit standard RFC 5322 `.eml` files and collect many messages into a single portable `.mbox` archive

## Installation

```bash
pip install pymsgkit
```

Or install from source:

```bash
git clone https://github.com/yourusername/pymsgkit.git
cd pymsgkit
pip install -e .
```

## Validation

PyMsgKit creates fully compliant MSG files according to the MS-OXMSG specification. The files can be opened by:
- Microsoft Outlook (Windows, Mac, Web)
- olefile Python library (for CFB structure validation)
- extract-msg Python library (use `ErrorBehavior.STANDARDS_VIOLATION` flag for strict validation)

**Note**: Some MSG validators like extract-msg have strict default validation. Use relaxed error handling if needed:

```python
import extract_msg
from extract_msg.enums import ErrorBehavior

msg = extract_msg.Message('file.msg', errorBehavior=ErrorBehavior.STANDARDS_VIOLATION)
```

## Quick Start

### Simple Email

```python
from pymsgkit import create_email

msg = create_email(
    subject="Hello World",
    body="This is a test email",
    sender_email="sender@example.com",
    sender_name="John Doe",
    to_recipients=[("recipient@example.com", "Jane Smith")]
)
msg.save("email.msg")
```

### Advanced Usage

```python
from pymsgkit import MSGWriter, RecipientType

# Create message
msg = MSGWriter()
msg.set_subject("Project Update")
msg.set_sender("manager@company.com", "Project Manager")
msg.set_body("<h1>Status Report</h1><p>All systems operational.</p>", is_html=True)

# Add recipients
msg.add_recipient("team@company.com", "Development Team", RecipientType.TO)
msg.add_recipient("stakeholder@company.com", "Stakeholder", RecipientType.CC)

# Add attachment
with open("report.pdf", "rb") as f:
    msg.add_attachment("report.pdf", f.read(), mime_type="application/pdf")

# Save
msg.save("project_update.msg")
```

### Email Threading

```python
from pymsgkit import MSGWriter, PropertyTag

# First message
msg1 = MSGWriter()
msg1.set_subject("Initial Message")
msg1.set_sender("alice@example.com", "Alice")
msg1.add_recipient("bob@example.com", "Bob")
msg1.set_body("Starting a conversation thread")
msg1.set_conversation_index()  # Create new thread
msg1.save("thread_01.msg")

# Reply
conversation_index = msg1.properties[PropertyTag.PR_CONVERSATION_INDEX].value

msg2 = MSGWriter()
msg2.set_subject("RE: Initial Message")
msg2.set_sender("bob@example.com", "Bob")
msg2.add_recipient("alice@example.com", "Alice")
msg2.set_body("Reply to the conversation")
msg2.set_conversation_index(conversation_index)  # Link to thread
msg2.save("thread_02.msg")
```

### Inline Images (HTML Email)

```python
from pymsgkit import MSGWriter

msg = MSGWriter()
msg.set_subject("Newsletter")
msg.set_sender("marketing@company.com", "Marketing Team")
msg.set_body(
    '<html><body><p>Check out our logo:</p><img src="cid:logo" /></body></html>',
    is_html=True
)
msg.add_recipient("customer@example.com", "Customer")

# Add inline image
with open("logo.png", "rb") as f:
    msg.add_attachment(
        "logo.png",
        f.read(),
        content_id="logo",
        mime_type="image/png",
        is_inline=True
    )

msg.save("newsletter.msg")
```

### Export to EML and MBOX

Besides Outlook `.msg`, PyMsgKit can emit standard internet-mail formats. A
single message becomes an RFC 5322 `.eml`, and many messages can be collected
into one `.mbox` archive that Outlook, Thunderbird, Apple Mail, and eDiscovery
tools import directly:

```python
from pymsgkit import MSGWriter, MboxWriter

msg = MSGWriter()
msg.set_subject("Welcome")
msg.set_sender("welcome@company.com", "Welcome Team")
msg.add_recipient("newuser@example.com", "New User")
msg.set_body("Thanks for joining!")

# Single message -> .eml
msg.save_eml("welcome.eml")
# ...and still available as .msg
msg.save("welcome.msg")

# Many messages -> one .mbox archive
box = MboxWriter()
box.add(msg)
box.add(another_msg)
box.save("archive.mbox")
```

### What about `.pst`?

A `.pst` file is **not** "a bunch of MSG files in a container" — it is a full
on-disk database (the MS-PST NDB/LTP layers: CRC-protected pages, node/block
B-trees, heap-on-node, table and property contexts, encoded blocks, a folder
hierarchy, and a named-property map). Writing one correctly from scratch is an
order of magnitude more work than this entire MSG library, which is why no
pure-Python PST *writer* exists in the wild.

For the real need behind "make a PST" — **collecting many messages into one
importable archive** — use the MBOX/EML export above. Outlook and every major
mail client can import MBOX, so it covers the practical use case without the
MS-PST complexity. Native PST authoring is intentionally out of scope.

## Use Cases

### eDiscovery & Forensics

PyMsgKit is designed for eDiscovery workflows where you need to reconstruct emails with specific sender addresses that standard APIs won't allow:

```python
# Reconstruct email from archived data with original sender preserved
msg = MSGWriter()
msg.set_sender("original.sender@company.com", "Original Sender")  # ✅ Works!
msg.set_subject("Reconstructed Email")
msg.set_body("This email was reconstructed from archive data.")
msg.add_recipient("legal@company.com", "Legal Team")
msg.save("evidence_email.msg")
```

Standard Outlook APIs won't let you set arbitrary senders, but MSG files created with PyMsgKit bypass this restriction by writing directly to the file structure.

### Automated Email Generation

Generate templated emails programmatically:

```python
import csv
from pymsgkit import create_email

with open("recipients.csv") as f:
    for row in csv.DictReader(f):
        msg = create_email(
            subject=f"Welcome, {row['name']}!",
            body=f"Dear {row['name']},\n\nWelcome to our service...",
            sender_email="welcome@company.com",
            sender_name="Welcome Team",
            to_recipients=[(row['email'], row['name'])]
        )
        msg.save(f"welcome_{row['id']}.msg")
```

## API Reference

### MSGWriter

Main class for creating MSG files.

**Methods:**

- `set_subject(subject: str)` - Set email subject
- `set_body(body: str, is_html: bool = False)` - Set email body
- `set_sender(email: str, name: str = "", addr_type: str = "SMTP")` - Set sender
- `add_recipient(email: str, name: str = "", recipient_type: RecipientType = RecipientType.TO)` - Add recipient
- `add_attachment(filename: str, data: bytes, content_id: str = None, mime_type: str = None, is_inline: bool = False)` - Add attachment
- `set_conversation_index(parent_index: bytes = None)` - Set threading
- `set_property(prop_tag: int, prop_type: int, value: Any)` - Set custom MAPI property
- `save(filepath: str)` - Save to MSG file
- `save_eml(filepath: str)` - Export the message as an RFC 5322 `.eml` file
- `to_eml_bytes() -> bytes` - Return the message serialized as `.eml` bytes

### MboxWriter

Collect multiple messages into a single `.mbox` archive.

- `add(msg)` - Add an `MSGWriter` (or a pre-built `email.message.EmailMessage`)
- `save(filepath: str)` - Write all collected messages to an mbox file
- `len(box)` - Number of messages currently collected

### Helper Functions

- `create_email(...)` - Quick email creation with sensible defaults
- `save_eml(msg, filepath)` / `msg_to_eml_bytes(msg)` / `msg_to_email_message(msg)` - EML export helpers

### Enums

- `RecipientType.TO`, `RecipientType.CC`, `RecipientType.BCC`
- `PropertyType` - MAPI property types
- `AttachMethod` - Attachment methods

## Technical Details

PyMsgKit implements the Microsoft specifications:

- **MS-CFB**: Compound File Binary Format
- **MS-OXMSG**: Outlook MSG File Format
- **MS-OXPROPS**: Exchange Server Protocols Property Tags

The library creates valid MSG files by:

1. Building a Compound File Binary (CFB) container structure
2. Encoding MAPI properties according to type specifications
3. Creating proper directory hierarchies for recipients and attachments
4. Writing property streams with correct naming conventions

## Requirements

- Python 3.7+
- No external dependencies (pure Python standard library)

## Testing

Install the test extras (pytest + olefile), then run the suite. The tests read
generated files back with the independent `olefile` parser and assert the
decoded MAPI properties match the input, so they validate real output rather
than just that a file was written.

```bash
# Install test dependencies
pip install -e ".[test]"

# Run all tests
python -m pytest tests/

# Run a specific module
python -m pytest tests/test_roundtrip.py -v
```

Test modules:

- `test_basic.py` - smoke tests and CFB signature checks
- `test_roundtrip.py` - write then read back with olefile; verify subject, body, sender, recipients, attachments, flags, timestamps
- `test_cfb.py` - Compound File Binary structure (mini + regular sectors, nested storages, many siblings, large streams)
- `test_properties.py` - MAPI property encoders (strings, ints, booleans, FILETIME, entryid, search key, headers)
- `test_export.py` - EML and MBOX export

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Acknowledgments

- Microsoft for documenting the MSG and CFB specifications
- The eDiscovery community for use case feedback
- MsgKit (C#) for architectural inspiration

## Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/pymsgkit/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/pymsgkit/discussions)

## Changelog

### v1.1.0 (2026-07-01)
- Added EML export (`save_eml`, `to_eml_bytes`, `msg_to_email_message`)
- Added `MboxWriter` for collecting many messages into one `.mbox` archive
- Documented why native `.pst` writing is out of scope and MBOX/EML is the practical alternative
- Made `MSGWriter.save()` idempotent (a single writer can be saved to multiple paths without duplicating streams)
- Fixed `filetime_to_datetime` (used a non-existent `timezone.timedelta`)
- Hardened string encoding so non-ASCII senders/recipients/subjects no longer raise
- Rebuilt the test suite around independent olefile round-trip validation (52 tests)

### v1.0.0 (2025-01-XX)
- Initial release
- Full MSG file creation support
- Custom sender control
- Threading support
- Attachment support (regular and inline)
- HTML and plain text bodies
